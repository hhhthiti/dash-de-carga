const DEFAULT_SUPABASE_URL = "https://pwjatxqtkvvcmzmjjvbi.supabase.co";
const DEFAULT_TIME_ZONE = "America/Sao_Paulo";

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    try {
      if (url.pathname === "/api/send-report") return await handleSendReport(request, env);
      if (url.pathname === "/api/report-config") return await handleReportConfig(request, env);
      if (url.pathname === "/api/cron-report-final") return await handleCronReportFinal(request, env);
      if (url.pathname === "/") return env.ASSETS.fetch(new Request(new URL("/index.html", url), request));
      return env.ASSETS.fetch(request);
    } catch (err) {
      return json({ error: err.message || "Erro interno." }, err.status || 500);
    }
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(runDailyReport(env));
  },
};

function json(payload, status = 200) {
  return Response.json(payload, {
    status,
    headers: { "cache-control": "no-store" },
  });
}

function requiredEnv(env, name) {
  const value = env[name];
  if (!value) throw new Error(`${name} nao configurado nas variaveis de ambiente.`);
  return value;
}

function supabaseUrl(env) {
  return String(env.SUPABASE_URL || DEFAULT_SUPABASE_URL).replace(/\/$/, "");
}

function parseList(value) {
  const raw = Array.isArray(value) ? value.join(";") : String(value || "");
  return raw.split(/[;,]/).map(v => v.trim()).filter(Boolean);
}

function validEmails(value) {
  return parseList(value).filter(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v));
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;",
  }[c]));
}

async function sb(env, path, options = {}) {
  const key = requiredEnv(env, "SUPABASE_SERVICE_ROLE_KEY");
  const response = await fetch(`${supabaseUrl(env)}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: key,
      authorization: `Bearer ${key}`,
      "content-type": "application/json",
      accept: "application/json",
      ...(options.headers || {}),
    },
  });
  const text = await response.text();
  let body = null;
  try { body = text ? JSON.parse(text) : null; } catch { body = text; }
  if (!response.ok) {
    const details = typeof body === "string" ? body : JSON.stringify(body);
    throw new Error(`Supabase ${response.status}: ${details}`);
  }
  return body;
}

function fromDbConfig(env, row = {}) {
  return {
    emailProvider: row.email_provider || "brevo",
    emailFrom: row.email_from || env.REPORT_FROM_EMAIL || "",
    emailEndpoint: "/api/send-report",
    emailTo: Array.isArray(row.email_to) ? row.email_to : validEmails(env.REPORT_DEFAULT_TO || ""),
    emailCc: Array.isArray(row.email_cc) ? row.email_cc : validEmails(env.REPORT_DEFAULT_CC || ""),
    autoEmailEnabled: row.auto_email_enabled !== undefined ? !!row.auto_email_enabled : true,
    autoEmailIntervalMinutes: Number(row.auto_email_interval_minutes || 60),
    reportInicial: row.report_inicial || "00:00",
    reportFinal: row.report_final || "11:59",
    sendNextDayAtFinal: row.send_next_day_at_final !== false,
  };
}

function toDbConfig(env, config = {}) {
  return {
    id: "default",
    email_provider: config.emailProvider || "brevo",
    email_from: String(config.emailFrom || env.REPORT_FROM_EMAIL || "").trim(),
    email_to: validEmails(config.emailTo),
    email_cc: validEmails(config.emailCc),
    auto_email_enabled: !!config.autoEmailEnabled,
    auto_email_interval_minutes: Math.max(5, Number(config.autoEmailIntervalMinutes || 60)),
    report_inicial: String(config.reportInicial || "00:00").slice(0, 5),
    report_final: String(config.reportFinal || "11:59").slice(0, 5),
    send_next_day_at_final: config.sendNextDayAtFinal !== false,
    updated_at: new Date().toISOString(),
  };
}

async function getConfig(env) {
  try {
    const rows = await sb(env, "report_delivery_config?id=eq.default&select=*&limit=1");
    if (Array.isArray(rows) && rows[0]) return fromDbConfig(env, rows[0]);
  } catch {
    // Keep manual send usable before the migration is applied.
  }
  return fromDbConfig(env, {});
}

async function saveConfig(env, config) {
  const dbConfig = toDbConfig(env, config);
  const rows = await sb(env, "report_delivery_config?on_conflict=id", {
    method: "POST",
    headers: { prefer: "resolution=merge-duplicates,return=representation" },
    body: JSON.stringify(dbConfig),
  });
  return fromDbConfig(env, Array.isArray(rows) ? rows[0] : dbConfig);
}

function localParts(env, date = new Date()) {
  const timeZone = env.REPORT_TIME_ZONE || DEFAULT_TIME_ZONE;
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(date).reduce((acc, part) => {
    if (part.type !== "literal") acc[part.type] = part.value;
    return acc;
  }, {});
  return {
    year: Number(parts.year),
    month: Number(parts.month),
    day: Number(parts.day),
    hour: Number(parts.hour),
    minute: Number(parts.minute),
  };
}

function dateKey(env, offsetDays = 0, base = new Date()) {
  const p = localParts(env, base);
  const utc = Date.UTC(p.year, p.month - 1, p.day + offsetDays, 12, 0, 0);
  const shifted = new Date(utc);
  return `${String(shifted.getUTCDate()).padStart(2, "0")}/${String(shifted.getUTCMonth() + 1).padStart(2, "0")}/${shifted.getUTCFullYear()}`;
}

function parseWeight(raw) {
  let s = String(raw ?? "").trim();
  if (!s) return 0;
  s = s.replace(/\s/g, "");
  if (s.includes(",") && s.includes(".")) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", ".");
    else s = s.replace(/,/g, "");
  } else if (s.includes(",")) {
    s = s.replace(",", ".");
  } else if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, "");
  }
  const n = Number.parseFloat(s.replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? Math.floor(n) : 0;
}

function countsByStatus(rows) {
  return rows.reduce((acc, row) => {
    const key = String(row.status || "SEM STATUS").trim() || "SEM STATUS";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

async function latestPlanning(env, dataRef) {
  try {
    const rows = await sb(env, `planejamento_suzano_snapshots?data_ref=eq.${encodeURIComponent(dataRef)}&select=*&order=snapshot_at.desc&limit=1`);
    return Array.isArray(rows) && rows[0] ? rows[0] : null;
  } catch {
    return null;
  }
}

async function buildReport(env, dataRef) {
  const rows = await sb(env, `reporte_carga?data_ref=eq.${encodeURIComponent(dataRef)}&select=*`);
  const cargaRows = Array.isArray(rows) ? rows : [];
  const statusCounts = countsByStatus(cargaRows);
  const validRows = cargaRows.filter(row => !["NO SHOW", "VEICULO RECUSADO"].includes(String(row.status || "").trim().toUpperCase()));
  const realizadoRows = cargaRows.filter(row => ["EXPEDIDO", "EM FATURAMENTO"].includes(String(row.status || "").trim().toUpperCase()));
  const nossaGrade = validRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const realizadoGrade = realizadoRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const planning = await latestPlanning(env, dataRef);
  const planejadoSuzano = Number(planning?.planejado_suzano_kg || 0);
  return {
    dataRef,
    planejadoSuzano,
    nossaGrade,
    realizadoGrade,
    pendente: Math.max(0, nossaGrade - realizadoGrade),
    dtsTotal: cargaRows.length,
    statusCounts,
  };
}

function fmtCarga(raw) {
  const n = Number(raw) || 0;
  if (Math.abs(n) >= 1000) {
    return `${(Math.trunc((n / 1000) * 10) / 10).toLocaleString("pt-BR", {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })} t`;
  }
  return `${Math.round(n).toLocaleString("pt-BR")} kg`;
}

function reportText(report) {
  return [
    `Reporte de Status - ${report.dataRef}`,
    "",
    `Planejado Suzano: ${fmtCarga(report.planejadoSuzano)}`,
    `Nossa grade: ${fmtCarga(report.nossaGrade)}`,
    `Realizado: ${fmtCarga(report.realizadoGrade)}`,
    `Pendente: ${fmtCarga(report.pendente)}`,
    "",
    `DTs total: ${report.dtsTotal}`,
    `DTs expedidas: ${report.statusCounts?.EXPEDIDO || 0}`,
    `No-shows: ${report.statusCounts?.["NO SHOW"] || 0}`,
    `Recusas: ${report.statusCounts?.["VEICULO RECUSADO"] || 0}`,
  ].join("\n");
}

function reportHtml(env, report) {
  const text = reportText(report);
  const dashboardUrl = env.REPORT_DASHBOARD_URL || "";
  return `<!doctype html><html><body style="font-family:Arial,sans-serif;background:#f8fafc;color:#0f172a;padding:24px;">
    <h2 style="margin:0 0 12px;">Reporte de Status - ${escapeHtml(report.dataRef)}</h2>
    <pre style="white-space:pre-wrap;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:16px;">${escapeHtml(text)}</pre>
    ${dashboardUrl ? `<p><a href="${escapeHtml(dashboardUrl)}" target="_blank" rel="noopener noreferrer">Abrir dashboard de carga</a></p>` : ""}
  </body></html>`;
}

async function sendEmail(env, { report, config }) {
  const to = validEmails(config.emailTo);
  const cc = validEmails(config.emailCc);
  if (!to.length) throw new Error("Nenhum destinatario configurado.");
  const senderEmail = env.REPORT_FROM_EMAIL || config.emailFrom;
  if (!senderEmail) throw new Error("REPORT_FROM_EMAIL nao configurado.");

  const payload = {
    sender: { name: env.REPORT_FROM_NAME || "dash de carga", email: senderEmail },
    to: to.map(email => ({ email })),
    subject: `Reporte de Status - ${report.dataRef}`,
    htmlContent: reportHtml(env, report),
    textContent: reportText(report),
  };
  if (cc.length) payload.cc = cc.map(email => ({ email }));

  const response = await fetch("https://api.brevo.com/v3/smtp/email", {
    method: "POST",
    headers: {
      accept: "application/json",
      "api-key": requiredEnv(env, "BREVO_API_KEY"),
      "content-type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`Brevo ${response.status}: ${body.slice(0, 500)}`);
  return body ? JSON.parse(body) : { ok: true };
}

function normalizeIncomingReport(report) {
  const dataRef = report.dataRef || report.data_ref || "";
  if (!dataRef) return {};
  return {
    ...report,
    dataRef,
    planejadoSuzano: Number(report.planejadoSuzano ?? report.planejado_suzano_kg ?? 0),
    nossaGrade: Number(report.nossaGrade ?? report.nossa_grade_kg ?? 0),
    realizadoGrade: Number(report.realizadoGrade ?? report.realizado_kg ?? 0),
    pendente: Number(report.pendente ?? report.pendente_kg ?? Math.max(0, Number(report.nossaGrade || 0) - Number(report.realizadoGrade || 0))),
    dtsTotal: Number(report.dtsTotal ?? report.dts_total ?? 0),
    statusCounts: report.statusCounts || report.status_counts || {},
  };
}

async function handleSendReport(request, env) {
  if (request.method !== "POST") return json({ error: "Metodo nao permitido." }, 405);
  const body = await request.json();
  const savedConfig = await getConfig(env);
  const config = { ...savedConfig, ...(body.config || {}) };
  const incomingReport = normalizeIncomingReport(body.report || {});
  const dataRef = incomingReport.dataRef || body.dataRef || dateKey(env, 0);
  const report = incomingReport.dataRef ? incomingReport : await buildReport(env, dataRef);
  const brevo = await sendEmail(env, { report, config });
  return json({ ok: true, brevo });
}

async function handleReportConfig(request, env) {
  if (request.method === "GET") return json({ ok: true, config: await getConfig(env) });
  if (request.method === "POST") return json({ ok: true, config: await saveConfig(env, await request.json()) });
  return json({ error: "Metodo nao permitido." }, 405);
}

function assertCronSecret(request, env) {
  if (!env.CRON_SECRET) return;
  const auth = request.headers.get("authorization") || "";
  if (auth !== `Bearer ${env.CRON_SECRET}`) {
    const err = new Error("Unauthorized");
    err.status = 401;
    throw err;
  }
}

async function runDailyReport(env) {
  const config = await getConfig(env);
  if (!config.autoEmailEnabled) return { skipped: "auto_disabled" };
  const results = [];
  results.push(await sendEmail(env, { report: await buildReport(env, dateKey(env, 0)), config }));
  if (config.sendNextDayAtFinal !== false) {
    results.push(await sendEmail(env, { report: await buildReport(env, dateKey(env, 1)), config }));
  }
  return { ok: true, results };
}

async function handleCronReportFinal(request, env) {
  if (request.method !== "GET") return json({ error: "Metodo nao permitido." }, 405);
  assertCronSecret(request, env);
  return json(await runDailyReport(env));
}
