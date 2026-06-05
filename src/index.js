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
    ctx.waitUntil(runScheduledReport(env, event));
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
    emailProvider: row.email_provider || "auto",
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
    email_provider: config.emailProvider || "auto",
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

function parseBrDateTime(raw) {
  const s = String(raw || "").trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
  return Number.isNaN(d.getTime()) ? null : d;
}

function parseBrDate(raw) {
  const d = parseBrDateTime(raw);
  if (!d) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function rowDateTime(row) {
  return parseBrDateTime(row.fim_carregamento || row.fim_agenda || row.grade_carregamento || row.data_ref || "");
}

function normalizeShift(raw) {
  const s = String(raw || "").trim().toUpperCase();
  return ["T1", "T2", "T3"].includes(s) ? s : "";
}

function shiftLabel(shift) {
  return {
    T1: "T1 07-14h",
    T2: "T2 15-22h",
    T3: "T3 23-06h",
  }[shift] || "Dia completo";
}

function inferScheduledShift(env, date = new Date()) {
  const p = localParts(env, date);
  const mins = p.hour * 60 + p.minute;
  if (Math.abs(mins - 14 * 60) <= 20) return "T1";
  if (Math.abs(mins - 22 * 60) <= 20) return "T2";
  if (Math.abs(mins - 6 * 60) <= 20) return "T3";
  return "";
}

function reportDateForShift(env, shift, date = new Date()) {
  return dateKey(env, shift === "T3" ? -1 : 0, date);
}

function rowMatchesShift(row, dataRef, shift) {
  const code = normalizeShift(shift);
  if (!code) return true;
  const base = parseBrDate(dataRef);
  const dt = rowDateTime(row);
  if (!base || !dt) return false;
  const start = new Date(base);
  const end = new Date(base);
  if (code === "T1") {
    start.setHours(7, 0, 0, 0);
    end.setHours(14, 59, 59, 999);
  } else if (code === "T2") {
    start.setHours(15, 0, 0, 0);
    end.setHours(22, 59, 59, 999);
  } else {
    start.setHours(23, 0, 0, 0);
    end.setTime(addDays(base, 1).getTime());
    end.setHours(6, 59, 59, 999);
  }
  return dt >= start && dt <= end;
}

function modality(row) {
  const op = String(row.tipo_operacao || row.descricao_documento || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  if (op.includes("PRE")) return "PRE-FATURA";
  if (op.includes("TRANSF") || op.includes("TNF") || op.includes("FILIAL") || op.includes("INTERCOMP")) return "TRANSFERENCIA";
  if (op.includes("MOGI") || String(row.centro || "").trim() === "1110") return "VENDA MOGI";
  return "VENDA ARUJA";
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

async function loadCargaRows(env, dataRef) {
  const rows = await sb(env, `reporte_carga?data_ref=eq.${encodeURIComponent(dataRef)}&select=*`);
  return Array.isArray(rows) ? rows : [];
}

async function buildReport(env, dataRef, options = {}) {
  const shift = normalizeShift(options.shift);
  const cargaRowsAll = await loadCargaRows(env, dataRef);
  const cargaRows = shift ? cargaRowsAll.filter(row => rowMatchesShift(row, dataRef, shift)) : cargaRowsAll;
  const statusCounts = countsByStatus(cargaRows);
  const validRows = cargaRows.filter(row => !["NO SHOW", "VEICULO RECUSADO"].includes(String(row.status || "").trim().toUpperCase()));
  const realizadoRows = cargaRows.filter(row => ["EXPEDIDO", "EM FATURAMENTO"].includes(String(row.status || "").trim().toUpperCase()));
  const nossaGrade = validRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const realizadoGrade = realizadoRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const planning = await latestPlanning(env, dataRef);
  const planejadoSuzano = Number(planning?.planejado_suzano_kg || 0);
  const tipos = {};
  const planejadoSuzanoDetalhado = planning?.detalhes && typeof planning.detalhes === "object" ? planning.detalhes : {};
  validRows.forEach(row => {
    const tipo = modality(row);
    const weight = parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso);
    const realizado = ["EXPEDIDO", "EM FATURAMENTO"].includes(String(row.status || "").trim().toUpperCase());
    if (!tipos[tipo]) tipos[tipo] = { tipo, planejado: 0, realizado: 0, pendente: 0 };
    tipos[tipo].planejado += weight;
    if (realizado) tipos[tipo].realizado += weight;
  });
  Object.values(tipos).forEach(item => {
    item.pendente = Math.max(0, item.planejado - item.realizado);
  });
  return {
    dataRef,
    shift,
    shiftLabel: shiftLabel(shift),
    planejadoSuzano,
    planejadoSuzanoDetalhado,
    nossaGrade,
    realizadoGrade,
    pendente: Math.max(0, nossaGrade - realizadoGrade),
    variacaoSuzanoGrade: nossaGrade - planejadoSuzano,
    dtsTotal: cargaRows.length,
    statusCounts,
    tipos: Object.values(tipos),
    rows: cargaRows,
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
    `Reporte de Status - ${report.dataRef}${report.shift ? ` - ${report.shiftLabel || report.shift}` : ""}`,
    "",
    `Planejado Suzano: ${fmtCarga(report.planejadoSuzano)}`,
    `Nossa grade x Suzano: ${fmtCarga(report.variacaoSuzanoGrade || 0)}`,
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
  const card = (title, value, color, sub = "") => `<td style="padding:8px;width:25%;">
    <div style="border:1px solid ${color}55;background:#111c2e;border-radius:8px;padding:14px;min-height:92px;">
      <div style="font-size:11px;letter-spacing:1px;color:${color};font-weight:800;">${escapeHtml(title)}</div>
      <div style="font-size:24px;line-height:1.35;color:${color};font-weight:900;">${escapeHtml(value)}</div>
      <div style="font-size:11px;color:#94a3b8;">${escapeHtml(sub)}</div>
    </div>
  </td>`;
  const tipos = Array.isArray(report.tipos) ? report.tipos : [];
  return `<!doctype html><html><body style="margin:0;font-family:Arial,sans-serif;background:#0b1220;color:#e2e8f0;padding:24px;">
    <div style="max-width:980px;margin:0 auto;">
      <h2 style="margin:0 0 4px;font-size:24px;">Reporte de Status</h2>
      <div style="color:#94a3b8;margin-bottom:16px;">Dia ${escapeHtml(report.dataRef)}${report.shift ? ` | ${escapeHtml(report.shiftLabel || report.shift)}` : ""}</div>
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin-bottom:14px;"><tr>
        ${card("PLANEJADO SUZANO", fmtCarga(report.planejadoSuzano), "#38bdf8", "Pedido informado")}
        ${card("NOSSA GRADE", fmtCarga(report.nossaGrade), "#60a5fa", `vs Suzano: ${fmtCarga(report.variacaoSuzanoGrade || 0)}`)}
        ${card("REALIZADO", fmtCarga(report.realizadoGrade), "#22c55e", "Expedido + faturamento")}
        ${card("PENDENTE", fmtCarga(report.pendente), "#f59e0b", "Nossa grade - realizado")}
      </tr></table>
      ${tipos.length ? `<div style="background:#111c2e;border:1px solid #334155;border-radius:8px;padding:14px;margin-bottom:14px;">
        <div style="font-size:12px;letter-spacing:1px;color:#94a3b8;font-weight:800;margin-bottom:10px;">PLANEJADO SUZANO x NOSSA GRADE x REALIZADO</div>
        <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;color:#e2e8f0;font-size:13px;">
          <tr style="color:#64748b;text-align:left;"><th style="padding:7px;">Tipo</th><th style="padding:7px;text-align:right;">Nossa grade</th><th style="padding:7px;text-align:right;">Realizado</th><th style="padding:7px;text-align:right;">Pendente</th></tr>
          ${tipos.map(t => `<tr>
            <td style="padding:7px;border-top:1px solid #263244;">${escapeHtml(t.tipo)}</td>
            <td style="padding:7px;border-top:1px solid #263244;text-align:right;color:#60a5fa;font-weight:800;">${escapeHtml(fmtCarga(t.planejado))}</td>
            <td style="padding:7px;border-top:1px solid #263244;text-align:right;color:#22c55e;font-weight:800;">${escapeHtml(fmtCarga(t.realizado))}</td>
            <td style="padding:7px;border-top:1px solid #263244;text-align:right;color:#f59e0b;font-weight:800;">${escapeHtml(fmtCarga(t.pendente))}</td>
          </tr>`).join("")}
        </table>
      </div>` : ""}
      <pre style="white-space:pre-wrap;background:#111c2e;border:1px solid #334155;border-radius:8px;padding:16px;color:#cbd5e1;">${escapeHtml(text)}</pre>
      ${dashboardUrl ? `<p><a href="${escapeHtml(dashboardUrl)}" target="_blank" rel="noopener noreferrer" style="color:#38bdf8;">Abrir dashboard de carga</a></p>` : ""}
    </div>
  </body></html>`;
}

function csvEscape(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function reportCsv(report) {
  const cols = [
    "DIA",
    "TURNO",
    "DT",
    "TRANSPORTADORA",
    "GRADE",
    "FIM",
    "HORA CHEGADA",
    "N PORTARIA",
    "STATUS",
    "DESC DOCUMENTO",
    "PESO LIQUIDO",
    "TIPO OPERACAO",
  ];
  const rows = (Array.isArray(report.rows) ? report.rows : []).map(row => [
    row.dia_ref || report.dataRef || "",
    report.shiftLabel || report.shift || "",
    row.dt || "",
    row.transportadora || "",
    row.grade_carregamento || "",
    row.fim_carregamento || row.fim_agenda || "",
    row.hora_chegada || "",
    row.n_portaria || "",
    row.status || "",
    row.descricao_documento || "",
    row.peso_liquido ?? row.toneladas ?? row.peso ?? "",
    row.tipo_operacao || "",
  ]);
  return "\uFEFF" + [cols, ...rows].map(row => row.map(csvEscape).join(";")).join("\n");
}

function base64Utf8(value) {
  const bytes = new TextEncoder().encode(String(value || ""));
  let binary = "";
  bytes.forEach(byte => { binary += String.fromCharCode(byte); });
  return btoa(binary);
}

function attachmentFilename(report) {
  const date = String(report.dataRef || "reporte").replace(/\//g, "-");
  const shift = report.shift ? `_${report.shift}` : "";
  return `cargas_${date}${shift}.csv`;
}

function emailArtifacts(report) {
  const csv = reportCsv(report);
  return {
    csv,
    csvBase64: base64Utf8(csv),
    csvFilename: attachmentFilename(report),
  };
}

async function sendEmail(env, { report, config }) {
  const to = validEmails(config.emailTo);
  const cc = validEmails(config.emailCc);
  if (!to.length) throw new Error("Nenhum destinatario configurado.");
  const senderEmail = env.REPORT_FROM_EMAIL || config.emailFrom;
  if (!senderEmail) throw new Error("REPORT_FROM_EMAIL nao configurado.");

  const from = env.REPORT_FROM_NAME
    ? `${env.REPORT_FROM_NAME} <${senderEmail}>`
    : senderEmail;
  const subject = `Reporte de Status - ${report.dataRef}${report.shift ? ` - ${report.shift}` : ""}`;
  const html = reportHtml(env, report);
  const text = reportText(report);
  const artifacts = emailArtifacts(report);
  const provider = String(config.emailProvider || "auto").toLowerCase();
  const order = provider === "brevo" ? ["brevo"] : provider === "resend" ? ["resend"] : ["resend", "brevo"];
  const errors = [];

  for (const item of order) {
    try {
      if (item === "resend") {
        return { provider: "resend", result: await sendResendEmail(env, { from, to, cc, subject, html, text, artifacts }) };
      }
      if (item === "brevo") {
        return { provider: "brevo", result: await sendBrevoEmail(env, { fromName: env.REPORT_FROM_NAME || "", senderEmail, to, cc, subject, html, text, artifacts }) };
      }
    } catch (err) {
      errors.push(`${item}: ${err.message}`);
    }
  }
  throw new Error(errors.join(" | ") || "Nenhum provedor de e-mail configurado.");
}

async function sendResendEmail(env, { from, to, cc, subject, html, text, artifacts }) {
  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      accept: "application/json",
      authorization: `Bearer ${requiredEnv(env, "RESEND_API_KEY")}`,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      from,
      to,
      cc,
      subject,
      html,
      text,
      attachments: [{
        filename: artifacts.csvFilename,
        content: artifacts.csvBase64,
      }],
    }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`Resend ${response.status}: ${body.slice(0, 500)}`);
  return body ? JSON.parse(body) : { ok: true };
}

async function sendBrevoEmail(env, { fromName, senderEmail, to, cc, subject, html, text, artifacts }) {
  const apiKey = requiredEnv(env, "BREVO_API_KEY");
  const response = await fetch("https://api.brevo.com/v3/smtp/email", {
    method: "POST",
    headers: {
      accept: "application/json",
      "api-key": apiKey,
      "content-type": "application/json",
    },
    body: JSON.stringify({
      sender: { email: senderEmail, name: fromName || senderEmail },
      to: to.map(email => ({ email })),
      cc: cc.map(email => ({ email })),
      subject,
      htmlContent: html,
      textContent: text,
      attachment: [{
        name: artifacts.csvFilename,
        content: artifacts.csvBase64,
      }],
    }),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`Brevo ${response.status}: ${body.slice(0, 500)}`);
  return body ? JSON.parse(body) : { ok: true };
}

function normalizeIncomingReport(report) {
  const dataRef = report.dataRef || report.data_ref || "";
  if (!dataRef) return {};
  const shift = normalizeShift(report.shift || report.turno || report.shift_code);
  return {
    ...report,
    dataRef,
    shift,
    shiftLabel: report.shiftLabel || report.shift_label || shiftLabel(shift),
    planejadoSuzano: Number(report.planejadoSuzano ?? report.planejado_suzano_kg ?? 0),
    nossaGrade: Number(report.nossaGrade ?? report.nossa_grade_kg ?? 0),
    realizadoGrade: Number(report.realizadoGrade ?? report.realizado_kg ?? 0),
    pendente: Number(report.pendente ?? report.pendente_kg ?? Math.max(0, Number(report.nossaGrade || 0) - Number(report.realizadoGrade || 0))),
    variacaoSuzanoGrade: Number(report.variacaoSuzanoGrade ?? report.variacao_suzano_grade ?? (Number(report.nossaGrade || 0) - Number(report.planejadoSuzano || 0))),
    dtsTotal: Number(report.dtsTotal ?? report.dts_total ?? 0),
    statusCounts: report.statusCounts || report.status_counts || {},
    tipos: Array.isArray(report.tipos) ? report.tipos : [],
    rows: Array.isArray(report.rows) ? report.rows : [],
  };
}

async function handleSendReport(request, env) {
  if (request.method !== "POST") return json({ error: "Metodo nao permitido." }, 405);
  const body = await request.json();
  const savedConfig = await getConfig(env);
  const config = { ...savedConfig, ...(body.config || {}) };
  const incomingReport = normalizeIncomingReport(body.report || {});
  const dataRef = incomingReport.dataRef || body.dataRef || dateKey(env, 0);
  const shift = normalizeShift(body.shift || incomingReport.shift || config.reportShift);
  const report = incomingReport.dataRef
    ? { ...incomingReport, rows: incomingReport.rows.length ? incomingReport.rows : await loadCargaRows(env, dataRef) }
    : await buildReport(env, dataRef, { shift });
  const email = await sendEmail(env, { report, config });
  return json({ ok: true, email });
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

async function runScheduledReport(env, event = {}) {
  const config = await getConfig(env);
  if (!config.autoEmailEnabled) return { skipped: "auto_disabled" };
  const now = event.scheduledTime ? new Date(event.scheduledTime) : new Date();
  const shift = inferScheduledShift(env, now);
  const dataRef = reportDateForShift(env, shift, now);
  const report = await buildReport(env, dataRef, { shift });
  const result = await sendEmail(env, { report, config });
  return { ok: true, shift, dataRef, result };
}

async function handleCronReportFinal(request, env) {
  if (request.method !== "GET") return json({ error: "Metodo nao permitido." }, 405);
  assertCronSecret(request, env);
  const url = new URL(request.url);
  const shift = normalizeShift(url.searchParams.get("shift"));
  if (shift) {
    const dataRef = url.searchParams.get("dataRef") || reportDateForShift(env, shift);
    const config = await getConfig(env);
    if (!config.autoEmailEnabled) return json({ skipped: "auto_disabled" });
    return json({ ok: true, shift, dataRef, result: await sendEmail(env, { report: await buildReport(env, dataRef, { shift }), config }) });
  }
  return json(await runScheduledReport(env));
}
