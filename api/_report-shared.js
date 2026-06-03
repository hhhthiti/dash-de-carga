const DEFAULT_SUPABASE_URL = 'https://pwjatxqtkvvcmzmjjvbi.supabase.co';
const TIME_ZONE = process.env.REPORT_TIME_ZONE || 'America/Sao_Paulo';

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`${name} nao configurado nas variaveis de ambiente.`);
  return value;
}

function supabaseUrl() {
  return (process.env.SUPABASE_URL || DEFAULT_SUPABASE_URL).replace(/\/$/, '');
}

function supabaseKey() {
  return requiredEnv('SUPABASE_SERVICE_ROLE_KEY');
}

function brevoKey() {
  return requiredEnv('BREVO_API_KEY');
}

function json(res, status, payload) {
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', chunk => {
      raw += chunk;
      if (raw.length > 1_000_000) {
        reject(new Error('Payload muito grande.'));
        req.destroy();
      }
    });
    req.on('end', () => {
      try { resolve(raw ? JSON.parse(raw) : {}); }
      catch { reject(new Error('JSON invalido.')); }
    });
    req.on('error', reject);
  });
}

async function sb(path, options = {}) {
  const response = await fetch(`${supabaseUrl()}/rest/v1/${path}`, {
    ...options,
    headers: {
      apikey: supabaseKey(),
      authorization: `Bearer ${supabaseKey()}`,
      'content-type': 'application/json',
      accept: 'application/json',
      ...(options.headers || {}),
    },
  });
  const text = await response.text();
  let body = null;
  try { body = text ? JSON.parse(text) : null; } catch { body = text; }
  if (!response.ok) {
    const details = typeof body === 'string' ? body : JSON.stringify(body);
    throw new Error(`Supabase ${response.status}: ${details}`);
  }
  return body;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function parseList(value) {
  const raw = Array.isArray(value) ? value.join(';') : String(value || '');
  return raw.split(/[;,]/).map(v => v.trim()).filter(Boolean);
}

function validEmails(value) {
  return parseList(value).filter(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v));
}

function toDbConfig(config = {}) {
  return {
    id: 'default',
    email_provider: config.emailProvider || 'brevo',
    email_from: String(config.emailFrom || process.env.REPORT_FROM_EMAIL || '').trim(),
    email_to: validEmails(config.emailTo),
    email_cc: validEmails(config.emailCc),
    auto_email_enabled: !!config.autoEmailEnabled,
    auto_email_interval_minutes: Math.max(5, Number(config.autoEmailIntervalMinutes || 60)),
    report_inicial: String(config.reportInicial || '00:00').slice(0, 5),
    report_final: String(config.reportFinal || '11:59').slice(0, 5),
    send_next_day_at_final: config.sendNextDayAtFinal !== false,
    updated_at: new Date().toISOString(),
  };
}

function fromDbConfig(row = {}) {
  return {
    emailProvider: row.email_provider || 'brevo',
    emailFrom: row.email_from || process.env.REPORT_FROM_EMAIL || '',
    emailEndpoint: '/api/send-report',
    emailTo: Array.isArray(row.email_to) ? row.email_to : validEmails(process.env.REPORT_DEFAULT_TO || ''),
    emailCc: Array.isArray(row.email_cc) ? row.email_cc : validEmails(process.env.REPORT_DEFAULT_CC || ''),
    autoEmailEnabled: row.auto_email_enabled !== undefined ? !!row.auto_email_enabled : true,
    autoEmailIntervalMinutes: Number(row.auto_email_interval_minutes || 60),
    reportInicial: row.report_inicial || '00:00',
    reportFinal: row.report_final || '11:59',
    sendNextDayAtFinal: row.send_next_day_at_final !== false,
  };
}

async function getConfig() {
  try {
    const rows = await sb('report_delivery_config?id=eq.default&select=*&limit=1');
    if (Array.isArray(rows) && rows[0]) return fromDbConfig(rows[0]);
  } catch {
    // The migration may not be applied yet. Fall back to env defaults.
  }
  return fromDbConfig({});
}

async function saveConfig(config) {
  const dbConfig = toDbConfig(config);
  const rows = await sb('report_delivery_config?on_conflict=id', {
    method: 'POST',
    headers: { prefer: 'resolution=merge-duplicates,return=representation' },
    body: JSON.stringify(dbConfig),
  });
  return fromDbConfig(Array.isArray(rows) ? rows[0] : dbConfig);
}

function localParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(date).reduce((acc, part) => {
    if (part.type !== 'literal') acc[part.type] = part.value;
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

function dateKey(offsetDays = 0, base = new Date()) {
  const p = localParts(base);
  const utc = Date.UTC(p.year, p.month - 1, p.day + offsetDays, 12, 0, 0);
  const shifted = new Date(utc);
  const day = String(shifted.getUTCDate()).padStart(2, '0');
  const month = String(shifted.getUTCMonth() + 1).padStart(2, '0');
  return `${day}/${month}/${shifted.getUTCFullYear()}`;
}

function timeToMinutes(value, fallback) {
  const m = String(value || '').match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return fallback;
  return Math.min(23, Number(m[1]) || 0) * 60 + Math.min(59, Number(m[2]) || 0);
}

function insideWindow(config, now = new Date()) {
  const p = localParts(now);
  const current = p.hour * 60 + p.minute;
  const start = timeToMinutes(config.reportInicial, 0);
  const end = timeToMinutes(config.reportFinal, 719);
  if (start <= end) return current >= start && current <= end;
  return current >= start || current <= end;
}

function parseWeight(raw) {
  let s = String(raw ?? '').trim();
  if (!s) return 0;
  s = s.replace(/\s/g, '');
  if (s.includes(',') && s.includes('.')) {
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  } else if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');
  }
  const n = Number.parseFloat(s.replace(/[^\d.-]/g, ''));
  return Number.isFinite(n) ? Math.floor(n) : 0;
}

function fmtCarga(raw) {
  const n = Number(raw) || 0;
  if (Math.abs(n) >= 1000) {
    return `${(Math.trunc((n / 1000) * 10) / 10).toLocaleString('pt-BR', {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1,
    })} t`;
  }
  return `${Math.round(n).toLocaleString('pt-BR')} kg`;
}

function countsByStatus(rows) {
  return rows.reduce((acc, row) => {
    const key = String(row.status || 'SEM STATUS').trim() || 'SEM STATUS';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

async function latestPlanning(dataRef) {
  try {
    const rows = await sb(`planejamento_suzano_snapshots?data_ref=eq.${encodeURIComponent(dataRef)}&select=*&order=snapshot_at.desc&limit=1`);
    return Array.isArray(rows) && rows[0] ? rows[0] : null;
  } catch {
    return null;
  }
}

async function buildReport(dataRef) {
  const rows = await sb(`reporte_carga?data_ref=eq.${encodeURIComponent(dataRef)}&select=*`);
  const cargaRows = Array.isArray(rows) ? rows : [];
  const statusCounts = countsByStatus(cargaRows);
  const validRows = cargaRows.filter(row => !['NO SHOW', 'VEICULO RECUSADO'].includes(String(row.status || '').trim().toUpperCase()));
  const realizadoRows = cargaRows.filter(row => ['EXPEDIDO', 'EM FATURAMENTO'].includes(String(row.status || '').trim().toUpperCase()));
  const nossaGrade = validRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const realizadoGrade = realizadoRows.reduce((sum, row) => sum + parseWeight(row.peso_liquido ?? row.toneladas ?? row.peso), 0);
  const planning = await latestPlanning(dataRef);
  const planejadoSuzano = Number(planning?.planejado_suzano_kg || 0);
  return {
    dataRef,
    planejadoSuzano,
    nossaGrade,
    realizadoGrade,
    pendente: Math.max(0, nossaGrade - realizadoGrade),
    dtsTotal: cargaRows.length,
    statusCounts,
    detalhes: {
      planejadoSuzanoDetalhado: planning?.detalhes || {},
    },
  };
}

function buildReportText(report) {
  return [
    `Reporte de Status - ${report.dataRef}`,
    '',
    `Planejado Suzano: ${fmtCarga(report.planejadoSuzano)}`,
    `Nossa grade: ${fmtCarga(report.nossaGrade)}`,
    `Realizado: ${fmtCarga(report.realizadoGrade)}`,
    `Pendente: ${fmtCarga(report.pendente)}`,
    '',
    `DTs total: ${report.dtsTotal}`,
    `DTs expedidas: ${report.statusCounts.EXPEDIDO || 0}`,
    `No-shows: ${report.statusCounts['NO SHOW'] || 0}`,
    `Recusas: ${report.statusCounts['VEICULO RECUSADO'] || 0}`,
  ].join('\n');
}

function buildReportHtml(report, dashboardUrl) {
  const text = buildReportText(report);
  return `<!doctype html><html><body style="font-family:Arial,sans-serif;background:#f8fafc;color:#0f172a;padding:24px;">
    <h2 style="margin:0 0 12px;">Reporte de Status - ${escapeHtml(report.dataRef)}</h2>
    <pre style="white-space:pre-wrap;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:16px;">${escapeHtml(text)}</pre>
    ${dashboardUrl ? `<p><a href="${escapeHtml(dashboardUrl)}" target="_blank" rel="noopener noreferrer">Abrir dashboard de carga</a></p>` : ''}
    <p style="color:#64748b;font-size:12px;">Gerado automaticamente pela Vercel em ${escapeHtml(new Date().toLocaleString('pt-BR', { timeZone: TIME_ZONE }))}.</p>
  </body></html>`;
}

async function sendEmail({ report, config, source = 'MANUAL' }) {
  const to = validEmails(config.emailTo);
  const cc = validEmails(config.emailCc);
  if (!to.length) throw new Error('Nenhum destinatario configurado.');
  const senderEmail = process.env.REPORT_FROM_EMAIL || config.emailFrom;
  if (!senderEmail) throw new Error('REPORT_FROM_EMAIL nao configurado.');
  const dashboardUrl = process.env.REPORT_DASHBOARD_URL || 'https://dash-de-carga.vercel.app/cco.html';
  const subject = `Reporte de Status - ${report.dataRef}`;
  const textContent = buildReportText(report);
  const payload = {
    sender: { name: process.env.REPORT_FROM_NAME || 'dash de carga', email: senderEmail },
    to: to.map(email => ({ email })),
    subject,
    htmlContent: buildReportHtml(report, dashboardUrl),
    textContent,
  };
  if (cc.length) payload.cc = cc.map(email => ({ email }));

  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': brevoKey(),
      'content-type': 'application/json',
    },
    body: JSON.stringify(payload),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`Brevo ${response.status}: ${body.slice(0, 500)}`);
  await recordReport(report, source).catch(() => null);
  return { subject, brevo: body ? JSON.parse(body) : { ok: true } };
}

async function recordReport(report, tipo) {
  await sb('reportes_diarios', {
    method: 'POST',
    headers: { prefer: 'return=minimal' },
    body: JSON.stringify([{
      data_ref: report.dataRef,
      tipo,
      hora_reporte: new Date().toISOString(),
      planejado_suzano_kg: report.planejadoSuzano || null,
      grade_atual_kg: report.nossaGrade || 0,
      realizado_kg: report.realizadoGrade || 0,
      pendente_kg: report.pendente || 0,
      variacao_suzano_kg: (report.nossaGrade || 0) - (report.planejadoSuzano || 0),
      dts_expedidas: Number(report.statusCounts.EXPEDIDO || 0),
      dts_no_show: Number(report.statusCounts['NO SHOW'] || 0),
      dts_recusadas: Number(report.statusCounts['VEICULO RECUSADO'] || 0),
      detalhes: report,
    }]),
  });
}

async function alreadyRan(slotKey) {
  try {
    const rows = await sb(`report_delivery_runs?slot_key=eq.${encodeURIComponent(slotKey)}&select=slot_key&limit=1`);
    return Array.isArray(rows) && rows.length > 0;
  } catch {
    return false;
  }
}

async function markRan(slotKey, details = {}) {
  await sb('report_delivery_runs?on_conflict=slot_key', {
    method: 'POST',
    headers: { prefer: 'resolution=merge-duplicates,return=minimal' },
    body: JSON.stringify([{ slot_key: slotKey, details, sent_at: new Date().toISOString() }]),
  }).catch(() => null);
}

module.exports = {
  buildReport,
  dateKey,
  getConfig,
  insideWindow,
  json,
  localParts,
  markRan,
  readJson,
  saveConfig,
  sendEmail,
  alreadyRan,
};
