const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const ENV_PATH = path.join(ROOT, '.env.local');
const PORT = Number(process.env.REPORT_SERVER_PORT || 8787);

loadEnvFile(ENV_PATH);

const BREVO_API_KEY = process.env.BREVO_API_KEY || '';
const BREVO_EVENT_FALLBACK_EMAIL = process.env.BREVO_EVENT_FALLBACK_EMAIL || '';
const REPORT_FROM_EMAIL = process.env.REPORT_FROM_EMAIL || '';
const REPORT_FROM_NAME = process.env.REPORT_FROM_NAME || 'Reporte Operacional';
const REPORT_DEFAULT_TO = parseEmailList(process.env.REPORT_DEFAULT_TO || '');
const REPORT_DEFAULT_CC = parseEmailList(process.env.REPORT_DEFAULT_CC || '');
const REPORT_SERVER_TOKEN = process.env.REPORT_SERVER_TOKEN || '';
const REPORT_IMAGE_TITLE = process.env.REPORT_IMAGE_TITLE || 'Dashboard de Carga';
const REPORT_PUBLIC_BASE_URL = process.env.REPORT_PUBLIC_BASE_URL || '';

let lastReportSnapshot = null;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml; charset=utf-8',
};

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'OPTIONS') return sendCors(res, 204);
    if (req.method === 'GET' && req.url === '/health') {
      return sendJson(res, 200, { ok: true, brevoConfigured: !!BREVO_API_KEY, fromConfigured: !!REPORT_FROM_EMAIL });
    }
    if (req.method === 'GET' && req.url && req.url.startsWith('/api/relatorio')) {
      return serveReportImage(req, res);
    }
    if (req.method === 'POST' && req.url === '/brevo-event') {
      return await handleBrevoEvent(req, res);
    }
    if (req.method === 'POST' && req.url === '/send-report') {
      return await handleSendReport(req, res);
    }
    if (req.method === 'GET' || req.method === 'HEAD') {
      return serveStatic(req, res);
    }
    sendJson(res, 405, { error: 'Metodo nao permitido.' });
  } catch (err) {
    sendJson(res, 500, { error: err.message || 'Erro interno.' });
  }
});

server.listen(PORT, () => {
  console.log(`Servidor do reporte em http://localhost:${PORT}`);
  console.log(`Endpoint de e-mail: http://localhost:${PORT}/send-report`);
});

async function handleSendReport(req, res) {
  if (!BREVO_API_KEY) return sendJson(res, 500, { error: 'BREVO_API_KEY nao configurada em .env.local.' });
  if (!REPORT_FROM_EMAIL) return sendJson(res, 500, { error: 'REPORT_FROM_EMAIL nao configurado em .env.local.' });
  if (REPORT_SERVER_TOKEN) {
    const got = req.headers['x-report-server-token'];
    if (got !== REPORT_SERVER_TOKEN) return sendJson(res, 401, { error: 'Token do servidor invalido.' });
  }

  const payload = await readJson(req);
  const config = payload.config || {};
  const to = parseEmailList(config.emailTo).length ? parseEmailList(config.emailTo) : REPORT_DEFAULT_TO;
  const cc = parseEmailList(config.emailCc).length ? parseEmailList(config.emailCc) : REPORT_DEFAULT_CC;
  if (!to.length) return sendJson(res, 400, { error: 'Nenhum destinatario informado.' });

  const subject = String(payload.subject || 'Reporte Operacional').slice(0, 180);
  const textContent = String(payload.body || '').trim() || buildPlainText(payload);
  const reportImageUrl = resolveReportImageUrl(payload);
  const htmlContent = buildHtml(payload, textContent, reportImageUrl);
  lastReportSnapshot = {
    ...payload,
    subject,
    textContent,
    reportImageUrl,
    sentAt: new Date().toISOString(),
  };

  const brevoPayload = {
    sender: { name: REPORT_FROM_NAME, email: REPORT_FROM_EMAIL },
    to: to.map(email => ({ email })),
    subject,
    htmlContent,
    textContent,
  };
  if (cc.length) brevoPayload.cc = cc.map(email => ({ email }));

  const response = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': BREVO_API_KEY,
      'content-type': 'application/json',
    },
    body: JSON.stringify(brevoPayload),
  });

  const body = await response.text();
  if (!response.ok) {
    return sendJson(res, response.status, { error: 'Brevo recusou o envio.', details: safeBrevoBody(body) });
  }

  const eventEmail = pickEventEmail(payload, to);
  if (eventEmail) {
    await postBrevoEvent({
      email: eventEmail,
      eventName: 'report_sent',
      contactProperties: {
        FIRSTNAME: REPORT_FROM_NAME,
      },
      eventProperties: {
        subject,
        recipients: to.length,
        cc: cc.length,
        report_image_url: reportImageUrl || '',
        report_source: payload?.config?.emailProvider || 'brevo',
      },
    }).catch(() => null);
  }

  sendJson(res, 200, { ok: true, brevo: safeBrevoBody(body) });
}

async function handleBrevoEvent(req, res) {
  if (!BREVO_API_KEY) return sendJson(res, 500, { error: 'BREVO_API_KEY nao configurada em .env.local.' });
  const payload = await readJson(req);
  const email = String(payload.email || payload.identifiers?.email_id || BREVO_EVENT_FALLBACK_EMAIL || '').trim();
  const eventName = sanitizeEventName(payload.eventName || payload.event_name || '');
  if (!email) return sendJson(res, 400, { error: 'Email do contato nao informado.' });
  if (!eventName) return sendJson(res, 400, { error: 'Nome do evento invalido.' });
  const result = await postBrevoEvent({
    email,
    eventName,
    contactProperties: payload.contactProperties || payload.contact_properties || {},
    eventProperties: payload.eventProperties || payload.event_properties || {},
    eventDate: payload.eventDate || payload.event_date || undefined,
  });
  return sendJson(res, 200, { ok: true, brevo: result });
}

function serveStatic(req, res) {
  const urlPath = decodeURIComponent((req.url || '/').split('?')[0]);
  const clean = path.normalize(urlPath).replace(/^(\.\.[/\\])+/, '');
  let filePath = path.join(ROOT, clean === '/' ? 'index.html' : clean);
  if (!filePath.startsWith(ROOT)) return sendJson(res, 403, { error: 'Acesso negado.' });
  if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) filePath = path.join(filePath, 'index.html');
  if (!fs.existsSync(filePath)) return sendJson(res, 404, { error: 'Arquivo nao encontrado.' });
  const ext = path.extname(filePath).toLowerCase();
  sendCors(res, 200, MIME[ext] || 'application/octet-stream');
  if (req.method === 'HEAD') return res.end();
  fs.createReadStream(filePath).pipe(res);
}

function serveReportImage(req, res) {
  const url = new URL(req.url, 'http://localhost');
  const query = Object.fromEntries(url.searchParams.entries());
  const snapshot = getReportSnapshotForImage(query);
  const svg = renderReportSvg(snapshot, query);
  sendCors(res, 200, 'image/svg+xml; charset=utf-8');
  res.setHeader('cache-control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('pragma', 'no-cache');
  res.setHeader('expires', '0');
  res.end(svg);
}

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const idx = trimmed.indexOf('=');
    if (idx < 0) continue;
    const key = trimmed.slice(0, idx).trim();
    const val = trimmed.slice(idx + 1).trim().replace(/^["']|["']$/g, '');
    if (key && process.env[key] === undefined) process.env[key] = val;
  }
}

function parseEmailList(value) {
  const raw = Array.isArray(value) ? value.join(';') : String(value || '');
  return raw.split(/[;,]/).map(v => v.trim()).filter(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v));
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', chunk => {
      raw += chunk;
      if (raw.length > 1_000_000) {
        req.destroy();
        reject(new Error('Payload muito grande.'));
      }
    });
    req.on('end', () => {
      try { resolve(raw ? JSON.parse(raw) : {}); }
      catch (err) { reject(new Error('JSON invalido.')); }
    });
    req.on('error', reject);
  });
}

function buildPlainText(payload) {
  const report = payload.report || {};
  return [
    payload.subject || 'Reporte Operacional',
    '',
    `Nossa grade: ${report.nossaGrade || 0}`,
    `Realizado: ${report.realizadoGrade || 0}`,
    `Pendente: ${Math.max(0, (report.nossaGrade || 0) - (report.realizadoGrade || 0))}`,
  ].join('\n');
}

function buildHtml(payload, textContent, reportImageUrl = '') {
  const report = payload.report || {};
  const det = payload.planejadoSuzanoDetalhado || {};
  const reportHref = String(payload.reportLink || REPORT_PUBLIC_BASE_URL || '').trim();
  return `<!doctype html>
<html><body style="font-family:Arial,sans-serif;background:#f8fafc;color:#0f172a;padding:24px;">
  <h2 style="margin:0 0 12px;">${escapeHtml(payload.subject || 'Reporte Operacional')}</h2>
  ${reportImageUrl ? `${reportHref ? `<a href="${escapeHtml(reportHref)}" target="_blank" rel="noopener noreferrer" style="display:block;margin:0 0 16px;">` : '<div style="display:block;margin:0 0 16px;">'}
    <img src="${escapeHtml(reportImageUrl)}" alt="${escapeHtml(payload.reportImageAlt || REPORT_IMAGE_TITLE)}" style="display:block;width:100%;max-width:900px;border:1px solid #e2e8f0;border-radius:12px;"/>
  ${reportHref ? '</a>' : '</div>'}` : ''}
  <pre style="white-space:pre-wrap;background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:16px;">${escapeHtml(textContent)}</pre>
  <h3 style="margin-top:22px;">Detalhe Suzano</h3>
  <ul>
    <li>Venda Aruja: ${escapeHtml(det['VENDA ARUJA'] || 0)}</li>
    <li>Venda Mogi: ${escapeHtml(det['VENDA MOGI'] || 0)}</li>
    <li>Transferencia: ${escapeHtml(det.TRANSFERENCIA || 0)}</li>
    <li>Pre-fatura: ${escapeHtml(det['PRE-FATURA'] || 0)}</li>
  </ul>
  <p style="color:#64748b;font-size:12px;">Gerado pelo dashboard de reporte em ${escapeHtml(new Date().toLocaleString('pt-BR'))}.</p>
</body></html>`;
}

function resolveReportImageUrl(payload) {
  const explicit = String(payload?.reportImageUrl || '').trim();
  if (explicit) return explicit;
  if (!REPORT_PUBLIC_BASE_URL) return '';
  const stamp = Date.now();
  return `${REPORT_PUBLIC_BASE_URL.replace(/\/$/, '')}/api/relatorio?ts=${stamp}`;
}

function getReportSnapshotForImage(query = {}) {
  const report = lastReportSnapshot?.report || {};
  const fallback = {
    title: REPORT_IMAGE_TITLE,
    dataRef: report.dataRef || '',
    planejadoSuzano: Number(report.planejadoSuzano || 0),
    nossaGrade: Number(report.nossaGrade || 0),
    realizadoGrade: Number(report.realizadoGrade || 0),
    pendente: Math.max(0, Number(report.nossaGrade || 0) - Number(report.realizadoGrade || 0)),
    dtsTotal: Number(report.statusCounts ? Object.values(report.statusCounts).reduce((a, b) => a + Number(b || 0), 0) : 0),
    dtsExpedidas: Number(report.statusCounts?.EXPEDIDO || 0),
    dtsNoShow: Number(report.statusCounts?.['NO SHOW'] || 0),
    dtsRecusadas: Number(report.statusCounts?.['VEICULO RECUSADO'] || 0),
    sentAt: lastReportSnapshot?.sentAt || new Date().toISOString(),
  };

  return {
    title: query.title || fallback.title,
    subtitle: query.subtitle || lastReportSnapshot?.subject || 'Resumo do reporte enviado',
    dataRef: query.data_ref || query.dataRef || fallback.dataRef,
    planejadoSuzano: toNumber(query.planejado_suzano_kg, fallback.planejadoSuzano),
    nossaGrade: toNumber(query.nossa_grade_kg, fallback.nossaGrade),
    realizadoGrade: toNumber(query.realizado_kg, fallback.realizadoGrade),
    pendente: toNumber(query.pendente_kg, fallback.pendente),
    dtsTotal: toNumber(query.dts_total, fallback.dtsTotal),
    dtsExpedidas: toNumber(query.dts_expedidas, fallback.dtsExpedidas),
    dtsNoShow: toNumber(query.dts_no_show, fallback.dtsNoShow),
    dtsRecusadas: toNumber(query.dts_recusadas, fallback.dtsRecusadas),
    sentAt: query.sent_at || fallback.sentAt,
  };
}

function renderReportSvg(snapshot, query = {}) {
  const W = 1200;
  const H = 630;
  const bg = String(query.theme || 'dark');
  const isDark = bg !== 'light';
  const colors = isDark
    ? { bg1: '#081120', bg2: '#0f172a', card: '#111827', text: '#f8fafc', muted: '#94a3b8', line: '#243244' }
    : { bg1: '#f8fafc', bg2: '#eef2ff', card: '#ffffff', text: '#0f172a', muted: '#475569', line: '#dbe4f0' };
  const metrics = [
    { label: 'Planejado Suzano', value: formatMetric(snapshot.planejadoSuzano), color: '#38bdf8' },
    { label: 'Nossa grade', value: formatMetric(snapshot.nossaGrade), color: '#60a5fa' },
    { label: 'Realizado', value: formatMetric(snapshot.realizadoGrade), color: '#22c55e' },
    { label: 'Pendente', value: formatMetric(snapshot.pendente), color: '#f59e0b' },
  ];
  const bars = [
    ['Expedidas', snapshot.dtsExpedidas, '#22c55e'],
    ['No show', snapshot.dtsNoShow, '#ef4444'],
    ['Recusas', snapshot.dtsRecusadas, '#dc2626'],
  ];
  const maxBar = Math.max(1, ...bars.map(([, value]) => Number(value) || 0));
  const barSvg = bars.map(([label, value, color], index) => {
    const barH = Math.max(18, Math.round(((Number(value) || 0) / maxBar) * 180));
    const x = 720 + index * 120;
    const y = 370 - barH;
    return `
      <rect x="${x}" y="${y}" width="70" height="${barH}" rx="10" fill="${color}" opacity="0.95"/>
      <text x="${x + 35}" y="392" text-anchor="middle" fill="${colors.muted}" font-size="16" font-family="Arial">${escapeXml(label)}</text>
      <text x="${x + 35}" y="${y - 12}" text-anchor="middle" fill="${colors.text}" font-size="22" font-weight="700" font-family="Arial">${escapeXml(String(value || 0))}</text>
    `;
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-label="${escapeXml(snapshot.title)}">
  <defs>
    <linearGradient id="bg" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0%" stop-color="${colors.bg1}"/>
      <stop offset="100%" stop-color="${colors.bg2}"/>
    </linearGradient>
    <filter id="shadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="18" stdDeviation="18" flood-color="#000000" flood-opacity="0.18"/>
    </filter>
  </defs>
  <rect width="1200" height="630" fill="url(#bg)"/>
  <circle cx="1050" cy="90" r="120" fill="#3b82f622"/>
  <circle cx="1090" cy="520" r="160" fill="#22c55e18"/>
  <rect x="42" y="42" width="1116" height="546" rx="28" fill="${colors.card}" filter="url(#shadow)"/>
  <text x="82" y="110" fill="${colors.text}" font-size="38" font-weight="700" font-family="Arial">${escapeXml(snapshot.title)}</text>
  <text x="82" y="150" fill="${colors.muted}" font-size="18" font-family="Arial">${escapeXml(snapshot.subtitle || '')}</text>
  <text x="82" y="188" fill="${colors.muted}" font-size="15" font-family="Arial">Data: ${escapeXml(snapshot.dataRef || 'sem data')}${snapshot.sentAt ? ` • ${escapeXml(new Date(snapshot.sentAt).toLocaleString('pt-BR'))}` : ''}</text>
  ${metrics.map((metric, index) => {
    const x = 82 + index * 235;
    return `
      <rect x="${x}" y="230" width="205" height="126" rx="20" fill="none" stroke="${colors.line}" stroke-width="2"/>
      <text x="${x + 18}" y="268" fill="${metric.color}" font-size="14" font-weight="700" font-family="Arial">${escapeXml(metric.label)}</text>
      <text x="${x + 18}" y="320" fill="${colors.text}" font-size="36" font-weight="700" font-family="Arial">${escapeXml(metric.value)}</text>
    `;
  }).join('')}
  <text x="720" y="214" fill="${colors.text}" font-size="24" font-weight="700" font-family="Arial">Status do dia</text>
  ${barSvg}
  <text x="82" y="486" fill="${colors.text}" font-size="22" font-weight="700" font-family="Arial">Resumo operacional</text>
  <text x="82" y="522" fill="${colors.muted}" font-size="16" font-family="Arial">Use esta imagem no Brevo com ?ts=TIMESTAMP para evitar cache e refletir a versao mais recente.</text>
</svg>`;
}

async function postBrevoEvent({ email, eventName, contactProperties = {}, eventProperties = {}, eventDate }) {
  const payload = {
    event_name: sanitizeEventName(eventName),
    identifiers: { email_id: email },
  };
  if (eventDate) payload.event_date = eventDate;
  if (contactProperties && Object.keys(contactProperties).length) payload.contact_properties = contactProperties;
  if (eventProperties && Object.keys(eventProperties).length) payload.event_properties = eventProperties;

  const response = await fetch('https://api.brevo.com/v3/events', {
    method: 'POST',
    headers: {
      accept: 'application/json',
      'api-key': BREVO_API_KEY,
      'content-type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  const body = await response.text();
  if (!response.ok) {
    throw new Error(`Brevo events API: ${response.status} ${body.slice(0, 400)}`);
  }
  return body ? safeBrevoBody(body) : { ok: true };
}

function sanitizeEventName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_-]/g, '')
    .slice(0, 255);
}

function pickEventEmail(payload, toList) {
  const candidates = [
    payload?.eventEmail,
    payload?.brevoEmail,
    payload?.config?.brevoTrackEmail,
    payload?.config?.emailFrom,
    Array.isArray(toList) && toList[0],
    BREVO_EVENT_FALLBACK_EMAIL,
  ];
  return candidates.map(v => String(v || '').trim()).find(Boolean) || '';
}

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : Number(fallback) || 0;
}

function formatMetric(value) {
  const n = Number(value) || 0;
  if (Math.abs(n) >= 1000) {
    const t = (n / 1000).toFixed(1).replace(/\.0$/, '');
    return `${t} t`;
  }
  return `${Math.round(n)} kg`;
}

function escapeXml(value) {
  return String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function sendCors(res, status, contentType = 'application/json; charset=utf-8') {
  res.writeHead(status, {
    'access-control-allow-origin': '*',
    'access-control-allow-methods': 'GET,POST,OPTIONS',
    'access-control-allow-headers': 'content-type,x-report-server-token',
    'content-type': contentType,
  });
}

function sendJson(res, status, payload) {
  sendCors(res, status);
  res.end(JSON.stringify(payload));
}

function safeBrevoBody(body) {
  try { return JSON.parse(body); }
  catch { return String(body || '').slice(0, 500); }
}
