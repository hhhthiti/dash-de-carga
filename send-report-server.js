const http = require('http');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const ENV_PATH = path.join(ROOT, '.env.local');
const PORT = Number(process.env.REPORT_SERVER_PORT || 8787);

loadEnvFile(ENV_PATH);

const BREVO_API_KEY = process.env.BREVO_API_KEY || '';
const REPORT_FROM_EMAIL = process.env.REPORT_FROM_EMAIL || '';
const REPORT_FROM_NAME = process.env.REPORT_FROM_NAME || 'Reporte Operacional';
const REPORT_DEFAULT_TO = parseEmailList(process.env.REPORT_DEFAULT_TO || '');
const REPORT_DEFAULT_CC = parseEmailList(process.env.REPORT_DEFAULT_CC || '');
const REPORT_SERVER_TOKEN = process.env.REPORT_SERVER_TOKEN || '';

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
  const htmlContent = buildHtml(payload, textContent);

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
  sendJson(res, 200, { ok: true, brevo: safeBrevoBody(body) });
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

function buildHtml(payload, textContent) {
  const report = payload.report || {};
  const det = payload.planejadoSuzanoDetalhado || {};
  return `<!doctype html>
<html><body style="font-family:Arial,sans-serif;background:#f8fafc;color:#0f172a;padding:24px;">
  <h2 style="margin:0 0 12px;">${escapeHtml(payload.subject || 'Reporte Operacional')}</h2>
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
