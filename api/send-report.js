const {
  buildReport,
  dateKey,
  getConfig,
  json,
  readJson,
  sendEmail,
} = require('./_report-shared');

module.exports = async function handler(req, res) {
  try {
    if (req.method !== 'POST') return json(res, 405, { error: 'Metodo nao permitido.' });
    const body = await readJson(req);
    const savedConfig = await getConfig();
    const config = { ...savedConfig, ...(body.config || {}) };
    const dataRef = body?.report?.dataRef || body?.dataRef || dateKey(0);
    const report = body.report && body.report.dataRef ? body.report : await buildReport(dataRef);
    const result = await sendEmail({ report, config, source: 'MANUAL' });
    return json(res, 200, { ok: true, ...result });
  } catch (err) {
    return json(res, 500, { error: err.message || 'Erro interno.' });
  }
};
