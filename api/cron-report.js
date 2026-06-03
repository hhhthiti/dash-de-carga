const {
  alreadyRan,
  buildReport,
  dateKey,
  getConfig,
  insideWindow,
  json,
  localParts,
  markRan,
  sendEmail,
} = require('./_report-shared');

function requireCron(req) {
  if (!process.env.CRON_SECRET) return;
  const auth = req.headers.authorization || '';
  if (auth !== `Bearer ${process.env.CRON_SECRET}`) {
    const err = new Error('Unauthorized');
    err.status = 401;
    throw err;
  }
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== 'GET') return json(res, 405, { error: 'Metodo nao permitido.' });
    requireCron(req);

    const config = await getConfig();
    if (!config.autoEmailEnabled) return json(res, 200, { ok: true, skipped: 'auto_disabled' });
    if (!insideWindow(config)) return json(res, 200, { ok: true, skipped: 'outside_window' });

    const now = localParts(new Date());
    const dataRef = dateKey(0);
    const slotKey = `hourly:${dataRef}:${now.hour}`;
    if (await alreadyRan(slotKey)) return json(res, 200, { ok: true, skipped: 'already_sent', slotKey });

    const report = await buildReport(dataRef);
    const result = await sendEmail({ report, config, source: 'AUTO_HOURLY' });
    await markRan(slotKey, { dataRef, source: 'AUTO_HOURLY' });
    return json(res, 200, { ok: true, slotKey, result });
  } catch (err) {
    return json(res, err.status || 500, { error: err.message || 'Erro interno.' });
  }
};
