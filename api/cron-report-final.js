const {
  alreadyRan,
  buildReport,
  dateKey,
  getConfig,
  json,
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

async function sendOnce(dataRef, source, config) {
  const slotKey = `${source}:${dataRef}`;
  if (await alreadyRan(slotKey)) return { dataRef, skipped: 'already_sent', slotKey };
  const report = await buildReport(dataRef);
  const result = await sendEmail({ report, config, source });
  await markRan(slotKey, { dataRef, source });
  return { dataRef, slotKey, result };
}

module.exports = async function handler(req, res) {
  try {
    if (req.method !== 'GET') return json(res, 405, { error: 'Metodo nao permitido.' });
    requireCron(req);

    const config = await getConfig();
    if (!config.autoEmailEnabled) return json(res, 200, { ok: true, skipped: 'auto_disabled' });

    const results = [];
    results.push(await sendOnce(dateKey(0), 'AUTO_FINAL_CURRENT', config));
    if (config.sendNextDayAtFinal !== false) {
      results.push(await sendOnce(dateKey(1), 'AUTO_FINAL_NEXT_DAY', config));
    }
    return json(res, 200, { ok: true, results });
  } catch (err) {
    return json(res, err.status || 500, { error: err.message || 'Erro interno.' });
  }
};
