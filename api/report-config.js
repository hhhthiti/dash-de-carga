const { getConfig, json, readJson, saveConfig } = require('./_report-shared');

module.exports = async function handler(req, res) {
  try {
    if (req.method === 'GET') {
      return json(res, 200, { ok: true, config: await getConfig() });
    }
    if (req.method === 'POST') {
      const body = await readJson(req);
      return json(res, 200, { ok: true, config: await saveConfig(body || {}) });
    }
    return json(res, 405, { error: 'Metodo nao permitido.' });
  } catch (err) {
    return json(res, 500, { error: err.message || 'Erro interno.' });
  }
};
