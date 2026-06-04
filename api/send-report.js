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
    const incomingReport = normalizeIncomingReport(body.report || {});
    const dataRef = incomingReport.dataRef || body?.dataRef || dateKey(0);
    const report = incomingReport.dataRef ? incomingReport : await buildReport(dataRef);
    const result = await sendEmail({ report, config, source: 'MANUAL' });
    return json(res, 200, { ok: true, ...result });
  } catch (err) {
    return json(res, 500, { error: err.message || 'Erro interno.' });
  }
};

function normalizeIncomingReport(report) {
  const dataRef = report.dataRef || report.data_ref || '';
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
