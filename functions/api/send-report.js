import { buildReport, dateKey, getConfig, json, sendEmail } from "./_shared";

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

export async function onRequestPost({ request, env }) {
  try {
    const body = await request.json();
    const savedConfig = await getConfig(env);
    const config = { ...savedConfig, ...(body.config || {}) };
    const incomingReport = normalizeIncomingReport(body.report || {});
    const dataRef = incomingReport.dataRef || body.dataRef || dateKey(env, 0);
    const report = incomingReport.dataRef ? incomingReport : await buildReport(env, dataRef);
    const brevo = await sendEmail(env, { report, config });
    return json({ ok: true, brevo });
  } catch (err) {
    return json({ error: err.message || "Erro interno." }, 500);
  }
}

export async function onRequestGet() {
  return json({ error: "Metodo nao permitido." }, 405);
}
