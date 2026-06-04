import { buildReport, dateKey, getConfig, json, sendEmail } from "./_shared";

function assertCronSecret(request, env) {
  if (!env.CRON_SECRET) return;
  const auth = request.headers.get("authorization") || "";
  if (auth !== `Bearer ${env.CRON_SECRET}`) {
    const err = new Error("Unauthorized");
    err.status = 401;
    throw err;
  }
}

async function sendForDate(env, config, offsetDays) {
  const dataRef = dateKey(env, offsetDays);
  const report = await buildReport(env, dataRef);
  const brevo = await sendEmail(env, { report, config });
  return { dataRef, brevo };
}

export async function onRequestGet({ request, env }) {
  try {
    assertCronSecret(request, env);
    const config = await getConfig(env);
    if (!config.autoEmailEnabled) return json({ ok: true, skipped: "auto_disabled" });
    const results = [await sendForDate(env, config, 0)];
    if (config.sendNextDayAtFinal !== false) results.push(await sendForDate(env, config, 1));
    return json({ ok: true, results });
  } catch (err) {
    return json({ error: err.message || "Erro interno." }, err.status || 500);
  }
}
