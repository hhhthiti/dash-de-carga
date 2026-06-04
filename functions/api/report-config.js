import { getConfig, json, saveConfig } from "./_shared";

export async function onRequestGet({ env }) {
  try {
    return json({ ok: true, config: await getConfig(env) });
  } catch (err) {
    return json({ error: err.message || "Erro interno." }, 500);
  }
}

export async function onRequestPost({ request, env }) {
  try {
    const body = await request.json();
    return json({ ok: true, config: await saveConfig(env, body || {}) });
  } catch (err) {
    return json({ error: err.message || "Erro interno." }, 500);
  }
}
