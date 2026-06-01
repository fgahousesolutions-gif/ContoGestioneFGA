export async function onRequestGet({ env }) {
  if (!env.APP_STATE) {
    return Response.json({ state: null, error: "Binding KV APP_STATE mancante su Cloudflare." }, { status: 500 });
  }
  const value = await env.APP_STATE.get("current_state");
  return Response.json({ state: value ? JSON.parse(value) : null });
}
