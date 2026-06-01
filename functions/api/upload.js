import { parseWorkbook } from "../_lib/parser.js";

function base64ToBytes(value) {
  const binary = atob(value || "");
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export async function onRequestPost({ request, env }) {
  if (!env.APP_STATE) {
    return Response.json({ error: "Binding KV APP_STATE mancante su Cloudflare." }, { status: 500 });
  }
  const payload = await request.json();
  const filename = String(payload.filename || "");
  if (!filename.toLowerCase().endsWith(".xlsx")) {
    return Response.json({ error: "Carica un file .xlsx." }, { status: 400 });
  }
  const bytes = base64ToBytes(String(payload.contentBase64 || ""));
  if (!bytes.length) {
    return Response.json({ error: "Il file caricato e vuoto." }, { status: 400 });
  }
  const state = parseWorkbook(bytes, filename);
  await env.APP_STATE.put("current_state", JSON.stringify(state));
  return Response.json(state);
}
