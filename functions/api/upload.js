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
  let payload = {};
  try {
    payload = await request.json();
  } catch (error) {
    return Response.json({ error: "Richiesta non valida: carica il file dalla webapp, non come form generico." }, { status: 400 });
  }
  const filename = String(payload.filename || "");
  if (!filename) {
    return Response.json({ error: "Nome file mancante." }, { status: 400 });
  }
  if (!filename.toLowerCase().endsWith(".xlsx")) {
    return Response.json({ error: "Carica un file Excel .xlsx. Se e un vecchio .xls o viene da Numbers, esportalo prima come .xlsx." }, { status: 400 });
  }
  let bytes;
  try {
    bytes = base64ToBytes(String(payload.contentBase64 || ""));
  } catch (error) {
    return Response.json({ error: "File non leggibile: contenuto base64 non valido." }, { status: 400 });
  }
  if (!bytes.length) {
    return Response.json({ error: "Il file caricato e vuoto." }, { status: 400 });
  }
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    return Response.json({ error: "Il file non sembra un .xlsx valido. Aprilo con Excel/Numbers ed esportalo come Excel .xlsx." }, { status: 400 });
  }
  try {
    const state = parseWorkbook(bytes, filename);
    await env.APP_STATE.put("current_state", JSON.stringify(state));
    return Response.json(state);
  } catch (error) {
    return Response.json({ error: `Errore durante la lettura del file: ${error.message || error}` }, { status: 500 });
  }
}
