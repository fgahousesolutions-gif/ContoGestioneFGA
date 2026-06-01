# Rendiconto Sandu Francesco

Webapp locale per caricare il file Excel degli appartamenti e mantenere salvato l'ultimo aggiornamento.

## Importante per GitHub

Questa non e una pagina statica: l'upload Excel funziona solo se gira anche `server.py`.

GitHub Pages da solo non basta, perche non esegue Python e non puo salvare `data/current_state.json`.

Per pubblicarla online su Cloudflare Pages:

1. Carica questa cartella su GitHub.
2. Collega la repo a Cloudflare Pages.
3. Framework preset: `None`.
4. Build command: `npm install`.
5. Build output directory: `public`.
6. Crea un KV namespace e collegalo al progetto Pages con binding `APP_STATE`.

Le API Cloudflare sono in:

```text
functions/api/current.js
functions/api/upload.js
```

La versione locale con Python resta disponibile usando `server.py`.

## Avvio

```bash
cd ~/Downloads/rendiconto-sandu-francesco
python3 server.py
```

Poi aprire:

```text
http://127.0.0.1:8788
```

## File Excel

La webapp legge file `.xlsx`. Se il file e un vecchio `.xls`, salvarlo prima come `.xlsx`.

Fogli analizzati:

- `ID003`
- `ID004`
- `ID005`
- `ID006`
- `TOTALE SPESE Sandu Francesco`

L'ultimo stato calcolato viene salvato in:

```text
data/current_state.json
```

## Regole quote

- `ID003`: 25% gestione, 20% Francesco e 5% Sandu.
- `ID004`: 25% gestione, 15% Francesco e 10% Sandu.
- `ID005`: 10% gestione, 5% Francesco e 5% Sandu.
- `ID006`: 15% gestione, 10% Francesco e 5% Sandu.

La quota viene calcolata sul netto proprieta dopo OTA e pulizie quando non sono presenti colonne gia calcolate.
