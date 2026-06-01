# Rendiconto Sandu Francesco

Webapp locale per caricare il file Excel degli appartamenti e mantenere salvato l'ultimo aggiornamento.

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
