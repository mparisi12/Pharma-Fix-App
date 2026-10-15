# Test di Pharma-Fix

Prove automatiche da lanciare dopo ogni modifica, prima di pubblicare. Usano **dati inventati** e un finto Google Sheet: non toccano il foglio vero.

```bash
cd test
npm install
npx playwright install chromium   # una volta sola
npm test
```

- `server.test.js`: lo script (Code.gs): versione dei dati e copia pronta, documenti, coordinatori.
- `app.test.js`: l'app nel browser: accesso, apri fermo, documenti, compiti (popup e presa in carico), vestizione.

Ogni riga stampa `ok` o `NO`; alla fine *tutto ok* oppure quanti test sono falliti.
