# Pharma-Fix Fermi — installazione

App per registrare i fermi macchina da telefono: QR sulla macchina → foto → apertura fermo → chiusura con foto della zona e dati del template. Ogni fermo diventa una riga in un **Google Sheet** che si aggiorna da solo, con le stesse colonne del "Template fermi meccanici".

```
Telefono (app su GitHub Pages)  ──►  Google Apps Script  ──►  Google Sheet "Fermi" + foto in Google Drive
```

Tempo per installare tutto: circa 30 minuti.

---

## 1. Il Google Sheet (database)

1. Su Google Drive crea un nuovo foglio Google e chiamalo **Pharma-Fix Fermi**.
2. Menu **Estensioni → Apps Script**.
3. Cancella il contenuto di `Codice.gs` e incolla tutto il file `apps-script/Code.gs`.
4. Icona ingranaggio **Impostazioni progetto** → spunta **Mostra il file manifest "appsscript.json"**. Torna all'editor, apri `appsscript.json` e sostituiscilo con il file `apps-script/appsscript.json`.
5. Salva (icona dischetto). In alto scegli la funzione **setup** e premi **Esegui**.
   - Google chiede le autorizzazioni: **Rivedi autorizzazioni** → il tuo account → *Avanzate* → *Vai a … (non sicuro)* → **Consenti**. È normale: lo script è tuo e non è verificato da Google.
6. Apri **Registro di esecuzione** in basso: trovi la riga `Chiave da copiare in config.js → KEY: xxxxxxxxxxxx`. **Copia la chiave.**
7. Nel foglio ora ci sono 4 schede:
   - **Fermi**: colonne A–L identiche al template, poi Linea, Turno, Codice pezzo, Foto apertura, Foto chiusura, Stato, ID fermo, Equipment, Aperto da. Le righe gialle sono fermi ancora aperti.
   - **Equipment**: elenco macchine con il codice che finisce nel QR. **Controllalo**: ho messo Riordinatore + Filling su tutte le linee con filling e le 6 macchine del secondario su tutte le linee con packaging. Correggi nomi, aggiungi o togli righe (colonna Attivo = NO per nasconderle).
   - **Liste**: categorie, tipi intervento, guasti sterile/secondario, manutentori. Quello che scrivi qui compare come pulsante nell'app.
   - **KPI**: per linea n° fermi, tempo di fermo, MTTR, n° lotti, Occurrancy, MTBF (inserisci tu le ore di funzionamento nella colonna gialla), più Pareto guasti con RFR e classifica macchine. Si aggiorna da solo.
8. In Apps Script: **Esegui il deployment → Nuovo deployment** → tipo **App web**.
   - Esegui come: **Me**
   - Chi ha accesso: **Chiunque**
   - **Esegui il deployment** → copia l'**URL dell'app web** (finisce con `/exec`).

> Se in futuro modifichi `Code.gs`: **Esegui il deployment → Gestisci deployment → matita → Versione: Nuova versione**. Così l'URL resta lo stesso.

## 2. L'app su GitHub Pages

1. L'app sta nella cartella `fermi/` del repository **Pharma-Fix-App**, accanto all'app Pharma-Fix. In Pharma-Fix c'è un nuovo pulsante **⏱️ FERMI MACCHINA** che la apre.
3. Apri `config.js` su GitHub (matita) e incolla URL e chiave:
   ```js
   window.PF_CONFIG = {
     API_URL: 'https://script.google.com/macros/s/……/exec',
     KEY: 'la-chiave-del-passo-1.6'
   };
   ```
4. Se GitHub Pages è già attivo per Pharma-Fix non serve altro: dopo un minuto l'app è su
   `https://mparisi12.github.io/Pharma-Fix-App/fermi/`
   (altrimenti **Settings → Pages → Branch: main / (root) → Save**).

Finché `API_URL` è vuoto l'app funziona in **modalità demo** (dati solo sul telefono): utile per provarla.

## 3. Le etichette QR

1. Apri `https://mparisi12.github.io/Pharma-Fix-App/fermi/qr.html` dal PC.
2. Scegli una linea (o tutte) → **Stampa** al 100%, su carta adesiva o da plastificare.
3. Attacca ogni etichetta sulla macchina giusta, in un punto visibile dal lato operatore.

Il QR contiene l'indirizzo `…/?eq=ST3-AST`: si legge dallo scanner dell'app **e** dalla fotocamera normale del telefono. Se aggiungi macchine nel foglio Equipment, ristampa solo quelle.

## 4. I telefoni del personale

1. Apri `https://mparisi12.github.io/Pharma-Fix-App/fermi/` in **Chrome** (Android) o **Safari** (iPhone).
2. Installa: Chrome → menu ⋮ → **Installa app** / **Aggiungi a schermata Home**; Safari → Condividi → **Aggiungi alla schermata Home**.
3. Al primo fermo l'app chiede **Chi sei?**: si sceglie il nome una volta sola (si cambia toccando il nome in alto).
4. Consentire l'uso della fotocamera quando richiesto.

**Senza rete?** L'app registra lo stesso: i fermi restano in coda sul telefono (pallino arancione/rosso in alto) e partono da soli appena torna la connessione.

## 5. Usare i dati in Excel

- **Download veloce**: dal Google Sheet **File → Scarica → Microsoft Excel (.xlsx)**.
- **Excel sempre aggiornato** (anche con Microsoft 365): nel Google Sheet **File → Condividi → Pubblica sul web** → scheda **Fermi** → formato **CSV** → Pubblica, copia il link. In Excel: **Dati → Da Web** → incolla il link → Carica. Poi **Dati → Aggiorna tutto** (o imposta l'aggiornamento automatico ogni X minuti nelle proprietà della query). Le tue pivot e la dashboard possono puntare a questa tabella.
  ⚠️ Chi ha quel link può leggere i dati del foglio Fermi (non le foto): non diffonderlo.

## Come si usa (per il personale)

| Momento | Cosa fare | Tempo |
|---|---|---|
| La macchina si ferma | Apri l'app → **Scansiona QR** → **Apri fermo** → foto → controlla il lotto → **Apri fermo** | ~10 s |
| Hai finito l'intervento | Nella home o nella scheda macchina tocca **Chiudi** → foto della zona, tocca per cerchiarla → Categoria, Intervento, Guasto → descrizione se serve → **Chiudi fermo** | ~20–30 s |
| Aperto per sbaglio | **Chiudi** → in fondo **Annulla il fermo** | |

Regole per dati puliti: un guasto per riga (come il template); se scegli **ALTRO** la descrizione è obbligatoria; il codice pezzo va nel suo campo.

## Note di sicurezza

- La chiave in `config.js` impedisce scritture casuali ma è visibile a chi apre il codice della pagina: non pubblicare il link dell'app fuori dal reparto. Per cambiarla: Apps Script → Impostazioni progetto → Proprietà script → `KEY`, poi aggiorna `config.js`.
- Le foto restano private nella cartella Drive **Pharma-Fix Foto Fermi** del tuo account; nel foglio c'è il link.
- È uno strumento di analisi della manutenzione, non un sistema validato GMP.
