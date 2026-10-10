# Pharma-Fix · Fermi macchina

App per registrare i fermi macchina dal telefono in meno di 30 secondi: si scansiona il QR sulla macchina, si apre il fermo, a fine intervento si chiude con la foto della zona e i dati del template "fermi meccanici". Ogni fermo diventa una riga di un foglio Google che si aggiorna da solo e si può aprire anche in Excel.

Indirizzo dell'app: **https://mparisi12.github.io/Pharma-Fix-App/fermi/**
(Pharma-Fix si apre già qui; il pulsante **Menu** in alto riporta alle altre funzioni.)

---

## Indice

1. [Come funziona: il ruolo del Google Sheet](#1-come-funziona-il-ruolo-del-google-sheet)
2. [I file della cartella](#2-i-file-della-cartella)
3. [Installazione (una volta sola)](#3-installazione-una-volta-sola)
4. [Uso quotidiano — per i tecnici](#4-uso-quotidiano--per-i-tecnici)
5. [Gestione — per il responsabile](#5-gestione--per-il-responsabile)
6. [I dati in Excel](#6-i-dati-in-excel)
7. [Aggiornare l'app o lo script](#7-aggiornare-lapp-o-lo-script)
8. [Problemi frequenti](#8-problemi-frequenti)
9. [Sicurezza e limiti](#9-sicurezza-e-limiti)

---

## 1. Come funziona: il ruolo del Google Sheet

L'app è fatta di due parti che vivono in posti diversi:

```
 ┌───────────────────────┐        ┌──────────────────────────┐        ┌──────────────────────────────┐
 │  TELEFONO del tecnico │        │  GOOGLE APPS SCRIPT      │        │  GOOGLE SHEET "Pharma-Fix    │
 │  app web su GitHub    │ ─────► │  (lo "sportello" che     │ ─────► │  Fermi" = il DATABASE        │
 │  Pages (questa        │ ◄───── │  riceve e consegna i     │ ◄───── │  + cartella Drive con le foto│
 │  cartella)            │        │  dati)                   │        │                              │
 └───────────────────────┘        └──────────────────────────┘        └──────────────────────────────┘
```

- **GitHub Pages** ospita solo la *pagina* dell'app: grafica, pulsanti, scanner QR. Non conserva nessun dato.
- **Il Google Sheet è il database.** Ogni fermo aperto o chiuso diventa una riga del foglio **Fermi**. Da lì leggi tutto, filtri, fai pivot, scarichi in Excel.
- **Google Apps Script** è un piccolo programma *attaccato* al Google Sheet (Estensioni → Apps Script). Pubblicato come "App web", ha un indirizzo che finisce con `/exec`: l'app sul telefono manda lì i dati, lo script li scrive nel foglio e salva le foto in una cartella di Google Drive. È lo stesso meccanismo che Pharma-Fix usa già per Troubleshooting e Segnala anomalia.
- **Perché così:** costa zero, non serve un server, i tecnici non devono avere un account Google (scrive lo script a nome tuo), e tu vedi i dati in tempo reale in un foglio che conosci.

### I fogli che lo script crea

| Foglio | A cosa serve | Chi lo modifica |
|---|---|---|
| **Fermi** | Una riga per fermo. Colonne A–L **identiche al template** (Lotto, Data, Ora Inizio, Ora Fine, Manutentore, Macchina, Categoria, Tipo intervento, Tipo guasto, Altro, Durata MTTR, Intervento-Guasto), poi Linea, Turno, Codice pezzo, Foto apertura, Foto chiusura, Stato, ID fermo, Equipment, Aperto da. Le righe **gialle** sono fermi ancora aperti. | L'app. Tu puoi correggere a mano un errore, ma **non toccare le colonne Stato e ID fermo**. |
| **Equipment** | Elenco macchine: codice del QR, linea, area, nome macchina, attivo. | Tu. |
| **Liste** | I pulsanti dell'app: Categoria, Tipo intervento, Guasto Sterile, Guasto Secondario, **Manutentori abilitati**. | Tu. |
| **KPI** | Indicatori calcolati da soli con formule: per linea n° fermi, tempo di fermo, MTTR, n° lotti, Occurrancy, MTBF; Pareto guasti con RFR; classifica macchine. | Solo la colonna gialla "Ore funzionamento" (per l'MTBF). Il resto è automatico. |

Le foto finiscono nella cartella Drive **Pharma-Fix Foto Fermi**; nel foglio Fermi c'è il link "Foto".

### Chiave e indirizzo

- **API_URL**: l'indirizzo `/exec` dello script. Dice all'app *dove* mandare i dati.
- **KEY**: una parola segreta generata da `setup()`. Lo script accetta dati solo se arrivano con la chiave giusta, così nessuno può scrivere nel foglio per caso.

Entrambi vanno in `config.js`. Finché `API_URL` è vuoto l'app gira in **modalità demo**: funziona, ma i dati restano solo sul telefono (scritta "MODALITÀ DEMO" in alto).

---

## 2. I file della cartella

| File | Cosa contiene | Va modificato? |
|---|---|---|
| `index.html` | L'app: accesso, scanner QR, scheda macchina, apertura e chiusura fermo, coda offline. | No |
| `config.js` | Indirizzo dello script (`API_URL`) e chiave (`KEY`). | **Sì, una volta**, dopo l'installazione |
| `defaults.js` | Liste e macchine "di partenza", usate in modalità demo e finché il foglio non risponde. Con il foglio collegato vale quello che c'è nel foglio. | No (le liste si cambiano nel foglio) |
| `qr.html` | Pagina per stampare le etichette QR, una per macchina, prese dal foglio Equipment. | No |
| `manifest.webmanifest` | Nome e icona quando si installa l'app sul telefono. | No |
| `sw.js` | Permette all'app di aprirsi anche senza rete. | No |
| `icon-192.png`, `icon-512.png` | Icone dell'app. | No |
| `apps-script/Code.gs` | Il programma da incollare nel Google Sheet (Estensioni → Apps Script). Crea i fogli, riceve i dati, salva le foto, controlla i manutentori. | No, si **copia** nel Google Sheet |
| `apps-script/appsscript.json` | Impostazioni dello script (fuso orario Europe/Rome, accesso app web). | No, si **copia** nel Google Sheet |
| `README.md` | Questo file. | — |

Nella cartella principale del repository, `index.html` è Pharma-Fix: all'apertura rimanda qui; con `?menu=1` (pulsante **Menu**) mostra le funzioni di sempre.

---

## 3. Installazione (una volta sola)

Tempo: circa 30 minuti. Serve un account Google (quello che usi per Pharma-Fix va bene) e il PC.

### 3.1 Creare il Google Sheet e lo script

1. Su Google Drive: **Nuovo → Fogli Google**. Chiamalo **Pharma-Fix Fermi**.
2. Menu **Estensioni → Apps Script**. Si apre l'editor dello script.
3. Cancella il contenuto di `Codice.gs` e incolla **tutto** il file `apps-script/Code.gs` (aprilo su GitHub → pulsante "Copy raw file").
4. Icona ingranaggio **Impostazioni progetto** → spunta **Mostra il file manifest "appsscript.json" nell'editor**. Torna all'editor (icona `< >`), apri `appsscript.json` e sostituiscilo con il contenuto di `apps-script/appsscript.json`.
5. Salva (icona dischetto).
6. In alto, nel menu a tendina delle funzioni, scegli **setup** e premi **Esegui**.
   - Google chiede le autorizzazioni: **Rivedi autorizzazioni** → scegli il tuo account → **Avanzate** → **Vai a Pharma-Fix Fermi (non sicuro)** → **Consenti**. È normale: lo script è tuo e Google non l'ha "verificato".
7. In basso si apre il **Registro di esecuzione** con la riga
   `FATTO. Chiave da copiare in config.js → KEY: xxxxxxxxxxxx` → **copia la chiave**.
8. Torna al Google Sheet: ora ci sono i fogli **Fermi, Equipment, Liste, KPI**.

### 3.2 Pubblicare lo script come App web

1. In Apps Script: **Esegui il deployment → Nuovo deployment**.
2. Icona ingranaggio accanto a "Seleziona tipo" → **App web**.
3. Esegui come: **Me** · Chi ha accesso: **Chiunque**.
4. **Esegui il deployment** → copia l'**URL dell'app web** (finisce con `/exec`).

### 3.3 Collegare l'app al foglio

1. Su GitHub apri `fermi/config.js` → matita ✏️.
2. Incolla i due valori:
   ```js
   window.PF_CONFIG = {
     API_URL: 'https://script.google.com/macros/s/AKfy……/exec',
     KEY: 'xxxxxxxxxxxx'
   };
   ```
3. **Commit changes**. Dopo un minuto l'app è collegata: in alto compare "Sincronizzato hh:mm" al posto di "MODALITÀ DEMO".

### 3.4 Preparare liste, macchine ed etichette

1. Foglio **Liste** → colonna **Manutentori**: controlla che ci siano tutti i tecnici, scritti `COGNOME Nome` (vedi [5.1](#51-tecnici-abilitati)).
2. Foglio **Equipment**: controlla e completa le macchine (vedi [5.2](#52-lista-macchine-foglio-equipment)).
3. Dal PC apri **https://mparisi12.github.io/Pharma-Fix-App/fermi/qr.html** → scegli la linea → **Stampa** al 100% su carta adesiva o da plastificare → attacca ogni etichetta sulla sua macchina, lato operatore.

### 3.5 Installare l'app sui telefoni

1. Apri **https://mparisi12.github.io/Pharma-Fix-App/fermi/** con **Chrome** (Android) o **Safari** (iPhone).
2. Chrome: menu ⋮ → **Installa app** / **Aggiungi a schermata Home**. Safari: **Condividi → Aggiungi alla schermata Home**.
3. Alla prima scansione consenti l'uso della **fotocamera**.

---

## 4. Uso quotidiano — per i tecnici

### Accesso
All'apertura l'app chiede **Nome** e **Cognome**. Entra solo chi è tra i manutentori abilitati (maiuscole, accenti e ordine non contano: "nicolo greco" va bene). Ogni fermo viene registrato a nome di chi è entrato.
L'accesso dura un turno (9 ore). Se passi il telefono a un collega: tocca il tuo nome in alto → **Esci**.

### Aprire un fermo (~10 secondi)
1. **Scansiona QR** e inquadra l'etichetta sulla macchina. (In alternativa: fotocamera normale del telefono sul QR, oppure **Fotografa QR**, oppure **Scegli macchina a mano**.)
2. Si apre **Apri fermo**: controlla il **lotto** (è già scritto quello dell'ultimo fermo sulla linea) e, se il fermo è iniziato prima, scegli −5 / −10 / −15 / −30 min.
3. **Apri fermo**. Il cronometro parte e il fermo compare nella home tra i **Fermi aperti**.

### Chiudere un fermo (~20–30 secondi)
1. **Scansiona di nuovo il QR** della stessa macchina (oppure, nella home, **Chiudi** sul fermo).
2. **Foto zona di intervento** → scatta → **tocca la foto per cerchiare** il punto in cui hai lavorato (tocca il cerchio per toglierlo).
3. Scegli **Categoria**, **Tipo intervento**, **Tipo guasto** (uno solo, come nel template).
4. **Altro (descrizione breve)**: obbligatoria se hai scelto ALTRO. Per chiusure astucci scrivi zona, lato macchina/OP, n° nonio.
5. **Codice pezzo sostituito** se hai cambiato un ricambio.
6. **Chiudi fermo**.

Fermo aperto per sbaglio? **Chiudi** → in fondo **Annulla il fermo** (tocca due volte).

### Il pallino in alto
| Pallino | Significa |
|---|---|
| 🟢 Sincronizzato | Tutto è arrivato nel foglio. |
| 🟠 Invio in corso | Ci sono registrazioni in coda, partono da sole. |
| 🔴 Non connesso | Niente rete: **continua pure a lavorare**, i fermi restano sul telefono e partono appena torna la connessione. **Non disinstallare l'app** finché ci sono dati in coda. |
| ⚪ Modalità demo | L'app non è ancora collegata al foglio. |

---

## 5. Gestione — per il responsabile

Tutto si gestisce dal **Google Sheet**: l'app rilegge liste e macchine a ogni apertura, non serve toccare GitHub.

### 5.1 Tecnici abilitati
Foglio **Liste**, colonna **Manutentori**: un nome per cella, formato `COGNOME Nome` (es. `GRECO Nicolò`).
- **Nuovo tecnico** → aggiungi una cella. Può entrare subito.
- **Tecnico che lascia** → cancella la cella. Lo storico resta; non potrà più registrare (lo blocca anche lo script).
- Il nome scritto qui è quello che finisce nella colonna **Manutentore** del foglio Fermi.

### 5.2 Lista macchine (foglio Equipment)

Una riga per ogni macchina che deve avere il suo QR:

| Codice (QR) | Linea | Area | Macchina | Attivo |
|---|---|---|---|---|
| ST3-AST | ST3 | Secondario | Astucciatrice | SI |
| ST3-FIL | ST3 | Sterile | Filling - Tappatrice | SI |

- **Codice (QR)**: unico, senza spazi, è quello stampato nell'etichetta. Consigliato `LINEA-SIGLA`. **Non cambiarlo** dopo aver attaccato l'etichetta.
- **Linea**: come vuoi vederla nei KPI (ST1 MAR, ST3, SPL1…).
- **Area**: `Sterile` o `Secondario` → decide quale lista di guasti compare alla chiusura.
- **Macchina**: il nome che finisce nella colonna Macchina del foglio Fermi (per le pivot usa sempre gli stessi nomi).
- **Attivo**: `NO` per toglierla dall'app senza perdere lo storico.

Puoi incollare l'elenco direttamente da Excel. Dopo aver aggiunto macchine, da `qr.html` stampa solo le etichette nuove.

> L'elenco iniziale creato da `setup()` mette Riordinatore + Filling su ogni linea con filling e le 6 macchine del secondario su ogni linea con packaging: **va controllato**, in particolare per le linee con macchine diverse (es. ST4, ST5).

### 5.3 Liste dei pulsanti
Foglio **Liste**: colonne **Categoria**, **Tipo intervento**, **Guasto Sterile**, **Guasto Secondario**. Quello che scrivi diventa un pulsante nell'app, nello stesso ordine (metti in alto le voci più usate). Per tenere puliti i KPI evita due voci che significano la stessa cosa.

### 5.4 KPI
Foglio **KPI**, tutto automatico:
- **Per linea**: n° fermi, tempo di fermo, **MTTR** (tempo fermo ÷ n° fermi), n° lotti, **Occurrancy** (fermi ÷ lotti).
- **MTBF**: inserisci nella colonna gialla le **ore di funzionamento** della linea nel periodo; l'MTBF (ore ÷ n° fermi) si calcola da solo.
- **Pareto guasti** con **RFR** (n° di quel guasto ÷ totale fermi).
- **Macchine** ordinate per tempo di fermo.

Contano solo i fermi **chiusi**. I KPI considerano tutto lo storico: per un periodo (es. mese) filtra il foglio Fermi per Data o lavora in Excel con le pivot.

---

## 6. I dati in Excel

- **Copia veloce**: dal Google Sheet **File → Scarica → Microsoft Excel (.xlsx)**.
- **Excel sempre aggiornato** (anche con Microsoft 365):
  1. Google Sheet → **File → Condividi → Pubblica sul web** → scegli il foglio **Fermi** e il formato **Valori separati da virgola (.csv)** → **Pubblica** → copia il link.
  2. Excel → **Dati → Da Web** → incolla il link → **Carica**.
  3. **Dati → Aggiorna tutto** quando vuoi i dati nuovi (o, nelle proprietà della query, aggiornamento automatico ogni X minuti).
  4. Le pivot e la dashboard del vecchio template possono puntare a questa tabella: le colonne A–L sono le stesse.

  ⚠️ Chi ha quel link legge i dati del foglio Fermi (non le foto): non diffonderlo.

---

## 7. Aggiornare l'app o lo script

- **App** (file in questa cartella): ogni modifica unita su GitHub (`main`) è online in un minuto; i telefoni la prendono alla successiva apertura con rete.
- **Script** (`apps-script/Code.gs`): se cambia, va ricopiato nel Google Sheet:
  1. Estensioni → Apps Script → incolla il nuovo `Code.gs` → salva.
  2. **Esegui il deployment → Gestisci deployment** → matita ✏️ → Versione: **Nuova versione** → **Esegui il deployment**.
  3. L'indirizzo `/exec` **resta lo stesso**: non serve toccare `config.js`.
  - **Non rieseguire `setup()`**: non tocca Fermi, Equipment e Liste, ma ricostruisce il foglio KPI e cancella le ore di funzionamento che hai inserito.

---

## 8. Problemi frequenti

| Problema | Causa / soluzione |
|---|---|
| "Non trovo … tra i manutentori abilitati" | Il nome non è nella colonna Manutentori del foglio Liste, o è scritto diverso. Aggiungilo/correggilo; il tecnico riapre l'app. |
| La fotocamera non si apre per lo scanner | Permesso negato: impostazioni del telefono → app/browser → Fotocamera → Consenti. Intanto usa **Fotografa QR** o **Scegli macchina a mano**. |
| "QR non riconosciuto" | L'etichetta è di un'altra app o la macchina non è nel foglio Equipment (o ha Attivo = NO). |
| Pallino rosso fisso con rete funzionante | `config.js` sbagliato (URL o KEY), oppure lo script non è pubblicato con accesso **Chiunque**. Apri nel browser `API_URL?action=ping&key=LA_CHIAVE`: deve rispondere `{"ok":true}`. |
| "Chiave non valida" | La KEY in `config.js` è diversa da quella nelle **Proprietà script** (Apps Script → Impostazioni progetto). |
| Un fermo è stato chiuso con dati sbagliati | Correggilo a mano nel foglio Fermi (non toccare Stato e ID fermo). |
| Fermo rimasto aperto (giallo) da giorni | Qualcuno ha dimenticato di chiuderlo: chiudilo dall'app (scansione QR o **Chiudi** nella home) e correggi poi a mano l'Ora Fine nel foglio se serve. |
| L'app mostra ancora la versione vecchia | Chiudila del tutto e riaprila con rete. |

---

## 9. Sicurezza e limiti

- La **chiave** in `config.js` blocca scritture casuali, ma è leggibile da chi apre il codice della pagina: non diffondere il link fuori dal reparto. Per cambiarla: Apps Script → Impostazioni progetto → **Proprietà script** → `KEY`, poi aggiorna `config.js`.
- L'**accesso** con nome e cognome serve a sapere chi ha registrato cosa, non è una password: un tecnico potrebbe scrivere il nome di un collega.
- Le **foto** restano private nella cartella Drive del tuo account; nel foglio c'è solo il link.
- Il foglio è sul tuo account Google: se cambi account o lasci l'incarico, trasferisci la proprietà del foglio e della cartella foto.
- È uno **strumento di analisi della manutenzione**, non un sistema validato GMP: non sostituisce le registrazioni ufficiali.
