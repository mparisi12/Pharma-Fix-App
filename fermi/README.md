# Pharma-Fix

App di reparto costruita attorno al **QR della macchina**:

- **Fermi macchina** in meno di 30 secondi: si scansiona il QR, si apre il fermo, a fine intervento si chiude con la foto della zona e i dati del template "fermi meccanici".
- **Scheda macchina**: fermi, soluzioni, manuali/schemi, schede formato e anomalie aperte di quella macchina.
- **Soluzioni** (troubleshooting): quelle verificate e quelle nate dai fermi chiusi.
- **Documenti**: manuali, schemi e schede formato collegati alla macchina o alla linea.
- **Anomalie**: segnalazioni con urgenza 1 critico / 2 urgente / 3 programmabile, da prendere in carico e chiudere.

Tutto finisce in un foglio Google che si aggiorna da solo e si può aprire anche in Excel.

Indirizzo dell'app: **https://mparisi12.github.io/Pharma-Fix-App/fermi/**
(Anche il vecchio indirizzo https://mparisi12.github.io/Pharma-Fix-App/ porta qui.)

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
| **Fermi** | Una riga per fermo. Colonne A–L **identiche al template** (Lotto, Data, Ora Inizio, Ora Fine, Manutentore, Macchina, Categoria, Tipo intervento, Tipo guasto, Altro, Durata MTTR, Intervento-Guasto), poi Linea, Turno, Codice pezzo, Foto apertura, Foto chiusura, Stato, ID fermo, Equipment, Aperto da, **Orari** (TEMPO REALE / DICHIARATI), **Registrato alle** (ora vera dell'inserimento), **Telefono**. Le righe **gialle** sono fermi ancora aperti. | L'app. Tu puoi correggere a mano un errore, ma **non toccare le colonne Stato e ID fermo**. |
| **Equipment** | Elenco macchine: codice del QR, linea, area (**Filling** / **Packaging**), nome macchina, attivo. | Tu. |
| **Liste** | I pulsanti uguali per tutti: Categoria, Tipo intervento, **Manutentori abilitati**. | Tu. |
| **Guasti** | I tipi di guasto **per macchina**: una riga per codice QR, i guasti uno per cella verso destra. Righe speciali `TUTTE` (valgono per ogni macchina) e nome linea (es. `ST3`, valgono per tutta la linea). | Tu. |
| **Soluzioni** | Troubleshooting: dove vale (codice macchina, linea o `TUTTE`), tipo guasto, problema, soluzione, codici ricambio, allegato. | L'app (pulsante *Aggiungi soluzione*) e tu. |
| **Documenti** | Manuali, schemi, schede formato: dove vale, tipo, titolo, descrizione, **link**. | Tu (vedi 5.6). |
| **Anomalie** | Segnalazioni: macchina, urgenza (cella colorata), descrizione, cosa serve, foto, stato APERTA / IN CORSO / CHIUSA, chi l'ha gestita, nota. | L'app; tu puoi cambiare stato e note. |
| **Telefoni** | Il telefono di ogni tecnico. Il primo si registra da solo; un telefono nuovo, o il telefono di un collega, resta **DA APPROVARE** (vedi 5.9). | Tu (colonna Stato). |
| **Controlli** | Casi da guardare (telefoni da approvare, stesso telefono per più tecnici, fermi sovrapposti, orari dichiarati inseriti tardi): menu **Pharma-Fix → Controlla registrazioni**. | Lo script. |
| **KPI** | Indicatori separati per Filling e Packaging, con filtri per linea, macchina e periodo: n° fermi, tempo di fermo, MTTR, n° lotti, Occurrancy, MTBF, RFR, Pareto doppio (frequenza e durata) con grafici, fermi per macchina. | Solo le celle gialle (filtri e ore di funzionamento). |

Le foto finiscono nella cartella Drive **Pharma-Fix Foto Fermi**; nel foglio Fermi c'è il link "Foto".

### Indirizzo e codice reparto

- **API_URL**: l'indirizzo `/exec` dello script. Dice all'app *dove* mandare i dati. Sta in `config.js` (pubblico: da solo non serve a niente).
- **Codice reparto**: la parola d'ordine che apre il foglio. È salvata **solo nello script** (Apps Script → Impostazioni progetto → Proprietà script → `KEY`) e la conoscono i tecnici. **Non è scritta da nessuna parte su GitHub.** Ogni tecnico la inserisce una sola volta sul suo telefono, al primo accesso; senza il codice giusto lo script non legge e non scrive niente.

Finché `API_URL` è vuoto l'app gira in **modalità demo**: i dati restano solo sul telefono (scritta "MODALITÀ DEMO" in alto).

---

## 2. I file della cartella

| File | Cosa contiene | Va modificato? |
|---|---|---|
| `index.html` | L'app: accesso, scanner QR, scheda macchina, apertura e chiusura fermo, coda offline. | No |
| `config.js` | Indirizzo dello script (`API_URL`). Il codice reparto **non** va qui. | **Sì, una volta**, dopo l'installazione |
| `defaults.js` | Liste e macchine "di partenza", usate finché il foglio non risponde. Non contiene i nomi dei tecnici (stanno solo nel foglio). Con il foglio collegato vale quello che c'è nel foglio. | No (le liste si cambiano nel foglio) |
| `qr.html` | Pagina per stampare le etichette QR, una per macchina, prese dal foglio Equipment. | No |
| `manifest.webmanifest` | Nome e icona quando si installa l'app sul telefono. | No |
| `sw.js` | Permette all'app di aprirsi anche senza rete. | No |
| `icon-192.png`, `icon-512.png` | Icone dell'app. | No |
| `apps-script/Code.gs` | Il programma da incollare nel Google Sheet (Estensioni → Apps Script). Crea i fogli, riceve i dati, salva le foto, controlla i manutentori. | No, si **copia** nel Google Sheet |
| `apps-script/appsscript.json` | Impostazioni dello script (fuso orario Europe/Rome, accesso app web). | No, si **copia** nel Google Sheet |
| `README.md` | Questo file. | — |

Nella cartella principale del repository, `index.html` è la vecchia Pharma-Fix: ora rimanda sempre qui (i suoi dati sono stati importati, vedi 5.7).

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
7. **Scegli il codice reparto**: icona ⚙️ **Impostazioni progetto** → in fondo **Proprietà script** → riga **KEY** → **Modifica** → scrivi il codice che vuoi (almeno 6 caratteri, meglio lettere e numeri) → **Salva**. È il codice che darai ai tecnici.
8. Torna al Google Sheet: ora ci sono i fogli **Fermi, Equipment, Liste, KPI**.

### 3.2 Pubblicare lo script come App web

1. In Apps Script: **Esegui il deployment → Nuovo deployment**.
2. Icona ingranaggio accanto a "Seleziona tipo" → **App web**.
3. Esegui come: **Me** · Chi ha accesso: **Chiunque**.
4. **Esegui il deployment** → copia l'**URL dell'app web** (finisce con `/exec`).

### 3.3 Collegare l'app al foglio

1. Su GitHub apri `fermi/config.js` → matita ✏️.
2. Incolla l'indirizzo:
   ```js
   window.PF_CONFIG = {
     API_URL: 'https://script.google.com/macros/s/AKfy……/exec'
   };
   ```
3. **Commit changes**. Dopo un minuto l'app è collegata: in alto compare "Sincronizzato hh:mm" al posto di "MODALITÀ DEMO".

### 3.4 Preparare liste, macchine ed etichette

1. Foglio **Liste** → colonna **Manutentori**: controlla che ci siano tutti i tecnici, scritti `COGNOME Nome` (vedi [5.1](#51-tecnici-abilitati)).
2. Foglio **Equipment**: controlla e completa le macchine (vedi [5.2](#52-lista-macchine-foglio-equipment)).
3. Nell'app tocca il tuo nome in alto → **Stampa etichette QR** (oppure dal PC apri **https://mparisi12.github.io/Pharma-Fix-App/fermi/qr.html**, scrivi il codice reparto → **Carica macchine**) → scegli la linea → **Stampa** al 100% su carta adesiva o da plastificare → attacca ogni etichetta sulla sua macchina, lato operatore.

### 3.5 Installare l'app sui telefoni

1. Apri **https://mparisi12.github.io/Pharma-Fix-App/fermi/** con **Chrome** (Android) o **Safari** (iPhone).
2. Chrome: menu ⋮ → **Installa app** / **Aggiungi a schermata Home**. Safari: **Condividi → Aggiungi alla schermata Home**.
3. Alla prima scansione consenti l'uso della **fotocamera**.

---

## 4. Uso quotidiano — per i tecnici

### Accesso
All'apertura l'app chiede **Nome** e **Cognome**. Entra solo chi è tra i manutentori abilitati (maiuscole, accenti e ordine non contano: "mario rossi" va bene per `ROSSI Mario`). Ogni fermo viene registrato a nome di chi è entrato.
La **prima volta** su un telefono chiede anche il **Codice reparto** (serve la rete); poi non lo chiede più, finché il responsabile non lo cambia.
**Ogni nome è legato al telefono personale del tecnico**: il primo telefono con cui entri diventa il tuo. Con il nome di un collega, o da un telefono nuovo, l'app risponde *"Telefono non autorizzato"* finché il responsabile non approva.
L'accesso dura un turno (9 ore). Se passi il telefono a un collega: tocca il tuo nome in alto → **Esci**.

### Aprire un fermo (~10 secondi)
1. **Scansiona QR** e inquadra l'etichetta sulla macchina. (In alternativa: fotocamera normale del telefono sul QR, oppure **Fotografa QR**, oppure **Scegli macchina a mano**.) Si apre la **pagina della macchina**, con lo storico degli interventi sotto.
2. Tocca **Apri fermo** (il pulsante rosso in alto): controlla il **lotto** (è già scritto quello dell'ultimo fermo sulla linea) e, se il fermo è iniziato prima, scegli −5 / −10 / −15 / −30 min. **Il lotto è obbligatorio** (solo cifre): senza, il fermo non si apre.
3. **Apri fermo**. Il cronometro parte e il fermo compare nella home tra i **Fermi aperti**.

### Macchine FILLING (sterile): registrare il fermo all'uscita
In sterile il telefono non entra, quindi il fermo si registra **dopo**, uscendo dal reparto:
1. Scansiona o **fotografa il QR** della macchina (nella scheda macchina il pulsante si chiama **Registra fermo**).
2. Scrivi **Inizio** e **Fine** del fermo, **gli stessi orari del batch record**.
3. Compila come una chiusura: foto (facoltativa), categoria, tipo intervento, guasto, descrizione, lotto, pezzo → **Registra fermo**.

Si possono registrare solo fermi **del turno in corso** (T1 dalle 6:00, T2 dalle 14:00, T3 dalle 22:00). Nel foglio la riga ha **Orari = DICHIARATI** e l'ora vera di inserimento in **Registrato alle**.

### Durante il fermo: "Cosa provare"
Con un fermo aperto, nella pagina della macchina (sotto **Chiudi fermo**) c'è **Cosa provare**: i tipi di guasto della macchina dal più frequente (il numero = quante volte negli ultimi 90 giorni). Tocca il guasto che vedi: compaiono le **soluzioni verificate** da provare. Il guasto scelto qui è già selezionato quando chiudi il fermo.

### Chiudere un fermo (~20–30 secondi)
1. **Scansiona di nuovo il QR** della stessa macchina → pulsante verde **Chiudi fermo** (oppure, nella home, **Chiudi** sul fermo).
2. **Foto zona di intervento** → scatta → **tocca la foto per cerchiare** il punto in cui hai lavorato (tocca il cerchio per toglierlo).
3. Scegli **Categoria**, **Tipo intervento**, **Tipo guasto** (uno solo, come nel template). Sotto il guasto compaiono le **soluzioni verificate** più adatte per quella macchina (*Prova prima*): leggile prima di cercare altrove.
4. **Altro (descrizione breve)**: obbligatoria se hai scelto ALTRO. Per chiusure astucci scrivi zona, lato macchina/OP, n° nonio.
5. **Lotto**: è già scritto quello dell'apertura; correggilo se nel frattempo è cambiato. **Obbligatorio**: senza, il fermo non si chiude.
6. **Codice pezzo sostituito** se hai cambiato un ricambio.
7. **Chiudi fermo**.

Fermo aperto per sbaglio? **Chiudi** → in fondo **Annulla il fermo** (tocca due volte).

### Scheda macchina, soluzioni, documenti, anomalie
- **Pagina della macchina** (si apre dopo il QR o scegliendo la macchina): in alto il pulsante **Apri fermo** (Packaging), **Registra fermo** (Filling) o **Chiudi fermo** verde con il cronometro se un fermo è in corso; poi Soluzioni, Manuali, Schede, Anomalia e i numeri degli ultimi 30 giorni. **Scorrendo**: gli interventi degli ultimi 90 giorni dal più recente, divisi per giorno, con orari, durata (arancione oltre l'ora), guasto, descrizione, tecnico e lotto; 20 alla volta (**Mostra altri interventi**). Tocca un intervento per il dettaglio completo e la foto.
- **Soluzioni**: cerca per parola (pinze, ventose…) o per tipo guasto. Compaiono solo le soluzioni **verificate** del foglio Soluzioni (lo storico dei fermi è nella pagina della macchina). **+ Aggiungi soluzione**: tipo guasto, problema, soluzione, ricambi, vale per questa macchina o tutta la linea, foto o PDF facoltativo.
- **Documenti**: due schede, *Manuali e schemi* e *Schede formato*; filtro *Questa macchina* / *Tutta la linea*. Si aprono con un tocco.
- **Segnala anomalia**: macchina (dal QR o *Cambia*), **urgenza 1 / 2 / 3**, cosa hai visto, foto con cerchio, cosa serve → **Invia segnalazione**.
- **Anomalie** (barra in basso): *Da gestire* e *Chiuse* (ultimi 30 giorni). Tocca una anomalia → **Prendo in carico** oppure **Chiudi anomalia** con una nota.

Anche queste funzioni lavorano senza rete: partono appena torna la connessione.

### Il pallino in alto
| Pallino | Significa |
|---|---|
| 🟢 Sincronizzato | Tutto è arrivato nel foglio. |
| 🟠 Invio in corso | Ci sono registrazioni in coda, partono da sole. |
| 🔴 Telefono non autorizzato | Il nome non è legato a questo telefono: le registrazioni restano in coda e partono appena il responsabile approva il telefono nel foglio Telefoni. |
| 🔴 Non connesso | Niente rete: **continua pure a lavorare**, i fermi restano sul telefono e partono appena torna la connessione. **Non disinstallare l'app** finché ci sono dati in coda (le foto in coda stanno nell'archivio del telefono, non si perdono chiudendo l'app). |
| ⚪ Modalità demo | L'app non è ancora collegata al foglio. |

---

## 5. Gestione — per il responsabile

Tutto si gestisce dal **Google Sheet**: l'app rilegge liste e macchine a ogni apertura, non serve toccare GitHub.

### 5.1 Tecnici abilitati
Foglio **Liste**, colonna **Manutentori**: un nome per cella, formato `COGNOME Nome` (es. `ROSSI Mario`).
- **Nuovo tecnico** → aggiungi una cella. Può entrare subito.
- **Tecnico che lascia** → cancella la cella. Lo storico resta; non potrà più registrare (lo blocca anche lo script).
- Il nome scritto qui è quello che finisce nella colonna **Manutentore** del foglio Fermi.

### 5.2 Lista macchine (foglio Equipment)

Una riga per ogni macchina che deve avere il suo QR:

| Codice (QR) | Linea | Area | Macchina | Attivo |
|---|---|---|---|---|
| ST3-04 | ST3 | Packaging | ASTUCCIATRICE ROMACO PC4400 | SI |
| ST3-02 | ST3 | Filling | RIEMPITRICE + TAPPATRICE FLEXFILL F940 | SI |

- **Codice (QR)**: unico, senza spazi, è quello stampato nell'etichetta. Consigliato `LINEA-SIGLA`. **Non cambiarlo** dopo aver attaccato l'etichetta.
- **Linea**: come vuoi vederla nei KPI (ST1 MAR, ST3, SPL1…).
- **Area**: `Filling` (in sterile: il fermo si registra all'uscita con gli orari del batch record) o `Packaging` (fermo aperto e chiuso in tempo reale).
- **Macchina**: il nome che finisce nella colonna Macchina del foglio Fermi (per le pivot usa sempre gli stessi nomi).
- **Attivo**: `NO` per toglierla dall'app senza perdere lo storico.

Puoi incollare l'elenco direttamente da Excel. **Non rinumerare** le macchine che hanno già l'etichetta: aggiungi quelle nuove con un codice libero.

**Stampare le etichette QR** (si trovano anche dall'app):
- **Tutte o alcune**: nell'app tocca il tuo nome in alto → **Stampa etichette QR**. Scegli la linea; **tocca le etichette** che ti servono (es. le macchine nuove) e premi **Stampa N selezionate**. Senza selezione si stampano tutte.
- **Una sola**: dalla scheda della macchina, in fondo → **Stampa etichetta QR di questa macchina**.

Stampa al 100% (senza adattare alla pagina) su carta adesiva o da plastificare.

> L'elenco iniziale creato da `setup()` mette Riordinatore + Filling su ogni linea con filling e le 6 macchine del secondario su ogni linea con packaging: **va controllato**, in particolare per le linee con macchine diverse (es. ST4, ST5).

### 5.3 Liste dei pulsanti
Foglio **Liste**: colonne **Categoria** e **Tipo intervento**, uguali per tutte le macchine. Quello che scrivi diventa un pulsante nell'app, nello stesso ordine (metti in alto le voci più usate). Le vecchie colonne Guasto Sterile / Guasto Secondario non servono più quando il foglio Guasti è compilato: puoi cancellarle.

### 5.4 Tipi di guasto per macchina (foglio Guasti)

| A: Codice (QR) / Linea / TUTTE | B: Macchina (promemoria) | C, D, E… : tipi di guasto |
|---|---|---|
| TUTTE | vale per tutte le macchine | FORMATO · SENSORE · USURA · SPIKE CORRENTE · ARIA COMPRESSA · ALTRO |
| ST3-04 | ST3 · ASTUCCIATRICE ROMACO PC4400 | MAG. ISTRUZIONI · CHIUSURE AST · MAG. ASTUCCI · … |
| ST6 | (riga facoltativa: vale per tutte le macchine di ST6) | … |

Alla chiusura di un fermo l'app mostra: i guasti della macchina → quelli della sua linea → quelli di `TUTTE`, con **ALTRO sempre in fondo**.
- Metti in alto i guasti più frequenti: compaiono per primi.
- Scrivi sempre allo stesso modo (stesse maiuscole, stessa sigla): è quello che finisce nel Pareto.
- Non mettere tra i guasti RIAVVIO, PULIZIA, SOSTITUZIONE, REGOLAZIONE: sono **tipi di intervento**, non guasti.
- Macchina nuova: aggiungi la riga anche qui. Macchina senza riga o con la riga vuota: compaiono solo i guasti di `TUTTE`.
- Per costruire le liste delle altre linee guarda le descrizioni dei fermi chiusi con **ALTRO**: le voci che si ripetono diventano un guasto.

### 5.5 KPI
Foglio **KPI**. In alto le celle gialle sono i **filtri**:

| Filtro | Cosa scegli |
|---|---|
| **Area** | **PACKAGING** oppure **FILLING**: sono sempre separati, non si sommano mai. Linee e macchine nei menu sotto sono solo quelle dell'area scelta |
| **Linea** | TUTTE oppure una linea |
| **Macchina** | TUTTE oppure una macchina della linea scelta (se cambi linea, rimetti TUTTE) |
| **Periodo** | Ultimi 7 / 30 / 90 giorni, Mese corrente, Mese scorso, Anno corrente, Tutto, Personalizzato (date nelle celle gialle a destra) |

Tutto il foglio si ricalcola sulla scelta. L'area di ogni fermo si prende dal foglio Equipment (dal codice macchina, o da linea + nome macchina per i fermi importati): se una macchina non si trova, il fermo viene escluso e sotto gli indicatori compare quanti sono.
- **Indicatori**: n° fermi, tempo di fermo, **MTTR** (tempo ÷ n° fermi), n° lotti, **Occurrancy** (fermi ÷ lotti), **guasto n°1 con il suo RFR** (n° fermi di quel guasto ÷ totale fermi).
- **MTBF**: scrivi nella cella gialla le **ore di funzionamento** della linea/macchina nel periodo scelto (dai dati di produzione); l'MTBF (ore ÷ n° fermi) si calcola da solo.
- **Pareto doppio**, due grafici affiancati su una scala unica 0–100% (barre = quota di ogni guasto, linea rossa = % cumulata; i guasti a sinistra dell'80% sono quelli su cui intervenire). I numeri assoluti (n° fermi, ore) sono nelle tabelle sotto:
  - **per frequenza**: quante volte succede ogni guasto, con % cumulata e **RFR**;
  - **per durata**: quante ore di fermo causa ogni guasto. Un guasto raro ma lungo (es. elettrico) qui sale in cima.
  - Nei grafici ci sono al massimo 20 barre: se i tipi di guasto sono di più, gli ultimi sono raggruppati in **ALTRI**, così la linea arriva sempre al 100%. Le tabelle sotto li elencano tutti.
- **Fermi per macchina** con n° fermi, tempo totale, MTTR e **guasto n°1 · RFR** di quella macchina.

Contano solo i fermi **chiusi**. I valori li calcola lo script: si aggiornano quando cambi una cella gialla, quando si chiude un fermo, oppure dal menu **Pharma-Fix → Aggiorna KPI** (in alto nel foglio, accanto a Guida). Se il foglio KPI si rovina: **Pharma-Fix → Ricrea foglio KPI**. Dopo un aggiornamento dello script che cambia la scheda KPI, il foglio si ricrea da solo al primo **Aggiorna KPI**, tenendo filtri e ore di funzionamento.

### 5.6 Documenti (foglio Documenti)
Una riga per documento:

| Colonna | Cosa scrivere |
|---|---|
| **Dove** | Codice QR della macchina (`ST3-04`), oppure la linea (`ST3`) per tutta la linea, oppure `TUTTE` |
| **Tipo** | `MANUALE`, `SCHEMA`, `SCHEDA FORMATO` o `ALTRO` (le *Schede formato* vanno nella seconda scheda dell'app) |
| **Titolo** / **Descrizione** | Quello che vede il tecnico (es. *Schema elettrico* · *rev. 3*) |
| **Link** | Link al file su Drive/SharePoint: in Drive tasto destro → Condividi → *Chiunque nell'organizzazione con il link* → Copia link |

Le colonne ID e Data si possono lasciare vuote. Il file resta dove è: l'app mostra solo il link, quindi lo apre solo chi ha già i permessi.

### 5.7 Portare i dati della vecchia Pharma-Fix
**Già preparato:** il file **Import_PharmaFix_pronto.xlsx** contiene il vecchio database convertito: 50 soluzioni, 11 documenti, 1 anomalia, ognuno collegato al codice QR della macchina nuova. Si carica una volta sola:
1. Nel Google Sheet *Pharma-Fix Fermi*: **File → Importa → Carica** → scegli il file → **Inserisci nuovi fogli**.
2. Dalla scheda importata *Soluzioni* copia dalla riga 2 in giù (colonne A–K) e incolla in **Soluzioni**, cella A2.
3. Allo stesso modo *Documenti* (A–H) → **Documenti** e *Anomalie* (A–N) → **Anomalie**.
4. Elimina le schede importate (compresa *Mappa*, che serviva solo a controllare gli abbinamenti).

> Non eseguire anche **importaPharmaFix** sugli stessi dati: le righe raddoppierebbero. (importaPharmaFix resta per un eventuale altro file della vecchia app: incollalo in una scheda **Import_PharmaFix**, esegui una volta, controlla i codici in **Mappa_import**, esegui di nuovo.)

**Allegati della vecchia app:** i link del vecchio database puntano ai file nella tua cartella Drive. **Non cancellare quei file e non togliere la condivisione con link**, altrimenti nell'app non si aprono più. Spostarli in un'altra cartella invece va bene: il link non cambia.

### 5.8 Dove finiscono i file nuovi
| Cosa | Chi lo carica | Dove finisce |
|---|---|---|
| Foto dei fermi, foto anomalie, allegati delle soluzioni | Il tecnico dall'app | Da solo nella cartella Drive **Pharma-Fix Foto Fermi** + link nel foglio |
| Manuali, schemi, schede formato | Tu | Carichi il file in Drive, lo condividi con link e aggiungi una riga nel foglio **Documenti** (vedi 5.6) |

### 5.9 Telefoni dei tecnici e controlli
Foglio **Telefoni**: una riga per ogni coppia tecnico + telefono, con modello, primo accesso, ultimo uso e **Stato**.
- **Primo telefono** di un tecnico: si registra da solo con Stato **OK**.
- **Telefono nuovo** (cambio telefono) o **telefono di un collega**: riga gialla con Stato **DA APPROVARE**; il tecnico non può registrare finché non scrivi **OK**. Se è un tentativo di usare il nome di un altro, scrivi **BLOCCATO**.
- **Telefono perso o cambiato**: metti **BLOCCATO** sulla riga vecchia.

Menu **Pharma-Fix → Controlla registrazioni**: crea il foglio **Controlli** con i casi da guardare — telefoni da approvare, stesso telefono usato da più tecnici, fermi sovrapposti dello stesso tecnico, orari dichiarati inseriti più di 4 ore dopo la fine. Fallo ogni tanto (es. una volta a settimana).

### 5.10 Sveglia dello script
Google "addormenta" lo script quando nessuno lo usa per un po': al risveglio la prima risposta può metterci 20–40 secondi (l'app scrive *"Il foglio risponde lentamente"* e intanto si lavora lo stesso: i dati restano in coda).
Per evitarlo: menu **Pharma-Fix → Attiva sveglia dello script** (una volta sola; Google chiede di autorizzare *"collegarsi a un servizio esterno"* e *"eseguire in tua assenza"*: consenti). Da lì lo script si interroga da solo ogni 10 minuti da lunedì 5:00 a sabato 6:00 (sempre, se nel foglio Avvisi c'è una linea a ciclo continuo, vedi 5.12). Per toglierla: **Disattiva sveglia dello script**.

### 5.11 Avvisi via mail
Foglio **Avvisi** (lo crea `setup`): nelle celle gialle
| Impostazione | Esempio |
|---|---|
| **Email destinatari** (separate da virgola) | `tuo.nome@azienda.it, capoturno@azienda.it` |
| Anomalia critica (urgenza 1) | `SI`: mail appena viene segnalata |
| Fermo aperto da più di (ore) | `2`: mail se un fermo resta aperto oltre 2 ore (fermo lungo o dimenticato). `0` = mai |
| Stesso guasto sulla stessa macchina: quante volte | `3`: mail quando lo stesso guasto si ripete 3 volte (e poi 6, 9…). `0` = mai. ALTRO non conta |
| …nell'arco di quanti giorni | `7` |

| Telegram: token del bot | vedi sotto |
| Telegram: ID chat o gruppo | lo scrive lo script |

**Avvisi su Telegram** (consigliato: arrivano subito, anche in un gruppo con i capiturno, e non passano dai filtri della posta aziendale):
1. In Telegram cerca **@BotFather** → scrivi `/newbot` → dai un nome (es. *Pharma-Fix Manutenzione*) e un nome utente che finisce con `bot` (es. `pharmafix_reparto_bot`). BotFather risponde con il **token** (una riga tipo `123456789:AA…`): copialo nella cella **B7** del foglio Avvisi. Il token è come una password: non condividerlo (se esce, in BotFather `/revoke`).
2. Crea un **gruppo** Telegram (es. *Manutenzione Pharma-Fix*) con te e i capiturno, **aggiungi il bot** al gruppo e scrivi `/start` nel gruppo. (Solo per te: apri la chat del bot e premi **Avvia**.)
3. Menu **Pharma-Fix → Telegram: trova ID chat**: scrive da solo l'ID nella cella **B8**.
4. **Pharma-Fix → Prova avvisi**: il messaggio arriva nel gruppo.

Mail e Telegram funzionano anche insieme: si usa quello che è compilato.

Sotto, il **registro** degli avvisi inviati (con il canale usato). Prova: menu **Pharma-Fix → Prova avvisi (mail di test)**. Le mail partono dal tuo account Google, come una mail scritta da te: solo testo, oggetto che inizia con **"Pharma-Fix ·"** (la prima volta Google chiede di autorizzare *"inviare email a tuo nome"*). Se all'indirizzo aziendale non arrivano, controlla la quarantena/spam o usa Telegram.
Il controllo dei **fermi aperti da troppo tempo** gira insieme alla **sveglia** (5.10): serve che sia attiva.

### 5.12 Report automatici e backup
Nel foglio **Avvisi**, colonne **E–F** (celle gialle), si accendono e spengono con SI / NO:
- **Riepilogo settimanale**: ogni **lunedì alle 7**, mail + Telegram con la settimana precedente, **divisa in due blocchi, PACKAGING e FILLING** (mai sommati): per ognuno n° fermi e tempo di fermo (con confronto sulla settimana prima), MTTR, fermi per linea, i 5 guasti più frequenti, i 3 fermi più lunghi. In fondo, una parte generale: anomalie da gestire / nuove / chiuse e fermi ancora aperti (con [F] o [P] davanti). Se un fermo ha una macchina che non si trova in Equipment, compare in un blocco a parte "AREA NON TROVATA".
- **Passaggio di turno**: a **fine di ogni turno** (6, 14, 22, solo da lunedì 14:00 a sabato 6:00), su Telegram (e per mail se lo attivi): fermi del turno con orari, durata, guasto, tecnico e descrizione, fermi ancora aperti, anomalie nuove/chiuse/da gestire. È **costruito dai dati registrati nell'app**, non scritto dai tecnici: compare tutto, anche "(senza descrizione)" e i fermi lasciati aperti.
- **Linee a ciclo continuo** (es. `ST3`, separate da virgola): per queste linee il passaggio di turno arriva anche nel weekend (solo con i loro dati) e la sveglia / l'avviso dei fermi lunghi restano attivi 7 giorni su 7. Quando la linea torna sui 3 turni, svuota la cella **F7** del foglio Avvisi: tutto torna lunedì–sabato da solo, senza toccare lo script.
- **Backup**: ogni **domenica sera** una copia Excel del foglio nella cartella Drive **Pharma-Fix Backup**; si tengono le ultime N (8 = due mesi).

Per partire: menu **Pharma-Fix → Attiva report e backup automatici** (una volta sola; Google può chiedere di nuovo le autorizzazioni). Prove subito: **Prova riepilogo settimanale**, **Prova passaggio di turno**, **Fai un backup adesso**.

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
  - Se l'aggiornamento cambia il foglio KPI, esegui **aggiornaKpi** (menu delle funzioni → aggiornaKpi → Esegui): ricrea il foglio KPI; va reinserito solo il valore delle ore di funzionamento.
  - `setup()` si può rieseguire senza rischi: crea solo i fogli che mancano (es. Guasti, Soluzioni, Documenti, Anomalie, Telefoni), aggiunge le colonne nuove di Fermi e trasforma le aree Sterile/Secondario in Filling/Packaging e risistema la formula della durata; non tocca Fermi, Equipment, Liste, Guasti e KPI già esistenti.

---

## 8. Problemi frequenti

| Problema | Causa / soluzione |
|---|---|
| "Non trovo … tra i manutentori abilitati" | Il nome non è nella colonna Manutentori del foglio Liste, o è scritto diverso. Aggiungilo/correggilo; il tecnico riapre l'app. |
| La fotocamera non si apre per lo scanner | Permesso negato: impostazioni del telefono → app/browser → Fotocamera → Consenti. Intanto usa **Fotografa QR** o **Scegli macchina a mano**. |
| "QR non riconosciuto" | L'etichetta è di un'altra app o la macchina non è nel foglio Equipment (o ha Attivo = NO). |
| Pallino rosso fisso con rete funzionante | `API_URL` in `config.js` sbagliato, oppure lo script non è pubblicato con accesso **Chiunque**. Apri nel browser `API_URL?action=ping&key=CODICE_REPARTO`: deve rispondere `{"ok":true}`. |
| "Codice reparto errato" / "Il codice reparto è cambiato" | Il codice inserito non corrisponde a **KEY** nelle Proprietà script. Dopo averlo cambiato, ogni tecnico lo reinserisce una volta. |
| Un fermo è stato chiuso con dati sbagliati | Correggilo a mano nel foglio Fermi (non toccare Stato e ID fermo). |
| Fermo rimasto aperto (giallo) da giorni | Qualcuno ha dimenticato di chiuderlo: chiudilo dall'app (scansione QR o **Chiudi** nella home) e correggi poi a mano l'Ora Fine nel foglio se serve. |
| `#ERROR!` nelle colonne Durata, Intervento-Guasto, Foto o nel foglio KPI | Formule scritte con il separatore sbagliato (fogli in italiano usano `;`). In Apps Script esegui **riparaFormule**: riscrive le formule e ricrea il foglio KPI. |
| Un documento non si apre ("Richiedi l'accesso") | Il file in Drive non è condiviso: tasto destro → Condividi → *Chiunque abbia il link* (o *chiunque nell'organizzazione*). |
| Un documento non compare sulla macchina | Nel foglio Documenti la colonna **Dove** non contiene il codice della macchina né la sua linea, oppure è scritta diversa (es. `ST3 04` invece di `ST3-04`). |
| "Il foglio risponde lentamente" | Lo script si stava svegliando: i dati arrivano da soli. Se succede spesso: **Pharma-Fix → Attiva sveglia dello script** (5.10). |
| L'app mostra ancora la versione vecchia | Chiudila del tutto e riaprila con rete. |

---

## 9. Sicurezza e limiti

- Il **repository è pubblico**: chiunque vede il codice e l'indirizzo dello script, ma **senza il codice reparto non può leggere né scrivere** nel foglio. Il codice non è su GitHub: sta nelle Proprietà script e sui telefoni dei tecnici.
- **Cambia il codice reparto** quando un tecnico lascia l'azienda o se pensi che sia uscito dal reparto: Apps Script → Impostazioni progetto → **Proprietà script** → `KEY` → Modifica → Salva. Effetto immediato, senza nuove versioni; i tecnici lo reinseriscono al prossimo uso.
- Nel repository non ci sono nomi di persone: l'elenco dei tecnici sta solo nel foglio Liste.
- L'**accesso** con nome e cognome è legato al **telefono personale** del tecnico (foglio Telefoni) e protetto dal blocco schermo del telefono. Non è una firma elettronica: chi conosce bene il sistema potrebbe aggirarlo, ma ogni registrazione lascia traccia (telefono, ora vera di inserimento) e i casi sospetti finiscono nel foglio Controlli.
- Le **foto** restano private nella cartella Drive del tuo account; nel foglio c'è solo il link.
- Il foglio è sul tuo account Google: se cambi account o lasci l'incarico, trasferisci la proprietà del foglio e della cartella foto.
- È uno **strumento di analisi della manutenzione**, non un sistema validato GMP: non sostituisce le registrazioni ufficiali.
