# Pharma-Fix — Guida per il coordinatore

Il coordinatore usa l'app come i tecnici (vedi la *Guida per il tecnico*): può ricevere compiti e ha le sue scadenze di vestizione e convalide. In più **tiene d'occhio il reparto**: riceve gli avvisi, segue le anomalie, controlla che i fermi siano registrati bene. Gli strumenti sono due: il **gruppo Telegram** e la sezione **Anomalie** dell'app.

## 1. Cosa ti arriva su Telegram e cosa fare

| Messaggio | Quando arriva | Cosa fare |
|---|---|---|
| 🔴 **Anomalia CRITICA** | Appena un tecnico segnala un'anomalia con urgenza 1 | Valuta subito: serve fermare la linea, chiamare un fornitore, ordinare un ricambio? Prendila in carico dall'app |
| ⏱️ **Fermo APERTO DA…** | Un fermo resta aperto oltre le ore impostate (es. 2) | È un fermo lungo o qualcuno si è dimenticato di chiuderlo. Chiama il tecnico: se ha finito, lo chiude dall'app |
| 🔁 **Guasto RIPETUTO** | Lo stesso guasto sulla stessa macchina si ripete (es. 3 volte in 7 giorni) | Non è più un caso: serve una soluzione vera (regolazione, ricambio, preventiva). Parlane con il responsabile |
| 🔔 **Passaggio turno** | A fine di ogni turno (6, 14, 22) | Leggilo con chi entra: ore di fermo per reparto (linea), divise in Filling e Packaging (senza totali tra linee diverse), i 3 fermi più lunghi, quante anomalie nuove e quante da gestire (il dettaglio è nell'app). È scritto **dai dati registrati**: in **Da sistemare** trovi i fermi lasciati aperti, quelli senza descrizione e quelli di Packaging **registrati dopo** (non aperti col QR), con il nome; chiedi al tecnico di completare |
| 🔔 **Riepilogo settimanale** | Lunedì alle 7 | Settimana separata **Packaging** e **Filling**: fermi, ore e confronto con la settimana prima, MTTR, ore di tutte le linee, guasti che si ripetono, fermi più lunghi, anomalie e compiti |

Su ST3, finché è a ciclo continuo, il passaggio di turno arriva anche sabato e domenica (solo con i dati di ST3).

## 2. Anomalie: prenderle in carico e chiuderle

Dall'app, barra in basso → **Anomalie**:
- **Da gestire**: ordinate per urgenza (1 critico, 2 urgente, 3 programmabile).
- Tocca un'anomalia → **Prendo in carico** quando qualcuno ci sta lavorando (anche tu per conto di un tecnico).
- A lavoro finito → **Chiudi anomalia** con una nota breve su cosa è stato fatto.

Obiettivo: **nessuna anomalia critica aperta a fine turno** senza che qualcuno l'abbia presa in carico.

## 3. Fermi: controllare che siano registrati bene

- **Fermi aperti**: li vedi nella home dell'app. A fine turno non ce ne devono essere di dimenticati.
- **Filling**: il fermo si apre prima di entrare in sterile (così lo vedi tra i fermi aperti) e si chiude all'uscita con gli orari del batch record. Se nel batch record c'è un fermo che non compare nel passaggio di turno, va registrato (al massimo entro 2 ore dal cambio turno).
- **Descrizioni**: una descrizione corta ma vera serve a tutti (e alle soluzioni future). I fermi senza descrizione compaiono nel passaggio di turno, in **Da sistemare**, con il nome di chi li ha registrati: sono quelli da far completare.

## 3b. Fermi da rivedere

Nella home, **"N fermi da rivedere"**: non tutti i fermi, solo quelli degli ultimi 30 giorni in cui qualcosa non torna:
- guasto **ALTRO** (con la voce proposta dal tecnico);
- la descrizione dice "sostituito" ma l'intervento non è SOSTITUZIONE, o il contrario (regolato/reset/pulito ma SOSTITUZIONE);
- "reset/sblocco" registrato come REGOLAZIONE (forse è RIPRISTINO);
- un guasto mai visto prima su quella macchina;
- una durata più di 3 volte quella solita per quel guasto.

Per ogni fermo: **Va bene** (è giusto così) oppure **Correggi** (lo sistemi tu: guasto, intervento, causa, descrizione…). Il tecnico vede la correzione nella sua home e impara la scelta giusta. In alto trovi anche le **voci proposte dai tecnici** con ALTRO: quelle che tornano vanno aggiunte al foglio Guasti (dillo al responsabile).

## 4. Soluzioni: sei tu ad approvarle

Quando un tecnico risolve qualcosa di nuovo, chiedigli di aggiungerlo in **Soluzioni → + Aggiungi soluzione**. Le soluzioni nuove **non le vede nessuno finché non le approvi tu**: solo dopo compaiono a tutti in **Soluzioni**, **Cosa provare** e **Prova prima**.

- Ti arriva un messaggio Telegram **SOLUZIONE DA APPROVARE**, e nella home dell'app compare *"N soluzioni da approvare"*.
- Apri **Soluzioni**: in alto c'è **Da approvare**. Leggi problema e soluzione:
  - **Approva** se è giusta, chiara e ripetibile da un collega;
  - **Scarta** (tocca due volte) se è sbagliata, troppo vaga o un doppione. Se è buona ma scritta male, scartala e chiedi al tecnico di riscriverla.
- Le soluzioni che aggiungi tu sono già approvate.

Se una soluzione già approvata è sbagliata o superata, segnalalo al responsabile, che la corregge nel foglio.

## 5. Compiti: assegnarli e seguirli

Tocca il **tuo nome** in alto → **Compiti del reparto** → **+ Assegna un compito**:
- **A chi**: un tecnico o un coordinatore, oppure **Turno T1 / T2 / T3** (lo vede chi è in turno; il primo che tocca *Lo prendo io* ne diventa responsabile).
- **Titolo**, **cosa fare**, **macchina** (facoltativa), **priorità** (normale / urgente), **entro il**.

Il tecnico vede il numero rosso sul suo nome e, al primo ingresso nell'app, la finestra **Hai un nuovo compito!**. In **Compiti del reparto** trovi gli aperti (in rosso quelli in ritardo), i fatti, il riepilogo **per persona** e lo storico (tocca un nome per vedere solo i suoi).
- **Annullare** un compito che non serve più o **riaprirne** uno fatto male: dal dettaglio, con il **motivo** (obbligatorio, resta scritto).
- Tutto resta nel foglio **Compiti**: chi l'ha assegnato, preso in carico e fatto, quando, con la nota. Nulla si cancella.
- I compiti non compaiono nel passaggio di turno: il dettaglio (anche chi è in ritardo) è nel foglio Compiti e in *Compiti del reparto*; il lunedì il riepilogo ne riporta i numeri.

Per i tecnici, **Stampa etichette QR** non c'è più nel menu del nome: resta solo ai coordinatori.

### Programma produzione
Dal menu del nome, **Programma produzione**: tutte le linee oggi, domani e dopodomani, cambi formato in giallo. Il passaggio di turno riporta i cambi formato del giorno dopo (*📦 Domani cambi formato: …*): è il momento di organizzare chi li segue.

### Turni
**Chi c'è** (menu del nome) mostra chi è in turno oggi, domani e dopodomani e chi è assente. Il passaggio di turno inizia con *In turno: …*. I compiti assegnati a un turno li vede chi quel giorno è davvero in quel turno (foglio Turni, aggiornato dal responsabile ogni mese).

### Scadenze del reparto (vestizioni e convalide GMP)
Nel menu del tuo nome, **Scadenze del reparto**: una riga per persona con la vestizione sterile e le convalide GMP per categoria (G, H…), i più urgenti in alto. Filtri **Da rinnovare** e **Scadute**; tocca un nome per il dettaglio.
- **MANCA** = la persona lavora in sterile ma quella convalida non è nello scadenziario: vale come scaduta, gli interventi di quella categoria li fa un collega abilitato.
- Chi non lavora in sterile (*Area sterile = NO* nel foglio Vestizioni) non compare.
- Le date le aggiorna il responsabile (foglio Vestizioni e scadenziario nel foglio Convalide). Il lunedì il riepilogo riporta chi è in scadenza o scaduto.

## 6. Telefoni dei tecnici

Ogni tecnico è legato al suo telefono. Se un tecnico vede **"Telefono non autorizzato"** (telefono nuovo o cambiato), avvisa il responsabile: lo approva nel foglio Telefoni. Intanto il tecnico può continuare a lavorare: i dati restano in coda e partono dopo l'approvazione.
Se un telefono viene perso, avvisa subito il responsabile per bloccarlo.
Il **codice reparto** per i tecnici nuovi lo dà solo il responsabile: non passarlo tu e non scriverlo in chat o sui gruppi.

## 7. La tua routine

| Quando | Cosa |
|---|---|
| **A ogni passaggio di turno** | Leggi il messaggio, sistema con i tecnici quello che è in **Da sistemare**, anomalie critiche |
| **Ogni giorno** | Anomalie da gestire: chi le ha in carico? Ci sono critiche ferme? Soluzioni da approvare. Compiti in ritardo |
| **Lunedì mattina** | Riepilogo settimanale: guasti che si ripetono e fermi lunghi da discutere con il responsabile |
