# Pharma-Fix — Guida per il coordinatore

Il coordinatore usa l'app come i tecnici (vedi la *Guida per il tecnico*) e in più **tiene d'occhio il reparto**: riceve gli avvisi, segue le anomalie, controlla che i fermi siano registrati bene. Gli strumenti sono due: il **gruppo Telegram** e la sezione **Anomalie** dell'app.

## 1. Cosa ti arriva su Telegram e cosa fare

| Messaggio | Quando arriva | Cosa fare |
|---|---|---|
| 🔴 **Anomalia CRITICA** | Appena un tecnico segnala un'anomalia con urgenza 1 | Valuta subito: serve fermare la linea, chiamare un fornitore, ordinare un ricambio? Prendila in carico dall'app |
| ⏱️ **Fermo APERTO DA…** | Un fermo resta aperto oltre le ore impostate (es. 2) | È un fermo lungo o qualcuno si è dimenticato di chiuderlo. Chiama il tecnico: se ha finito, lo chiude dall'app |
| 🔁 **Guasto RIPETUTO** | Lo stesso guasto sulla stessa macchina si ripete (es. 3 volte in 7 giorni) | Non è più un caso: serve una soluzione vera (regolazione, ricambio, preventiva). Parlane con il responsabile |
| 🔔 **Passaggio turno** | A fine di ogni turno (6, 14, 22) | Leggilo con chi entra. È scritto **dai dati registrati**, non dai tecnici: se vedi "(senza descrizione)" o "guasto non indicato", chiedi al tecnico di completare |
| 🔔 **Riepilogo settimanale** | Lunedì alle 7 | Numeri della settimana, separati **Packaging** e **Filling**: fermi, tempo, MTTR, guasti più frequenti, fermi più lunghi |

Su ST3, finché è a ciclo continuo, il passaggio di turno arriva anche sabato e domenica (solo con i dati di ST3).

## 2. Anomalie: prenderle in carico e chiuderle

Dall'app, barra in basso → **Anomalie**:
- **Da gestire**: ordinate per urgenza (1 critico, 2 urgente, 3 programmabile).
- Tocca un'anomalia → **Prendo in carico** quando qualcuno ci sta lavorando (anche tu per conto di un tecnico).
- A lavoro finito → **Chiudi anomalia** con una nota breve su cosa è stato fatto.

Obiettivo: **nessuna anomalia critica aperta a fine turno** senza che qualcuno l'abbia presa in carico.

## 3. Fermi: controllare che siano registrati bene

- **Fermi aperti**: li vedi nella home dell'app. A fine turno non ce ne devono essere di dimenticati.
- **Filling**: i fermi si registrano all'uscita dallo sterile **entro fine turno**, con gli orari del batch record. Se nel batch record c'è un fermo che non compare nel passaggio di turno, va registrato.
- **Descrizioni**: una descrizione corta ma vera serve a tutti (e alle soluzioni future). Le righe "(senza descrizione)" nel passaggio di turno sono quelle da far completare.

## 4. Soluzioni: sei tu ad approvarle

Quando un tecnico risolve qualcosa di nuovo, chiedigli di aggiungerlo in **Soluzioni → + Aggiungi soluzione**. Le soluzioni nuove **non le vede nessuno finché non le approvi tu**: solo dopo compaiono a tutti in **Soluzioni**, **Cosa provare** e **Prova prima**.

- Ti arriva un messaggio Telegram **SOLUZIONE DA APPROVARE**, e nella home dell'app compare *"N soluzioni da approvare"*.
- Apri **Soluzioni**: in alto c'è **Da approvare**. Leggi problema e soluzione:
  - **Approva** se è giusta, chiara e ripetibile da un collega;
  - **Scarta** (tocca due volte) se è sbagliata, troppo vaga o un doppione. Se è buona ma scritta male, scartala e chiedi al tecnico di riscriverla.
- Le soluzioni che aggiungi tu sono già approvate.

Se una soluzione già approvata è sbagliata o superata, segnalalo al responsabile, che la corregge nel foglio.

## 5. Telefoni dei tecnici

Ogni tecnico è legato al suo telefono. Se un tecnico vede **"Telefono non autorizzato"** (telefono nuovo o cambiato), avvisa il responsabile: lo approva nel foglio Telefoni. Intanto il tecnico può continuare a lavorare: i dati restano in coda e partono dopo l'approvazione.
Se un telefono viene perso, avvisa subito il responsabile per bloccarlo.
Il **codice reparto** per i tecnici nuovi lo dà solo il responsabile: non passarlo tu e non scriverlo in chat o sui gruppi.

## 6. La tua routine

| Quando | Cosa |
|---|---|
| **A ogni passaggio di turno** | Leggi il messaggio, controlla fermi aperti e "(senza descrizione)", anomalie critiche |
| **Ogni giorno** | Anomalie da gestire: chi le ha in carico? Ci sono critiche ferme? Soluzioni da approvare |
| **Lunedì mattina** | Riepilogo settimanale: guasti che si ripetono e fermi lunghi da discutere con il responsabile |
