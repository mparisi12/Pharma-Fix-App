# Pharma-Fix — Guida per il tecnico

Pharma-Fix è l'app per registrare i fermi macchina partendo dal **QR della macchina**. Con lo stesso QR trovi anche lo storico degli interventi, le soluzioni già provate dai colleghi, i manuali e le schede formato, e puoi segnalare un'anomalia.

> Regola d'oro: **scansiona il QR quando inizi e quando finisci**. Il resto lo fa l'app.

## 1. Prima volta: installare l'app

1. Apri **https://mparisi12.github.io/Pharma-Fix-App/fermi/** con **Chrome** (Android) o **Safari** (iPhone).
2. Chrome: menu ⋮ → **Installa app** / **Aggiungi a schermata Home**. Safari: **Condividi → Aggiungi alla schermata Home**.
3. Apri l'app dall'icona, scrivi **Nome** e **Cognome** e il **Codice reparto** che ti dà il responsabile della manutenzione.
4. Alla prima scansione consenti l'uso della **fotocamera**.

**iPhone:** apri Pharma-Fix **sempre dall'icona sulla schermata Home** e per i QR usa il pulsante **QR dentro l'app**, non la fotocamera dell'iPhone. Aperta da Safari o da Chrome l'app non ricorda accesso e codice reparto (lo vedi dall'avviso giallo *Stai usando il browser, non l'app*). Dall'icona invece l'accesso resta: si rinnova a ogni apertura e si rifà solo dopo 30 giorni senza usarla.

**Il tuo nome è legato al tuo telefono personale.** Usa sempre il tuo telefono e il tuo nome. Se cambi telefono, avvisa il coordinatore: il telefono nuovo va approvato prima di poter registrare.

## 2. Macchine PACKAGING: aprire e chiudere un fermo

### Aprire (~10 secondi)
1. Nella home tocca **APRI FERMO MACCHINA** e inquadra il QR (oppure *Scegli la macchina a mano* / *Fotografa QR*).
2. Si apre subito il riquadro **Apri fermo**: controlla la macchina e il **lotto** (obbligatorio).
3. **Quando si è fermata?** **ADESSO**, oppure **SCEGLI ORARIO** se si è fermata prima (scrivi l'ora).
4. **Apri fermo**: parte il cronometro. (*Non ora* chiude il riquadro e ti lascia sulla scheda della macchina.)

### Mentre lavori: "Cosa provare"
Con il fermo aperto, in cima alla pagina della macchina c'è il grande pulsante verde **PREMI QUI PER CHIUDERE FERMO** con il cronometro. Sotto trovi **Cosa provare**: tocca il guasto che vedi e leggi le **soluzioni verificate** dei colleghi. Spesso il problema è già stato risolto da qualcuno.

Conosci una soluzione nuova? In fondo alla pagina c'è il link **Conosci una soluzione per questa macchina? Aggiungila**. Attenzione: **aggiungere una soluzione non chiude il fermo**, il fermo si chiude solo dal pulsante verde.

### Chiudere (~30 secondi)
1. **Scansiona di nuovo il QR** → pulsante verde **PREMI QUI PER CHIUDERE FERMO** (oppure **Chiudi** sul fermo nella home). In alto vedi quando è stato aperto, da chi, la durata e il lotto.
   **Quando è ripartita?** **ADESSO**, oppure **SCEGLI ORARIO** se la macchina è ripartita prima e chiudi solo ora (l'orario viene segnato come dichiarato).
2. Tre domande, in ordine:
   - **Cosa si è guastato**: il componente o la zona (i più frequenti su quella macchina sono in cima). La **categoria** (meccanico, elettrico, automation) la mette l'app da sola; la scegli tu solo con ALTRO.
   - **Cosa hai fatto**, con una regola sola:
     - **REGOLAZIONE**: hai spostato o registrato qualcosa (nonio, guida, parametro);
     - **SOSTITUZIONE**: hai montato un pezzo nuovo → codice pezzo obbligatorio. Niente pezzo nuovo = non è una sostituzione;
     - **RIPRISTINO**: reset, riavvio, sblocco, homing: non hai regolato né cambiato niente;
     - **PULIZIA**.
     L'app propone già scelto l'intervento più usato per quel guasto su quella macchina: se hai fatto altro, tocca quello giusto.
   - **Perché** (facoltativo, se lo sai): usura, formato / regolazione persa, materiale fuori specifica, errore operatore, elettrica / alimentazione, non nota. Tocca di nuovo per togliere la scelta.
3. **Foto** della zona di intervento (consigliata) → **tocca la foto per cerchiare** il punto.
4. **Descrizione breve** (**obbligatoria**): cosa hai trovato e cosa hai fatto, anche corta.
5. Hai risolto qualcosa che può servire ai colleghi? Spunta **Proponi come soluzione** e scrivi il problema in una riga: diventa una soluzione (la approva il coordinatore).
6. Controlla il **lotto**.
7. **Chiudi fermo**. Se il fermo è durato meno di 3 minuti l'app chiede conferma: se è giusto tocca di nuovo, altrimenti correggi l'orario.

Il fermo lo chiude **chi l'ha aperto** (o un coordinatore). Hai sbagliato qualcosa in un fermo già chiuso? Toccalo nello storico o in *Chiusi oggi* → **Correggi il fermo**: la **prima** correzione la fai tu, anche degli orari. Per correggere ancora → **Chiedi correzione** e scrivi cosa va cambiato: lo sistema il responsabile.

Fermo aperto per sbaglio? **Chiudi** → in fondo, sotto la riga tratteggiata, **Annulla il fermo** (tocca due volte).

Ti sei **dimenticato di aprirlo** e il fermo è già finito? Nella pagina della macchina tocca **Ti sei dimenticato di aprirlo? Registra un fermo già finito**, scrivi inizio e fine e compila come sempre. Si può fare solo per il turno in corso (o entro 2 ore dal cambio turno) e nel passaggio di turno compare tra i *Registrati dopo*: meglio aprirlo sempre col QR.

## 3. Macchine FILLING: apri prima di entrare, chiudi all'uscita

In sterile il telefono non entra, ma il fermo si apre lo stesso: così tutti sanno che c'è un fermo in corso.
1. **Prima di entrare**: apri la macchina (QR, **Fotografa QR** o **Scegli macchina a mano**) → **Apri fermo**.
2. **All'uscita**: QR o fermo nella home → **Chiudi fermo**. Inizio e Fine sono già scritti: **correggili con gli orari del batch record**, poi compila come sempre.

Fermo già finito e non l'avevi aperto? Nella pagina della macchina tocca **Fermo già finito? Registralo con gli orari del batch record**.

Gli orari scritti a mano valgono per il **turno in corso**, o per quello appena finito se registri **entro 2 ore dal cambio turno**.

## 4. Le altre cose che trovi sulla macchina

| Cosa | A cosa serve |
|---|---|
| **In produzione oggi** (sotto il nome della macchina) | Cosa sta girando oggi su quella linea e cosa c'è domani, dal programma di produzione |
| **Storico** (scorrendo la pagina della macchina) | Gli ultimi interventi, 5 alla volta (**Mostra altri** per i precedenti): chi, quando, quanto è durato, cosa ha fatto. Tocca un intervento per vedere la foto |
| **Soluzioni** (barra in basso) | L'archivio di tutte le macchine: cerca per parola (pinze, ventose…) o per guasto. **+ Aggiungi soluzione**: la vedi subito tu (*DA APPROVARE*), la vedranno tutti dopo l'approvazione del coordinatore |
| **Manuali / Schede formato** | Manuali, schemi elettrici e schede formato della macchina o della linea |
| **Segnala anomalia** | Qualcosa da sistemare ma che non ferma la macchina adesso: urgenza **1** critico, **2** urgente, **3** programmabile, foto e cosa serve |

## 5. Il menu del tuo nome

Tocca il **tuo nome** in alto a destra. Il **numero rosso** sul nome ti dice che c'è qualcosa da fare (compiti o vestizione in scadenza).

### I miei compiti
Il coordinatore può assegnarti dei lavori (es. un inventario, una verifica, una pulizia). Il **numero rosso sul tuo nome**, in alto, dice quanti compiti hai da fare.

Quando entri nell'app e c'è un compito nuovo per te compare al centro la finestra **Hai un nuovo compito!**: **Vedi il compito** lo apre, **Dopo** la chiude ma te la ripropone al prossimo ingresso finché non lo guardi.
1. **I miei compiti**: quelli assegnati a te e quelli del **tuo turno** (secondo il foglio turni).
2. Tocca un compito → **Prendo in carico** (per un compito del turno: **Lo prendo io**, così i colleghi sanno che è tuo).
3. A lavoro finito scrivi una **nota** (cosa hai fatto; se serve, il riferimento al logbook) → **Fatto**.

Il coordinatore vede chi l'ha fatto e quando. Il compito non sostituisce logbook e moduli.

### Programma produzione
Nella scheda della macchina vedi cosa è in produzione oggi su quella linea; aprendo un fermo i lotti del giorno sono già pronti da toccare. **Programma produzione** mostra tutte le linee (oggi, domani, dopodomani) con i cambi formato in giallo.

### Turni
**I miei turni** (i prossimi 14 giorni, con i cambi turno) e **Chi c'è** (chi è in turno oggi, domani e dopodomani, chi è assente o in ferie).

### Vestizione sterile e convalide GMP
Nel menu del nome, **Le mie abilitazioni** mostra la tua **vestizione** (accesso all'area sterile) e le **convalide GMP** per categoria di intervento (es. Cat. G regolazione, Cat. H sostituzione componente), con la scadenza di ognuna.
- **In scadenza** (da 60 giorni prima, o i giorni decisi dal responsabile): avviso giallo nel menu e pallino rosso sul nome. Prenota il rinnovo con il responsabile. Il pallino si spegne quando apri il menu e si riaccende a 14 giorni, a 7 giorni e alla scadenza.
- **Scaduta**: avviso rosso. Vestizione scaduta = **non entrare in sterile**. Convalida scaduta (o *non trovata* nello scadenziario) = in sterile quegli interventi li fa un collega abilitato finché non è rinnovata.
- Se non lavori in area sterile non vedi niente di tutto questo.

## 6. Il pallino in alto

L'app si aggiorna da sola ogni minuto. Vuoi vedere **subito** se un collega ha aperto un fermo? **Tocca la scritta in alto** (quella col pallino e la freccia ↻): si aggiorna in quel momento e ti dice quanti fermi sono aperti.

| Pallino | Cosa fare |
|---|---|
| 🟢 Sincronizzato | Niente, è tutto arrivato |
| 🟠 Invio in corso | Niente, parte da solo |
| 🟠 Il foglio risponde lentamente | Niente: lo script si sta svegliando, riprova da solo |
| 🔴 Non connesso da N min · dati delle hh:mm | **Continua a lavorare**: i dati restano sul telefono e partono quando torna la rete (l'app riprova da sola). L'ora dice a quando risalgono i dati che vedi. Non disinstallare l'app |
| 🔴 Telefono non autorizzato | Avvisa il coordinatore: deve approvare il tuo telefono. I dati restano in coda e partono dopo |

## 7. Cose da ricordare

- **Un fermo per volta per macchina**, aperto e chiuso con il QR.
- **Mai lasciare un fermo aperto** a fine turno: arriva un avviso al coordinatore.
- **Descrizione sempre**: il passaggio di turno si scrive da solo con quello che registri, e i fermi senza descrizione compaiono con il tuo nome in "Da sistemare".
- Il **lotto** è obbligatorio.
- Problemi con l'app? Chiudila del tutto e riaprila con la rete. Se continua, avvisa il coordinatore.
