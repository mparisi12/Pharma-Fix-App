# Pharma-Fix — Scheda per la direzione

**Cos'è.** Pharma-Fix è l'app della manutenzione: ogni fermo macchina viene registrato dal tecnico partendo dal QR della macchina, con orari, guasto, descrizione, lotto, foto e ricambio. I dati finiscono in un foglio Google che calcola gli indicatori da solo. Nessuno ricopia nulla a mano: i numeri vengono da quello che succede in reparto.

**Filling e Packaging sono sempre separati.** Sono due mondi diversi (sterile e confezionamento): gli indicatori non vengono mai sommati.

## Cosa ricevi

| Cosa | Quando | Contenuto |
|---|---|---|
| **Riepilogo settimanale** | Lunedì alle 7 (mail o Telegram) | Per Packaging e per Filling: n° fermi e tempo di fermo con confronto sulla settimana prima, MTTR, fermi per linea, guasti più frequenti, fermi più lunghi. In fondo: anomalie aperte e fermi ancora in corso |
| **Scheda KPI** (nel foglio, su richiesta) | Quando serve | Indicatori per area, linea, macchina e periodo, con i grafici di Pareto |

## Come leggere gli indicatori

| Indicatore | Cosa dice | Meglio se |
|---|---|---|
| **N° fermi** | Quante volte si è fermata la macchina/linea | Scende |
| **Tempo di fermo** | Ore perse per guasti | Scende |
| **MTTR** | Tempo medio per rimettere in marcia (tempo di fermo ÷ n° fermi): misura la **rapidità della manutenzione** | Scende |
| **MTBF** | Ore di funzionamento medie tra un fermo e l'altro: misura l'**affidabilità della macchina**. Le ore di funzionamento vengono dai dati di produzione | Sale |
| **Occurrancy** | Fermi per lotto (n° fermi ÷ n° lotti): confronta periodi con produzioni diverse | Scende |
| **RFR** | Peso di un guasto sul totale (fermi di quel guasto ÷ tutti i fermi): dice **dove intervenire** | Il guasto n°1 scende |

## I due Pareto

- **Per frequenza**: i guasti che capitano più spesso. Riducono le interruzioni e il lavoro dei tecnici.
- **Per durata**: i guasti che fanno perdere più ore. Spesso sono diversi dai primi: un guasto raro ma lungo (es. elettrico) qui sale in cima.

La linea rossa è la percentuale cumulata: i guasti a sinistra dell'80% sono le priorità.

## Perché i dati sono affidabili

- Ogni registrazione è legata al **telefono personale** del tecnico, con l'ora vera di inserimento.
- Il **passaggio di turno** è generato dai dati registrati, non scritto a mano: compaiono anche i fermi senza descrizione o lasciati aperti.
- In **Filling** gli orari sono quelli del **batch record**, registrati entro fine turno.
- Il foglio viene salvato in **backup** ogni settimana.
