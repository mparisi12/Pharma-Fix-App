/**
 * Pharma-Fix — backend Google Apps Script
 * --------------------------------------------------------------
 * Il Google Sheet che contiene questo script È il database.
 * Fogli:
 *   Fermi      → una riga per fermo, colonne A–L identiche al template
 *                "Template fermi meccanici", più colonne M–U aggiuntive
 *   Equipment  → anagrafica macchine (codice nel QR, linea, area, macchina)
 *   Liste      → categorie, tipi intervento, tipi guasto, manutentori
 *   KPI        → indicatori calcolati con formule (si aggiornano da soli)
 *
 * Installazione: vedi LEGGIMI.md. In breve: esegui setup() una volta,
 * poi Distribuisci → Nuova distribuzione → App web.
 */

const SH_FERMI = 'Fermi', SH_EQ = 'Equipment', SH_LISTE = 'Liste', SH_KPI = 'KPI', SH_GUASTI = 'Guasti';

// Tipi di guasto di partenza. Chiave = codice macchina (QR), nome linea, oppure TUTTE.
// Solo voci generiche: quelle specifiche delle macchine si compilano nel foglio Guasti (privato), mai qui nel codice pubblico.
const DEF_GUASTI = {
  'TUTTE': ['FORMATO', 'SENSORE', 'USURA', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO']
};
const TZ = 'Europe/Rome';

// Colonne del foglio Fermi (A–L = template originale)
const HEAD = [
  'Lotto', 'Data', 'Ora Inizio', 'Ora Fine', 'Manutentore ', 'Macchina', 'Categoria',
  'Tipo intervento', 'Tipo guasto', 'Altro (descrizione breve)', 'Durata intervento MTTR',
  'Intervento-Guasto',
  'Linea', 'Turno', 'Codice pezzo', 'Foto apertura', 'Foto chiusura', 'Stato', 'ID fermo',
  'Equipment', 'Aperto da',
  'Orari', 'Registrato alle', 'Telefono'
];
const C = { LOTTO: 1, DATA: 2, INIZIO: 3, FINE: 4, MAN: 5, MACC: 6, CAT: 7, INT: 8, GUASTO: 9, ALTRO: 10,
  DUR: 11, IG: 12, LINEA: 13, TURNO: 14, PEZZO: 15, FOTO_A: 16, FOTO_C: 17, STATO: 18, ID: 19, EQ: 20, APERTO_DA: 21, ORARI: 22, REG: 23, TEL: 24, CAUSA: 25, PROP: 26, RIV: 27, GRUPPO: 28, GIRO: 29, MULTI: 30, FOTO_X: 31 }; // CAUSA e seguenti sono fuori da HEAD: le colonne si aggiungono solo quando servono

const DEF_LISTE = {
  'Categoria': ['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO'],
  'Tipo intervento': ['REGOLAZIONE', 'SOSTITUZIONE', 'RIPRISTINO', 'PULIZIA'],
  'Guasto Sterile': ['SENSORE', 'FORMATO', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO'], // voci vere: foglio Liste
  'Guasto Secondario': ['GUIDE', 'INCASTRO', 'FORMATO', 'REGOLAZIONE DA PANNELLO', 'USURA', 'SENSORE', 'PULIZIA',
    'SOSTITUZIONE', 'RIAVVIO', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO'], // voci vere: foglio Liste
  'Manutentori': [] // compilare nel foglio Liste: COGNOME Nome
};

// Linee: [codice, nome, ha filling (sterile), ha packaging (secondario)]
const DEF_LINEE = [
  ['ST1MAR', 'ST1 MAR', 1, 1], ['ST1TGM', 'ST1 TGM', 1, 1], ['ST2', 'ST2', 1, 1], ['ORL', 'ORL', 1, 0],
  ['ST3', 'ST3', 1, 1], ['ST4', 'ST4', 1, 1], ['ST5', 'ST5', 1, 0], ['SPL1', 'SPL1', 0, 1],
  ['SPL2', 'SPL2', 0, 1], ['ST6', 'ST6', 1, 1]
];
const DEF_MACC = {
  Sterile: [['RIO', 'Riordinatore'], ['FIL', 'Filling - Tappatrice']],
  Secondario: [['PGR', 'Polmone Grande'], ['ETI', 'Etichettatrice'], ['PPI', 'Polmone Piccolo'],
    ['AST', 'Astucciatrice'], ['BIL', 'Bilancia'], ['CAR', 'Cartonatrice']]
};


/* ---------- Separatore delle formule ----------
   Un foglio in italiano vuole il punto e virgola tra gli argomenti (=SE(A1="";…)), uno in inglese la virgola.
   Lo script scrive le formule con la virgola e F_() le converte se serve. */
function sep_() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('SEP');
  if (s) return s;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = ss.insertSheet('_prova_formule');
  t.getRange('A1').setFormula('=SUM(1,2)');
  SpreadsheetApp.flush();
  s = t.getRange('A1').getValue() === 3 ? ',' : ';';
  ss.deleteSheet(t);
  props.setProperty('SEP', s);
  return s;
}
function F_(f) {
  if (!f || sep_() === ',') return f;
  let out = '', q = false;
  for (const ch of String(f)) { if (ch === '"') q = !q; out += (!q && ch === ',') ? ';' : ch; }
  return out;
}

// Ripara le formule già scritte con il separatore sbagliato (#ERROR!) e ricrea il foglio KPI.
function riparaFormule() {
  PropertiesService.getScriptProperties().deleteProperty('SEP');
  const convertite = sistemaDate_();
  Logger.log('Righe con date scritte come testo convertite: ' + convertite);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_FERMI);
  const last = sh.getLastRow();
  if (last >= 2) {
    const k = [], l = [];
    for (let r = 2; r <= last; r++) {
      k.push([F_('=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))')]);
      l.push([F_('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')')]);
    }
    sh.getRange(2, C.DUR, last - 1, 1).setFormulas(k).setNumberFormat('[h]:mm');
    sh.getRange(2, C.IG, last - 1, 1).setFormulas(l);
    [C.FOTO_A, C.FOTO_C].forEach(col => {
      const rg = sh.getRange(2, col, last - 1, 1);
      rg.setFormulas(rg.getFormulas().map(r => [r[0] ? F_(r[0].replace(/;/g, ',')) : '']));
    });
  }
  aggiornaKpi();
  Logger.log('Formule riparate. Separatore del foglio: "' + sep_() + '"');
}

/* =================== SETUP (da eseguire una volta) =================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  const props = PropertiesService.getScriptProperties();

  // Fermi
  let sh = ss.getSheetByName(SH_FERMI) || ss.insertSheet(SH_FERMI, 0);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#dfe7ec');
    sh.getRange(1, C.LINEA, 1, HEAD.length - C.LINEA + 1).setBackground('#fff3d6');
    sh.setFrozenRows(1);
    sh.setColumnWidth(C.ALTRO, 280);
    sh.setColumnWidth(C.GUASTO, 220);
  }

  // Equipment
  sh = ss.getSheetByName(SH_EQ) || ss.insertSheet(SH_EQ);
  if (sh.getLastRow() === 0) {
    const rows = [['Codice (QR)', 'Linea', 'Area', 'Macchina', 'Attivo']];
    DEF_LINEE.forEach(([cod, nome, f, p]) => {
      if (f) DEF_MACC.Sterile.forEach(([m, n]) => rows.push([cod + '-' + m, nome, 'Filling', n, 'SI']));
      if (p) DEF_MACC.Secondario.forEach(([m, n]) => rows.push([cod + '-' + m, nome, 'Packaging', n, 'SI']));
    });
    sh.getRange(1, 1, rows.length, 5).setValues(rows);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#dfe7ec');
    sh.setFrozenRows(1);
  }

  // Liste
  sh = ss.getSheetByName(SH_LISTE) || ss.insertSheet(SH_LISTE);
  if (sh.getLastRow() === 0) {
    const keys = Object.keys(DEF_LISTE);
    const max = Math.max(...keys.map(k => DEF_LISTE[k].length));
    const rows = [keys];
    for (let i = 0; i < max; i++) rows.push(keys.map(k => DEF_LISTE[k][i] || ''));
    sh.getRange(1, 1, rows.length, keys.length).setValues(rows);
    sh.getRange(1, 1, 1, keys.length).setFontWeight('bold').setBackground('#dfe7ec');
    sh.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SH_KPI)) buildKpi_(ss); // il foglio KPI non viene più ricostruito se esiste già
  if (!ss.getSheetByName(SH_GUASTI)) buildGuasti_(ss);
  creaFogliArchivio_(ss);
  aggiornaV3_(ss);
  sistemaDurate();

  // Cartella foto
  if (!props.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('Pharma-Fix Foto Fermi');
    props.setProperty('FOLDER_ID', f.getId());
  }
  // Codice reparto: chiave che i tecnici inseriscono al primo accesso (cambiala in Proprietà script → KEY)
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 12));

  const s1 = ss.getSheetByName('Foglio1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  Logger.log('FATTO. Codice reparto attuale (KEY): ' + props.getProperty('KEY') + ' — per cambiarlo: Impostazioni progetto → Proprietà script → KEY');
}

// Foglio Guasti: una riga per macchina (codice QR), i tipi di guasto uno per cella verso destra.
// Righe speciali: TUTTE (vale per ogni macchina) e il nome di una linea (vale per tutte le macchine della linea).
function buildGuasti_(ss) {
  const sh = ss.insertSheet(SH_GUASTI, ss.getSheets().length);
  const rows = [['Codice (QR) / Linea / TUTTE', 'Macchina (promemoria)', 'Tipi di guasto → uno per cella, verso destra']];
  rows.push(['TUTTE', 'vale per tutte le macchine'].concat(DEF_GUASTI['TUTTE']));
  equipment_(ss).forEach(e => rows.push([e.code, e.linea + ' · ' + e.macchina].concat(DEF_GUASTI[e.code] || [])));
  const w = Math.max(...rows.map(r => r.length));
  sh.getRange(1, 1, rows.length, w).setValues(rows.map(r => r.concat(new Array(w - r.length).fill(''))));
  sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#dfe7ec');
  sh.getRange(2, 1, 1, w).setBackground('#fff3d6');
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 260);
}

function guasti_(ss) {
  const sh = ss.getSheetByName(SH_GUASTI);
  if (!sh || sh.getLastRow() < 2) return {};
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(r => {
    const k = String(r[0]).trim().toUpperCase();
    if (k) out[k] = r.slice(2).map(v => String(v).trim()).filter(Boolean);
  });
  return out;
}

// Riscrive la formula della durata su tutte le righe (funziona sia con data+ora sia con sola ora).
function sistemaDurate() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const last = sh.getLastRow();
  if (last < 2) return;
  const f = [];
  for (let r = 2; r <= last; r++) f.push(['=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))']);
  sh.getRange(2, C.DUR, last - 1, 1).setFormulas(f.map(r => r.map(F_))).setNumberFormat('[h]:mm');
}

// ---------- KPI con filtri (linea, macchina, periodo) ----------
// I valori li calcola lo script (niente formule complicate): si aggiornano da soli quando cambi un filtro
// giallo, quando si chiude un fermo, oppure dal menu Pharma-Fix → Aggiorna KPI.
const PERIODI = ['Ultimi 7 giorni', 'Ultimi 30 giorni', 'Ultimi 90 giorni', 'Mese corrente', 'Mese scorso', 'Anno corrente', 'Tutto', 'Personalizzato'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Pharma-Fix')
    .addItem('Aggiorna KPI', 'aggiornaValoriKpi')
    .addItem('Ricrea foglio KPI', 'aggiornaKpi')
    .addItem('Aggiorna KPI Risorse', 'aggiornaKpiRisorse')
    .addItem('Alleggerisci foglio Programma (colonne vuote)', 'alleggerisciProgramma')
    .addSeparator()
    .addItem('Controlla registrazioni', 'controllaRegistrazioni')
    .addItem('Ricodifica fermi vecchi: prepara', 'ricodificaPrepara')
    .addItem('Ricodifica fermi vecchi: applica', 'ricodificaApplica')
    .addItem('Sistema lotti (zeri iniziali)', 'sistemaLotti')
    .addItem('Completa fermi scritti a mano (da carta)', 'sistemaFermiCarta')
    .addItem('Aggiorna elenchi foglio Documenti', 'aggiornaElenchiDocumenti')
    .addItem('Aggiorna menu a tendina foglio Fermi', 'aggiornaTendineFermi')
    .addItem('Voci guasto: applica unioni (colonna E)', 'applicaUnioniVoci')
    .addSeparator()
    .addItem('Attiva sveglia dello script', 'attivaSveglia')
    .addItem('Disattiva sveglia dello script', 'disattivaSveglia')
    .addSeparator()
    .addItem('Prova avvisi (mail / Telegram)', 'provaAvvisi')
    .addItem('Telegram: trova ID chat', 'trovaChatTelegram')
    .addSeparator()
    .addItem('Attiva report e backup automatici', 'attivaReport')
    .addItem('Prova riepilogo settimanale', 'riepilogoSettimanale')
    .addItem('Prova passaggio di turno', 'provaPassaggioTurno')
    .addItem('Fai un backup adesso', 'backupFoglio')
    .addSeparator()
    .addItem('Formati: collega il file Base Gestione Formati', 'collegaFormati')
    .addItem('Versione dello script', 'infoVersione')
    .addToUi();
  // Foglio Convalide MF: nomi nuovi, impostazioni e colori subito all'apertura, senza aspettare l'app
  try { vestSetup_(SpreadsheetApp.getActiveSpreadsheet()); } catch (e) { Logger.log(e); }
}

function onEdit(e) {
  try {
    if (e.range.getSheet().getName() === SH_KR) { if (e.range.getRow() <= 4) aggiornaKpiRisorse(); return; } // filtri del foglio KPI Risorse
    if (e.range.getSheet().getName() === SH_FERMI && e.range.getRow() > 1) { // cambiata la macchina di una riga: tendina Tipo guasto di quella macchina
      const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
      if ([C.MACC, C.LINEA, C.EQ].some(c => c >= c1 && c <= c2)) try { const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
        if (c1 === c2 && e.range.getNumRows() === 1) eqCoerente_(ss, e.range.getSheet(), e.range.getRow(), c1);
        fermiGuastoDV_(ss, e.range.getSheet(), e.range.getRow(), e.range.getNumRows()); } catch (x) { Logger.log(x); }
    }
    if (e.range.getSheet().getName() !== SH_KPI) { bumpVer_(); return; } // modifica a mano: i telefoni si aggiornano
    if (['B2', 'B3', 'B4', 'B5', 'D6', 'D7', 'F10'].indexOf(e.range.getA1Notation()) >= 0) aggiornaValoriKpi();
  } catch (x) { Logger.log(x); }
}

// Ricrea da zero il foglio KPI (impaginazione, menu, grafico) e lo calcola
function aggiornaKpi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['KPI_dati', SH_KPI].forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  buildKpi_(ss);
}

const KPI_AREE = ['PACKAGING', 'FILLING'];
const KPI_T = 31, KPI_N = 20; // riga intestazioni tabelle, righe mostrate nei grafici
const KPI_C1 = 27, KPI_C2 = 33, MICRO_MIN = 3; // colonne delle tabelle per causa e per intervento (dopo quelle nascoste dei grafici); micro-fermo = meno di 3 minuti
let KPI_NUOVO = false;
// Pareto e guasto n°1 contano la coppia intervento · voce: "REG · MAG. ASTUCCI" e "SOS · MAG. ASTUCCI" sono problemi diversi
const KPI_GH = 'Intervento · guasto', INT_BREVE = { REGOLAZIONE: 'REG', SOSTITUZIONE: 'SOS', RIPRISTINO: 'RIP', PULIZIA: 'PUL' };
const intGuasto_ = (int, g) => (int && int.charAt(0) !== '(' ? (INT_BREVE[int] || int.slice(0, 3)) + ' · ' : '') + g;
const ORE_GIORNO = 24; // stima MTBF: ore per giorno di produzione (3 turni)
function buildKpi_(ss, keep) {
  const sh = ss.getSheetByName(SH_KPI) || ss.insertSheet(SH_KPI, 1);
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getCharts().forEach(c => sh.removeChart(c));
  const H = '#dfe7ec', Y = '#fff3d6', G = '#5a6873';
  sh.getRange('A1').setValue('KPI manutenzione — fermi chiusi').setFontWeight('bold').setFontSize(14);
  sh.getRange('C1').setValue('Cambia le celle gialle: i valori si ricalcolano da soli (oppure menu Pharma-Fix → Aggiorna KPI)').setFontColor(G);
  sh.getRange('A2').setValue('Area').setFontWeight('bold');
  sh.getRange('B2').setValue('PACKAGING').setBackground(Y).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KPI_AREE, true).setAllowInvalid(false).build());
  sh.getRange('C2').setValue('Filling e Packaging sono sempre separati: scegli quale guardare').setFontColor(G);
  sh.getRange('A3:A7').setValues([['Linea'], ['Macchina'], ['Periodo'], ['Dal'], ['Al']]).setFontWeight('bold');
  sh.getRange('B3:B5').setValues([['TUTTE'], ['TUTTE'], ['Ultimi 30 giorni']]).setBackground(Y);
  sh.getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PERIODI, true).setAllowInvalid(false).build());
  sh.getRange('B6:B7').setNumberFormat('dd/MM/yyyy');
  sh.getRange('C6:C7').setValues([['Personalizzato: dal →'], ['Personalizzato: al →']]).setFontColor(G).setHorizontalAlignment('right');
  sh.getRange('D6:D7').setBackground(Y).setNumberFormat('dd/MM/yyyy');

  if (sh.getMaxColumns() < KPI_C2 + 4) sh.insertColumnsAfter(sh.getMaxColumns(), KPI_C2 + 4 - sh.getMaxColumns());
  sh.getRange('A9:K9').setValues([['N° fermi (macchina ferma)', 'Tempo di fermo', 'MTTR per intervento (senza micro)', 'N° lotti', 'Occurrancy (fermi/lotti)',
    'Ore funzionamento (vuoto = stima)', 'MTBF (h, sui fermi)', 'Guasto n°1 · RFR', 'Micro-interventi (< ' + MICRO_MIN + ' min)', 'N° interventi', 'Durata media di un fermo']])
    .setFontWeight('bold').setBackground(H).setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A10:K10').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F10').setBackground(Y);
  sh.getRange('F11').setValue('vuota = stima dal Programma (giorni con lotti × 24 h); scrivi le ore vere se le hai').setFontColor(G).setFontSize(9).setWrap(true);
  sh.getRange('H11').setValue('RFR = interventi su quel guasto (stesso tipo di intervento) / totale interventi').setFontColor(G).setFontSize(9);

  // grafici in alto (righe 12-29), tabelle sotto (da riga KPI_T)
  const t = KPI_T;
  sh.getRange(t - 1, 1).setValue('Pareto per FREQUENZA — quante volte succede').setFontWeight('bold');
  sh.getRange(t, 1, 1, 5).setValues([[KPI_GH, 'N° interventi', 'Tempo', '% cumulata', 'RFR']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, 7).setValue('Pareto per DURATA — quanto ferma la linea').setFontWeight('bold');
  sh.getRange(t, 7, 1, 5).setValues([[KPI_GH, 'Ore di fermo', 'N° interventi', '% cumulata', '% del tempo']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, 13).setValue('Fermi per macchina').setFontWeight('bold');
  sh.getRange(t, 13, 1, 6).setValues([['Linea', 'Macchina', 'N° fermi', 'Tempo', 'MTTR per intervento', 'Guasto n°1 · RFR']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, KPI_C1).setValue('Per CAUSA — perché si guasta').setFontWeight('bold');
  sh.getRange(t, KPI_C1, 1, 5).setValues([['Causa', 'N° interventi', 'Ore di fermo', '% degli interventi', '% del tempo']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, KPI_C2).setValue('Per COSA HAI FATTO').setFontWeight('bold');
  sh.getRange(t, KPI_C2, 1, 4).setValues([['Intervento', 'N° interventi', 'Ore di fermo', '% degli interventi']]).setFontWeight('bold').setBackground(H);
  [[KPI_C1, 220], [KPI_C1 + 1, 80], [KPI_C1 + 2, 90], [KPI_C1 + 3, 90], [KPI_C1 + 4, 90], [KPI_C2, 160], [KPI_C2 + 1, 80], [KPI_C2 + 2, 90], [KPI_C2 + 3, 90]].forEach(([c, w]) => sh.setColumnWidth(c, w));

  [[1, 260], [2, 260], [3, 110], [4, 110], [5, 110], [6, 150], [7, 230], [8, 200], [9, 90], [10, 90], [11, 90], [12, 30],
   [13, 70], [14, 280], [15, 80], [16, 80], [17, 80], [18, 220]].forEach(([c, w]) => sh.setColumnWidth(c, w));
  sh.setFrozenRows(10);

  // dati dei grafici in colonne nascoste T:Y: al massimo KPI_N barre, il resto raggruppato in "ALTRI" così la linea arriva al 100%
  // un solo asse 0-100%: barre = quota di ogni guasto, linea = % cumulata (Google non tiene il secondo asse e la linea restava schiacciata in basso)
  sh.getRange(t, 20, 1, 6).setValues([['Grafico %: guasto', '% dei fermi (RFR)', '% cumulata', 'Grafico %: guasto', '% del tempo di fermo', '% cumulata']]);
  sh.hideColumns(20, 6);
  const pareto = (r1, c1, titolo, asse, colore, col, w) => sh.newChart().asComboChart()
    .addRange(sh.getRange(t, c1, KPI_N + 1, 3))
    .setNumHeaders(1)
    .setOption('title', titolo)
    .setOption('legend', { position: 'bottom' })
    .setOption('series', { 0: { type: 'bars', color: colore }, 1: { type: 'line', color: '#c23b2a', pointSize: 5, lineWidth: 2 } })
    .setOption('vAxis', { title: asse, format: 'percent', minValue: 0, maxValue: 1, viewWindow: { min: 0, max: 1 }, gridlines: { count: 6 } })
    .setPosition(r1, col, 0, 0).setOption('width', w).setOption('height', 360)
    .build();
  try {
    sh.insertChart(pareto(12, 20, 'Pareto per frequenza (% dei fermi)', '% dei fermi', '#245a7a', 1, 840));
    sh.insertChart(pareto(12, 23, 'Pareto per durata (% del tempo di fermo)', '% del tempo', '#b5761b', 7, 800));
  } catch (e) { Logger.log('Grafici non creati: ' + e); }

  KPI_NUOVO = true; // appena ricreato: niente controllo dell'impaginazione (evita di ricrearlo all'infinito)
  if (keep) { if (keep.area) sh.getRange('B2').setValue(keep.area); sh.getRange('B3:B5').setValues(keep.filtri); sh.getRange('D6:D7').setValues(keep.date); sh.getRange('F10').setValue(keep.ore); }
  aggiornaValoriKpi();
}

// Date del periodo scelto (inizio giorno "from", fine giorno "to")
function periodo_(nome, d6, d7, oggi) {
  const g = (y, m, d) => new Date(y, m, d);
  const y = oggi.getFullYear(), m = oggi.getMonth(), d = oggi.getDate();
  let from, to = g(y, m, d);
  switch (nome) {
    case 'Ultimi 7 giorni': from = g(y, m, d - 6); break;
    case 'Ultimi 90 giorni': from = g(y, m, d - 89); break;
    case 'Mese corrente': from = g(y, m, 1); break;
    case 'Mese scorso': from = g(y, m - 1, 1); to = g(y, m, 0); break;
    case 'Anno corrente': from = g(y, 0, 1); break;
    case 'Tutto': from = g(2000, 0, 1); break;
    case 'Personalizzato':
      from = d6 instanceof Date ? g(d6.getFullYear(), d6.getMonth(), d6.getDate()) : g(y, m, d - 29);
      if (d7 instanceof Date) to = g(d7.getFullYear(), d7.getMonth(), d7.getDate());
      break;
    default: from = g(y, m, d - 29); // Ultimi 30 giorni
  }
  return { from: from, to: to, toEnd: new Date(to.getFullYear(), to.getMonth(), to.getDate() + 1) };
}

// Accetta una data vera oppure un testo "gg/mm/aaaa hh:mm" (es. righe incollate); con solo "hh:mm" usa il giorno della riga
function dataOra_(v, giorno) {
  if (v instanceof Date) return v;
  const s = String(v == null ? '' : v).trim();
  let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:[ T]+(\d{1,2})[:.](\d{2}))?/);
  if (m) { const y = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]); return new Date(y, m[2] - 1, m[1], Number(m[4] || 0), Number(m[5] || 0)); }
  m = s.match(/^(\d{1,2})[:.](\d{2})$/);
  if (m && giorno instanceof Date) return new Date(giorno.getFullYear(), giorno.getMonth(), giorno.getDate(), Number(m[1]), Number(m[2]));
  return null;
}

// Converte in date vere Data, Ora Inizio e Ora Fine scritte come testo nel foglio Fermi
function sistemaDate_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const rg = sh.getRange(2, C.DATA, last - 1, 3), v = rg.getValues();
  let n = 0;
  const out = v.map(r => {
    const g = dataOra_(r[0], null), gg = g ? new Date(g.getFullYear(), g.getMonth(), g.getDate()) : r[0];
    const a = r[1] === '' ? '' : (dataOra_(r[1], gg) || r[1]);
    const b = r[2] === '' ? '' : (dataOra_(r[2], gg) || r[2]);
    if (!(r[0] instanceof Date) || (r[1] !== '' && !(r[1] instanceof Date)) || (r[2] !== '' && !(r[2] instanceof Date))) n++;
    return [gg || r[0], a, b];
  });
  rg.setValues(out);
  sh.getRange(2, C.DATA, last - 1, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(2, C.INIZIO, last - 1, 2).setNumberFormat('HH:mm');
  return n;
}

// Ore di funzionamento stimate dal Programma: giorni (nel periodo, fino a oggi) in cui la linea ha qualcosa in programma × ORE_GIORNO
function oreStimate_(ss, p, area, linea) {
  const prog = programmaTutto_(ss), eq = equipment_(ss).filter(e => String(e.area).toUpperCase() === area);
  const linee = [...new Set(eq.map(e => String(e.linea).trim()))].filter(l => linea === 'TUTTE' || l === linea);
  const da = Utilities.formatDate(p.from, TZ, 'yyyy-MM-dd'), fine = new Date(Math.min(p.toEnd.getTime(), Date.now()));
  const a = Utilities.formatDate(new Date(fine.getTime() - 1), TZ, 'yyyy-MM-dd');
  let giorni = 0, nl = 0;
  linee.forEach(l => { const g = Object.keys(prog[l] || {}).filter(k => k >= da && k <= a && String(prog[l][k]).trim()).length; if (g) { giorni += g; nl++; } });
  return { giorni: giorni, ore: giorni * ORE_GIORNO, linee: nl };
}
// Calcolo puro (testabile): righe del foglio Fermi + filtri → numeri e tabelle
const kpiChiave_ = (linea, macc) => (String(linea) + '|' + String(macc)).toUpperCase().replace(/\s+/g, ' ').replace(/ ?\| ?/, '|').trim();
function calcolaKpi_(vals, f) {
  const rows = [];
  const senzaArea = [];
  vals.forEach((r, ix) => {
    if (r[C.STATO - 1] !== 'CHIUSO') return;
    const g0 = dataOra_(r[C.DATA - 1], null);
    const ini = dataOra_(r[C.INIZIO - 1], g0), fin = dataOra_(r[C.FINE - 1], g0);
    if (!ini || !fin) return;
    const giorno = g0 || ini;
    if (giorno < f.from || giorno >= f.toEnd) return;
    if (f.area) {
      const a = f.aree[String(r[C.EQ - 1]).trim()] || f.aree[kpiChiave_(r[C.LINEA - 1], r[C.MACC - 1])] || '';
      if (!a) {
        const g = dataOra_(r[C.DATA - 1], null) || dataOra_(r[C.INIZIO - 1], null);
        senzaArea.push((g ? ('0' + g.getDate()).slice(-2) + '/' + ('0' + (g.getMonth() + 1)).slice(-2) + '/' + g.getFullYear() + ' ' : '') +
          [r[C.LINEA - 1], r[C.MACC - 1]].map(x => String(x).trim()).filter(Boolean).join(' · ') +
          (String(r[C.EQ - 1]).trim() ? ' (codice ' + String(r[C.EQ - 1]).trim() + ')' : ' (senza codice)'));
        return;
      }
      if (a !== f.area) return;
    }
    if (f.linea !== 'TUTTE' && String(r[C.LINEA - 1]) !== f.linea) return;
    if (f.code && String(r[C.EQ - 1]) !== f.code) return;
    let dur = (fin - ini) / 864e5; if (dur < 0) dur = ((dur % 1) + 1) % 1; // ore scritte senza data
    rows.push({ lotto: String(r[C.LOTTO - 1]).trim(), linea: String(r[C.LINEA - 1]), macc: String(r[C.MACC - 1]),
      guasto: String(r[C.GUASTO - 1]).trim() || '(vuoto)', dur: dur,
      g: String((f.gruppi && f.gruppi[ix]) || '').trim() || 'r' + ix,
      causa: String((f.cause && f.cause[ix]) || '').trim() || '(non indicata)', int: (u => u === 'RIAVVIO' ? 'RIPRISTINO' : u || '(vuoto)')(String(r[C.INT - 1]).trim().toUpperCase()) });
  });
  const n = rows.length, tot = rows.reduce((a, x) => a + x.dur, 0), nF = new Set(rows.map(x => x.g)).size; // n = interventi, nF = fermi (macchina ferma)
  const lotti = new Set(rows.map(x => x.lotto).filter(Boolean)).size;
  const grp = (key) => { const m = {}; rows.forEach(x => { const k = key(x); m[k] = m[k] || { n: 0, t: 0 }; m[k].n++; m[k].t += x.dur; }); return m; };
  const g = grp(x => intGuasto_(x.int, x.guasto));
  let cum = 0;
  const pareto = Object.keys(g).map(k => [k, g[k].n, g[k].t]).sort((a, b) => b[1] - a[1] || b[2] - a[2])
    .map(r => { cum += r[1]; return [r[0], r[1], r[2], cum / n, r[1] / n]; });
  let cumT = 0;
  const paretoT = Object.keys(g).map(k => [k, g[k].t * 24, g[k].n]).sort((a, b) => b[1] - a[1] || b[2] - a[2])
    .map(r => { cumT += r[1]; return [r[0], r[1], r[2], tot ? cumT / (tot * 24) : 0, tot ? r[1] / (tot * 24) : 0]; });
  const mm = {};
  rows.forEach(x => {
    const k = x.linea + '\u0001' + x.macc;
    const m = mm[k] = mm[k] || { n: 0, t: 0, g: {}, f: new Set() };
    const kg = intGuasto_(x.int, x.guasto); m.n++; m.f.add(x.g); m.t += x.dur; m.g[kg] = (m.g[kg] || 0) + 1;
  });
  const top1 = (gg, tn) => { const k = Object.keys(gg).sort((a, b) => gg[b] - gg[a])[0]; return k ? k + ' · ' + Math.round(gg[k] / tn * 100) + '%' : ''; };
  const macchine = Object.keys(mm).map(k => { const p = k.split('\u0001'), m = mm[k]; return [p[0], p[1], m.f.size, m.t, m.t / m.n, top1(m.g, m.n)]; })
    .sort((a, b) => b[3] - a[3]);
  const micro = rows.filter(x => x.dur * 1440 < MICRO_MIN), norm = rows.filter(x => x.dur * 1440 >= MICRO_MIN), totN = norm.reduce((a, x) => a + x.dur, 0);
  const tab = (key, conT) => { const m = grp(key); return Object.keys(m).map(k => conT ? [k, m[k].n, m[k].t * 24, m[k].n / n, tot ? m[k].t / tot : 0] : [k, m[k].n, m[k].t * 24, m[k].n / n])
    .sort((a, b) => (a[0].charAt(0) === '(') - (b[0].charAt(0) === '(') || b[2] - a[2]); };
  const oreOk = f.ore !== '' && f.ore != null && !isNaN(Number(f.ore)) && Number(f.ore) > 0;
  return { n: nF, nInt: n, durFermo: nF ? tot / nF : '', tot: tot, mttr: norm.length ? totN / norm.length : '', lotti: lotti, occ: lotti ? nF / lotti : '', micro: micro.length,
    perCausa: n ? tab(x => x.causa, true) : [], perInt: n ? tab(x => x.int, false) : [],
    mtbf: (nF && oreOk) ? Number(f.ore) / nF : '',
    top: pareto.length ? pareto[0][0] + ' · ' + Math.round(pareto[0][4] * 100) + '%' : '',
    pareto: pareto, paretoT: paretoT, macchine: macchine, senzaArea: senzaArea };
}

function aggiornaValoriKpi() {
  try { completaFermiCarta_(); } catch (e) { Logger.log(e); } // righe ricopiate a mano: contano anche loro
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_KPI);
  if (!sh) return;
  // foglio KPI della versione precedente: lo ricrea tenendo filtri e ore di funzionamento
  if (!KPI_NUOVO && (sh.getRange(KPI_T, 7).getValue() !== KPI_GH || sh.getRange('A2').getValue() !== 'Area' || sh.getRange(KPI_T, 20).getValue() !== 'Grafico %: guasto' ||
      sh.getMaxColumns() < KPI_C2 + 4 || sh.getRange(KPI_T, KPI_C1).getValue() !== 'Causa' || sh.getRange('J9').getValue() !== 'N° interventi')) {
    const area = KPI_AREE.indexOf(String(sh.getRange('B2').getValue()).toUpperCase()) >= 0 ? String(sh.getRange('B2').getValue()).toUpperCase() : '';
    const filtri = sh.getRange('B3:B5').getValues(), date = sh.getRange('D6:D7').getValues(), ore = sh.getRange('F10').getValue();
    ss.deleteSheet(sh);
    buildKpi_(ss, { area: area, filtri: filtri, date: date, ore: ore });
    return;
  }
  // menu Linea e Macchina dal foglio Equipment
  const eqAll = equipment_(ss);
  let area = String(sh.getRange('B2').getValue()).toUpperCase();
  if (KPI_AREE.indexOf(area) < 0) { area = 'PACKAGING'; sh.getRange('B2').setValue(area); }
  const eq = eqAll.filter(e => e.area.toUpperCase() === area);
  const linee = ['TUTTE'].concat(Array.from(new Set(eq.map(e => e.linea))).sort());
  let linea = String(sh.getRange('B3').getValue() || 'TUTTE');
  if (linee.indexOf(linea) < 0) { linea = 'TUTTE'; sh.getRange('B3').setValue(linea); }
  sh.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(linee, true).setAllowInvalid(false).build());
  const macc = ['TUTTE'].concat(eq.filter(e => linea === 'TUTTE' || e.linea === linea).map(e => e.code + ' · ' + e.macchina));
  let mSel = String(sh.getRange('B4').getValue() || 'TUTTE');
  if (macc.indexOf(mSel) < 0) { mSel = 'TUTTE'; sh.getRange('B4').setValue(mSel); }
  sh.getRange('B4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(macc.slice(0, 500), true).setAllowInvalid(false).build());

  const p = periodo_(String(sh.getRange('B5').getValue()), sh.getRange('D6').getValue(), sh.getRange('D7').getValue(), new Date());
  sh.getRange('B6:B7').setValues([[p.from], [p.to]]);

  const fsh = ss.getSheetByName(SH_FERMI), last = fsh.getLastRow();
  const vals = last > 1 ? fsh.getRange(2, 1, last - 1, HEAD.length).getValues() : [];
  // area di ogni fermo: anche le macchine segnate NO (non più attive) in Equipment
  const aree = {};
  ss.getSheetByName(SH_EQ).getDataRange().getValues().slice(1).forEach(r => {
    const a = String(r[2]).trim().toUpperCase(); if (!r[0] || !a) return;
    aree[String(r[0]).trim()] = a; aree[kpiChiave_(r[1], r[3])] = a;
  });
  const cause = last > 1 && fsh.getMaxColumns() >= C.CAUSA ? fsh.getRange(2, C.CAUSA, last - 1, 1).getValues().map(x => x[0]) : [];
  const oreIn = sh.getRange('F10').getValue(), manuali = oreIn !== '' && oreIn != null && !isNaN(Number(oreIn)) && Number(oreIn) > 0;
  let ore = manuali ? Number(oreIn) : '', stima = null;
  if (!manuali) { try { stima = oreStimate_(ss, p, area, linea); ore = stima.ore || ''; } catch (e) { Logger.log(e); } }
  const k = calcolaKpi_(vals, { from: p.from, toEnd: p.toEnd, linea: linea, area: area, aree: aree, code: mSel === 'TUTTE' ? '' : mSel.split(' · ')[0],
    ore: ore, cause: cause, gruppi: gruppiCol_(fsh, last) });
  sh.getRange('F11').setValue(manuali ? 'ore scritte a mano (cancella la cella per tornare alla stima)'
    : stima && stima.ore ? 'stima: ' + stima.giorni + ' giorni con lotti nel Programma × ' + ORE_GIORNO + ' h = ' + stima.ore + ' h' + (stima.linee > 1 ? ' (' + stima.linee + ' linee)' : '') + '. Scrivi qui le ore vere se le hai.'
    : 'nessun lotto nel Programma per il periodo: scrivi qui le ore di funzionamento').setFontColor('#5a6873').setFontSize(9).setWrap(true);

  sh.getRange('A10:E10').setValues([[k.n, k.tot, k.mttr, k.lotti, k.occ]]);
  const ex = k.senzaArea.length;
  sh.getRange('A11').setValue(ex ? '⚠ ' + (ex === 1 ? '1 fermo escluso' : ex + ' fermi esclusi') + ', macchina non trovata in Equipment: ' +
    k.senzaArea.slice(0, 5).join('; ') + (ex > 5 ? '; …' : '') + '. Correggi Linea/Macchina nel foglio Fermi oppure aggiungi la macchina in Equipment.' : '')
    .setFontColor('#a3402d').setFontSize(9).setWrap(false);
  sh.getRange('G10:K10').setValues([[k.mtbf, k.top, k.micro, k.nInt, k.durFermo]]);
  sh.getRange('K10').setNumberFormat('[h]:mm');
  sh.getRange('J11').setValue(k.nInt > k.n ? 'più interventi nello stesso fermo: ' + (k.nInt - k.n) + ' in più dei fermi' : '').setFontColor('#5a6873').setFontSize(9).setWrap(true);
  sh.getRange('I11').setValue(k.micro ? 'contano nel n° fermi e nei Pareto, non nell\'MTTR: controlla che gli orari siano giusti' : '').setFontColor('#5a6873').setFontSize(9).setWrap(true);
  sh.getRange('B10:C10').setNumberFormat('[h]:mm'); sh.getRange('E10').setNumberFormat('0.00'); sh.getRange('G10').setNumberFormat('0.0');

  const t = KPI_T + 1;
  sh.getRange(t, 1, 300, 25).clearContent();
  const grafico = (rows, iv, ic) => {
    if (rows.length <= KPI_N) return rows.map(r => [r[0], r[iv], r[ic]]);
    const top = rows.slice(0, KPI_N - 1).map(r => [r[0], r[iv], r[ic]]);
    const resto = rows.slice(KPI_N - 1);
    return top.concat([['ALTRI (' + resto.length + ' tipi)', resto.reduce((a, r) => a + r[iv], 0), 1]]);
  };
  if (k.pareto.length) {
    sh.getRange(t, 1, k.pareto.length, 5).setValues(k.pareto);
    sh.getRange(t, 3, k.pareto.length, 1).setNumberFormat('[h]:mm');
    sh.getRange(t, 4, k.pareto.length, 2).setNumberFormat('0%');
    sh.getRange(t, 7, k.paretoT.length, 5).setValues(k.paretoT);
    sh.getRange(t, 8, k.paretoT.length, 1).setNumberFormat('0.0');
    sh.getRange(t, 10, k.paretoT.length, 2).setNumberFormat('0%');
    const gf = grafico(k.pareto, 4, 3), gd = grafico(k.paretoT, 4, 3);
    sh.getRange(t, 20, gf.length, 3).setValues(gf).setNumberFormat('0%');
    sh.getRange(t, 23, gd.length, 3).setValues(gd).setNumberFormat('0%');
  } else sh.getRange(t, 1).setValue('Nessun fermo chiuso con questi filtri');
  if (k.macchine.length) {
    sh.getRange(t, 13, k.macchine.length, 6).setValues(k.macchine);
    sh.getRange(t, 16, k.macchine.length, 2).setNumberFormat('[h]:mm');
  }
  sh.getRange(t, KPI_C1, 300, KPI_C2 + 4 - KPI_C1).clearContent();
  if (k.perCausa.length) { sh.getRange(t, KPI_C1, k.perCausa.length, 5).setValues(k.perCausa); sh.getRange(t, KPI_C1 + 2, k.perCausa.length, 1).setNumberFormat('0.0'); sh.getRange(t, KPI_C1 + 3, k.perCausa.length, 2).setNumberFormat('0%'); }
  if (k.perInt.length) { sh.getRange(t, KPI_C2, k.perInt.length, 4).setValues(k.perInt); sh.getRange(t, KPI_C2 + 2, k.perInt.length, 1).setNumberFormat('0.0'); sh.getRange(t, KPI_C2 + 3, k.perInt.length, 1).setNumberFormat('0%'); }
}

/* =================== SOLUZIONI, DOCUMENTI, ANOMALIE =================== */
// Colonna "Dove": codice macchina (es. ST3-04), nome linea (es. ST3) oppure TUTTE.
const SH_SOL = 'Soluzioni', SH_DOC = 'Documenti', SH_ANO = 'Anomalie', SH_IMP = 'Import_PharmaFix', SH_MAPPA = 'Mappa_import';
const SOL_H = ['ID', 'Data', 'Dove (codice / linea / TUTTE)', 'Linea', 'Macchina', 'Tipo guasto', 'Problema', 'Soluzione', 'Codici ricambio', 'Allegato', 'Autore',
  'Stato (VERIFICATA / DA APPROVARE / SCARTATA)', 'Approvata da', 'Motivo (se scartata)'];
// Soluzioni: le nuove dei tecnici restano DA APPROVARE finché un coordinatore non le approva (vuoto = VERIFICATA, per le righe già esistenti)
function solColonne_(sh) {
  if (sh.getMaxColumns() < SOL_H.length) sh.insertColumnsAfter(sh.getMaxColumns(), SOL_H.length - sh.getMaxColumns());
  if (String(sh.getRange(1, SOL_H.length).getValue()) !== SOL_H[SOL_H.length - 1]) {
    const cSt = SOL_H.length - 2; // Stato, Approvata da, Motivo
    sh.getRange(1, cSt, 1, 3).setValues([SOL_H.slice(-3)]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
    sh.getRange(2, cSt, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['VERIFICATA', 'DA APPROVARE', 'SCARTATA'], true).setAllowInvalid(true).build());
  }
}
function coordinatori_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LISTE);
  const leggi = (fresco) => { const lv = fresco ? sh.getDataRange().getValues() : listeVal_(); // tutte le colonne che si chiamano Coordinatori / Coordinatore (anche con spazi)
    const cols = lv[0].map((h, i) => /^\s*coordinator[ei]\s*$/i.test(String(h)) ? i : -1).filter(i => i >= 0);
    return { lv: lv, cols: cols }; };
  let x = leggi();
  if (!x.cols.length) { // colonna creata da sola la prima volta (con il lucchetto: niente doppioni se più telefoni aprono l'app insieme)
    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); x = leggi(true); if (!x.cols.length) { sh.getRange(1, x.lv[0].length + 1).setValue('Coordinatori').setFontWeight('bold'); LISTE_MEMO = null; } }
    catch (e) { Logger.log(e); } finally { try { lock.releaseLock(); } catch (e) {} }
    return [];
  }
  return [...new Set(x.lv.slice(1).flatMap(r => x.cols.map(c => String(r[c]).trim())).filter(Boolean))];
}
// Tecnici = colonna Manutentori + coordinatori (anche loro intervengono e ricevono compiti); esclusi solo dai KPI Risorse
function tecniciTutti_(ss) {
  const lv = listeVal_(ss), c = lv[0].indexOf('Manutentori');
  const man = c < 0 ? [] : lv.slice(1).map(r => String(r[c]).trim()).filter(Boolean);
  return [...new Set(man.concat(coordinatori_()))];
}
const isCoordinatore_ = nome => coordinatori_().some(c => c.toUpperCase() === String(nome || '').trim().toUpperCase());
const DOC_H = ['ID', 'Data', 'Dove (codice / linea / TUTTE)', 'Tipo (MANUALE / SCHEMA / SCHEDA FORMATO / ALTRO)', 'Titolo', 'Descrizione', 'Link', 'Autore'];
const ANO_H = ['ID', 'Data', 'Equipment', 'Linea', 'Macchina', 'Urgenza (1 critico, 2 urgente, 3 programmabile)', 'Descrizione', 'Cosa serve', 'Foto',
  'Segnalata da', 'Stato (APERTA / IN CORSO / CHIUSA)', 'Gestita da', 'Data chiusura', 'Note'];

function creaFogliArchivio_(ss) {
  [[SH_SOL, SOL_H, [7, 8]], [SH_DOC, DOC_H, [5, 6]], [SH_ANO, ANO_H, [7, 8]]].forEach(([n, h, wide]) => {
    if (ss.getSheetByName(n)) return;
    const sh = ss.insertSheet(n);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
    sh.setFrozenRows(1);
    wide.forEach(c => sh.setColumnWidth(c, 320));
    sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm');
  });
}

function righe_(ss, nome, n) {
  const sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues().filter(r => r[0] !== '' || r[2] !== '');
}
const ts_ = v => v instanceof Date ? v.getTime() : (dataOra_(v, null) ? dataOra_(v, null).getTime() : null);

function archivio_(ss) {
  const solSh = ss.getSheetByName(SH_SOL); if (solSh) solColonne_(solSh);
  coordinatori_(); // crea la colonna Coordinatori nel foglio Liste se manca
  const tutte = righe_(ss, SH_SOL, SOL_H.length).map(r => ({ id: String(r[0]), ts: ts_(r[1]), dove: String(r[2]).trim(), linea: String(r[3]), macchina: String(r[4]),
    guasto: String(r[5]), problema: String(r[6]), soluzione: String(r[7]), ricambi: String(r[8]), link: String(r[9]), autore: String(r[10]),
    stato: String(r[11] || '').trim().toUpperCase() || 'VERIFICATA', da: String(r[12] || '').split(' · ')[0], motivo: String(r[13] || '').trim() }));
  const sol = tutte.filter(x => x.stato !== 'SCARTATA').map(x => { delete x.motivo; delete x.da; return x; });
  // scartate degli ultimi 90 giorni: l'app le mostra solo a chi le ha proposte, con il motivo
  const da90 = Date.now() - 90 * 864e5;
  const solScartate = tutte.filter(x => x.stato === 'SCARTATA' && (x.ts || 0) >= da90).map(x => ({ id: x.id, ts: x.ts, dove: x.dove, macchina: x.macchina, guasto: x.guasto, problema: x.problema, autore: x.autore, motivo: x.motivo, da: x.da }));
  try { docElenchi_(ss); } catch (e) { Logger.log(e); }
  try { fermiConvalida_(ss); } catch (e) { Logger.log(e); }
  const dSh = ss.getSheetByName(SH_DOC);
  const doc = (dSh && dSh.getLastRow() > 1 ? dSh.getRange(2, 1, dSh.getLastRow() - 1, DOC_H.length).getValues() : [])
    .filter(r => String(r[4]).trim() || String(r[6]).trim()) // basta titolo o link: ID e Dove possono restare vuoti
    .map(r => ({ id: String(r[0]), ts: ts_(r[1]), dove: String(r[2]).trim() || 'TUTTE', tipo: String(r[3]).trim().toUpperCase() || 'MANUALE',
      titolo: String(r[4]).trim() || 'Documento', descr: String(r[5]), link: String(r[6]).trim() }));
  const from = Date.now() - 30 * 864e5;
  const ano = righe_(ss, SH_ANO, ANO_H.length).map(r => ({ id: String(r[0]), ts: ts_(r[1]), code: String(r[2]), linea: String(r[3]), macchina: String(r[4]),
    urgenza: Number(r[5]) || 3, descr: String(r[6]), serve: String(r[7]), foto: String(r[8]), da: String(r[9]), stato: String(r[10] || 'APERTA').toUpperCase(),
    gestita: String(r[11]), chiusa: ts_(r[12]), note: String(r[13]) }))
    .filter(a => a.stato !== 'CHIUSA' || (a.chiusa || a.ts || 0) >= from);
  return { soluzioni: sol, solScartate: solScartate, documenti: doc, anomalie: ano };
}

/* Documenti: nella colonna "Dove" si sceglie da un elenco (con suggerimenti mentre scrivi): TUTTE, le linee e i codici
   del foglio Equipment (solo attivi). Vuoto = TUTTE. L'elenco sta nel foglio nascosto Elenco_Dove e si aggiorna da solo
   quando cambia Equipment (all'apertura dell'app). Anche "Tipo" ha il suo elenco. */
const SH_ELENCO = 'Elenco_Dove', DOC_TIPI = ['MANUALE', 'SCHEMA', 'SCHEDA FORMATO', 'ALTRO'];
function docDove_(ss) {
  const eq = equipment_(ss);
  const linee = [...new Set(eq.map(e => e.linea.trim()).filter(Boolean))].sort();
  return ['TUTTE'].concat(linee, eq.map(e => e.code).sort());
}
function docElenchi_(ss, forza) {
  const props = PropertiesService.getScriptProperties();
  const elenco = docDove_(ss), firma = 'v1|' + elenco.join('|');
  let sh = ss.getSheetByName(SH_DOC);
  if (!sh) { creaFogliArchivio_(ss); sh = ss.getSheetByName(SH_DOC); }
  if (!forza && props.getProperty('DOC_ELENCO') === firma && ss.getSheetByName(SH_ELENCO)) return false;
  let el = ss.getSheetByName(SH_ELENCO);
  if (!el) { el = ss.insertSheet(SH_ELENCO); el.hideSheet(); }
  el.clearContents();
  el.getRange(1, 1).setValue('Dove (generato dallo script dal foglio Equipment: non modificare)');
  el.getRange(2, 1, elenco.length, 1).setNumberFormat('@').setValues(elenco.map(x => [x]));
  const n = Math.max(sh.getMaxRows() - 1, 999);
  sh.getRange(2, 3, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(el.getRange(2, 1, elenco.length, 1), true).setAllowInvalid(false)
    .setHelpText('Scegli TUTTE, una linea o un codice macchina del foglio Equipment. Vuoto = TUTTE.').build());
  sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(DOC_TIPI, true).setAllowInvalid(false).build());
  props.setProperty('DOC_ELENCO', firma);
  return true;
}
// Menu: rifà subito gli elenchi (se hai cambiato Equipment e vuoi vederli senza aprire l'app)
function aggiornaElenchiDocumenti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  docElenchi_(ss, true);
  ss.toast('Elenchi del foglio Documenti aggiornati: ' + docDove_(ss).length + ' voci in "Dove".', 'Pharma-Fix', 6);
}

/* Foglio Fermi: menu a tendina (convalida dati) per chi scrive o corregge a mano, come nel vecchio template.
   E Manutentore, G Categoria, H Tipo intervento → colonne del foglio Liste (solo valori dell'elenco);
   F Macchina → macchine attive del foglio Equipment; I Tipo guasto → tutti i tipi del foglio Guasti (TUTTE + macchine).
   Su F e I un valore fuori elenco è solo segnalato (triangolino), perché i fermi da carta hanno nomi liberi.
   Macchine e guasti stanno nel foglio nascosto Elenchi_Fermi, che si aggiorna da solo quando cambiano Equipment o Guasti. */
const SH_EL_FERMI = 'Elenchi_Fermi';

/* ---------- Classificazione dei fermi (dal 17/10/2026) ----------
   Il tecnico sceglie: 1) cosa si è guastato (componente/zona, foglio Guasti) 2) cosa ha fatto (4 interventi con regola fissa)
   3) perché (facoltativo). La CATEGORIA non la sceglie più: sta nel foglio "Voci guasto" (una riga per voce, la decide il responsabile). */
const INTERVENTI = ['REGOLAZIONE', 'SOSTITUZIONE', 'RIPRISTINO', 'PULIZIA'];
const CAUSE = ['USURA', 'FORMATO / REGOLAZIONE PERSA', 'MATERIALE FUORI SPEC.', 'ERRORE OPERATORE', 'ELETTRICA / ALIMENTAZIONE', 'NON NOTA'];
const INT_DESC = { REGOLAZIONE: 'spostato o registrato (nonio, guida, parametro)', SOSTITUZIONE: 'pezzo nuovo montato · codice obbligatorio',
  RIPRISTINO: 'reset, riavvio, sblocco, homing: niente regolato né cambiato', PULIZIA: 'pulito, tolto residui o inceppamenti' };
const LI_DESC = 'Tipo intervento · spiegazione', LI_CAUSA = 'Causa';
const NON_VOCI = ['FORMATO', 'USURA', 'SPIKE CORRENTE', 'RIPARAZIONE', 'SOSTITUZIONE', 'PULIZIA', 'RIAVVIO', 'REGOLAZIONE']; // cause o interventi, non componenti
const SH_VOCI = 'Voci guasto', VOCI_H = ['Voce (come nel foglio Guasti)', 'Categoria', 'Da dove viene (giallo = da controllare)', 'Quando si usa (la legge il tecnico)'];
const DEF_QUANDO = { 'INCASTRO': 'prodotto o astuccio bloccato nel trasporto, nessun pezzo rotto', 'SENSORE': 'sensore, fotocellula o fibra che non legge o legge male',
  'ARIA COMPRESSA': 'pressione bassa, perdite, tubi o raccordi dell\'aria', 'SICUREZZE / EMERGENZE': 'porte, micro, barriere o moduli di sicurezza che fermano la macchina',
  'PARAMETRI PANNELLO': 'valori del pannello o della ricetta da correggere' };
let VDESC_MEMO = null;
let VOCI_MEMO = null;
// Colonne in fondo al foglio Fermi (fuori da HEAD): Causa, Voce proposta (con ALTRO), Rivisto (coordinatore),
// e per i fermi con più interventi: Fermo n. (uguale per tutti gli interventi dello stesso fermo), Intervento n., Più interventi?
const EXTRA_H = ['Causa', 'Voce proposta (con ALTRO)', 'Rivisto (coordinatore)', 'Fermo n.', 'Intervento n.', 'Più interventi?', 'Altre foto chiusura'];
function causaCol_(sh) {
  const ultima = C.CAUSA + EXTRA_H.length - 1;
  if (sh.getMaxColumns() < ultima) sh.insertColumnsAfter(sh.getMaxColumns(), ultima - sh.getMaxColumns());
  const h = sh.getRange(1, C.CAUSA, 1, EXTRA_H.length), v = h.getValues()[0];
  if (v.join('|') === EXTRA_H.join('|')) return true;
  // AB–AD (Fermo n., Intervento n., Più interventi?) non sovrascrivono colonne già usate per altro
  const libere = v.slice(3).every((x, i) => !String(x).trim() || String(x).trim() === EXTRA_H[3 + i]);
  sh.getRange(1, C.CAUSA, 1, libere ? EXTRA_H.length : 3).setValues([libere ? EXTRA_H : EXTRA_H.slice(0, 3)]).setFontWeight('bold').setBackground('#fff3d6');
  if (!libere) Logger.log('Foglio Fermi: colonne AB–AD già usate, Fermo n./Intervento n. non scritti');
  return libere;
}
// Descrizione che racconta più operazioni (stessa regola dell'app): due azioni unite da "e" o dalla virgola, oppure tipi d'intervento diversi
const AZIONI_ = /\b(REGOLAT|SOSTITUIT|CAMBIAT|MONTAT|PULIT|RESETTAT|RIAVVIAT|SBLOCCAT|RIMESS|SISTEMAT|SPOSTAT|RIPRISTINAT|ALLINEAT|CENTRAT|FASAT|SERRAT|ABBASSAT|TARAT|RADDRIZZAT|RIMOSS|AGGIUNT|RIFATT)[OAIE]\b|\bALZAT[OI]\b/g;
function piuOperazioni_(txt) {
  const d = ' ' + String(txt || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[.,!]/g, m => m === ',' ? ' , ' : ' ').replace(/\s+/g, ' ') + ' ';
  const tipi = [/SOSTITU|CAMBIAT|MONTAT/, /REGOLA|REGISTRAT|CENTRAT|ALLINEAT/, /PULIT|PULIZ/, /RESET|RIAVVI|SBLOCC|HOMING|RIPRISTIN/].filter(re => re.test(d)).length;
  if (tipi >= 2) return 'la descrizione parla di interventi di tipo diverso';
  const az = []; let m; AZIONI_.lastIndex = 0; while ((m = AZIONI_.exec(d))) az.push(m);
  for (let i = 1; i < az.length; i++) if (/ E | ED | , | POI /.test(d.slice(az[i - 1].index, az[i].index)) && !/TROVAT\w* $/.test(d.slice(0, az[i - 1].index))) return 'la descrizione racconta più operazioni';
  return '';
}
// Colonna "Fermo n." (solo se è nostra): un valore per riga, '' se il fermo ha un solo intervento
function gruppiCol_(sh, last) {
  if (last < 2 || sh.getMaxColumns() < C.GRUPPO || String(sh.getRange(1, C.GRUPPO).getValue()).trim() !== EXTRA_H[3]) return [];
  return sh.getRange(2, C.GRUPPO, last - 1, 1).getValues().map(x => String(x[0] || '').trim());
}
// Fermo con più interventi: ogni intervento è una riga; "Fermo n." lega le righe dello stesso fermo
function gruppoFermo_(sh, r, p) {
  const g = String(p.gruppo || '').trim();
  if (!g) return;
  const giro = Math.max(1, Math.round(Number(p.giro) || 1));
  if (causaCol_(sh)) sh.getRange(r, C.GRUPPO, 1, 2).setValues([[g, giro]]);
}
// Categoria proposta per una voce: quella più usata nello storico (almeno 3 fermi), altrimenti dal nome
function catProposta_(voce, storico) {
  const c = storico[voce];
  if (c) { const t = Object.keys(c).filter(k => k && k !== 'ALTRO').sort((a, b) => c[b] - c[a])[0], n = Object.values(c).reduce((a, b) => a + b, 0);
    if (t && n >= 3) return [t, 'storico: ' + c[t] + ' fermi su ' + n]; }
  if (/CAM\.|VIDEOJET|PARAM|PANNELLO|RICETTA|STAMPA DATI|SEA VISION|COGNEX|ANTARES|CHECKWEIGHER|ZONA ROBOT|VELOCIT/.test(voce)) return ['AUTOMATION', 'dal nome'];
  if (/SENSOR|FOTOCELL|FINE CORSA|SICUREZZ|SELETTOR|TEMPERATUR|LETTURA|H10|BARRIERE|FIBRA/.test(voce)) return ['ELETTRICO', 'dal nome'];
  return ['MECCANICO', 'dal nome'];
}
// { VOCE: "quando si usa" } dal foglio Voci guasto
function vociDesc_(ss) {
  if (VDESC_MEMO) return VDESC_MEMO;
  const sh = ss.getSheetByName(SH_VOCI), out = {};
  if (sh && sh.getLastRow() > 1 && sh.getMaxColumns() >= 4) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => { const v = String(r[0]).trim().toUpperCase(), d = String(r[3]).trim(); if (v && d) out[v] = d; });
  return VDESC_MEMO = out;
}
// { VOCE: CATEGORIA } dal foglio Voci guasto
function vociCat_(ss) {
  if (VOCI_MEMO) return VOCI_MEMO;
  const sh = ss.getSheetByName(SH_VOCI), out = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { const v = String(r[0]).trim().toUpperCase(), c = String(r[1]).trim().toUpperCase(); if (v && c) out[v] = c; });
  return VOCI_MEMO = out;
}
// Una volta: Tipo intervento a 4 voci, RIAVVIO → RIPRISTINO nello storico, cause tolte da TUTTE, colonna Causa.
// Ogni volta (se cambia il foglio Guasti): le voci nuove entrano nel foglio Voci guasto con la categoria proposta.
function classificazione_(ss) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('CLASS_V1') !== '1') {
    const f = ss.getSheetByName(SH_FERMI);
    if (f) {
      causaCol_(f);
      if (f.getLastRow() > 1) { const rg = f.getRange(2, C.INT, f.getLastRow() - 1, 1), v = rg.getValues();
        if (v.some(x => String(x[0]).trim().toUpperCase() === 'RIAVVIO')) rg.setValues(v.map(x => [String(x[0]).trim().toUpperCase() === 'RIAVVIO' ? 'RIPRISTINO' : x[0]])); }
    }
    const li = ss.getSheetByName(SH_LISTE);
    if (li) { const testa = li.getRange(1, 1, 1, li.getLastColumn()).getValues()[0], c = testa.findIndex(h => String(h).trim().toUpperCase() === 'TIPO INTERVENTO') + 1;
      if (c) { li.getRange(2, c, Math.max(li.getLastRow() - 1, INTERVENTI.length), 1).clearContent(); li.getRange(2, c, INTERVENTI.length, 1).setValues(INTERVENTI.map(x => [x])); LISTE_MEMO = null; } }
    const g = ss.getSheetByName(SH_GUASTI);
    if (g && g.getLastRow() > 1) { const v = g.getDataRange().getValues(), i = v.findIndex(r => String(r[0]).trim().toUpperCase() === 'TUTTE');
      if (i > 0) { const resto = v[i].slice(2).filter(x => String(x).trim() && NON_VOCI.indexOf(String(x).trim().toUpperCase()) < 0);
        g.getRange(i + 1, 3, 1, v[0].length - 2).clearContent(); if (resto.length) g.getRange(i + 1, 3, 1, resto.length).setValues([resto]); } }
    props.setProperty('CLASS_V1', '1');
  }
  // V2: le liste di intervento (con la spiegazione per il tecnico) e di cause stanno nel foglio Liste, modificabili dal responsabile
  if (props.getProperty('CLASS_V2') !== '1') {
    const li = ss.getSheetByName(SH_LISTE);
    if (li) {
      const v = li.getDataRange().getValues(), testa = v[0].map(h => String(h).trim().toUpperCase());
      const cInt = testa.indexOf('TIPO INTERVENTO');
      let col = testa.length;
      if (cInt >= 0 && testa.indexOf(LI_DESC.toUpperCase()) < 0) {
        col++; li.getRange(1, col).setValue(LI_DESC).setFontWeight('bold');
        const righe = v.slice(1).map(r => [INT_DESC[String(r[cInt]).trim().toUpperCase()] || '']);
        if (righe.length) li.getRange(2, col, righe.length, 1).setValues(righe);
      }
      if (testa.indexOf(LI_CAUSA.toUpperCase()) < 0) {
        col++; li.getRange(1, col).setValue(LI_CAUSA).setFontWeight('bold');
        li.getRange(2, col, CAUSE.length, 1).setValues(CAUSE.map(x => [x]));
      }
      LISTE_MEMO = null;
    }
    props.setProperty('CLASS_V2', '1');
  }
  // V3: colonna "Quando si usa" nel foglio Voci guasto (definizione di ogni voce, la legge il tecnico)
  if (props.getProperty('CLASS_V3') !== '1') {
    const vs = ss.getSheetByName(SH_VOCI);
    if (vs) {
      if (/^Una riga per voce/.test(String(vs.getRange(1, 5).getValue()))) { vs.getRange(1, 6).setValue(vs.getRange(1, 5).getValue()).setWrap(true).setFontStyle('italic'); vs.getRange(1, 5).clearContent(); }
      if (String(vs.getRange(1, 4).getValue()).trim() !== VOCI_H[3]) { vs.getRange(1, 4).setValue(VOCI_H[3]).setFontWeight('bold').setBackground('#dfe7ec').setFontStyle('normal'); vs.setColumnWidth(4, 380); }
      if (vs.getLastRow() > 1) { const v = vs.getRange(2, 1, vs.getLastRow() - 1, 4).getValues();
        vs.getRange(2, 4, v.length, 1).setValues(v.map(r => [String(r[3]).trim() || DEF_QUANDO[String(r[0]).trim().toUpperCase()] || ''])); }
    }
    props.setProperty('CLASS_V3', '1');
  }
  // V4: categoria "A SCELTA" nel foglio Voci guasto (voci che possono essere meccaniche o elettriche: la sceglie il tecnico)
  if (props.getProperty('CLASS_V4') !== '1') {
    const vs = ss.getSheetByName(SH_VOCI);
    if (vs && vs.getMaxRows() > 1) {
      vs.getRange(2, 2, vs.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO', 'A SCELTA'], true).setAllowInvalid(false).build());
      vs.getRange(1, 2).setNote('Categoria della voce: l\'app la scrive da sola alla chiusura del fermo. A SCELTA = la voce può essere meccanica o elettrica (es. SENSORE: posizione o taratura): il tecnico deve sceglierla ogni volta.');
    }
    props.setProperty('CLASS_V4', '1');
  }
  vociSync_(ss);
}
/* =================== Unione voci guasto ===================
   Foglio Voci guasto, colonna E "Unisci in": accanto a una voce doppia scrivi la voce che resta (anche un nome nuovo = rinomina).
   Menu Pharma-Fix → "Voci guasto: applica unioni": la voce vecchia diventa quella nuova nel foglio Guasti (senza doppioni nella riga),
   nello storico del foglio Fermi (Tipo guasto) e nelle Soluzioni; la riga vecchia sparisce dal foglio Voci guasto.
   Ogni unione resta scritta nel foglio "Unioni voci" (data, da, a, righe cambiate). */
const VOCI_UNISCI = 'Unisci in (voce che resta)', SH_UNIONI = 'Unioni voci';
function unioniCol_(vs) {
  if (String(vs.getRange(1, 5).getValue()).trim() === VOCI_UNISCI) return;
  if (String(vs.getRange(1, 5).getValue()).trim()) return; // colonna E già usata per altro: non tocco
  vs.getRange(1, 5).setValue(VOCI_UNISCI).setFontWeight('bold').setBackground('#fde3c8')
    .setNote('Voce doppia o simile? Scrivi qui la voce che resta (anche un nome nuovo = rinomina), poi menu Pharma-Fix → "Voci guasto: applica unioni". Vale per foglio Guasti, storico Fermi e Soluzioni.');
  vs.setColumnWidth(5, 240);
}
function unioniLeggi_(ss) {
  const vs = ss.getSheetByName(SH_VOCI); if (!vs || vs.getLastRow() < 2 || String(vs.getRange(1, 5).getValue()).trim() !== VOCI_UNISCI) return { vs: vs, map: {}, righe: [] };
  const U = x => String(x || '').trim().toUpperCase(), v = vs.getRange(2, 1, vs.getLastRow() - 1, 5).getValues(), map = {}, righe = [];
  v.forEach((r, i) => { const da = U(r[0]), a = String(r[4]).trim(); if (da && a && U(a) !== da) { map[da] = a; righe.push(i + 2); } });
  // catene (A→B, B→C) e giri (A→B, B→A: si ferma)
  Object.keys(map).forEach(k => { let t = map[k], n = 0; while (map[U(t)] && n++ < 10 && U(map[U(t)]) !== k) t = map[U(t)]; map[k] = t; });
  return { vs: vs, map: map, righe: righe, voci: v.map(r => U(r[0])) };
}
function unioniApplica_(ss) {
  const u = unioniLeggi_(ss), map = u.map, U = x => String(x || '').trim().toUpperCase(), conta = {};
  if (!Object.keys(map).length) return null;
  // stesso nome scritto in modo diverso (maiuscole/spazi): diventa come nel foglio Voci guasto
  const giusta = {}; u.voci.forEach((k, i) => { if (k && !map[k] && !giusta[k]) giusta[k] = String(u.vs.getRange(i + 2, 1).getValue()).trim(); });
  Object.keys(map).forEach(k => { const t = U(map[k]); if (!giusta[t]) giusta[t] = String(map[k]).trim(); });
  const nuova = x => { const t = String(x).trim(), m = map[U(t)] || giusta[U(t)]; return m && m !== t ? m : null; };
  // Guasti: ogni riga, voci rinominate e senza doppioni
  const g = ss.getSheetByName(SH_GUASTI);
  if (g && g.getLastRow() > 1 && g.getLastColumn() > 2) {
    const rg = g.getRange(2, 3, g.getLastRow() - 1, g.getLastColumn() - 2), v = rg.getValues(); let cambia = false;
    const out = v.map(r => { const visti = {}, l = [];
      r.forEach(x => { const t = String(x).trim(); if (!t) return; const n = nuova(t) || t; if (n !== t) cambia = true; if (visti[U(n)]) { cambia = true; return; } visti[U(n)] = 1; l.push(n); });
      while (l.length < r.length) l.push(''); return l; });
    if (cambia) rg.setValues(out);
  }
  // Storico: Fermi (Tipo guasto) e Soluzioni (Tipo guasto)
  [[SH_FERMI, C.GUASTO], [SH_SOL, 6]].forEach(([nome, col]) => {
    const sh = ss.getSheetByName(nome); if (!sh || sh.getLastRow() < 2) return;
    const rg = sh.getRange(2, col, sh.getLastRow() - 1, 1), v = rg.getValues(); let n = 0;
    const out = v.map(r => { const t = String(r[0]).trim(), m = t && nuova(t); if (m) { n++; if (map[U(t)]) { conta[U(t)] = conta[U(t)] || {}; conta[U(t)][nome] = (conta[U(t)][nome] || 0) + 1; } return [m]; } return [r[0]]; });
    if (n) rg.setValues(out);
  });
  // Voci guasto: la riga vecchia sparisce se la voce che resta c'è già, altrimenti diventa quella nuova (rinomina)
  const del = [], prime = {};
  u.voci.forEach((k, i) => { if (!k || map[k]) return; if (prime[k]) del.push(i + 2); else prime[k] = 1; }); // righe doppie (stessa voce due volte)
  u.righe.forEach(rr => { const da = u.voci[rr - 2], a = map[da];
    if (u.voci.indexOf(U(a)) >= 0 && !map[U(a)]) del.push(rr); else { u.vs.getRange(rr, 1).setValue(a); u.vs.getRange(rr, 5).clearContent(); u.voci[rr - 2] = U(a); } });
  [...new Set(del)].sort((a, b) => b - a).forEach(rr => u.vs.deleteRow(rr));
  // Registro
  let lg = ss.getSheetByName(SH_UNIONI);
  if (!lg) { lg = ss.insertSheet(SH_UNIONI); lg.getRange(1, 1, 1, 5).setValues([['Data', 'Voce vecchia', 'Voce che resta', 'Righe Fermi cambiate', 'Righe Soluzioni cambiate']]).setFontWeight('bold').setBackground('#dfe7ec'); lg.setFrozenRows(1); }
  const ora = new Date(), log = Object.keys(map).map(k => [ora, k, map[k], (conta[k] || {})[SH_FERMI] || 0, (conta[k] || {})[SH_SOL] || 0]);
  lg.getRange(lg.getLastRow() + 1, 1, log.length, 5).setValues(log); lg.getRange(2, 1, lg.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  VOCI_MEMO = null; VDESC_MEMO = null;
  return log;
}
function applicaUnioniVoci() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi(), vs = ss.getSheetByName(SH_VOCI);
  if (vs) unioniCol_(vs);
  const u = unioniLeggi_(ss), k = Object.keys(u.map);
  if (!k.length) { ui.alert('Voci guasto', 'Nessuna unione da fare: nel foglio Voci guasto scrivi nella colonna E ("' + VOCI_UNISCI + '") la voce che resta, accanto a quella da togliere.', ui.ButtonSet.OK); return; }
  const ok = ui.alert('Unire ' + k.length + ' voci?', k.slice(0, 25).map(x => x + '  →  ' + u.map[x]).join('\n') + (k.length > 25 ? '\n…' : '') +
    '\n\nCambiano: foglio Guasti, storico Fermi (Tipo guasto) e Soluzioni. Ogni unione resta scritta nel foglio "' + SH_UNIONI + '". Fai prima un backup se vuoi tornare indietro.', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const log = unioniApplica_(ss);
  try { fermiConvalida_(ss, true); } catch (e) { Logger.log(e); }
  bumpVer_();
  ss.toast(log.length + ' voci unite. Righe Fermi cambiate: ' + log.reduce((a, r) => a + r[3], 0) + '. Dettaglio nel foglio "' + SH_UNIONI + '".', 'Pharma-Fix', 10);
}
function vociSync_(ss) {
  const g = guasti_(ss), voci = [...new Set(Object.keys(g).flatMap(k => g[k]).map(x => String(x).trim()).filter(x => x && x.toUpperCase() !== 'ALTRO' && NON_VOCI.indexOf(x.toUpperCase()) < 0))];
  let sh = ss.getSheetByName(SH_VOCI);
  const nuovo = !sh;
  if (nuovo) { sh = ss.insertSheet(SH_VOCI); sh.getRange(1, 1, 1, 4).setValues([VOCI_H]).setFontWeight('bold').setBackground('#dfe7ec'); sh.setFrozenRows(1); sh.setColumnWidth(1, 340); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 260); sh.setColumnWidth(4, 380);
    sh.getRange('F1').setValue('Una riga per voce del foglio Guasti: la categoria la decidi qui una volta e l\'app la scrive da sola alla chiusura del fermo. Le voci nuove del foglio Guasti compaiono qui da sole, in giallo.').setWrap(true).setFontStyle('italic'); sh.setColumnWidth(6, 380); }
  try { unioniCol_(sh); } catch (e) { Logger.log(e); }
  const gia = new Set(sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toUpperCase()) : []);
  const manca = voci.filter(x => !gia.has(x.toUpperCase()));
  if (!manca.length) return;
  const storico = {}, f = ss.getSheetByName(SH_FERMI);
  if (f && f.getLastRow() > 1) f.getRange(2, C.CAT, f.getLastRow() - 1, 3).getValues().forEach(r => { const v = String(r[2]).trim().toUpperCase(), c = String(r[0]).trim().toUpperCase(); if (!v) return; (storico[v] = storico[v] || {})[c] = (storico[v][c] || 0) + 1; });
  const righe = manca.map(x => { const [c, da] = catProposta_(x.toUpperCase(), storico); return [x, c, da, DEF_QUANDO[x.toUpperCase()] || '']; });
  const r0 = sh.getLastRow() + 1;
  sh.getRange(r0, 1, righe.length, 4).setValues(righe);
  righe.forEach((r, i) => sh.getRange(r0 + i, 2).setBackground(/^storico/.test(r[2]) ? null : '#fff2cc'));
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO', 'A SCELTA'], true).setAllowInvalid(false).build());
  VOCI_MEMO = null;
}
function fermiConvalida_(ss, forza) {
  const sh = ss.getSheetByName(SH_FERMI), li = ss.getSheetByName(SH_LISTE); if (!sh || !li) return false;
  const macc = [...new Set(equipment_(ss).map(e => String(e.macchina).trim()).filter(Boolean))].sort();
  const g = guasti_(ss), tipi = [...new Set(Object.keys(g).flatMap(k => g[k]).map(x => x.toUpperCase()))].sort();
  const testa = listeVal_(ss)[0].map(h => String(h).trim());
  const colL = nome => testa.findIndex(h => h.toUpperCase() === nome.toUpperCase()) + 1;
  const righe = Math.max(sh.getMaxRows(), sh.getLastRow() + 500);
  const props = PropertiesService.getScriptProperties();
  const tec = tecniciTutti_(ss).sort(); // Manutentori + Coordinatori
  const firma = 'v4|' + Math.ceil(righe / 500) + '|' + testa.join('/') + '|' + macc.join('|') + '#' + hash_(JSON.stringify(g)) + '#' + tec.join('|');
  if (!forza && props.getProperty('FERMI_DV') === firma && ss.getSheetByName(SH_EL_FERMI)) return false;
  if (sh.getMaxRows() < righe) sh.insertRowsAfter(sh.getMaxRows(), righe - sh.getMaxRows());
  let el = ss.getSheetByName(SH_EL_FERMI);
  if (!el) { el = ss.insertSheet(SH_EL_FERMI); el.hideSheet(); }
  el.clearContents();
  el.getRange(1, 1, 1, 3).setValues([['Macchine (da Equipment)', 'Tipi di guasto (da Guasti)', 'Tecnici (Manutentori + Coordinatori)']]);
  if (tec.length) el.getRange(2, 3, tec.length, 1).setNumberFormat('@').setValues(tec.map(x => [x]));
  if (macc.length) el.getRange(2, 1, macc.length, 1).setNumberFormat('@').setValues(macc.map(x => [x]));
  if (tipi.length) el.getRange(2, 2, tipi.length, 1).setNumberFormat('@').setValues(tipi.map(x => [x]));
  const dv = (rng, rigido, aiuto) => SpreadsheetApp.newDataValidation().requireValueInRange(rng, true).setAllowInvalid(!rigido).setHelpText(aiuto).build();
  const metti = (col, regola) => sh.getRange(2, col, righe - 1, 1).setDataValidation(regola);
  if (tec.length) metti(C.MAN, dv(el.getRange(2, 3, tec.length, 1), true, 'Scegli un tecnico: colonne Manutentori e Coordinatori del foglio Liste.'));
  [['Categoria', C.CAT], ['Tipo intervento', C.INT]].forEach(([nome, col]) => {
    const c = colL(nome); if (!c) return;
    metti(col, dv(li.getRange(2, c, li.getMaxRows() - 1, 1), col === C.CAT, 'Scegli dall\'elenco "' + nome + '" del foglio Liste.'));
  });
  if (macc.length) metti(C.MACC, dv(el.getRange(2, 1, macc.length, 1), false, 'Macchine del foglio Equipment.'));
  if (tipi.length) { metti(C.GUASTO, dv(el.getRange(2, 2, tipi.length, 1), false, 'Tipi di guasto del foglio Guasti.')); fermiGuastoDV_(ss, sh, 2, sh.getLastRow() - 1); }
  const cC = colL(LI_CAUSA);
  if (sh.getMaxColumns() >= C.CAUSA) metti(C.CAUSA, (cC ? SpreadsheetApp.newDataValidation().requireValueInRange(li.getRange(2, cC, li.getMaxRows() - 1, 1), true) : SpreadsheetApp.newDataValidation().requireValueInList(CAUSE, true)).setAllowInvalid(true).setHelpText('Perché si è guastato (facoltativo).').build());
  props.setProperty('FERMI_DV', firma);
  return true;
}
function hash_(t) { let h = 0; for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) | 0; return t.length + ':' + h; }
// Colonna Tipo guasto del foglio Fermi: in ogni riga solo le voci di QUELLA macchina (come nell'app: codice + linea + TUTTE + ALTRO).
// Righe senza macchina riconosciuta: tutte le voci. Si rifà quando cambiano codice, linea o macchina della riga (onEdit).
function vociMacchina_(g, eq) {
  const U = x => String(x || '').trim().toUpperCase(), out = [], visti = {};
  [...(g[U(eq.code)] || []), ...(g[U(eq.linea)] || []), ...(g.TUTTE || []), 'ALTRO'].forEach(v => { const k = U(v); if (k && !visti[k]) { visti[k] = 1; out.push(String(v).trim()); } });
  const i = out.findIndex(x => U(x) === 'ALTRO'); out.push(out.splice(i, 1)[0]);
  return out;
}
function fermiGuastoDV_(ss, sh, r0, n) {
  if (!sh || n < 1) return;
  const g = guasti_(ss), eqs = equipment_(ss), el = ss.getSheetByName(SH_EL_FERMI), tutte = el ? el.getRange(2, 2, Math.max(el.getLastRow() - 1, 1), 1) : null;
  const U = x => String(x || '').trim().toUpperCase(), regole = {};
  const tutteDV = tutte ? SpreadsheetApp.newDataValidation().requireValueInRange(tutte, true).setAllowInvalid(true).setHelpText('Macchina non riconosciuta: tutte le voci del foglio Guasti.').build() : null;
  const v = sh.getRange(r0, 1, n, C.EQ).getValues();
  const out = v.map(r => {
    const code = U(r[C.EQ - 1]), lin = U(r[C.LINEA - 1]), mac = U(r[C.MACC - 1]);
    const eq = eqs.find(e => code && U(e.code) === code) || eqs.find(e => lin && mac && U(e.linea) === lin && U(e.macchina) === mac);
    if (!eq) return [tutteDV];
    const k = U(eq.code);
    if (!regole[k]) { const l = vociMacchina_(g, eq);
      regole[k] = SpreadsheetApp.newDataValidation().requireValueInList(l, true).setAllowInvalid(true).setHelpText('Voci guasto di ' + eq.code + ' (' + eq.macchina + '): foglio Guasti.').build(); }
    return [regole[k]];
  });
  sh.getRange(r0, C.GUASTO, n, 1).setDataValidations(out);
}
// Menu: rimette subito i menu a tendina del foglio Fermi
function aggiornaTendineFermi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  fermiConvalida_(ss, true);
  ss.toast('Menu a tendina del foglio Fermi aggiornati (colonne E–I).', 'Pharma-Fix', 6);
}

// Nuovo documento dall'app (solo coordinatori): il file sta su Drive, qui arriva il link
function addDocumento_(p) {
  if (!isCoordinatore_(p.manutentore)) throw new Error('Solo i coordinatori possono aggiungere documenti');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SH_DOC);
  if (!sh) { creaFogliArchivio_(ss); sh = ss.getSheetByName(SH_DOC); }
  if (p.id && sh.getLastRow() > 1 && sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(p.id)).matchEntireCell(true).findNext()) return { dup: true };
  const elenco = docDove_(ss), dove = String(p.dove || '').trim() || 'TUTTE';
  const ok = elenco.find(x => x.toUpperCase() === dove.toUpperCase());
  if (!ok) throw new Error('"' + dove + '" non è nel foglio Equipment: scegli TUTTE, una linea o un codice macchina');
  const tipo = String(p.tipo || '').trim().toUpperCase();
  if (DOC_TIPI.indexOf(tipo) < 0) throw new Error('Tipo non valido');
  const titolo = String(p.titolo || '').trim(), link = String(p.link || '').trim();
  if (!titolo) throw new Error('Scrivi il titolo');
  if (!/^https?:\/\/\S+$/i.test(link)) throw new Error('Incolla un link valido (https://…)');
  sh.appendRow([p.id || ('D' + Date.now()), new Date(Number(p.ts) || Date.now()), ok, tipo, titolo, String(p.descr || '').trim(), link, p.manutentore]);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('dd/MM/yyyy HH:mm');
  return { ok: true };
}

function saveFile_(b64, name, mime) {
  if (!b64) return '';
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name)).getUrl();
}

function addSoluzione_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SH_SOL);
  if (!sh) throw new Error('Foglio Soluzioni mancante: esegui setup');
  if (findIn_(sh, p.id)) return { dup: true };
  if (!String(p.problema || '').trim() || !String(p.soluzione || '').trim()) throw new Error('Problema e soluzione sono obbligatori');
  const eq = equipment_(ss).find(x => x.code === p.code) || { linea: p.linea || '', macchina: '' };
  const d = new Date(Number(p.ts) || Date.now());
  const url = p.file ? saveFile_(p.file, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + (p.code || 'soluzione') + '_' + (p.fileName || 'allegato'), p.mime) : '';
  solColonne_(sh);
  const coord = isCoordinatore_(p.manutentore);
  sh.appendRow([p.id, d, p.code || p.linea || 'TUTTE', eq.linea, eq.macchina, p.guasto || '', p.problema, p.soluzione, p.ricambi || '', url, p.manutentore || '',
    coord ? 'VERIFICATA' : 'DA APPROVARE', coord ? p.manutentore : '']);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('dd/MM/yyyy HH:mm');
  if (!coord) try {
    mail_('Soluzione da approvare', 'SOLUZIONE DA APPROVARE · ' + (p.code || p.linea || 'TUTTE') + (eq.macchina ? ' ' + nice_(eq.macchina) : ''), [
      ['Nuova soluzione di ' + (p.manutentore || '?')],
      ['Macchina', (p.code || ('linea ' + (p.linea || 'TUTTE'))) + (eq.macchina ? ' · ' + nice_(eq.macchina) : '')],
      ['Guasto', p.guasto || '—'],
      ['Problema', h_(String(p.problema).slice(0, 200))],
      ['Soluzione', h_(String(p.soluzione).slice(0, 400))],
      ['Cosa fare', 'il coordinatore la approva dall\'app: Soluzioni → Da approvare']], { noMail: (c => !!(c.tgToken && c.tgChat))(avvisiCfg_()) });
  } catch (e) { Logger.log(e); }
  return { stato: coord ? 'VERIFICATA' : 'DA APPROVARE' };
}

function approvaSoluzione_(p) {
  if (!isCoordinatore_(p.manutentore)) throw new Error('Solo i coordinatori possono approvare le soluzioni');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SOL);
  solColonne_(sh);
  const r = findIn_(sh, p.id);
  if (!r) throw new Error('Soluzione non trovata');
  const no = p.esito === 'NO';
  if (no && String(p.motivo || '').trim().length < 5) throw new Error('Scrivi il motivo dello scarto');
  sh.getRange(r, SOL_H.length - 2, 1, 3).setValues([[no ? 'SCARTATA' : 'VERIFICATA',
    p.manutentore + ' · ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'), no ? String(p.motivo).trim() : '']]);
  return {};
}

function addAnomalia_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SH_ANO);
  if (!sh) throw new Error('Foglio Anomalie mancante: esegui setup');
  if (findIn_(sh, p.id)) return { dup: true };
  if (!String(p.descr || '').trim()) throw new Error('Descrizione anomalia obbligatoria');
  const eq = equipment_(ss).find(x => x.code === p.code) || { linea: '', macchina: '' };
  const d = new Date(Number(p.ts) || Date.now());
  const url = saveFile_(p.foto, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + p.code + '_anomalia.jpg', 'image/jpeg');
  sh.appendRow([p.id, d, p.code, eq.linea, eq.macchina, [1, 2, 3].indexOf(Number(p.urgenza)) >= 0 ? Number(p.urgenza) : 3,
    p.descr, p.serve || '', url, p.manutentore || '', 'APERTA', '', '', '']);
  const r = sh.getLastRow();
  sh.getRange(r, 2).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(r, 6).setBackground({ 1: '#f4c7c0', 2: '#fbeaa0', 3: '#cfe9d9' }[Number(p.urgenza)] || null);
  if (Number(p.urgenza) === 1) avvisoAnomalia_(p, eq, url);
  return {};
}

function statoAnomalia_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ANO);
  const r = findIn_(sh, p.id);
  if (!r) throw new Error('Anomalia ' + p.id + ' non trovata');
  const stato = String(p.stato || '').toUpperCase();
  if (['APERTA', 'IN CORSO', 'CHIUSA'].indexOf(stato) < 0) throw new Error('Stato non valido');
  sh.getRange(r, 11, 1, 2).setValues([[stato, p.manutentore || '']]);
  if (stato === 'CHIUSA') sh.getRange(r, 13).setValue(new Date(Number(p.ts) || Date.now())).setNumberFormat('dd/MM/yyyy HH:mm');
  if (p.note) sh.getRange(r, 14).setValue(p.note);
  return {};
}

function findIn_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const c = sh.getRange(2, 1, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}

/* ---------- Import dalla vecchia Pharma-Fix ----------
   1) Nel foglio della vecchia Pharma-Fix seleziona tutto (intestazioni comprese) e copia.
   2) Qui crea una scheda "Import_PharmaFix" e incolla in A1.
   3) Esegui importaPharmaFix(): la prima volta crea "Mappa_import" (reparto + macchina → codice QR) con i codici indovinati:
      controllali/correggili (vuoto = vale per tutta la linea indicata, oppure TUTTE).
   4) Esegui di nuovo importaPharmaFix(): copia guasti in Soluzioni, manuali/schede in Documenti, anomalie in Anomalie. Rilanciarlo non crea doppioni. */
function importaPharmaFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(SH_IMP);
  if (!src) throw new Error('Crea la scheda "' + SH_IMP + '" e incolla i dati della vecchia Pharma-Fix (con le intestazioni)');
  creaFogliArchivio_(ss);
  const v = src.getDataRange().getValues(), head = v[0].map(h => String(h).trim().toLowerCase());
  const col = n => head.findIndex(h => h === n || h.indexOf(n) === 0);
  const ix = { rep: col('reparto'), mac: col('macchina'), prob: col('problema'), sol: col('soluzione'), link: col('link'), cat: col('categoria'), data: col('data') };
  if (ix.rep < 0 || ix.mac < 0 || ix.prob < 0) throw new Error('Intestazioni non trovate: servono almeno Reparto, Macchina, Problema (e Soluzione, Link, Categoria)');
  const rows = v.slice(1).filter(r => String(r[ix.prob]).trim());
  const linkCell = (i, r) => ix.link < 0 ? '' : (function () { const rt = src.getRange(i + 2, ix.link + 1).getRichTextValue(); return (rt && rt.getLinkUrl()) || String(r[ix.link] || ''); })();
  const eq = equipment_(ss), linee = Array.from(new Set(eq.map(e => e.linea)));
  const key = r => String(r[ix.rep]).trim() + ' | ' + String(r[ix.mac]).trim();

  let mappa = ss.getSheetByName(SH_MAPPA);
  if (!mappa) {
    mappa = ss.insertSheet(SH_MAPPA);
    const uniq = Array.from(new Set(rows.map(key)));
    const out = [['Reparto | Macchina (vecchia Pharma-Fix)', 'Codice QR / linea / TUTTE (controlla!)', 'Righe']];
    uniq.forEach(k => {
      const [rep, mac] = k.split(' | ');
      out.push([k, indovinaCodice_(rep, mac, eq, linee), rows.filter(r => key(r) === k).length]);
    });
    mappa.getRange(1, 1, out.length, 3).setValues(out);
    mappa.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#dfe7ec');
    mappa.getRange(2, 2, Math.max(1, out.length - 1), 1).setBackground('#fff3d6');
    mappa.setColumnWidth(1, 380); mappa.setColumnWidth(2, 260); mappa.setFrozenRows(1);
    Logger.log('Creata la scheda ' + SH_MAPPA + ' con ' + (out.length - 1) + ' abbinamenti. Controlla la colonna B, poi esegui di nuovo importaPharmaFix.');
    return;
  }
  const map = {};
  mappa.getDataRange().getValues().slice(1).forEach(r => { map[String(r[0]).trim()] = String(r[1]).trim(); });

  const sol = ss.getSheetByName(SH_SOL), doc = ss.getSheetByName(SH_DOC), ano = ss.getSheetByName(SH_ANO);
  const gia = new Set([sol, doc, ano].reduce((a, sh) => a.concat(sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(x => String(x[0])) : []), []));
  const nSol = [], nDoc = [], nAno = [];
  rows.forEach((r, i) => {
    const id = 'PF-' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key(r) + '|' + r[ix.prob] + '|' + (ix.cat >= 0 ? r[ix.cat] : ''))
      .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 12);
    if (gia.has(id)) return;
    const dove = map[key(r)] || 'TUTTE';
    const e = eq.find(x => x.code === dove) || { linea: linee.indexOf(dove) >= 0 ? dove : '', macchina: '' };
    const cat = ix.cat >= 0 ? String(r[ix.cat]).toUpperCase() : 'GUASTO';
    const data = ix.data >= 0 && r[ix.data] instanceof Date ? r[ix.data] : '';
    const link = linkCell(i, r), sol0 = ix.sol >= 0 ? String(r[ix.sol] || '') : '';
    if (cat.indexOf('MANUAL') >= 0 || cat.indexOf('SCHEMA') >= 0) nDoc.push([id, data, dove, cat.indexOf('SCHEMA') >= 0 ? 'SCHEMA' : 'MANUALE', r[ix.prob], sol0, link, 'import']);
    else if (cat.indexOf('SCHEDA') >= 0) nDoc.push([id, data, dove, 'SCHEDA FORMATO', r[ix.prob], sol0, link, 'import']);
    else if (cat.indexOf('ANOMALIA') >= 0) nAno.push([id, data, e.code || dove, e.linea, e.macchina, 3, r[ix.prob], sol0.indexOf('IN ATTESA') >= 0 ? '' : sol0, link, 'import', 'APERTA', '', '', '']);
    else nSol.push([id, data, dove, e.linea, e.macchina, '', r[ix.prob], sol0, '', link, 'import']);
  });
  [[sol, nSol], [doc, nDoc], [ano, nAno]].forEach(([sh, rr]) => { if (rr.length) sh.getRange(sh.getLastRow() + 1, 1, rr.length, rr[0].length).setValues(rr); });
  Logger.log('Importati: ' + nSol.length + ' soluzioni, ' + nDoc.length + ' documenti, ' + nAno.length + ' anomalie (le righe già importate sono state saltate).');
}

function indovinaCodice_(rep, mac, eq, linee) {
  const n = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
  const R = n(rep), M = n(mac);
  const num = (R.match(/(\d+)/) || [])[1];
  let linea = linee.find(l => n(l) === R) || linee.find(l => R && (n(l).indexOf(R) >= 0 || R.indexOf(n(l)) >= 0));
  if (!linea && num) linea = linee.find(l => n(l).replace(/\D/g, '') === num && /^ST/.test(n(l)));
  if (!linea && /SPL\s*(\d)/.test(R)) linea = 'SPL' + R.match(/SPL\s*(\d)/)[1];
  const cand = eq.filter(e => !linea || e.linea === linea);
  const SIN = { FILLING: 'RIEMPITRICE', TAPPATRICE: 'RIEMPITRICE', RIEMPIMENTO: 'RIEMPITRICE', ASTUCCIAMENTO: 'ASTUCCIATRICE', CARTONATORE: 'CARTONATRICE', ETICHETTATURA: 'ETICHETTATRICE', PESATRICE: 'BILANCIA' };
  const w0 = M.split(' ')[0], w = SIN[w0] || w0;
  const hit = cand.find(e => n(e.macchina) === M) || (w && cand.find(e => n(e.macchina).split(' ')[0] === w)) || (w && cand.find(e => n(e.macchina).indexOf(w) >= 0));
  return hit ? hit.code : (linea || 'TUTTE');
}

/* ---------- Telefoni dei tecnici ----------
   Ogni nome è legato al suo telefono. Il primo telefono con cui un tecnico entra viene registrato in automatico;
   un telefono diverso, o un telefono già di un altro tecnico, resta DA APPROVARE finché il responsabile non scrive OK. */
const SH_TEL = 'Telefoni';
const TEL_H = ['Tecnico', 'ID telefono', 'Modello', 'Primo accesso', 'Ultimo uso', 'Stato (OK / DA APPROVARE / BLOCCATO)'];
function telSheet_(ss) {
  let sh = ss.getSheetByName(SH_TEL);
  if (!sh) {
    sh = ss.insertSheet(SH_TEL);
    sh.getRange(1, 1, 1, TEL_H.length).setValues([TEL_H]).setFontWeight('bold').setBackground('#dfe7ec');
    sh.setFrozenRows(1); sh.setColumnWidth(1, 180); sh.setColumnWidth(3, 220); sh.setColumnWidth(6, 230);
  }
  return sh;
}
function checkTelefono_(man, dev, devName) {
  man = String(man || '').trim(); dev = String(dev || '').trim();
  if (!dev) throw new Error('Telefono non riconosciuto: aggiorna l\'app');
  const sh = telSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, TEL_H.length).getValues() : [];
  const stato = r => String(r[5] || '').trim().toUpperCase();
  const mine = rows.findIndex(r => String(r[0]).trim() === man && String(r[1]).trim() === dev);
  if (mine >= 0) {
    const st = stato(rows[mine]);
    if (st === 'OK') {
      const cell = sh.getRange(mine + 2, 5); const prev = cell.getValue();
      if (!(prev instanceof Date) || Date.now() - prev.getTime() > 3600e3) cell.setValue(new Date()).setNumberFormat('dd/MM/yyyy HH:mm');
      return;
    }
    throw new Error(st === 'BLOCCATO' ? 'Telefono non autorizzato: bloccato dal responsabile' : 'Telefono non autorizzato per ' + man + ': in attesa di approvazione del responsabile');
  }
  const altro = rows.find(r => String(r[1]).trim() === dev && String(r[0]).trim() !== man && stato(r) === 'OK');
  const giaOk = rows.some(r => String(r[0]).trim() === man && stato(r) === 'OK');
  const auto = !altro && !giaOk; // primo telefono di questo tecnico e non usato da altri
  sh.appendRow([man, dev, String(devName || '').slice(0, 80), new Date(), new Date(), auto ? 'OK' : 'DA APPROVARE']);
  const r = sh.getLastRow();
  sh.getRange(r, 4, 1, 2).setNumberFormat('dd/MM/yyyy HH:mm');
  if (auto) return;
  sh.getRange(r, 1, 1, TEL_H.length).setBackground('#fbeaa0');
  throw new Error('Telefono non autorizzato per ' + man + ': ' + (altro ? 'è il telefono di ' + altro[0] : 'telefono nuovo') + '. Chiedi al responsabile di approvarlo');
}

/* Orari dichiarati (fermi FILLING registrati all'uscita dallo sterile): solo dentro il turno in cui si registra */
// Orari scritti a mano (FILLING, dal batch record): fine nel turno in corso, oppure nel turno appena finito
// se si registra entro TOLL_DICH minuti dal cambio turno; durata massima 12 ore.
const TOLL_DICH = 120;
function limiteDichiarati_(reg) {
  let avvio = inizioTurno_(new Date(reg)).getTime();
  if (reg - avvio < TOLL_DICH * 60000) avvio = inizioTurno_(new Date(avvio - 60000)).getTime(); // tolleranza dopo il cambio turno
  return avvio;
}
function controllaDichiarati_(p, inizio) {
  const reg = Number(p.reg) || Date.now(), ts = Number(p.ts);
  if (ts > reg + 120000) throw new Error('Orario non valido: nel futuro');
  if (ts < limiteDichiarati_(reg) - 60000) throw new Error('Orario non valido: fuori dal turno in corso');
  const ini = inizio instanceof Date ? inizio.getTime() : Number(inizio) || null;
  if (ini != null && ts <= ini) throw new Error('Orario non valido: la fine è prima dell\'inizio');
  if (ini != null && ts - ini > 12 * 3600e3) throw new Error('Orario non valido: fermo più lungo di 12 ore');
}
function inizioTurno_(d) {
  const h = Number(Utilities.formatDate(d, TZ, 'H'));
  const base = dateOnly_(d);
  const ore = h >= 6 && h < 14 ? 6 : h >= 14 && h < 22 ? 14 : 22;
  const t = new Date(base.getTime() + ore * 3600e3);
  return h < 6 ? new Date(t.getTime() - 864e5) : t; // T3 dopo mezzanotte: iniziato alle 22 del giorno prima
}

/* Menu Pharma-Fix → Controlla registrazioni: elenca i casi da guardare nel foglio "Controlli" */
function controllaRegistrazioni() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [['Tipo', 'Dettaglio', 'Riferimento']];
  // 1) telefoni in attesa e telefoni usati da più tecnici
  const tv = telSheet_(ss).getDataRange().getValues().slice(1);
  tv.filter(r => String(r[5]).toUpperCase().trim() === 'DA APPROVARE')
    .forEach(r => out.push(['Telefono da approvare', r[0] + ' ha provato a entrare da un telefono nuovo (' + (r[2] || r[1]) + ')', 'Telefoni: ' + r[1]]));
  const perTel = {};
  tv.forEach(r => { const k = String(r[1]); (perTel[k] = perTel[k] || new Set()).add(String(r[0])); });
  Object.keys(perTel).forEach(k => { if (perTel[k].size > 1) out.push(['Stesso telefono, più tecnici', [...perTel[k]].join(', '), 'Telefoni: ' + k]); });
  // 2) fermi: sovrapposti per lo stesso tecnico, orari dichiarati registrati molto dopo
  const sh = ss.getSheetByName(SH_FERMI), last = sh.getLastRow();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, HEAD.length).getValues().map((r, i) => ({ r: i + 2, man: String(r[C.MAN - 1]), a: r[C.INIZIO - 1], b: r[C.FINE - 1], orari: r[C.ORARI - 1], reg: r[C.REG - 1], id: r[C.ID - 1] }))
      .filter(x => x.a instanceof Date && x.b instanceof Date && x.man);
    const per = {};
    v.forEach(x => (per[x.man] = per[x.man] || []).push(x));
    Object.keys(per).forEach(m => {
      const l = per[m].sort((x, y) => x.a - y.a);
      for (let i = 1; i < l.length; i++) if (l[i].a < l[i - 1].b) out.push(['Fermi sovrapposti', m + ': righe ' + l[i - 1].r + ' e ' + l[i].r, 'Fermi riga ' + l[i].r]);
    });
    // fermi registrati da chi quel giorno risultava a riposo, in ferie o assente (foglio Turni)
    let tt = {}; try { tt = turniTutti_(ss); } catch (e) { Logger.log(e); }
    const nomiT = {}; Object.keys(tt).forEach(n => { nomiT[n.toUpperCase()] = tt[n]; });
    v.forEach(x => {
      const g = new Date(x.a.getTime() - (Number(Utilities.formatDate(x.a, TZ, 'H')) < 6 ? 864e5 : 0)), p = nomiT[x.man.trim().toUpperCase()];
      const cod = p ? codTurno_(p[Utilities.formatDate(g, TZ, 'yyyy-MM-dd')]) : null;
      if (cod && cod.c && !/^[123GC]$/.test(cod.c)) out.push(['Fermo di chi non era in turno', x.man + ' risultava ' + ({ R: 'a riposo', F: 'in ferie' }[cod.c] || 'assente') + ' il ' + Utilities.formatDate(g, TZ, 'dd/MM') + ' (nome o telefono sbagliato? turni da aggiornare?)', 'Fermi riga ' + x.r]);
    });
    try { const U = s2 => String(s2 || '').trim().toUpperCase(), eqs = equipment_(ss);
      sh.getRange(2, 1, last - 1, C.EQ).getValues().forEach((r, i) => { const code = U(r[C.EQ - 1]); if (!code) return; const e = eqs.find(x => U(x.code) === code);
        if (e && (U(e.macchina) !== U(r[C.MACC - 1]) || U(e.linea) !== U(r[C.LINEA - 1]))) out.push(['Macchina, linea e codice non coincidono', r[C.LINEA - 1] + ' · ' + r[C.MACC - 1] + ' ma codice ' + e.code + ' = ' + e.linea + ' · ' + e.macchina + ': riscegli la macchina giusta', 'Fermi riga ' + (i + 2)]); });
    } catch (e) { Logger.log(e); }
    v.filter(x => x.orari === 'DICHIARATI' && x.reg instanceof Date && x.reg - x.b > 4 * 3600e3)
      .forEach(x => out.push(['Orari dichiarati registrati tardi', x.man + ': registrato ' + Math.round((x.reg - x.b) / 3600e3) + ' ore dopo la fine', 'Fermi riga ' + x.r]));
    // più interventi chiusi in un fermo solo: l'app ha avvisato e il tecnico ha risposto "è uno solo", oppure la descrizione racconta più operazioni
    const nostra = sh.getMaxColumns() >= C.MULTI && String(sh.getRange(1, C.MULTI).getValue()).trim() === EXTRA_H[5];
    const mv = nostra ? sh.getRange(2, C.MULTI, last - 1, 1).getValues() : [], vv = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    vv.forEach((r, i) => { if (!r[C.ID - 1] || String(r[C.ORARI - 1]) === ORA_CARTA) return;
      const m = String((mv[i] || [])[0] || '').trim(), d = piuOperazioni_(r[C.ALTRO - 1]);
      if (m || d) out.push(['Più interventi in un fermo?', String(r[C.MAN - 1]) + ': ' + (m || d + ' — "' + String(r[C.ALTRO - 1]).trim().slice(0, 80) + '"'), 'Fermi riga ' + (i + 2)]); });
  }
  if (out.length === 1) out.push(['Nessun caso da controllare', '', '']);
  let c = ss.getSheetByName('Controlli'); if (c) c.clear(); else c = ss.insertSheet('Controlli');
  c.getRange(1, 1, out.length, 3).setValues(out);
  c.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#dfe7ec');
  c.setColumnWidth(1, 240); c.setColumnWidth(2, 460); c.setColumnWidth(3, 160); c.setFrozenRows(1);
  c.getRange(1, 4).setValue('Aggiornato: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
  ss.setActiveSheet(c);
}

/* Aggiornamento v3 (rieseguibile): aree Filling/Packaging, colonne di traccia in Fermi, foglio Telefoni */
function aggiornaV3_(ss) {
  const eq = ss.getSheetByName(SH_EQ);
  if (eq && eq.getLastRow() > 1) {
    const rg = eq.getRange(2, 3, eq.getLastRow() - 1, 1), v = rg.getValues();
    rg.setValues(v.map(([a]) => [/^steril/i.test(String(a).trim()) ? 'Filling' : /^secondar/i.test(String(a).trim()) ? 'Packaging' : a]));
  }
  const f = ss.getSheetByName(SH_FERMI);
  if (f) {
    const h = f.getRange(1, C.ORARI, 1, 3);
    if (!String(h.getValues()[0][0])) h.setValues([['Orari', 'Registrato alle', 'Telefono']]).setFontWeight('bold').setBackground('#fff3d6');
  }
  telSheet_(ss);
  avvisiSheet_(ss);
}

/* ---------- Avvisi via mail ----------
   Foglio "Avvisi": destinatari e soglie (celle gialle), sotto il registro delle mail inviate.
   - anomalia con urgenza 1 → mail subito
   - fermo aperto da più di N ore → mail (controllo ogni 10 minuti insieme alla sveglia)
   - stesso guasto sulla stessa macchina N volte in G giorni → mail alla chiusura del fermo */
const SH_AVV = 'Avvisi';
const AVV_RIGHE = [
  ['Email destinatari (separate da virgola)', ''],
  ['Avvisa per anomalia critica (urgenza 1): SI / NO', 'SI'],
  ['Avvisa per fermo aperto da più di (ore, 0 = mai)', 2],
  ['Avvisa per stesso guasto sulla stessa macchina: quante volte (0 = mai)', 3],
  ['…nell\'arco di quanti giorni', 7],
  ['Telegram: token del bot (da @BotFather)', ''],
  ['Telegram: ID della chat o del gruppo (menu Pharma-Fix → Telegram: trova ID chat)', '']
];
function avvisiSheet_(ss) {
  let sh = ss.getSheetByName(SH_AVV);
  if (sh) { // foglio già esistente: aggiunge le righe nuove (Telegram) senza toccare i valori
    if (String(sh.getRange(2 + AVV_RIGHE.length - 1, 1).getValue()) !== AVV_RIGHE[AVV_RIGHE.length - 1][0]) {
      sh.getRange(2, 1, AVV_RIGHE.length, 1).setValues(AVV_RIGHE.map(r => [r[0]]));
      sh.getRange(2, 2, AVV_RIGHE.length, 1).setBackground('#fff3d6');
    }
    return sh;
  }
  sh = ss.insertSheet(SH_AVV);
  sh.getRange(1, 1, 1, 2).setValues([['Impostazione', 'Valore']]).setFontWeight('bold').setBackground('#dfe7ec');
  sh.getRange(2, 1, AVV_RIGHE.length, 2).setValues(AVV_RIGHE);
  sh.getRange(2, 2, AVV_RIGHE.length, 1).setBackground('#fff3d6');
  sh.getRange(9, 1, 1, 3).setValues([['Registro avvisi inviati', 'Tipo', 'Messaggio']]).setFontWeight('bold').setBackground('#dfe7ec');
  sh.setColumnWidth(1, 430); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 520); sh.setFrozenRows(1);
  return sh;
}
function avvisiCfg_() {
  const sh = avvisiSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const v = sh.getRange(2, 2, AVV_RIGHE.length, 1).getValues().map(r => r[0]);
  return { to: String(v[0] || '').split(/[,;\s]+/).filter(x => /@/.test(x)).join(','), critica: !/^no$/i.test(String(v[1]).trim()),
    ore: Number(v[2]) || 0, volte: Number(v[3]) || 0, giorni: Number(v[4]) || 7,
    tgToken: String(v[5] || '').trim(), tgChat: String(v[6] || '').trim(), sh: sh };
}
// Invia l'avviso per mail e/o Telegram (quello che è configurato nel foglio Avvisi). true = almeno un canale ok
function mail_(tipo, oggetto, righe, opt) {
  opt = opt || {};
  let c; try { c = avvisiCfg_(); } catch (e) { Logger.log(e); return false; }
  const canali = [];
  if (c.to && !opt.noMail) try {
    // mail semplice, come una mail scritta a mano: solo testo e mittente col nome dell'account
    // (nome "Pharma-Fix" e impaginazione HTML facevano scattare i filtri antispam aziendali)
    const corpo = typeof righe === 'string' ? oggetto + '\n\n' + righe : righe.map(([k, v]) => v === undefined ? String(k).toUpperCase() + '\n' : k + ': ' + testo_(v)).join('\n'); // niente link al foglio: il filtro aziendale lo rende illeggibile
    MailApp.sendEmail(c.to, 'Pharma-Fix · ' + oggetto, corpo);
    canali.push('mail');
  } catch (e) { Logger.log('Mail non inviata: ' + e); }
  if (c.tgToken && c.tgChat) {
    const r = telegram_(c.tgToken, c.tgChat, oggetto, righe);
    if (r === true) canali.push('Telegram'); else Logger.log('Telegram non inviato: ' + r);
  }
  if (!canali.length) return false;
  try {
    c.sh.insertRowAfter(9);
    c.sh.getRange(10, 1, 1, 3).setValues([[new Date(), tipo + ' (' + canali.join(' + ') + ')', oggetto]]);
    c.sh.getRange(10, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  } catch (e) { Logger.log(e); }
  return true;
}
// Messaggio Telegram in testo semplice (i valori HTML delle mail diventano testo)
const testo_ = v => String(v).replace(/<br\s*\/?>/gi, '\n   ').replace(/<a [^>]*href="([^"]+)"[^>]*>[^<]*<\/a>/gi, '$1').replace(/<[^>]+>/g, '')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
function telegram_(token, chat, oggetto, righe) {
  const txt = testo_;
  const icona = /CRITICA/.test(oggetto) ? '🔴 ' : /RIPETUTO/.test(oggetto) ? '🔁 ' : /APERTO DA/.test(oggetto) ? '⏱️ ' : '🔔 ';
  const testo = typeof righe === 'string' ? icona + oggetto + '\n\n' + righe : icona + 'PHARMA-FIX · ' + oggetto + '\n\n' + righe.filter((r, i) => r[1] !== undefined || i > 0).map(([k, v]) => v === undefined ? '\n' + k : k + ': ' + txt(v)).join('\n'); // righe senza valore dopo la prima = titoli di sezione
  try {
    const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chat, text: testo.slice(0, 4000), disable_web_page_preview: true }), muteHttpExceptions: true });
    const j = JSON.parse(r.getContentText());
    return j.ok ? true : (j.description || 'errore ' + r.getResponseCode());
  } catch (e) { return String(e); }
}
// Menu: trova l'ID della chat. Prima scrivi /start al bot (o nel gruppo dove l'hai aggiunto), poi esegui.
function trovaChatTelegram() {
  const c = avvisiCfg_(), ui = (() => { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } })();
  const say = m => ui ? ui.alert(m) : Logger.log(m);
  if (!c.tgToken) return say('Scrivi prima il token del bot nel foglio Avvisi (cella B7).');
  let j;
  try { j = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + c.tgToken + '/getUpdates', { muteHttpExceptions: true }).getContentText()); }
  catch (e) { return say('Telegram non raggiungibile: ' + e); }
  if (!j.ok) return say('Token non valido: ricopialo da @BotFather. (' + (j.description || '') + ')');
  const chats = [];
  (j.result || []).forEach(u => { const m = u.message || u.channel_post || u.my_chat_member; const ch = m && m.chat; if (ch && !chats.some(x => x.id === ch.id)) chats.push(ch); });
  if (!chats.length) return say('Nessun messaggio trovato. Apri la chat con il bot (o il gruppo dove l\'hai aggiunto), scrivi /start e riprova.');
  const ch = chats.find(x => x.type !== 'private') || chats[chats.length - 1]; // se c'è un gruppo, preferisce il gruppo
  c.sh.getRange(8, 2).setValue(String(ch.id));
  say('Trovata: ' + (ch.title || [ch.first_name, ch.last_name].filter(Boolean).join(' ')) + ' (' + (ch.type === 'private' ? 'chat privata' : 'gruppo') + ').\nID scritto nella cella B8. Ora prova: Pharma-Fix → Prova avvisi.');
}
const fmtOra_ = d => Utilities.formatDate(d, TZ, 'dd/MM HH:mm');
const h_ = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function avvisoAnomalia_(p, eq, foto) {
  if (!avvisiCfg_().critica) return;
  mail_('Anomalia critica', 'ANOMALIA CRITICA · ' + eq.linea + ' · ' + eq.macchina, [
    ['Anomalia critica (urgenza 1)'], ['Macchina', eq.linea + ' · ' + eq.macchina + ' (' + p.code + ')'],
    ['Cosa è stato visto', h_(p.descr)], ['Cosa serve', h_(p.serve || '—')], ['Segnalata da', h_(p.manutentore || '—')],
    ['Quando', fmtOra_(new Date(Number(p.ts) || Date.now()))], ['Foto', foto ? '<a href="' + foto + '">apri</a>' : '—']]);
}
function avvisoRipetuto_(sh, code, guasto) {
  try {
    const c = avvisiCfg_();
    if (!c.volte || !guasto || /^ALTRO$/i.test(String(guasto).trim())) return;
    const from = Date.now() - c.giorni * 864e5, last = sh.getLastRow();
    const v = sh.getRange(2, 1, last - 1, HEAD.length).getValues()
      .filter(r => r[C.EQ - 1] === code && r[C.GUASTO - 1] === guasto && r[C.STATO - 1] === 'CHIUSO' && r[C.INIZIO - 1] instanceof Date && r[C.INIZIO - 1].getTime() >= from);
    const n = v.length;
    if (n < c.volte || (n - c.volte) % c.volte !== 0) return; // avvisa alla soglia e ai multipli (3, 6, 9…)
    const min = v.reduce((a, r) => a + (r[C.FINE - 1] instanceof Date ? (r[C.FINE - 1] - r[C.INIZIO - 1]) / 60000 : 0), 0);
    const r0 = v[v.length - 1];
    mail_('Guasto ripetuto', 'GUASTO RIPETUTO ' + n + ' volte · ' + r0[C.LINEA - 1] + ' · ' + r0[C.MACC - 1], [
      ['Lo stesso guasto si è ripetuto ' + n + ' volte in ' + c.giorni + ' giorni'], ['Macchina', r0[C.LINEA - 1] + ' · ' + r0[C.MACC - 1] + ' (' + code + ')'],
      ['Tipo guasto', guasto], ['Tempo di fermo totale', Math.round(min) + ' minuti'],
      ['Interventi', v.map(r => fmtOra_(r[C.INIZIO - 1]) + ' · ' + (r[C.MAN - 1] || '') + (r[C.ALTRO - 1] ? ' · ' + h_(r[C.ALTRO - 1]) : '')).join('<br>')],
      ['Suggerimento', 'valutare un intervento correttivo invece della sola regolazione']]);
  } catch (e) { Logger.log(e); }
}
function avvisoFermiLunghi_() {
  try {
    const c = avvisiCfg_();
    if (!c.ore || !(c.to || (c.tgToken && c.tgChat))) return;
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI), last = sh.getLastRow();
    if (last < 2) return;
    const props = PropertiesService.getScriptProperties();
    let fatti = []; try { fatti = JSON.parse(props.getProperty('AVV_APERTI') || '[]'); } catch (e) {}
    const lim = Date.now() - c.ore * 3600e3;
    sh.getRange(2, 1, last - 1, HEAD.length).getValues()
      .filter(r => r[C.STATO - 1] === 'APERTO' && r[C.INIZIO - 1] instanceof Date && r[C.INIZIO - 1].getTime() < lim && fatti.indexOf(String(r[C.ID - 1])) < 0)
      .forEach(r => {
        const ore = Math.round((Date.now() - r[C.INIZIO - 1].getTime()) / 36e5 * 10) / 10;
        if (mail_('Fermo lungo', 'FERMO APERTO DA ' + ore + ' ORE · ' + r[C.LINEA - 1] + ' · ' + r[C.MACC - 1], [
          ['Fermo aperto da ' + ore + ' ore'], ['Macchina', r[C.LINEA - 1] + ' · ' + r[C.MACC - 1] + ' (' + r[C.EQ - 1] + ')'],
          ['Aperto alle', fmtOra_(r[C.INIZIO - 1])], ['Aperto da', r[C.APERTO_DA - 1] || '—'], ['Lotto', r[C.LOTTO - 1] || '—'],
          ['Nota', 'se il fermo è già finito, qualcuno ha dimenticato di chiuderlo nell\'app']])) fatti.push(String(r[C.ID - 1]));
      });
    props.setProperty('AVV_APERTI', JSON.stringify(fatti.slice(-300)));
  } catch (e) { Logger.log(e); }
}
function provaAvvisi() {
  const c = avvisiCfg_();
  const ok = mail_('Prova', 'avviso di prova', [['Gli avvisi funzionano'], ['Mail a', c.to || '—'], ['Telegram', c.tgToken && c.tgChat ? 'attivo' : 'non configurato'],
    ['Anomalia critica', c.critica ? 'SI' : 'NO'], ['Fermo aperto oltre', c.ore ? c.ore + ' ore' : 'mai'],
    ['Guasto ripetuto', c.volte ? c.volte + ' volte in ' + c.giorni + ' giorni' : 'mai']]);
  const msg = ok ? 'Avviso di prova inviato (' + [c.to ? 'mail' : '', c.tgToken && c.tgChat ? 'Telegram' : ''].filter(Boolean).join(' + ') + '). Controlla il registro nel foglio Avvisi.'
    : (c.to || (c.tgToken && c.tgChat) ? 'Invio non riuscito: apri Apps Script → Esecuzioni per vedere il motivo.' : 'Compila prima il foglio Avvisi: email (B2) e/o Telegram (B7, B8).');
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ---------- Report automatici e backup ----------
   Impostazioni nel foglio Avvisi, colonne E–F (celle gialle). attivaReport() crea i timer (una volta sola):
   - backup: domenica sera, copia .xlsx del foglio nella cartella Drive "Pharma-Fix Backup", tiene le ultime N
   - riepilogo settimanale: lunedì alle 7, settimana precedente (mail + Telegram)
   - passaggio di turno: a fine turno (6, 14, 22) solo nei giorni lavorativi, costruito dai dati registrati (Telegram) */
const REP_RIGHE = [
  ['Report automatici', 'Valore'],
  ['Riepilogo settimanale il lunedì alle 7: SI / NO', 'SI'],
  ['Passaggio di turno a fine turno su Telegram: SI / NO', 'SI'],
  ['Passaggio di turno anche per mail: SI / NO', 'NO'],
  ['Backup settimanale del foglio (domenica sera): SI / NO', 'SI'],
  ['Backup da conservare (quanti)', 8],
  ['Linee a ciclo continuo, 7 giorni su 7 (separate da virgola; vuoto = nessuna)', 'ST3']
];
function repCfg_() {
  const sh = avvisiSheet_(SpreadsheetApp.getActiveSpreadsheet());
  let nuovo = false;
  if (String(sh.getRange(1, 5).getValue()) !== REP_RIGHE[0][0]) { sh.getRange(1, 5, REP_RIGHE.length, 2).setValues(REP_RIGHE); nuovo = true; }
  else { // righe aggiunte in seguito: scrive le etichette nuove e il valore di partenza solo dove la cella è vuota
    const cur = sh.getRange(1, 5, REP_RIGHE.length, 2).getValues();
    REP_RIGHE.forEach((r, i) => { if (String(cur[i][0]) !== r[0]) { sh.getRange(i + 1, 5).setValue(r[0]); if (cur[i][1] === '') sh.getRange(i + 1, 6).setValue(r[1]); nuovo = true; } });
  }
  if (nuovo) {
    sh.getRange(1, 5, 1, 2).setFontWeight('bold').setBackground('#dfe7ec');
    sh.getRange(2, 6, REP_RIGHE.length - 1, 1).setBackground('#fff3d6');
    sh.setColumnWidth(5, 420);
  }
  const v = sh.getRange(2, 6, REP_RIGHE.length - 1, 1).getValues().map(r => r[0]);
  const si = x => !/^no$/i.test(String(x).trim());
  return { settimanale: si(v[0]), turno: si(v[1]), turnoMail: /^si$/i.test(String(v[2]).trim()), backup: si(v[3]), tieni: Number(v[4]) || 8,
    continuo: String(v[5] || '').split(/[,;]+/).map(x => x.trim().toUpperCase()).filter(Boolean) };
}
function attivaReport() {
  repCfg_();
  ScriptApp.getProjectTriggers().filter(t => ['backupSettimanale_', 'riepilogoSettimanale_', 'passaggioTurno_'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupSettimanale_').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).create();
  ScriptApp.newTrigger('riepilogoSettimanale_').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  [6, 14, 22].forEach(h => ScriptApp.newTrigger('passaggioTurno_').timeBased().atHour(h).nearMinute(10).everyDays(1).create());
  try { SpreadsheetApp.getUi().alert('Attivati: backup (domenica sera), riepilogo settimanale (lunedì 7:00), passaggio di turno (fine di ogni turno).\nSi accendono e spengono dalle celle gialle E–F del foglio Avvisi.'); } catch (e) {}
}
const durTxt_ = m => { m = Math.round(m); return m < 60 ? m + '′' : Math.floor(m / 60) + 'h' + (m % 60 ? ' ' + (m % 60) + '′' : ''); };
const hm_ = d => Utilities.formatDate(d, TZ, 'HH:mm'), gm_ = d => Utilities.formatDate(d, TZ, 'dd/MM');
const nice_ = s => String(s || '').split(' ').map(w => /\d/.test(w) || w.length <= 3 ? w : w.charAt(0) + w.slice(1).toLowerCase()).join(' '); // sigle e modelli (PC4400, M&H) restano maiuscoli
function fermiTutti_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI), last = sh.getLastRow();
  if (last < 2) return [];
  const gr = gruppiCol_(sh, last);
  return sh.getRange(2, 1, last - 1, HEAD.length).getValues().map((r, i) => [r, gr[i] || '']).filter(([r]) => r[C.INIZIO - 1] instanceof Date).map(([r, g]) => ({
    id: String(r[C.ID - 1] || ''), gruppo: g, a: r[C.INIZIO - 1], b: r[C.FINE - 1] instanceof Date ? r[C.FINE - 1] : null, stato: r[C.STATO - 1], linea: String(r[C.LINEA - 1] || ''),
    macc: String(r[C.MACC - 1] || ''), guasto: String(r[C.GUASTO - 1] || ''), man: String(r[C.MAN - 1] || r[C.APERTO_DA - 1] || ''),
    altro: String(r[C.ALTRO - 1] || ''), lotto: r[C.LOTTO - 1], orari: r[C.ORARI - 1], eq: String(r[C.EQ - 1] || '').trim() }));
}
// Area (FILLING / PACKAGING) di un fermo: dal codice macchina, oppure da linea + nome macchina (fermi importati)
function areaDi_() {
  const aree = {};
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_EQ).getDataRange().getValues().slice(1).forEach(r => {
    const a = String(r[2]).trim().toUpperCase(); if (!r[0] || !a) return;
    aree[String(r[0]).trim()] = a; aree[kpiChiave_(r[1], r[3])] = a;
  });
  return f => aree[f.eq] || aree[kpiChiave_(f.linea, f.macc)] || '';
}
const minDi_ = f => f.b ? (f.b - f.a) / 60000 : 0;
function anomalieTutte_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return righe_(ss, SH_ANO, ANO_H.length).map(r => ({ ts: r[1] instanceof Date ? r[1] : null, linea: r[3], macc: r[4], urg: Number(r[5]) || 3,
    descr: String(r[6] || ''), da: r[9], stato: String(r[10] || 'APERTA').toUpperCase(), chiusa: r[12] instanceof Date ? r[12] : null }));
}

/* ---------- Riepiloghi brevi (passaggio di turno e settimanale) ----------
   Testo già impaginato: ore per area (PACKAGING / FILLING, mai sommate) e per linea, i fermi più lunghi,
   solo le cose da sistemare. Il dettaglio completo resta nell'app e nel foglio Fermi. */
const areaStd_ = a => /^(FILL|STERIL)/i.test(String(a || '').trim()) ? 'FILLING' : /^(PACK|SECOND)/i.test(String(a || '').trim()) ? 'PACKAGING' : '';
const nFermi_ = n => n + (n === 1 ? ' fermo' : ' fermi');
// Nome corto della macchina: la prima parola ("Astucciatrice"); se sulla stessa linea due macchine iniziano uguale, il nome intero
function macchinaBreve_() {
  const conta = {};
  equipment_(SpreadsheetApp.getActiveSpreadsheet()).forEach(e => { const k = String(e.linea).trim().toUpperCase() + '|' + String(e.macchina).trim().split(/\s+/)[0].toUpperCase(); conta[k] = (conta[k] || 0) + 1; });
  return (linea, macc) => {
    const m = String(macc || '').trim(), p = m.split(/\s+/)[0] || '';
    return conta[String(linea).trim().toUpperCase() + '|' + p.toUpperCase()] > 1 ? nice_(m).slice(0, 28) : nice_(p);
  };
}
// Ore per linea di un'area. tutteLinee: elenco delle linee da mostrare anche a zero (riepilogo settimanale)
/* Fermi per reparto (linea), divisi in Filling e Packaging: niente somme tra linee diverse.
   "ST3 · Filling 1h 10′ (2) · Packaging 7′ (1)"; con prec (settimana prima) anche ▲/▼ per area. In fondo le linee senza fermi. */
const NOME_AREA_ = { FILLING: 'Filling', PACKAGING: 'Packaging', '': 'area?' };
// Un fermo = la macchina ferma: gli interventi con lo stesso "Fermo n." contano una volta sola
let UNICO_ = 0;
const chiaveFermo_ = f => f.gruppo || f.id || (f._k = f._k || 'u' + (++UNICO_)); // senza ID (righe di prova): ognuna è un fermo a sé
const contaTxt_ = (nI, nF) => nF === nI ? String(nI) : nF + (nF === 1 ? ' fermo' : ' fermi') + ' · ' + nI + ' int.';
// Interventi dello stesso fermo uniti: inizio del primo, fine dell'ultimo, guasti in fila
function unisciFermi_(lista) {
  const m = {};
  lista.forEach(f => { const k = chiaveFermo_(f), x = m[k];
    if (!x) { m[k] = Object.assign({}, f, { nInt: 1 }); return; }
    if (f.a < x.a) x.a = f.a; if (f.b && (!x.b || f.b > x.b)) x.b = f.b; x.nInt++;
    if (f.guasto && String(x.guasto).split(' + ').indexOf(f.guasto) < 0) x.guasto = x.guasto ? x.guasto + ' + ' + f.guasto : f.guasto;
    if (f.orari === 'DICHIARATI') x.orari = 'DICHIARATI'; });
  return Object.keys(m).map(k => m[k]);
}
function righeReparti_(lista, tutteLinee, prec) {
  const per = {}, ordA = ['FILLING', 'PACKAGING', ''];
  const mett = (dove, f) => { const l = String(f.linea || '?').trim(), a = ordA.indexOf(f.area) >= 0 ? f.area : ''; const x = dove[l] = dove[l] || {}; const y = x[a] = x[a] || { n: 0, m: 0, k: new Set() }; y.n++; y.k.add(chiaveFermo_(f)); y.m += minDi_(f); };
  lista.forEach(f => mett(per, f));
  const pp = {}; (prec || []).forEach(f => mett(pp, f));
  const tot = l => Object.keys(per[l]).reduce((x, a) => x + per[l][a].m, 0), ord = tutteLinee || [];
  const righe = Object.keys(per).sort((x, y) => tot(y) - tot(x) || ord.indexOf(x) - ord.indexOf(y)).map(l => l + ' · ' + ordA.filter(a => per[l][a]).map(a => {
    const y = per[l][a], q = prec ? ((pp[l] || {})[a] || { m: 0 }).m : null, d = q === null ? 0 : y.m - q;
    return NOME_AREA_[a] + ' ' + durTxt_(y.m) + ' (' + contaTxt_(y.n, y.k.size) + ')' + (prec && q && Math.abs(d) >= 1 ? ' ' + (d > 0 ? '▲+' : '▼−') + durTxt_(Math.abs(d)) : '');
  }).join(' · '));
  const senza = ord.filter(l => !per[l]);
  if (senza.length) righe.push('Senza fermi: ' + senza.join(', '));
  return righe;
}
function lineeTutte_(soloLinee) {
  const out = [];
  equipment_(SpreadsheetApp.getActiveSpreadsheet()).forEach(e => { const l = String(e.linea).trim(); if (l && out.indexOf(l) < 0 && (!soloLinee || soloLinee.indexOf(l.toUpperCase()) >= 0)) out.push(l); });
  return out;
}
function piuLunghi_(lista, breve, conData) {
  const ch = unisciFermi_(lista.filter(f => f.b)).sort((x, y) => minDi_(y) - minDi_(x)).slice(0, 3);
  return ch.length ? ['', '🔝 PIÙ LUNGHI'].concat(ch.map(f => '  ' + (conData ? gm_(f.a) + ' ' : '') + f.linea + ' ' + breve(f.linea, f.macc) + ' · ' +
    (f.guasto || 'guasto non indicato') + ' · ' + durTxt_(minDi_(f)) + (f.nInt > 1 ? ' (' + f.nInt + ' interventi)' : '') + (f.orari === 'DICHIARATI' ? ' [dich.]' : ''))) : [];
}
// Cognome da "COGNOME Nome": tutte le parole in maiuscolo all'inizio (DI MAIO, DE LUCA, LA ROSA…)
// Le particelle (Di, De, La, Del, D'…) restano attaccate alla parola dopo anche se scritte in minuscolo: "Di Bello Mario" → Di Bello
const PARTICELLE_ = /^(DI|DE|DA|DEL|DELLA|DELLE|DEGLI|DEI|DELLO|DAL|DALLA|DALLE|LA|LE|LO|LI|D'|DELL'|VAN|VON|MC|MAC|SAN|SANTA)$/i;
function cognome_(s) {
  const w = String(s || '').trim().split(/\s+/).filter(Boolean);
  const c = []; for (const x of w) { if (/[A-ZÀ-Ý]/.test(x) && x === x.toUpperCase()) c.push(x); else break; }
  if (!c.length && w.length) c.push(w[0]);
  while (c.length < w.length && PARTICELLE_.test(c[c.length - 1])) c.push(w[c.length]); // "DI" da solo non è un cognome
  return c.join(' ') || '?';
}
function rigaAnomalie_(nuove, aperte, chiuse) {
  const crit = aperte.filter(x => x.urg === 1).length;
  return '🔺 ANOMALIE · ' + (nuove.length === 1 ? '1 nuova' : nuove.length + ' nuove') + (chiuse != null ? ' · ' + chiuse + (chiuse === 1 ? ' chiusa' : ' chiuse') : '') +
    ' · ' + aperte.length + ' da gestire' + (crit ? ' (' + crit + (crit === 1 ? ' critica' : ' critiche') + ')' : '');
}

/* Riepilogo della settimana precedente (lun–dom) */
function riepilogoSettimanale_() { if (repCfg_().settimanale) riepilogoSettimanale(); }
// Guasti cronici: le 3 coppie guasto + macchina con più ore nelle ultime 4 settimane, con la tendenza sulle 4 precedenti
function croniciRighe_(tutti, fine, breve) {
  const d4 = new Date(fine.getTime() - 28 * 864e5), d8 = new Date(fine.getTime() - 56 * 864e5), m = {};
  tutti.forEach(f => { if (!f.guasto || String(f.guasto).trim().toUpperCase() === 'ALTRO' || !(f.a >= d8 && f.a < fine)) return;
    const k = f.guasto + ' · ' + f.linea + ' ' + breve(f.linea, f.macc), x = m[k] = m[k] || { n: 0, m: 0, p: 0 };
    if (f.a >= d4) { x.n++; x.m += minDi_(f); } else x.p += minDi_(f); });
  const top = Object.keys(m).filter(k => m[k].n >= 2).sort((a, b) => m[b].m - m[a].m).slice(0, 3);
  if (!top.length) return [];
  return ['', '📈 GUASTI CRONICI · ultime 4 settimane (tendenza sulle 4 prima)'].concat(top.map(k => { const x = m[k], dlt = x.m - x.p;
    return '  ' + k + ' ×' + x.n + ' · ' + durTxt_(x.m) + ' ' + (Math.abs(dlt) < 15 ? '= stabile' : dlt > 0 ? '▲ +' + durTxt_(dlt) : '▼ −' + durTxt_(-dlt)); }));
}
function riepilogoSettimanale() {
  try { completaFermiCarta_(); } catch (e) { Logger.log(e); }
  const oggi = dateOnly_(new Date()), dow = Number(Utilities.formatDate(oggi, TZ, 'u'));
  const fine = new Date(oggi.getTime() - (dow - 1) * 864e5), inizio = new Date(fine.getTime() - 7 * 864e5), prima = new Date(inizio.getTime() - 7 * 864e5);
  const areaDi = areaDi_(), breve = macchinaBreve_();
  const tutti = fermiTutti_().filter(f => f.stato === 'CHIUSO' && f.b);
  tutti.forEach(f => { f.area = areaStd_(areaDi(f)); });
  const sett = tutti.filter(f => f.a >= inizio && f.a < fine), prec = tutti.filter(f => f.a >= prima && f.a < inizio);
  const out = ['', '⏱ FERMI PER REPARTO · confronto con la settimana prima'];
  out.push.apply(out, righeReparti_(sett, lineeTutte_(), prec));
  const dopo = sett.filter(f => f.area !== 'FILLING' && String(f.orari) === 'DICHIARATI').length, nPack = sett.filter(f => f.area !== 'FILLING').length;
  const carta = sett.filter(f => String(f.orari) === ORA_CARTA).length;
  if (dopo || carta) out.push((dopo ? 'Packaging registrati dopo: ' + dopo + ' su ' + nPack : '') + (dopo && carta ? ' · ' : '') + (carta ? 'da carta: ' + carta : ''));
  const g = {}; sett.forEach(f => { if (!f.guasto) return; const k = f.guasto + ' · ' + f.linea + ' ' + breve(f.linea, f.macc); const x = g[k] = g[k] || { n: 0, m: 0 }; x.n++; x.m += minDi_(f); });
  const ric = Object.keys(g).filter(k => g[k].n > 1).sort((x, y) => g[y].n - g[x].n || g[y].m - g[x].m).slice(0, 3);
  if (ric.length) out.push('', '🔁 RICORRENTI'), ric.forEach(k => out.push('  ' + k + ' ×' + g[k].n + ' · ' + durTxt_(g[k].m)));
  out.push.apply(out, croniciRighe_(tutti, fine, breve));
  out.push.apply(out, piuLunghi_(sett, breve, true));
  const an = anomalieTutte_();
  out.push('', rigaAnomalie_(an.filter(a => a.ts && a.ts >= inizio && a.ts < fine), an.filter(a => a.stato !== 'CHIUSA'),
    an.filter(a => a.chiusa && a.chiusa >= inizio && a.chiusa < fine).length));
  try { // compiti: solo i numeri, il dettaglio è nel foglio Compiti
    const k = compitiTutti_();
    if (k.length) {
      const fatti = k.filter(x => x.stato === 'FATTO' && x.fatto >= inizio.getTime() && x.fatto < fine.getTime()), rit = k.filter(x => x.aperto && x.ritardo).length;
      out.push('📋 COMPITI · ' + fatti.length + (fatti.length === 1 ? ' fatto' : ' fatti') + (fatti.length ? ' (' + fatti.filter(x => !x.ritardo).length + ' in tempo)' : '') + ' · ' + rit + ' in ritardo aperti');
    }
  } catch (e) { Logger.log(e); }
  const rv = rigaVestizioni_(); if (rv) out.push(rv);
  const rc = rigaConvalide_(); if (rc) out.push(rc);
  const titolo = 'SETTIMANA ' + gm_(inizio) + '–' + gm_(new Date(fine.getTime() - 864e5));
  mail_('Riepilogo settimanale', titolo, out.slice(1).join('\n'));
}

/* Passaggio di turno: cosa è successo nel turno appena finito, solo dai dati registrati */
function passaggioTurno_() {
  const now = new Date(), g = Number(Utilities.formatDate(now, TZ, 'u')), h = Number(Utilities.formatDate(now, TZ, 'H'));
  // turni lun 6:00 → sab 6:00: fine turno valida lun 14 e 22, mar–ven 6/14/22, sab 6
  const ok = (g === 1 && h >= 13) || (g >= 2 && g <= 5) || (g === 6 && h < 12);
  const c = repCfg_();
  if (!c.turno) return;
  if (ok) passaggioTurno(c.turnoMail);
  else if (c.continuo.length) passaggioTurno(c.turnoMail, c.continuo); // weekend: solo le linee a ciclo continuo
}
function provaPassaggioTurno() { passaggioTurno(true); }
function passaggioTurno(conMail, soloLinee) {
  const ref = new Date(Date.now() - 30 * 60000), a = inizioTurno_(ref), b = new Date(a.getTime() + 8 * 3600e3);
  const linea = x => !soloLinee || soloLinee.indexOf(String(x.linea || '').trim().toUpperCase()) >= 0;
  const t = turno_(a), giorno = ['', 'lun', 'mar', 'mer', 'gio', 'ven', 'sab', 'dom'][Number(Utilities.formatDate(a, TZ, 'u'))] + ' ' + gm_(a);
  const areaDi = areaDi_(), breve = macchinaBreve_();
  const tutti = fermiTutti_().filter(linea);
  tutti.forEach(f => { f.area = areaStd_(areaDi(f)); });
  const nel = tutti.filter(f => f.a >= a && f.a < b).sort((x, y) => x.a - y.a);
  const aperti = tutti.filter(f => f.stato === 'APERTO');
  const an = anomalieTutte_().filter(linea), nuove = an.filter(x => x.ts && x.ts >= a && x.ts < b), aperte = an.filter(x => x.stato !== 'CHIUSA');
  const out = ['⏱ FERMI PER REPARTO'];
  const rt = rigaTurno_(a, t); if (rt) out.unshift(rt, ''); // chi era in turno (foglio Turni)
  out.push.apply(out, nel.length ? righeReparti_(nel, lineeTutte_(soloLinee)) : ['Nessun fermo nel turno']);
  out.push.apply(out, piuLunghi_(nel, breve, false));
  // Da sistemare: fermi lasciati aperti e fermi chiusi senza descrizione
  const senza = nel.filter(f => f.b && !String(f.altro || '').trim());
  // Packaging registrato dopo (orari scritti a mano): il fermo andava aperto col QR. In Filling è normale (sterile)
  const dopo = nel.filter(f => f.area !== 'FILLING' && String(f.orari) === 'DICHIARATI');
  const nomi = l => l.map(f => cognome_(f.man)).filter((x, i, q) => q.indexOf(x) === i).join(', ');
  if (aperti.length || senza.length || dopo.length) {
    out.push('', '⚠️ DA SISTEMARE');
    aperti.forEach(f => out.push('  APERTO: ' + f.linea + ' ' + breve(f.linea, f.macc) + ' dalle ' + (f.a < a ? gm_(f.a) + ' ' : '') + hm_(f.a) + ' · ' + (f.man || '?')));
    if (senza.length) out.push('  Senza descrizione: ' + senza.length + ' (' + nomi(senza) + ')');
    if (dopo.length) out.push('  Registrati dopo: ' + dopo.length + ' (' + nomi(dopo) + ')');
  }
  const rcf = rigaCambiFormato_(b); if (rcf) out.push('', rcf);
  out.push('', rigaAnomalie_(nuove, aperte, null));
  const titolo = 'PASSAGGIO TURNO ' + t + ' · ' + giorno + ' · ' + Utilities.formatDate(a, TZ, 'H') + '–' + Utilities.formatDate(b, TZ, 'H') + (soloLinee ? ' · solo ' + soloLinee.join(', ') : '');
  mail_('Passaggio di turno', titolo, out.join('\n'), { noMail: !conMail });
}

/* Backup: copia .xlsx del foglio in "Pharma-Fix Backup" (Drive), tiene le ultime N */
function backupSettimanale_() { if (repCfg_().backup) backupFoglio(); }
function backupFoglio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties();
  let folder = null;
  try { folder = DriveApp.getFolderById(props.getProperty('BACKUP_ID')); } catch (e) {}
  if (!folder) { folder = DriveApp.createFolder('Pharma-Fix Backup'); props.setProperty('BACKUP_ID', folder.getId()); }
  const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }).getBlob()
    .setName('Pharma-Fix Fermi - backup ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HHmm') + '.xlsx');
  folder.createFile(blob);
  const files = []; const it = folder.getFiles(); while (it.hasNext()) files.push(it.next());
  files.sort((x, y) => y.getDateCreated() - x.getDateCreated()).slice(repCfg_().tieni).forEach(f => f.setTrashed(true));
  try { SpreadsheetApp.getUi().alert('Backup fatto nella cartella Drive "Pharma-Fix Backup" (' + Math.min(files.length, repCfg_().tieni) + ' copie conservate).'); } catch (e) { Logger.log('Backup fatto'); }
}

/* ---------- Sveglia ----------
   Google "addormenta" l'app web quando nessuno la usa: la prima risposta può superare i 20 secondi.
   attivaSveglia() crea un timer che ogni 10 minuti, durante i turni (lun 6:00 – sab 6:00), interroga l'app
   come farebbe un telefono, così resta pronta. Si esegue una volta sola (menu Pharma-Fix o editor). */
function attivaSveglia() {
  disattivaSveglia();
  ScriptApp.newTrigger('tieniSveglio').timeBased().everyMinutes(10).create();
  tieniSveglio(true);
  try { SpreadsheetApp.getUi().alert('Sveglia attiva: ogni 10 minuti durante i turni.'); } catch (e) { Logger.log('Sveglia attiva'); }
}
function disattivaSveglia() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tieniSveglio').forEach(t => ScriptApp.deleteTrigger(t));
}
function tieniSveglio(forza) {
  const now = new Date(), g = Number(Utilities.formatDate(now, TZ, 'u')), h = Number(Utilities.formatDate(now, TZ, 'H'));
  const turni = (g >= 1 && g <= 5) && !(g === 1 && h < 5) || (g === 6 && h < 6); // da lunedì 5:00 (pronta per il T1) a sabato 6:00
  let continuo = false; try { continuo = repCfg_().continuo.length > 0; } catch (e) {}
  if (!turni && !continuo && forza !== true) return;
  avvisoFermiLunghi_();
  const url = ScriptApp.getService().getUrl(), key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!url || !key) return;
  try { UrlFetchApp.fetch(url + '?action=init&key=' + encodeURIComponent(key), { muteHttpExceptions: true, followRedirects: true }); }
  catch (e) { Logger.log('Sveglia: ' + e); }
}

/* =================== API =================== */
function doGet(e) {
  try {
    checkKey_(e.parameter.key);
    const action = e.parameter.action || 'init';
    if (action === 'init') return initCache_(e.parameter.v);
    if (action === 'ping') return json_({ ok: true });
    if (action === 'formati') { // solo coordinatori, dal loro telefono: i dati dei formati non arrivano mai ai tecnici
      checkTecnico_(e.parameter.man); checkTelefono_(e.parameter.man, e.parameter.dev, e.parameter.devName);
      if (!isCoordinatore_(e.parameter.man)) throw new Error('Sezione riservata ai coordinatori');
      return formatiJson_();
    }
    if (action === 'login') { // nome + telefono: il telefono deve essere quello del tecnico
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try { checkTecnico_(e.parameter.man); checkTelefono_(e.parameter.man, e.parameter.dev, e.parameter.devName); SpreadsheetApp.flush(); }
      finally { lock.releaseLock(); }
      return json_({ ok: true });
    }
    throw new Error('Azione sconosciuta');
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const p = JSON.parse(e.postData.contents);
    checkKey_(p.key);
    lock.waitLock(20000);
    if (['open', 'close', 'cancel', 'soluzione', 'anomalia', 'anomaliaStato', 'approvaSol', 'correzione', 'compito', 'compitoStato', 'modifica', 'foto'].indexOf(p.action) >= 0) { checkTecnico_(p.manutentore); checkTelefono_(p.manutentore, p.dev, p.devName); }
    let res;
    if (p.action === 'open') res = openFermo_(p);
    else if (p.action === 'close') res = closeFermo_(p);
    else if (p.action === 'foto') res = fotoFermo_(p);
    else if (p.action === 'cancel') res = cancelFermo_(p);
    else if (p.action === 'soluzione') res = addSoluzione_(p);
    else if (p.action === 'anomalia') res = addAnomalia_(p);
    else if (p.action === 'anomaliaStato') res = statoAnomalia_(p);
    else if (p.action === 'approvaSol') res = approvaSoluzione_(p);
    else if (p.action === 'correzione') res = correzione_(p);
    else if (p.action === 'modifica') res = modificaFermo_(p);
    else if (p.action === 'rivisto') res = rivistoFermo_(p);
    else if (p.action === 'documento') res = addDocumento_(p);
    else if (p.action === 'compito') res = addCompito_(p);
    else if (p.action === 'compitoStato') res = statoCompito_(p);
    else throw new Error('Azione sconosciuta');
    SpreadsheetApp.flush();
    bumpVer_(); // i telefoni scaricano i dati nuovi al prossimo giro
    return json_(Object.assign({ ok: true }, res));
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function checkKey_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('KEY');
  if (!key || String(k || '') !== key) throw new Error('Codice reparto non valido');
}
// Solo i nomi presenti nella colonna Manutentori del foglio Liste possono registrare fermi
function checkTecnico_(nome) {
  const lv = listeVal_();
  const col = lv[0].indexOf('Manutentori');
  const ok = (col >= 0 && lv.slice(1).some(r => String(r[col]).trim() !== '' && String(r[col]).trim() === String(nome || '').trim())) ||
    isCoordinatore_(nome); // i coordinatori entrano anche se non sono nella colonna Manutentori
  if (!ok) throw new Error('Manutentore non abilitato: ' + nome);
}
/* ---------- Versione dei dati e copia pronta ----------
   Ogni telefono chiede i dati ogni minuto mandando la versione che ha. Se nel frattempo non è cambiato niente
   risponde solo { same: true } (pochi byte, nessun foglio letto). La versione cambia quando arriva qualcosa dall'app
   (doPost), quando qualcuno modifica il foglio a mano (onEdit) e comunque ogni VER_MIN minuti (per sicurezza:
   cambi di giorno, righe cancellate, funzioni lanciate dal menu). I dati calcolati restano pronti nella cache
   dello script, così con molti telefoni il foglio si legge una volta sola per versione. */
const VER_MIN = 15;
function dataVer_() {
  const props = PropertiesService.getScriptProperties();
  let v = props.getProperty('DATA_VER');
  if (!v) { v = String(Date.now()); props.setProperty('DATA_VER', v); }
  return v + '-' + Math.floor(Date.now() / (VER_MIN * 60000)) + '-' + VERSIONE_SCRIPT.replace(/\D/g, ''); // script nuovo = dati ricalcolati subito
}
function bumpVer_() { try { const p = PropertiesService.getScriptProperties(); p.setProperty('DATA_VER', String(Math.max(Date.now(), Number(p.getProperty('DATA_VER')) + 1 || 0))); } catch (e) { Logger.log(e); } }
function initCache_(vApp) {
  const ver = dataVer_();
  if (vApp && String(vApp) === ver) return json_({ ok: true, same: true, ver: ver, serverTime: Date.now() });
  const cache = CacheService.getScriptCache(), k = 'init:' + ver;
  let txt = null;
  try { txt = cacheGet_(cache, k); } catch (e) { Logger.log(e); }
  if (!txt) {
    txt = JSON.stringify(Object.assign(init_(), { ver: ver }));
    try { cachePut_(cache, k, txt); } catch (e) { Logger.log(e); }
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}
// La cache tiene al massimo 100 KB per voce: il testo va compresso e spezzato in pezzi
function cachePut_(cache, k, txt) {
  const b = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(txt, 'application/json')).getBytes()), N = 90000;
  const n = Math.ceil(b.length / N), o = {};
  for (let i = 0; i < n; i++) o[k + ':' + i] = b.substr(i * N, N);
  o[k] = String(n);
  cache.putAll(o, VER_MIN * 60 + 120);
}
function cacheGet_(cache, k) {
  const n = Number(cache.get(k)); if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(k + ':' + i);
  const m = cache.getAll(keys), parts = keys.map(x => m[x]);
  if (parts.some(x => !x)) return null;
  return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(parts.join('')), 'application/x-gzip')).getDataAsString();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function init_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { classificazione_(ss); } catch (e) { Logger.log(e); }
  const equipment = equipment_(ss);
  // Liste
  const lv = listeVal_(ss);
  const liste = {};
  lv[0].forEach((h, i) => { if (h) liste[h] = lv.slice(1).map(r => String(r[i]).trim()).filter(Boolean); });
  // spiegazione accanto a ogni tipo intervento (stessa riga del foglio Liste)
  const intDesc = {}, hi = lv[0].map(h => String(h).trim().toUpperCase()), ci = hi.indexOf('TIPO INTERVENTO'), cd = hi.indexOf(LI_DESC.toUpperCase());
  if (ci >= 0 && cd >= 0) lv.slice(1).forEach(r => { const k = String(r[ci]).trim().toUpperCase(); if (k && String(r[cd]).trim()) intDesc[k] = String(r[cd]).trim(); });
  // Fermi: aperti + chiusi ultimi 90 giorni
  const sh = ss.getSheetByName(SH_FERMI);
  try { idDoppi_(sh); } catch (e) { Logger.log(e); } // righe copiate a mano con lo stesso ID fermo
  const last = sh.getLastRow();
  const fermi = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    const fotoC = sh.getRange(2, C.FOTO_C, last - 1, 1).getFormulas(); // link della foto di chiusura
    const nx = Math.min(EXTRA_H.length, sh.getMaxColumns() - C.CAUSA + 1), extra = nx > 0 ? sh.getRange(2, C.CAUSA, last - 1, nx).getValues() : [];
    const nostre = nx > 3 && String(sh.getRange(1, C.GRUPPO).getValue()).trim() === EXTRA_H[3];
    if (!nostre) extra.forEach(x => x.splice(3));
    const cause = extra, corr = correzioniDett_(ss);
    const from = Date.now() - 90 * 864e5, nCorr = correzioniPerFermo_(ss);
    vals.forEach((r, i) => {
      const start = r[C.INIZIO - 1] instanceof Date ? r[C.INIZIO - 1].getTime() : null;
      const stato = r[C.STATO - 1];
      if (!r[C.ID - 1] || !start) return;
      if (stato !== 'APERTO' && start < from) return;
      const end = r[C.FINE - 1] instanceof Date ? r[C.FINE - 1].getTime() : null;
      fermi.push({ id: r[C.ID - 1], code: r[C.EQ - 1], linea: r[C.LINEA - 1], macchina: r[C.MACC - 1], lotto: String(r[C.LOTTO - 1]),
        start: start, end: end, stato: stato, cat: r[C.CAT - 1], int: r[C.INT - 1], guasto: r[C.GUASTO - 1], causa: String((cause[i] || [])[0] || ''), prop: String((extra[i] || [])[1] || ''), rivisto: String((extra[i] || [])[2] || ''),
        gruppo: String((extra[i] || [])[3] || ''), giro: Number((extra[i] || [])[4]) || 0, multi: String((extra[i] || [])[5] || ''),
        fotoX: String((extra[i] || [])[6] || '').split(/\s+/).filter(u => /^https?:/.test(u)),
        altro: r[C.ALTRO - 1], man: r[C.MAN - 1], apertoDa: r[C.APERTO_DA - 1],
        pezzo: String(r[C.PEZZO - 1] || ''), orari: String(r[C.ORARI - 1] || ''), nCorr: nCorr[String(r[C.ID - 1])] || 0, corr: (c => c && String(c.da).toUpperCase() !== String(r[C.MAN - 1]).trim().toUpperCase() ? c : null)(corr[String(r[C.ID - 1])]), fotoC: ((fotoC[i][0] || '').match(/"(https?:[^"]+)"/) || [])[1] || '' });
    });
  }
  return Object.assign({ ok: true, equipment: equipment, liste: liste, guasti: guasti_(ss), intDesc: intDesc, vociCat: (() => { try { return vociCat_(ss); } catch (e) { Logger.log(e); return {}; } })(), vociDesc: (() => { try { return vociDesc_(ss); } catch (e) { Logger.log(e); return {}; } })(), fermi: fermi, compiti: compiti_(ss), vestizioni: (() => { try { return vestizioni_(ss); } catch (e) { Logger.log(e); return []; } })(), sterileNo: (() => { try { return esonerati_(ss); } catch (e) { Logger.log(e); return []; } })(), convalide: (() => { try { return convalide_(ss); } catch (e) { Logger.log(e); return { list: [], ignote: [] }; } })(), turni: (() => { try { return turni_(ss); } catch (e) { Logger.log(e); return {}; } })(), programma: (() => { try { return programma_(ss); } catch (e) { Logger.log(e); return {}; } })(), continuo: (() => { try { return repCfg_().continuo; } catch (e) { return []; } })(), avviso: (() => { try { const c = vestSetup_(ss); return c ? { vest: c.avvV, conv: c.avvC, tolV: c.tolV, tolC: c.tolC } : null; } catch (e) { Logger.log(e); return null; } })(), serverTime: Date.now() }, archivio_(ss));
}

// Equipment e Liste si leggono una volta sola per richiesta (ogni richiesta allo script riparte da zero)
let EQ_MEMO = null, LISTE_MEMO = null;
const listeVal_ = ss => LISTE_MEMO || (LISTE_MEMO = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(SH_LISTE).getDataRange().getValues());
function equipment_(ss) {
  if (EQ_MEMO) return EQ_MEMO;
  const eqv = ss.getSheetByName(SH_EQ).getDataRange().getValues().slice(1);
  return EQ_MEMO = eqv.filter(r => r[0] && String(r[4]).toUpperCase() !== 'NO')
    .map(r => ({ code: String(r[0]).trim(), linea: String(r[1]), area: String(r[2]), macchina: String(r[3]) }));
}

/* Riga del foglio Fermi copiata a mano (es. per aggiungere un 2° intervento): ha lo stesso "ID fermo" della riga da cui viene.
   Con due righe uguali, "Va bene" / "Correggi" dall'app finirebbero sempre sulla prima. La copia riceve un ID suo (ID-2, ID-3…)
   e "Fermo n." lega le righe dello stesso fermo, come per gli interventi registrati dall'app. */
function idDoppi_(sh) {
  const last = sh.getLastRow(); if (last < 3) return 0;
  const ids = sh.getRange(2, C.ID, last - 1, 1).getValues().map(x => String(x[0]).trim()), visti = {}, doppi = [];
  ids.forEach((id, i) => { if (!id) return; if (visti[id]) doppi.push(i); else visti[id] = 1; });
  if (!doppi.length) return 0;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(10000)) return 0;
  try {
    const usati = new Set(ids), gruppi = causaCol_(sh), g = gruppi ? sh.getRange(2, C.GRUPPO, last - 1, 1).getValues() : null;
    doppi.forEach(i => {
      const base = ids[i]; let n = 2; while (usati.has(base + '-' + n)) n++;
      const nuovo = base + '-' + n; usati.add(nuovo); ids[i] = nuovo;
      sh.getRange(i + 2, C.ID).setValue(nuovo);
      if (g) ids.forEach((x, j) => { if ((x === base || j === i) && !String(g[j][0]).trim()) { g[j][0] = base; sh.getRange(j + 2, C.GRUPPO).setValue(base); } });
    });
    bumpVer_();
  } finally { lock.releaseLock(); }
  return doppi.length;
}
// Foglio Fermi corretto a mano: vince la colonna modificata. Macchina → Codice e Linea; Codice → Macchina e Linea; Linea → Codice.
function eqCoerente_(ss, sh, row, col) {
  const U = x => String(x || '').trim().toUpperCase(), eqs = equipment_(ss);
  const v = sh.getRange(row, 1, 1, C.EQ).getValues()[0], mac = U(v[C.MACC - 1]), lin = U(v[C.LINEA - 1]), code = U(v[C.EQ - 1]);
  let e = null;
  if (col === C.EQ) e = eqs.find(x => U(x.code) === code);
  else { const conNome = eqs.filter(x => U(x.macchina) === mac);
    e = conNome.length === 1 ? conNome[0] : conNome.find(x => U(x.linea) === lin) || null;
    if (col === C.LINEA && !e) return; }
  if (!e) return;
  if (U(e.code) !== code) sh.getRange(row, C.EQ).setValue(e.code);
  if (U(e.linea) !== lin) sh.getRange(row, C.LINEA).setValue(e.linea);
  if (U(e.macchina) !== mac) sh.getRange(row, C.MACC).setValue(e.macchina);
}
function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const cell = sh.getRange(2, C.ID, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function turno_(d) {
  const h = Number(Utilities.formatDate(d, TZ, 'H'));
  return h >= 6 && h < 14 ? 'T1' : h >= 14 && h < 22 ? 'T2' : 'T3';
}
function dateOnly_(d) {
  const p = Utilities.formatDate(d, TZ, 'yyyy-M-d').split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function savePhoto_(b64, name) {
  if (!b64) return '';
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', name + '.jpg');
  return folder.createFile(blob).getUrl();
}
// Foto in più della chiusura (2ª e 3ª): arrivano dopo la chiusura, una alla volta; i link vanno in "Altre foto chiusura", uno per riga
function fotoFermo_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SH_FERMI), r = findRow_(sh, p.fermo);
  if (!r) throw new Error('Fermo non trovato per la foto');
  if (!p.foto) return { vuota: true };
  causaCol_(sh);
  const cell = sh.getRange(r, C.FOTO_X), prima = String(cell.getValue() || '').trim();
  const code = String(sh.getRange(r, C.EQ).getValue() || 'fermo'), d = new Date(Number(p.ts) || Date.now());
  const nome = Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + code + '_chiusura_' + (Number(p.n) || 2) + '_' + String(p.id || '').slice(-6);
  let url = '';
  try { const it = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID')).getFilesByName(nome + '.jpg'); if (it.hasNext()) url = it.next().getUrl(); } catch (e) { Logger.log(e); }
  if (!url) url = savePhoto_(p.foto, nome); // invio ripetuto: la foto c'è già, non la duplica
  if (prima.indexOf(url) < 0) cell.setValue((prima ? prima + '\n' : '') + url).setWrap(true);
  return { url: url };
}
function linkFormula_(url) {
  return url ? '=HYPERLINK("' + url + '","Foto")' : '';
}

function openFermo_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_FERMI);
  if (findRow_(sh, p.id)) return { dup: true }; // già registrato (invio ripetuto)
  const eq = equipment_(ss).find(x => x.code === p.code) || { linea: p.linea || '', macchina: p.macchina || '' };
  const d = new Date(Number(p.ts));
  const stamp = Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm');
  const lotto = lotto_(p.lotto);
  const url = savePhoto_(p.foto, stamp + '_' + p.code + '_apertura');
  const row = new Array(HEAD.length).fill('');
  row[C.LOTTO - 1] = lotto;
  row[C.DATA - 1] = dateOnly_(d);
  row[C.INIZIO - 1] = d;
  row[C.MAN - 1] = p.manutentore || '';
  row[C.MACC - 1] = eq.macchina;
  row[C.LINEA - 1] = eq.linea;
  row[C.TURNO - 1] = turno_(d);
  row[C.STATO - 1] = 'APERTO';
  row[C.ID - 1] = p.id;
  row[C.EQ - 1] = p.code;
  row[C.APERTO_DA - 1] = p.manutentore || '';
  row[C.ORARI - 1] = p.dichiarato ? 'DICHIARATI' : 'TEMPO REALE';
  row[C.REG - 1] = new Date(Number(p.reg) || Date.now());
  row[C.TEL - 1] = p.dev || '';
  if (p.dichiarato) controllaDichiarati_(p);
  sh.appendRow(row);
  const r = sh.getLastRow();
  try { fermiGuastoDV_(ss, sh, r, 1); } catch (e) { Logger.log(e); }
  sh.getRange(r, C.DATA).setNumberFormat('dd/MM/yyyy');
  sh.getRange(r, C.INIZIO, 1, 2).setNumberFormat('HH:mm');
  sh.getRange(r, C.REG).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(r, C.LOTTO).setNumberFormat('@').setValue(lotto);
  sh.getRange(r, C.DUR).setFormula(F_('=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))')).setNumberFormat('[h]:mm');
  sh.getRange(r, C.IG).setFormula(F_('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')'));
  if (url) sh.getRange(r, C.FOTO_A).setFormula(F_(linkFormula_(url)));
  sh.getRange(r, 1, 1, HEAD.length).setBackground('#fff3d6');
  gruppoFermo_(sh, r, p);
  return { row: r };
}

// Lotto obbligatorio: solo cifre, da 4 a 10
function lotto_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!/^\d{4,10}$/.test(s)) throw new Error('Lotto obbligatorio: inserisci il numero di lotto');
  return s; // testo: gli zeri iniziali (es. 012345) non si perdono
}
// Lotti scritti come numero dalle versioni precedenti: tornano testo a 6 cifre (012345). Gira da solo una volta.
const LOTTO_CIFRE = 6;
function sistemaLotti() {
  bumpVer_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI), last = sh.getLastRow();
  let n = 0;
  if (last >= 2) {
    const rg = sh.getRange(2, C.LOTTO, last - 1, 1), v = rg.getValues();
    const out = v.map(r => {
      const x = r[0];
      if (typeof x === 'number' && isFinite(x)) { const t = String(Math.round(x)); if (t.length < LOTTO_CIFRE) n++; return [t.padStart(LOTTO_CIFRE, '0')]; }
      return [x];
    });
    rg.setNumberFormat('@').setValues(out);
  }
  PropertiesService.getScriptProperties().setProperty('LOTTI_TESTO', '1');
  try { SpreadsheetApp.getUi().alert('Lotti sistemati: ' + n + ' con lo zero iniziale ripristinato. La colonna Lotto ora è in formato testo.'); } catch (e) {}
  return n;
}

function closeFermo_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const r = findRow_(sh, p.id);
  if (!r) throw new Error('Fermo ' + p.id + ' non trovato');
  if (sh.getRange(r, C.STATO).getValue() === 'CHIUSO') return { dup: true };
  const daAltri = chiCoordina_(sh, r, p.manutentore); // chi non l'ha aperto deve essere coordinatore
  lotto_(p.lotto);
  if (String(p.intervento || '').trim().toUpperCase() === 'SOSTITUZIONE' && !String(p.pezzo || '').trim())
    throw new Error('Con SOSTITUZIONE il codice pezzo sostituito è obbligatorio');
  const d = new Date(Number(p.ts));
  // FILLING: fermo aperto prima di entrare in sterile, chiuso con gli orari del batch record (p.inizio = inizio corretto)
  const nuovoInizio = p.inizio ? new Date(Number(p.inizio)) : null;
  if (p.dichiarato) controllaDichiarati_(p, nuovoInizio || sh.getRange(r, C.INIZIO).getValue());
  if (nuovoInizio) {
    const prima = sh.getRange(r, C.INIZIO).getValue();
    if (prima instanceof Date && Math.abs(prima - nuovoInizio) > 60000) // traccia: l'ora vera di apertura nell'app resta come nota
      sh.getRange(r, C.INIZIO).setNote('Aperto nell\'app alle ' + Utilities.formatDate(prima, TZ, 'HH:mm dd/MM') + '; inizio corretto dal batch record');
    sh.getRange(r, C.INIZIO).setValue(nuovoInizio).setNumberFormat('HH:mm');
    sh.getRange(r, C.DATA).setValue(dateOnly_(nuovoInizio)).setNumberFormat('dd/MM/yyyy');
    sh.getRange(r, C.TURNO).setValue(turno_(nuovoInizio));
  }
  if (p.dichiarato) sh.getRange(r, C.ORARI).setValue('DICHIARATI');
  const code = sh.getRange(r, C.EQ).getValue();
  const url = savePhoto_(p.foto, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + code + '_chiusura');
  sh.getRange(r, C.FINE).setValue(d).setNumberFormat('HH:mm');
  if (p.manutentore && !daAltri) sh.getRange(r, C.MAN).setValue(p.manutentore);
  if (daAltri) sh.getRange(r, C.MAN).setNote('Chiuso da ' + p.manutentore + ' (coordinatore) il ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
  const cat = String(p.categoria || '').trim() || (c0 => c0 === 'A SCELTA' ? '' : c0)(vociCat_(SpreadsheetApp.getActiveSpreadsheet())[String(p.guasto || '').trim().toUpperCase()] || '');
  sh.getRange(r, C.CAT, 1, 4).setValues([[cat, p.intervento || '', p.guasto || '', p.altro || '']]);
  if (p.causa || p.proposta) causaCol_(sh);
  if (p.causa) sh.getRange(r, C.CAUSA).setValue(String(p.causa).trim());
  if (p.proposta) sh.getRange(r, C.PROP).setValue(String(p.proposta).trim().toUpperCase());
  sh.getRange(r, C.PEZZO).setValue(p.pezzo || '');
  gruppoFermo_(sh, r, p); // primo intervento di un fermo diviso: diventa "intervento 1"
  if (p.multi && causaCol_(sh)) sh.getRange(r, C.MULTI).setValue(String(p.multi).trim().slice(0, 200));
  if (url) sh.getRange(r, C.FOTO_C).setFormula(F_(linkFormula_(url)));
  sh.getRange(r, C.STATO).setValue('CHIUSO');
  sh.getRange(r, C.REG).setValue(new Date(Number(p.reg) || Date.now())).setNumberFormat('dd/MM/yyyy HH:mm');
  const tel = String(sh.getRange(r, C.TEL).getValue() || '');
  if (p.dev && tel.indexOf(p.dev) < 0) sh.getRange(r, C.TEL).setValue(tel ? tel + ' / ' + p.dev : p.dev);
  sh.getRange(r, C.LOTTO).setNumberFormat('@').setValue(lotto_(p.lotto));
  sh.getRange(r, 1, 1, HEAD.length).setBackground(null);
  try { SpreadsheetApp.flush(); aggiornaValoriKpi(); } catch (e) { Logger.log(e); }
  avvisoRipetuto_(sh, code, p.guasto);
  return { row: r };
}

// Il fermo lo chiude/annulla chi l'ha aperto; un altro solo se è coordinatore (torna true in quel caso)
function chiCoordina_(sh, r, chi) {
  const aperto = String(sh.getRange(r, C.APERTO_DA).getValue() || sh.getRange(r, C.MAN).getValue() || '').trim();
  if (!aperto || aperto.toUpperCase() === String(chi || '').trim().toUpperCase()) return false;
  if (isCoordinatore_(chi)) return true;
  throw new Error('Può chiudere il fermo solo chi l\'ha aperto (' + aperto + ') o un coordinatore');
}
/* Correzioni: il tecnico chiede, il responsabile corregge a mano il foglio Fermi e mette FATTA */
const SH_CORR = 'Correzioni', CORR_H = ['Data richiesta', 'Fermo (ID)', 'Riga Fermi', 'Linea', 'Macchina', 'Data fermo', 'Chiesta da', 'Cosa correggere', 'Stato (DA FARE / FATTA)', 'Note', 'ID richiesta'];
function correzione_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), fsh = ss.getSheetByName(SH_FERMI);
  const r = findRow_(fsh, p.fermo);
  if (!r) throw new Error('Fermo ' + p.fermo + ' non trovato');
  if (!String(p.testo || '').trim()) throw new Error('Scrivi cosa va corretto: obbligatorio');
  let sh = ss.getSheetByName(SH_CORR);
  if (!sh) {
    sh = ss.insertSheet(SH_CORR);
    sh.getRange(1, 1, 1, CORR_H.length).setValues([CORR_H]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
    sh.setFrozenRows(1); sh.setColumnWidth(8, 380); sh.setColumnWidth(10, 260); sh.hideColumns(11);
  }
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 11, sh.getLastRow() - 1, 1).getValues().map(x => String(x[0])) : [];
  if (ids.indexOf(String(p.id)) >= 0) return { dup: true };
  const v = fsh.getRange(r, 1, 1, HEAD.length).getValues()[0];
  const chi = String(v[C.MAN - 1] || v[C.APERTO_DA - 1] || '').trim(), io = String(p.manutentore || '').trim();
  if (chi.toUpperCase() !== io.toUpperCase() && String(v[C.APERTO_DA - 1] || '').trim().toUpperCase() !== io.toUpperCase() && !isCoordinatore_(io))
    throw new Error('La correzione la può chiedere solo chi ha registrato il fermo o un coordinatore');
  const giorno = v[C.INIZIO - 1] instanceof Date ? v[C.INIZIO - 1] : v[C.DATA - 1];
  sh.appendRow([new Date(Number(p.ts) || Date.now()), p.fermo, r, v[C.LINEA - 1], v[C.MACC - 1], giorno, io, String(p.testo).trim(), 'DA FARE', '', p.id]);
  const n = sh.getLastRow();
  sh.getRange(n, 1).setNumberFormat('dd/MM/yyyy HH:mm'); sh.getRange(n, 6).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(n, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['DA FARE', 'FATTA'], true).build()).setBackground('#fff3d6');
  fsh.getRange(r, C.ALTRO).setNote('Correzione chiesta da ' + io + ' il ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + ': ' + String(p.testo).trim());
  try {
    mail_('Correzione richiesta', 'CORREZIONE RICHIESTA · ' + v[C.LINEA - 1] + ' ' + nice_(v[C.MACC - 1]), [
      ['Fermo del ' + (giorno instanceof Date ? Utilities.formatDate(giorno, TZ, 'dd/MM HH:mm') : '') + ' (riga ' + r + ' del foglio Fermi)'],
      ['Chiesta da', io], ['Cosa correggere', h_(String(p.testo).trim().slice(0, 400))],
      ['Cosa fare', 'correggi la riga ' + r + ' nel foglio Fermi, poi metti FATTA nel foglio Correzioni']], { noMail: (c => !!(c.tgToken && c.tgChat))(avvisiCfg_()) });
  } catch (e) { Logger.log(e); }
  return {};
}
function cancelFermo_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const r = findRow_(sh, p.id);
  if (!r) return { dup: true };
  if (sh.getRange(r, C.STATO).getValue() !== 'APERTO') throw new Error('Si possono annullare solo fermi aperti');
  chiCoordina_(sh, r, p.manutentore);
  sh.deleteRow(r);
  return {};
}

/* =================== Compiti ===================
   Il coordinatore assegna un compito a un tecnico o a un turno; il tecnico lo prende in carico e lo segna fatto.
   Una riga per compito nel foglio Compiti, mai cancellata: lo stato cambia e ogni passaggio finisce nella colonna Storia. */
const SH_COMP = 'Compiti';
const COMP_H = ['ID', 'Creato il', 'Assegnato da', 'Assegnato a (tecnico o Turno T1/T2/T3)', 'Linea', 'Macchina (codice)', 'Titolo', 'Cosa fare',
  'Priorità (NORMALE / URGENTE)', 'Entro il', 'Stato (DA FARE / IN CORSO / FATTO / ANNULLATO)', 'Preso in carico da', 'Preso il', 'Fatto da', 'Fatto il',
  'Nota di chiusura', 'In ritardo', 'Storia'];
const K = { ID: 1, CREATO: 2, DA: 3, A: 4, LINEA: 5, EQ: 6, TITOLO: 7, DESCR: 8, PR: 9, ENTRO: 10, STATO: 11, PRESO_DA: 12, PRESO: 13, FATTO_DA: 14, FATTO: 15, NOTA: 16, RIT: 17, STORIA: 18 };
const K_STATI = ['DA FARE', 'IN CORSO', 'FATTO', 'ANNULLATO'];
function compitiSheet_(ss) {
  let sh = ss.getSheetByName(SH_COMP);
  if (sh) return sh;
  sh = ss.insertSheet(SH_COMP);
  sh.getRange(1, 1, 1, COMP_H.length).setValues([COMP_H]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
  sh.setFrozenRows(1);
  sh.setColumnWidth(K.TITOLO, 240); sh.setColumnWidth(K.DESCR, 320); sh.setColumnWidth(K.NOTA, 280); sh.setColumnWidth(K.STORIA, 420);
  sh.getRange(2, K.ENTRO, 999, 1).setNumberFormat('dd/MM/yyyy');
  [K.CREATO, K.PRESO, K.FATTO].forEach(c => sh.getRange(2, c, 999, 1).setNumberFormat('dd/MM/yyyy HH:mm'));
  sh.getRange(2, K.STATO, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(K_STATI, true).build());
  return sh;
}
const turnoDi_ = a => ((String(a || '').match(/^\s*turno\s+(T[123])\s*$/i) || [])[1] || '').toUpperCase();
function entro_(v) {
  if (v instanceof Date) return dateOnly_(v);
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const n = Number(v); return n ? dateOnly_(new Date(n)) : null;
}
function addCompito_(p) {
  if (!isCoordinatore_(p.manutentore)) throw new Error('Solo i coordinatori possono assegnare compiti');
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = compitiSheet_(ss);
  if (findIn_(sh, p.id)) return { dup: true };
  const titolo = String(p.titolo || '').trim(), a = String(p.a || '').trim();
  if (!titolo) throw new Error('Il titolo del compito è obbligatorio');
  const turno = turnoDi_(a);
  const tecnici = tecniciTutti_(ss);
  if (!turno && !tecnici.some(t => t.toUpperCase() === a.toUpperCase())) throw new Error('Tecnico non valido per il compito: ' + a);
  const entro = entro_(p.entro);
  if (!entro) throw new Error('La data "entro il" è obbligatoria');
  const eq = p.code ? (equipment_(ss).find(x => x.code === p.code) || { linea: '' }) : { linea: '' };
  const d = new Date(Number(p.ts) || Date.now()), chi = String(p.manutentore).trim();
  const pr = /^urg/i.test(String(p.pr || '')) ? 'URGENTE' : 'NORMALE';
  sh.appendRow([p.id, d, chi, turno ? 'Turno ' + turno : a, eq.linea || '', p.code || '', titolo, String(p.descr || '').trim(), pr, entro, 'DA FARE',
    '', '', '', '', '', '', Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm') + ' assegnato da ' + chi + ' a ' + (turno ? 'Turno ' + turno : a)]);
  const r = sh.getLastRow();
  sh.getRange(r, K.CREATO).setNumberFormat('dd/MM/yyyy HH:mm'); sh.getRange(r, K.ENTRO).setNumberFormat('dd/MM/yyyy');
  // In ritardo: fatto dopo la scadenza, oppure ancora aperto e scaduto (si aggiorna da solo ogni giorno)
  sh.getRange(r, K.RIT).setFormulaR1C1('=IF(RC11="ANNULLATO","",IF(RC11="FATTO",IF(INT(RC15)>RC10,"SÌ","NO"),IF(TODAY()>RC10,"SÌ","")))');
  if (pr === 'URGENTE') sh.getRange(r, K.PR).setBackground('#fbeaa0');
  return {};
}
function statoCompito_(p) {
  const sh = compitiSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const r = findIn_(sh, p.id);
  if (!r) throw new Error('Compito ' + p.id + ' non trovato');
  const v = sh.getRange(r, 1, 1, COMP_H.length).getValues()[0];
  const stato = String(p.stato || '').toUpperCase(), prima = String(v[K.STATO - 1] || 'DA FARE').toUpperCase();
  if (K_STATI.indexOf(stato) < 0) throw new Error('Stato non valido');
  const chi = String(p.manutentore || '').trim(), CHI = chi.toUpperCase(), coord = isCoordinatore_(chi);
  const a = String(v[K.A - 1] || '').trim(), preso = String(v[K.PRESO_DA - 1] || '').trim();
  const mio = a.toUpperCase() === CHI || preso.toUpperCase() === CHI || (!!turnoDi_(a) && !preso);
  const nota = String(p.nota || '').trim(), quando = new Date(Number(p.ts) || Date.now());
  if (stato === prima) return { dup: true }; // reinvio dalla coda del telefono
  let log;
  if (stato === 'IN CORSO') {
    if (prima !== 'DA FARE') throw new Error('Il compito non è più da fare (' + prima + ')');
    if (!mio && !coord) throw new Error('Compito assegnato a ' + (preso || a) + ': lo può prendere solo lui o un coordinatore');
    sh.getRange(r, K.STATO).setValue('IN CORSO');
    sh.getRange(r, K.PRESO_DA, 1, 2).setValues([[chi, quando]]); sh.getRange(r, K.PRESO).setNumberFormat('dd/MM/yyyy HH:mm');
    log = 'preso in carico da ' + chi;
  } else if (stato === 'FATTO') {
    if (prima === 'ANNULLATO') throw new Error('Il compito è stato annullato');
    if (!mio && !coord) throw new Error('Compito assegnato a ' + (preso || a) + ': lo può segnare fatto solo lui o un coordinatore');
    if (!preso) { sh.getRange(r, K.PRESO_DA, 1, 2).setValues([[chi, quando]]); sh.getRange(r, K.PRESO).setNumberFormat('dd/MM/yyyy HH:mm'); }
    sh.getRange(r, K.STATO).setValue('FATTO');
    sh.getRange(r, K.FATTO_DA, 1, 3).setValues([[chi, quando, nota]]); sh.getRange(r, K.FATTO).setNumberFormat('dd/MM/yyyy HH:mm');
    log = 'fatto da ' + chi + (nota ? ': ' + nota : '');
  } else { // DA FARE (riaperto) o ANNULLATO: solo coordinatori, con il motivo
    if (!coord) throw new Error('Solo i coordinatori possono ' + (stato === 'ANNULLATO' ? 'annullare' : 'riaprire') + ' un compito');
    if (!nota) throw new Error('Il motivo è obbligatorio');
    sh.getRange(r, K.STATO).setValue(stato);
    if (stato === 'DA FARE') sh.getRange(r, K.PRESO_DA, 1, 5).clearContent(); // torna a chi era assegnato; la storia resta
    log = (stato === 'ANNULLATO' ? 'annullato' : 'riaperto') + ' da ' + chi + ': ' + nota;
  }
  const storia = String(v[K.STORIA - 1] || '');
  sh.getRange(r, K.STORIA).setValue((storia ? storia + '\n' : '') + Utilities.formatDate(quando, TZ, 'dd/MM/yyyy HH:mm') + ' ' + log);
  return {};
}
// Per l'app: tutti i compiti aperti + fatti/annullati negli ultimi 90 giorni
function compiti_(ss) {
  const from = Date.now() - 90 * 864e5;
  return righe_(ss, SH_COMP, COMP_H.length).map(r => ({ id: String(r[K.ID - 1]), ts: ts_(r[K.CREATO - 1]), da: String(r[K.DA - 1]), a: String(r[K.A - 1]).trim(),
    linea: String(r[K.LINEA - 1]), code: String(r[K.EQ - 1]), titolo: String(r[K.TITOLO - 1]), descr: String(r[K.DESCR - 1]),
    pr: String(r[K.PR - 1] || 'NORMALE').toUpperCase(), entro: r[K.ENTRO - 1] instanceof Date ? r[K.ENTRO - 1].getTime() : ts_(r[K.ENTRO - 1]),
    stato: String(r[K.STATO - 1] || 'DA FARE').toUpperCase(), presoDa: String(r[K.PRESO_DA - 1]), preso: ts_(r[K.PRESO - 1]),
    fattoDa: String(r[K.FATTO_DA - 1]), fatto: ts_(r[K.FATTO - 1]), nota: String(r[K.NOTA - 1]), storia: String(r[K.STORIA - 1]) }))
    .filter(k => k.id && (k.stato === 'DA FARE' || k.stato === 'IN CORSO' || (k.fatto || k.ts || 0) >= from));
}
// Per i riepiloghi (passaggio di turno e settimanale)
function compitiTutti_() {
  const oggi = dateOnly_(new Date()).getTime();
  return compiti_(SpreadsheetApp.getActiveSpreadsheet()).map(k => Object.assign(k, {
    aperto: k.stato === 'DA FARE' || k.stato === 'IN CORSO',
    ritardo: (k.stato === 'DA FARE' || k.stato === 'IN CORSO') ? !!k.entro && oggi > k.entro : k.stato === 'FATTO' && !!k.entro && !!k.fatto && dateOnly_(new Date(k.fatto)).getTime() > k.entro }));
}


/* =================== Fermi scritti a mano nel foglio (da carta) ===================
   Righe aggiunte a mano nel foglio Fermi (senza ID fermo), es. ricopiate dai vecchi fogli cartacei.
   Le completa perché contino come le altre (KPI, riepiloghi, storico nell'app):
   data+ora vere, Stato CHIUSO, Linea e codice Equipment, Turno, ID, Orari = CARTACEO, formule.
   Si lancia dal menu (con il resoconto) e da sola prima di KPI e riepilogo settimanale. */
const ORA_CARTA = 'CARTACEO';
function sistemaFermiCarta() {
  const r = completaFermiCarta_(); bumpVer_();
  const msg = r.fatte ? 'Fermi da carta completati: ' + r.fatte + '.' : 'Nessun fermo da completare.';
  const det = r.problemi.length ? '\n\nDa sistemare a mano (poi rilancia):\n' + r.problemi.slice(0, 25).join('\n') + (r.problemi.length > 25 ? '\n… e altri ' + (r.problemi.length - 25) : '') : '';
  try { SpreadsheetApp.getUi().alert(msg + det); } catch (e) { Logger.log(msg + det); }
}
function completaFermiCarta_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { fatte: 0, problemi: ['foglio occupato, riprova tra poco'] };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SH_FERMI), last = sh.getLastRow();
    if (last < 2) return { fatte: 0, problemi: [] };
    const v = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    const disp = sh.getRange(2, C.INIZIO, last - 1, 2).getDisplayValues(); // ore come si vedono (le sole ore del 1899 hanno fusi orari storici)
    const eq = equipment_(ss), linee = Array.from(new Set(eq.map(e => e.linea)));
    const U_ = s => String(s || '').trim().toUpperCase();
    // ora scritta nel foglio: data vera, sola ora (Sheets la salva come 30/12/1899) o testo "14:20"
    const ora = (x, testo, g) => {
      if (x instanceof Date && x.getFullYear() >= 1950) return x;
      const m = String(testo || x || '').trim().match(/(\d{1,2})[:.](\d{2})/);
      return m ? new Date(g.getFullYear(), g.getMonth(), g.getDate(), Number(m[1]), Number(m[2])) : null;
    };
    let fatte = 0; const problemi = [];
    v.forEach((row, i) => {
      const r = i + 2;
      if (String(row[C.ID - 1]).trim()) return;
      if (row[C.DATA - 1] === '' && row[C.INIZIO - 1] === '' && row[C.MACC - 1] === '') return; // riga vuota
      const g0 = dataOra_(row[C.DATA - 1], null) || (row[C.INIZIO - 1] instanceof Date && row[C.INIZIO - 1].getFullYear() >= 1950 ? row[C.INIZIO - 1] : null);
      const giorno = g0 ? new Date(g0.getFullYear(), g0.getMonth(), g0.getDate()) : null;
      const ini = giorno ? ora(row[C.INIZIO - 1], disp[i][0], giorno) : null;
      let fin = giorno ? ora(row[C.FINE - 1], disp[i][1], giorno) : null;
      if (fin && ini && fin <= ini) fin = new Date(fin.getTime() + 864e5); // a cavallo della mezzanotte
      let code = String(row[C.EQ - 1]).trim(), linea = String(row[C.LINEA - 1]).trim();
      const e0 = eq.find(x => U_(x.code) === U_(code));
      if (e0) { code = e0.code; linea = linea || e0.linea; }
      else {
        if (linea) { const l = linee.find(x => U_(x) === U_(linea)); if (l) linea = l; }
        // codice indovinato dal nome macchina, solo dentro la linea scritta (senza linea sarebbe un tiro a caso)
        const g = linea ? indovinaCodice_(linea, row[C.MACC - 1], eq, linee) : '', e1 = eq.find(x => x.code === g);
        code = e1 && e1.linea === linea ? e1.code : '';
        const nm = s => U_(s).replace(/[^A-Z0-9]+/g, ' ').trim(), w = nm(row[C.MACC - 1]).split(' ')[0];
        if (code && nm(e1.macchina) !== nm(row[C.MACC - 1]) && eq.filter(x => x.linea === linea && nm(x.macchina).split(' ')[0] === w).length > 1) code = ''; // es. 3 riordinatori sulla stessa linea
      }
      const manca = [!giorno && 'data', giorno && !ini && 'ora inizio', giorno && !fin && 'ora fine', !linea && 'linea', !code && 'macchina (codice Equipment)']
        .filter(Boolean);
      if (manca.length) { problemi.push('riga ' + r + ': manca ' + manca.join(', ') + (row[C.MACC - 1] ? ' (' + String(row[C.MACC - 1]).trim() + ')' : '')); return; }
      if (fin - ini > 24 * 3600e3) { problemi.push('riga ' + r + ': fermo più lungo di 24 ore, controlla gli orari'); return; }
      const e = eq.find(x => x.code === code);
      const id = 'CARTA-' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, [r, ini.getTime(), fin.getTime(), code, row[C.MAN - 1], Date.now()].join('|'))
        .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 10).toUpperCase();
      sh.getRange(r, C.DATA, 1, 3).setValues([[giorno, ini, fin]]);
      sh.getRange(r, C.DATA).setNumberFormat('dd/MM/yyyy'); sh.getRange(r, C.INIZIO, 1, 2).setNumberFormat('HH:mm');
      if (!String(row[C.MACC - 1]).trim()) sh.getRange(r, C.MACC).setValue(e.macchina);
      sh.getRange(r, C.LINEA, 1, 2).setValues([[linea, turno_(ini)]]);
      sh.getRange(r, C.STATO, 1, 4).setValues([['CHIUSO', id, code, row[C.APERTO_DA - 1] || row[C.MAN - 1] || '']]);
      sh.getRange(r, C.ORARI, 1, 2).setValues([[ORA_CARTA, new Date()]]); sh.getRange(r, C.REG).setNumberFormat('dd/MM/yyyy HH:mm');
      if (row[C.LOTTO - 1] !== '') sh.getRange(r, C.LOTTO).setNumberFormat('@').setValue(lotto_(row[C.LOTTO - 1]));
      sh.getRange(r, C.DUR).setFormula(F_('=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))')).setNumberFormat('[h]:mm');
      sh.getRange(r, C.IG).setFormula(F_('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')'));
      sh.getRange(r, C.ORARI).setBackground('#e8eef3');
      fatte++;
    });
    if (fatte) SpreadsheetApp.flush();
    return { fatte: fatte, problemi: problemi };
  } finally { lock.releaseLock(); }
}

/* =================== Convalide MF (vestizione, corsi e convalide per l'area sterile) ===================
   Foglio "Convalide MF" (prima si chiamava "Vestizioni"), una riga per tecnico:
   A Tecnico · B Area sterile (SI/NO): NO = esonerato da tutto (niente avvisi, fuori dagli elenchi) · vuoto = SI
   C–G "fatto il": MD_20, Vestizione, Partecipazione MF, Cat. G, Cat. H (giallo chiaro: le scrivi tu;
       Cat. G e H le riempie anche lo script dallo "Scadenziario MF": scadenza dello scadenziario − durata)
   H–L scadenze calcolate dallo script (viola: data, MANCA o esonerato) · M Note · a destra il riquadro Impostazioni.
   Scadenza = data "fatto il" + durata; per le categorie vale la più lontana tra questa e lo scadenziario.
   Una categoria nuova nello scadenziario (es. Cat. F) aggiunge le sue due colonne in fondo (il foglio si riordina da solo).
   L'app avvisa dal preavviso; scaduta ma entro la tolleranza = arancio (ancora valida). */
const SH_VEST = 'Convalide MF', SH_VEST_OLD = 'Vestizioni', VEST_AVVISO = 60;
const VEST_STER = 'Area sterile (SI/NO)', CONV_NOME_ = { G: 'regolazione', H: 'sostituzione', MD_20: 'corso', MF: 'partecipazione' };
// Scadenze che valgono per tutti quelli in area sterile anche se lo scadenziario non le riporta
const CONV_FISSE = ['MD_20', 'MF'];
const MF_BASE = ['MD_20', 'V', 'MF', 'G', 'H']; // ordine delle colonne
const mfNome_ = k => ({ MD_20: 'MD_20', V: 'Vestizione', MF: 'Partecipazione MF' })[k] || 'Cat. ' + k + (CONV_NOME_[k] ? ' ' + CONV_NOME_[k] : '');
const mfFatto_ = k => mfNome_(k) + '\n' + (k === 'V' || k === 'MF' ? 'fatta il' : k === 'MD_20' ? 'fatto il' : 'ultima convalida');
const mfScad_ = k => 'Scadenza\n' + ({ MD_20: 'MD_20', V: 'Vestizione', MF: 'Partecipazione MF' }[k] || 'Cat. ' + k);
const VEST_SET = [['Durata vestizione (mesi)', 12], ['Durata corso MD_20 (mesi)', 12], ['Durata partecipazione MF (mesi)', 12], ['Durata convalide Cat. G / H (mesi)', 12],
  ['Preavviso vestizione (giorni prima)', 60], ['Preavviso corsi e convalide (giorni prima)', 60], ['Tolleranza vestizione (giorni dopo la scadenza)', 30],
  ['Tolleranza corsi e convalide (giorni dopo la scadenza)', 30]];
const VEST_SET_RE = [/^durata (standard )?vestizione/i, /^durata corso md/i, /^durata partecipazione mf/i, /^durata (standard )?convalide/i, /^preavviso vestizione/i,
  /^preavviso (corsi|convalide)/i, /^tolleranza vestizione/i, /^tolleranza (corsi|convalide)/i];
const VEST_NOTA = 'Giallo chiaro = scrivi tu la data. Viola = calcolate dallo script (non modificare). Giallo = in preavviso, arancio = scaduta ma in tolleranza (ancora valida), rosso = oltre la tolleranza o MANCA. Area sterile NO = esonerato da tutto.';
let VEST_MEMO = null;
// Nomi nuovi dei fogli (una volta): "Convalide" → "Scadenziario MF", "Vestizioni" → "Convalide MF"
function rinominaMF_(ss) {
  try {
    if (!ss.getSheetByName(SH_CONV) && ss.getSheetByName(SH_CONV_OLD)) ss.getSheetByName(SH_CONV_OLD).setName(SH_CONV);
    if (!ss.getSheetByName(SH_VEST) && ss.getSheetByName(SH_VEST_OLD)) ss.getSheetByName(SH_VEST_OLD).setName(SH_VEST);
  } catch (e) { Logger.log(e); }
}
// Legge il foglio (anche nel formato vecchio "Vestizioni") per nome di colonna
function mfLeggi_(sh) {
  const lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lc < 1) return { h: [], rows: [], keys: [], set: {} };
  const h = sh.getRange(1, 1, 1, lc).getValues()[0].map(x => String(x).trim()), U1 = s => s.split('\n')[0].trim().toUpperCase();
  const cSet = h.findIndex(x => /^impostazioni$/i.test(x)) + 1, fine = cSet ? cSet - 1 : lc;
  const col = { fatto: {} }; let keys = [];
  h.slice(0, fine).forEach((x, i) => {
    const u = U1(x), full = x.toUpperCase().replace(/\s+/g, ' '); let m;
    if (/^TECNICO/.test(u)) col.man = i;
    else if (/AREA STERILE/.test(u)) col.ster = i;
    else if (/^NOTE/.test(u)) col.note = i;
    else if (/^DATA VESTIZIONE/.test(u) || /^VESTIZIONE/.test(u)) col.fatto.V = i;
    else if (/^MD_?20/.test(u)) col.fatto.MD_20 = i;
    else if (/^PARTECIPAZIONE MF/.test(u)) col.fatto.MF = i;
    else if ((m = full.match(/^ULTIMA CONVALIDA CAT\.?\s*([A-Z0-9_]+)/))) col.fatto[m[1]] = i;
    else if (!/^SCADENZA/.test(u) && (m = u.match(/^CAT\.?\s*([A-Z0-9_]+)/))) col.fatto[m[1]] = i;
  });
  keys = Object.keys(col.fatto);
  const rows = lr > 1 ? sh.getRange(2, 1, lr - 1, fine).getValues().filter(r => String(r[col.man || 0]).trim()).map(r => {
    const f = {}; keys.forEach(k => { f[k] = r[col.fatto[k]]; });
    return { man: String(r[col.man || 0]).trim(), ster: col.ster != null ? String(r[col.ster]).trim() : '', note: col.note != null ? r[col.note] : '', fatto: f };
  }) : [];
  const set = {};
  if (cSet && lr > 1) sh.getRange(2, cSet, Math.min(lr, 15), 2).getValues().forEach(r => { VEST_SET_RE.forEach((re, i) => { if (re.test(String(r[0]).trim()) && String(r[1]).trim() !== '') set[i] = r[1]; }); });
  const hh = h.slice(0, fine); while (hh.length && !hh[hh.length - 1]) hh.pop();
  return { h: hh, rows: rows, keys: keys, set: set, cSet: cSet };
}
const mfTesta_ = keys => ['Tecnico', VEST_STER, ...keys.map(mfFatto_), ...keys.map(mfScad_), 'Note'];
// Riscrive il foglio nel formato nuovo tenendo i dati (prima volta o categoria nuova)
function mfRiscrivi_(sh, L, keys) {
  const testa = mfTesta_(keys), n = keys.length, nc = testa.length, righe = L.rows.slice().sort((a, b) => a.man.localeCompare(b.man));
  sh.clearConditionalFormatRules(); sh.clear();
  try { sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2), Math.max(sh.getMaxColumns(), nc + 3)).clearDataValidations(); } catch (e) { Logger.log(e); }
  sh.getRange(1, 1, 1, nc).setValues([testa]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true).setVerticalAlignment('middle');
  sh.getRange(1, 3, 1, n).setBackground('#fff3d6').setNote('Ultima volta fatto: scrivi tu la data. Cat. G e H le scrive anche lo script dallo Scadenziario MF (scadenza − durata); se è stata rinnovata scrivi la data nuova.');
  sh.getRange(1, 3 + n, 1, n).setBackground('#e6e0f3').setNote('Scritte dallo script (data, MANCA o esonerato): non modificare.');
  if (righe.length) {
    sh.getRange(2, 1, righe.length, nc).setValues(righe.map(r => [r.man, r.ster, ...keys.map(k => r.fatto[k] instanceof Date ? r.fatto[k] : (dataConv_(r.fatto[k]) || '')), ...keys.map(() => ''), r.note || '']));
  }
  const nr = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 2, nr, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true).setAllowInvalid(false).build());
  sh.getRange(2, 3, nr, n).setNumberFormat('dd/MM/yyyy').setBackground('#fff8e6');
  sh.getRange(2, 3 + n, nr, n).setNumberFormat('dd/MM/yyyy').setBackground('#f5f2fb').setHorizontalAlignment('center');
  const cSet = nc + 2;
  sh.getRange(1, cSet, 1, 2).setValues([['Impostazioni', 'Valore']]).setFontWeight('bold').setBackground('#dfe7ec');
  sh.getRange(2, cSet, VEST_SET.length, 2).setValues(VEST_SET.map((r, i) => [r[0], L.set[i] != null ? L.set[i] : r[1]]));
  sh.getRange(2, cSet, VEST_SET.length, 1).setBackground('#dfe7ec').setFontWeight('bold');
  sh.getRange(2, cSet + 1, VEST_SET.length, 1).setBackground('#fff3d6').setFontWeight('bold');
  sh.getRange(VEST_SET.length + 3, cSet).setValue(VEST_NOTA).setWrap(true).setFontStyle('italic');
  sh.setFrozenRows(1); sh.setFrozenColumns(1); sh.setColumnWidth(1, 200); sh.setColumnWidth(cSet, 320); sh.setColumnWidth(nc, 220);
  PropertiesService.getScriptProperties().deleteProperty('VEST_CF');
}
function vestSetup_(ss) {
  if (VEST_MEMO) return VEST_MEMO;
  const sh = vestSheet_(ss); if (!sh) return null;
  const lock = LockService.getScriptLock(); let preso = false;
  try { preso = lock.tryLock(10000); } catch (e) {}
  try {
    const conv = convalideTutte_(ss).list;
    let L = mfLeggi_(sh);
    const extra = [...new Set(conv.map(x => x.cat).concat(L.keys))].filter(k => MF_BASE.indexOf(k) < 0).sort();
    const keys = MF_BASE.concat(extra), testa = mfTesta_(keys), n = keys.length;
    if (L.h.join('|') !== testa.join('|') || !L.cSet) { mfRiscrivi_(sh, L, keys); L = mfLeggi_(sh); }
    const cSet = L.cSet, nc = testa.length;
    // Impostazioni cercate per nome; quelle nuove nella prima riga libera
    const blocco = sh.getRange(2, cSet, 14, 2).getValues(), riga = [];
    VEST_SET_RE.forEach((re, i) => {
      let r = blocco.findIndex(x => re.test(String(x[0]).trim()));
      if (r < 0) { r = blocco.findIndex(x => !String(x[0]).trim() || /^giallo chiaro|^viola\b/i.test(String(x[0]).trim())); if (r < 0) r = blocco.length - 1;
        sh.getRange(2 + r, cSet, 1, 2).setValues([[VEST_SET[i][0], VEST_SET[i][1]]]); blocco[r] = [VEST_SET[i][0], VEST_SET[i][1]];
        sh.getRange(2 + r, cSet).setBackground('#dfe7ec').setFontWeight('bold'); sh.getRange(2 + r, cSet + 1).setBackground('#fff3d6').setFontWeight('bold'); }
      riga[i] = 2 + r;
    });
    const rNota = blocco.findIndex(x => /^giallo chiaro|^viola\b/i.test(String(x[0]).trim())), rGiusta = Math.max(...riga) + 2;
    if (rNota < 0 || 2 + rNota !== rGiusta) {
      if (rNota >= 0) sh.getRange(2 + rNota, cSet).clearContent().setFontStyle('normal');
      const qui = rGiusta - 2 < blocco.length ? String(blocco[rGiusta - 2][0]).trim() : '';
      if (!qui) sh.getRange(rGiusta, cSet).setValue(VEST_NOTA).setWrap(true).setFontStyle('italic');
    }
    const sv = riga.map(r => blocco[r - 2][1]), num = (x, d) => Number(x) > 0 ? Number(x) : d, tol = x => String(x).trim() === '' || !(Number(x) >= 0) ? 0 : Math.round(Number(x));
    const set = { durata: num(sv[0], 12), durMD: num(sv[1], 12), durMF: num(sv[2], 12), durC: num(sv[3], 12), avvV: num(sv[4], VEST_AVVISO), avvC: num(sv[5], VEST_AVVISO),
      tolV: tol(sv[6]), tolC: tol(sv[7]) };
    const durK = k => k === 'V' ? set.durata : k === 'MD_20' ? set.durMD : k === 'MF' ? set.durMF : set.durC;
    const piu = (d, m) => new Date(d.getFullYear(), d.getMonth() + m, d.getDate());
    // scadenze: "fatto il" + durata; categorie: la più lontana con lo scadenziario (e "fatto il" scritto dallo script se manca o è più vecchio)
    const nr = sh.getLastRow() - 1, eff = [], vest = [];
    if (nr > 0) {
      const rg = sh.getRange(2, 1, nr, nc), v = rg.getValues(), t = x => x instanceof Date ? x.getTime() : String(x);
      let cambia = false;
      v.forEach(r => {
        const nm = String(r[0]).trim(); if (!nm) return;
        const no = /^no$/i.test(String(r[1]).trim());
        keys.forEach((k, j) => {
          const cF = 2 + j, cS = 2 + n + j, prima = [r[cF], r[cS]];
          if (no) { r[cS] = 'esonerato'; }
          else {
            let d = r[cF] instanceof Date ? r[cF] : dataConv_(r[cF]);
            const x = k !== 'V' ? conv.find(c => c.cat === k && c.man.toUpperCase() === nm.toUpperCase()) : null;
            if (x) { const base = piu(new Date(x.scad), -durK(k)); if (!d || d.getTime() < base.getTime()) { d = base; r[cF] = base; } }
            let sc = d ? piu(d, durK(k)).getTime() : null;
            if (x && (!sc || x.scad > sc)) sc = x.scad;
            r[cS] = sc ? new Date(sc) : 'MANCA';
            if (sc) { if (k === 'V') vest.push({ man: nm, data: d ? d.getTime() : null, mesi: set.durata, scad: sc });
              else eff.push({ man: nm, cat: k, desc: x ? x.desc : (k === 'MD_20' ? 'Corso MD_20' : k === 'MF' ? 'Partecipazione MF' : ''), scad: sc, da: x && sc === x.scad ? 'scadenziario' : 'fatto il' }); }
          }
          if (t(prima[0]) !== t(r[cF]) || t(prima[1]) !== t(r[cS])) cambia = true;
        });
      });
      if (cambia) { sh.getRange(2, 3, nr, 2 * n).setValues(v.map(r => r.slice(2, 2 + 2 * n))); }
    }
    // colori: grigio esonerati; scadenze: rosso oltre tolleranza / MANCA, arancio in tolleranza, giallo in preavviso
    const firma = 'mf1|' + keys.join(',') + '|' + cSet + '|' + riga.join('.');
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('VEST_CF') !== firma) {
      const R = (f, bg, fg, rngs) => { let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(F_(f)).setRanges(rngs); if (bg) b = b.setBackground(bg); if (fg) b = b.setFontColor(fg); return b.build(); };
      const L1 = c => sh.getRange(1, c).getA1Notation().replace(/\d+$/, ''), abs = r => '$' + L1(cSet + 1) + '$' + r, righe = Math.max(sh.getMaxRows() - 1, 1);
      const cV = 3 + n + keys.indexOf('V'), altre = keys.map((k, j) => 3 + n + j).filter(c => c !== cV);
      const regole = [R('=TRIM($B2)="NO"', '#f1f3f4', '#9aa0a6', [sh.getRange(2, 1, righe, nc)])];
      const set3 = (c0, rngs, rTol, rAvv) => { const a = L1(c0) + '2';
        regole.push(R('=OR(' + a + '="MANCA",AND(ISNUMBER(' + a + '),' + a + '+N(' + abs(rTol) + ')<TODAY()))', '#f6c7bf', '#a3271a', rngs),
          R('=AND(ISNUMBER(' + a + '),' + a + '<TODAY())', '#fcd5b0', '#8a3f00', rngs), R('=AND(ISNUMBER(' + a + '),' + a + '-TODAY()<=' + abs(rAvv) + ')', '#fde9a9', '#6b4600', rngs)); };
      set3(cV, [sh.getRange(2, cV, righe, 1)], riga[6], riga[4]);
      set3(altre[0], altre.map(c => sh.getRange(2, c, righe, 1)), riga[7], riga[5]);
      sh.setConditionalFormatRules(regole);
      props.setProperty('VEST_CF', firma);
    }
    return VEST_MEMO = { sh: sh, cSter: 2, keys: keys, cSet: cSet, durata: set.durata, durMD: set.durMD, durMF: set.durMF, durC: set.durC, avvV: set.avvV, avvC: set.avvC,
      tolV: set.tolV, tolC: set.tolC, vest: vest, conv: eff.sort((a, b) => a.man.localeCompare(b.man) || a.cat.localeCompare(b.cat)) };
  } finally { if (preso) try { lock.releaseLock(); } catch (e) {} }
}
function esonerati_(ss) {
  const c = vestSetup_(ss); if (!c || c.sh.getLastRow() < 2) return [];
  return c.sh.getRange(2, 1, c.sh.getLastRow() - 1, 2).getValues().filter(r => String(r[0]).trim() && /^no$/i.test(String(r[1]).trim())).map(r => String(r[0]).trim());
}
function vestSheet_(ss) {
  rinominaMF_(ss);
  let sh = ss.getSheetByName(SH_VEST);
  if (sh) return sh;
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return null;
    sh = ss.getSheetByName(SH_VEST); if (sh) return sh;
    sh = ss.insertSheet(SH_VEST);
    const nomi = tecniciTutti_(ss).sort();
    sh.getRange(1, 1, 1, 2).setValues([['Tecnico', VEST_STER]]);
    if (nomi.length) sh.getRange(2, 1, nomi.length, 1).setValues(nomi.map(n => [n]));
    return sh;
  } catch (e) { Logger.log(e); return ss.getSheetByName(SH_VEST); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}
// Per l'app: tecnici con la data della vestizione e la scadenza calcolata qui
function vestizioni_(ss) { const c = vestSetup_(ss); return c ? c.vest : []; }
function rigaVestizioni_() {
  let v; try { const ss = SpreadsheetApp.getActiveSpreadsheet(), no = esonerati_(ss).map(x => x.toUpperCase()); v = vestizioni_(ss).filter(x => no.indexOf(x.man.toUpperCase()) < 0); } catch (e) { Logger.log(e); return ''; }
  const oggi = dateOnly_(new Date()).getTime(), gg = x => Math.round((x.scad - oggi) / 864e5);
  const corto = x => cognome_(x.man) + ' ' + gm_(new Date(x.scad));
  const cfg = (() => { try { return vestSetup_(SpreadsheetApp.getActiveSpreadsheet()) || {}; } catch (e) { return {}; } })(), avv = cfg.avvV || VEST_AVVISO, tl = cfg.tolV || 0;
  const inScad = v.filter(x => gg(x) >= 0 && gg(x) <= avv).sort((a, b) => a.scad - b.scad), inTol = v.filter(x => gg(x) < 0 && -gg(x) <= tl).sort((a, b) => a.scad - b.scad),
    scadute = v.filter(x => -gg(x) > tl).sort((a, b) => a.scad - b.scad);
  if (!inScad.length && !inTol.length && !scadute.length) return '';
  return '🧤 VESTIZIONI' + (inScad.length ? ' · in scadenza: ' + inScad.map(corto).join(', ') : '') + (inTol.length ? ' · in tolleranza fino al: ' + inTol.map(x => cognome_(x.man) + ' ' + gm_(piuGiorni_(x.scad, tl))).join(', ') : '') +
    (scadute.length ? ' · scadute: ' + scadute.map(corto).join(', ') : '');
}

function piuGiorni_(t, n) { const d = new Date(t); d.setDate(d.getDate() + n); return d; }
// Versione dello script: compare nel foglio KPI Risorse e nel menu, per controllare di aver incollato l'ultimo Code.gs
const VERSIONE_SCRIPT = '30/10/2026 21:40';
function infoVersione() { SpreadsheetApp.getActiveSpreadsheet().toast('Code.gs versione ' + VERSIONE_SCRIPT, 'Pharma-Fix', 8); }
/* =================== KPI Risorse (foglio "KPI Risorse") ===================
   Indicatori per tecnico, confrontati in modo equo: ogni intervento è paragonato agli interventi dello STESSO guasto
   sulla STESSA macchina nel periodo (MTTR relativo, 1,00 = come il reparto). Filtri in alto (periodo, area, minimo);
   si ricalcola da solo quando li cambi o dal menu. Per responsabile e coordinatori: è uno strumento per decidere
   affiancamenti e formazione, non una classifica. */
const SH_KR = 'KPI Risorse', KR_AREE = ['Tutte', 'Filling', 'Packaging'], KR_T = 7;
const KR_H = ['Tecnico', 'Interventi', 'Ore di intervento', 'Turni lavorati (foglio Turni)', 'Interventi per turno', 'MTTR relativo (1,00 = reparto)', 'Interventi confrontabili',
  'Ricorrenza entro 24h', 'Registrati in tempo (Packaging, col QR)', 'Con descrizione', 'Anomalie segnalate', 'Soluzioni approvate', 'Compiti fatti in tempo', 'Quadrante'];
function krSheet_(ss) {
  let sh = ss.getSheetByName(SH_KR);
  if (sh) return sh;
  sh = ss.insertSheet(SH_KR);
  sh.getRange('A1').setValue('KPI RISORSE').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('Confronti equi: ogni intervento è paragonato agli interventi dello stesso guasto sulla stessa macchina. Strumento per affiancamenti e formazione, non una classifica.').setFontStyle('italic').setFontColor('#5f6368');
  sh.getRange('A3:H3').setValues([['Periodo', 'Ultimi 90 giorni', '', 'Area', 'Tutte', '', 'Minimo interventi confrontabili', 5]]);
  sh.getRange('A4:E4').setValues([['Dal (personalizzato)', '', '', 'Al', '']]);
  ['A3', 'D3', 'G3', 'A4', 'D4'].forEach(c => sh.getRange(c).setFontWeight('bold'));
  ['B3', 'E3', 'H3', 'B4', 'E4'].forEach(c => sh.getRange(c).setBackground('#fff3d6'));
  sh.getRange('B3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PERIODI, true).setAllowInvalid(false).build());
  sh.getRange('E3').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(KR_AREE, true).setAllowInvalid(false).build());
  sh.getRange('B4:E4').setNumberFormat('dd/MM/yyyy');
  sh.setColumnWidth(1, 200); sh.setFrozenRows(0);
  return sh;
}
const medianaKr_ = a => { const v = a.slice().sort((x, y) => x - y), n = v.length; return n ? (n % 2 ? v[(n - 1) / 2] : (v[n / 2 - 1] + v[n / 2]) / 2) : null; };
function kpiRisorseCalcola_(ss, filtri) {
  const per = periodo_(filtri.periodo, filtri.dal, filtri.al, dateOnly_(new Date()));
  const areaDi = areaDi_(), A = String(filtri.area || 'Tutte').toUpperCase();
  const tutti = fermiTutti_().filter(f => f.b && f.stato === 'CHIUSO');
  tutti.forEach(f => { f.area = areaStd_(areaDi(f)); f.k = (f.eq || kpiChiave_(f.linea, f.macc)) + '|' + String(f.guasto).trim().toUpperCase(); f.m = minDi_(f); });
  const nel = tutti.filter(f => f.a >= per.from && f.a < per.toEnd && (A === 'TUTTE' || f.area === A));
  // mediana per macchina + guasto (almeno 3 interventi)
  const grp = {}; nel.forEach(f => { if (String(f.guasto).trim()) (grp[f.k] = grp[f.k] || []).push(f.m); });
  const med = {}; Object.keys(grp).forEach(k => { if (grp[k].length >= 3) med[k] = medianaKr_(grp[k]); });
  // ricorrenza: lo stesso guasto torna sulla stessa macchina entro 24 ore dalla fine dell'intervento
  const perK = {}; tutti.forEach(f => (perK[f.k] = perK[f.k] || []).push(f));
  const torna = f => String(f.guasto).trim() && (perK[f.k] || []).some(g => g !== f && g.a > f.b && g.a - f.b <= 864e5);
  // turni lavorati nel periodo (foglio Turni)
  let tt = {}; try { tt = turniTutti_(ss); } catch (e) { Logger.log(e); }
  const da = Utilities.formatDate(per.from, TZ, 'yyyy-MM-dd'), a = Utilities.formatDate(new Date(Math.min(per.toEnd.getTime(), Date.now()) - 1), TZ, 'yyyy-MM-dd');
  const turniDi = n => { const k = Object.keys(tt).find(x => x.toUpperCase() === n.toUpperCase()); if (!k) return null;
    const giorni = Object.keys(tt[k]).filter(d => d >= da && d <= a); if (!giorni.length) return null;
    return { lav: giorni.filter(d => /^[123G]$/.test(tt[k][d].split(' ')[0])), tutti: giorni }; };
  const ano = {}; anomalieTutte_().forEach(x => { if (x.ts && x.ts >= per.from && x.ts < per.toEnd && x.da) { const n = String(x.da).trim().toUpperCase(); ano[n] = (ano[n] || 0) + 1; } });
  const sol = {}; righe_(ss, SH_SOL, SOL_H.length).forEach(r => { const d = r[1] instanceof Date ? r[1] : null, st = String(r[11] || 'VERIFICATA').toUpperCase();
    if (d && d >= per.from && d < per.toEnd && st === 'VERIFICATA' && r[10]) { const n = String(r[10]).trim().toUpperCase(); sol[n] = (sol[n] || 0) + 1; } });
  const kt = {}; try { compitiTutti_().forEach(k => { if (k.stato === 'FATTO' && k.fatto >= per.from.getTime() && k.fatto < per.toEnd.getTime() && k.fattoDa) {
    const n = String(k.fattoDa).trim().toUpperCase(); kt[n] = kt[n] || { f: 0, t: 0 }; kt[n].f++; if (!k.ritardo) kt[n].t++; } }); } catch (e) { Logger.log(e); }
  // i coordinatori (colonna Coordinatori del foglio Liste) non compaiono: intervenire non è il loro ruolo principale.
  // I loro interventi restano nel calcolo delle mediane del reparto.
  const coord = coordinatori_().map(x => x.toUpperCase());
  const nomi = [...new Set(nel.map(f => String(f.man).trim()).filter(Boolean))].filter(n => coord.indexOf(n.toUpperCase()) < 0);
  const rows = nomi.map(n => {
    const fs = nel.filter(f => String(f.man).trim() === n), N = n.toUpperCase();
    const rel = fs.filter(f => med[f.k] > 0).map(f => f.m / med[f.k]);
    const app = fs.filter(f => String(f.orari) !== ORA_CARTA), pk = app.filter(f => f.area === 'PACKAGING');
    const tu = turniDi(n);
    const nT = tu ? fs.filter(f => tu.tutti.indexOf(Utilities.formatDate(f.a, TZ, 'yyyy-MM-dd')) >= 0).length : 0; // interventi nei giorni coperti dal foglio Turni
    return { n: n, c: fs.length, ore: fs.reduce((x, f) => x + f.m, 0) / 60, turni: tu ? tu.lav.length : null, xt: tu && tu.lav.length ? nT / tu.lav.length : null,
      rel: rel.length >= filtri.min ? medianaKr_(rel) : null, nrel: rel.length, ric: fs.filter(torna).length / fs.length,
      tempo: pk.length ? pk.filter(f => String(f.orari) !== 'DICHIARATI').length / pk.length : null,
      desc: app.length ? app.filter(f => String(f.altro).trim()).length / app.length : null,
      ano: ano[N] || 0, sol: sol[N] || 0, comp: kt[N] ? kt[N].t + ' su ' + kt[N].f : '' };
  }).sort((x, y) => y.c - x.c);
  const mx = medianaKr_(rows.filter(r => r.xt !== null).map(r => r.xt));
  rows.forEach(r => { r.q = r.rel === null || r.xt === null || mx === null ? '' : r.rel <= 1 ? (r.xt >= mx ? 'Riferimento' : 'Da coinvolgere di più') : (r.xt >= mx ? 'Carico alto: verificare tempi' : 'Affiancamento / formazione'); });
  // matrice competenze: interventi per linea
  const linee = lineeTutte_(), mat = rows.map(r => linee.map(l => nel.filter(f => String(f.man).trim() === r.n && String(f.linea).trim() === l).length));
  const esperti = linee.map((l, j) => mat.filter(x => x[j] >= 5).length);
  // ricorrenza per tipo di guasto (i 10 più frequenti)
  const pg = {}; nel.forEach(f => { const g = String(f.guasto).trim().toUpperCase(); if (!g) return; const x = pg[g] = pg[g] || { n: 0, r: 0 }; x.n++; if (torna(f)) x.r++; });
  const guasti = Object.keys(pg).sort((x, y) => pg[y].n - pg[x].n).slice(0, 10).map(g => ({ g: g, n: pg[g].n, r: pg[g].r / pg[g].n }));
  return { per: per, rows: rows, linee: linee, mat: mat, esperti: esperti, guasti: guasti, nFermi: nel.length };
}
// Grafico a quadranti. Entrambi gli assi sono "rispetto al reparto" (1 = come il reparto) e in scala logaritmica centrata su 1:
// così le due linee di riferimento passano esattamente al centro e i 4 quadranti hanno la stessa grandezza.
// X = MTTR relativo, Y = interventi per turno ÷ mediana del reparto. Una serie per tecnico (nome in legenda).
// Sotto il grafico, gli stessi 4 quadranti in celle colorate con i nomi. Dati nel foglio nascosto KPI_Risorse_dati.
const KR_Q = [['Riferimento', '↖ RIFERIMENTI · più veloci e più presenti', '#dff1e7'], ['Carico alto: verificare tempi', '↗ CARICO ALTO · più presenti ma più lenti', '#fde9a9'],
  ['Da coinvolgere di più', '↙ DA COINVOLGERE · veloci ma meno presenti', '#e3edf3'], ['Affiancamento / formazione', '↘ AFFIANCAMENTO · più lenti e meno presenti', '#fbe4e0']];
function krGrafico_(ss, sh, K) {
  sh.getCharts().forEach(c => sh.removeChart(c));
  const pts = K.rows.filter(r => r.rel !== null && r.xt !== null && r.xt > 0);
  let d = ss.getSheetByName('KPI_Risorse_dati');
  if (!d) { d = ss.insertSheet('KPI_Risorse_dati'); d.hideSheet(); }
  d.clear();
  const C0 = KR_H.length + 2;
  if (pts.length < 2) { sh.getRange(KR_T - 1, C0).setValue('Grafico a quadranti: servono almeno 2 tecnici con MTTR relativo e turni (foglio Turni).').setFontColor('#5f6368'); return; }
  const med = medianaKr_(pts.map(p => p.xt)), Y = p => p.xt / med;
  const k = Math.max(1.6, ...pts.map(p => Math.max(p.rel, 1 / p.rel, Y(p), 1 / Y(p)))) * 1.25; // stesso margine sui due lati di 1
  const nS = pts.length + 2, rows = [];
  rows.push(['MTTR relativo'].concat(pts.map(p => cognome_(p.n))).concat(['┃ MTTR come il reparto', '━ interventi come la mediana']));
  pts.forEach((p, i) => { const r = new Array(nS + 1).fill(''); r[0] = Math.round(p.rel * 100) / 100; r[i + 1] = Math.round(Y(p) * 100) / 100; rows.push(r); });
  // la croce dei quadranti: Google Sheets nel grafico a dispersione non disegna le linee impostate dallo script,
  // quindi ogni linea è fatta di tanti punti piccoli e fitti (da 1/k a k, equidistanti in scala logaritmica)
  const N = 120;
  for (let i = 0; i <= N; i++) { const v = Math.pow(k, 2 * i / N - 1);
    const a = new Array(nS + 1).fill(''); a[0] = 1; a[nS - 1] = v; rows.push(a);
    const b = new Array(nS + 1).fill(''); b[0] = v; b[nS] = 1; rows.push(b); }
  d.getRange(1, 1, rows.length, nS + 1).setValues(rows);
  const serie = {}; pts.forEach((p, i) => { serie[i] = { pointSize: 10, lineWidth: 0 }; });
  serie[pts.length] = { pointSize: 3, lineWidth: 0, color: '#757575' };
  serie[pts.length + 1] = { pointSize: 3, lineWidth: 0, color: '#757575' };
  const tick = [0.2, 0.4, 0.6, 0.8, 1, 2, 4].filter(t => t >= 1 / k && t <= k);
  const ch = sh.newChart().asScatterChart().addRange(d.getRange(1, 1, rows.length, nS + 1)).setNumHeaders(1)
    .setOption('title', '↖ Riferimenti   ↗ Carico alto   ↙ Da coinvolgere   ↘ Affiancamento')
    .setOption('titleTextStyle', { fontSize: 18, bold: true, color: '#5f6368' })
    .setOption('hAxis', { title: 'MTTR relativo · 1 = come il reparto  (← più veloce · più lento →)', logScale: true, viewWindow: { min: 1 / k, max: k }, ticks: tick, gridlines: { color: '#eeeeee' } })
    .setOption('vAxis', { title: 'Interventi per turno · 1 = mediana del reparto  (↑ più presente)', logScale: true, viewWindow: { min: 1 / k, max: k }, ticks: tick, gridlines: { color: '#eeeeee' } })
    .setOption('legend', { position: 'right' }).setOption('series', serie)
    .setOption('width', 760).setOption('height', 560)
    .setPosition(KR_T, C0, 0, 0).build();
  sh.insertChart(ch);
  // i 4 quadranti in celle colorate, sotto il grafico
  const r0 = KR_T + 29;
  sh.getRange(r0 - 1, C0).setValue('QUADRANTI').setFontWeight('bold');
  KR_Q.forEach((q, i) => {
    const nomi = K.rows.filter(r => r.q === q[0]).map(r => cognome_(r.n) + ' (' + String(Math.round(r.rel * 100) / 100).replace('.', ',') + ' · ' + String(Math.round(r.xt * 100) / 100).replace('.', ',') + ')');
    sh.getRange(r0 + Math.floor(i / 2), C0 + (i % 2)).setValue(q[1] + '\n' + (nomi.join('\n') || '—')).setBackground(q[2]).setWrap(true).setVerticalAlignment('top').setFontSize(10);
  });
  sh.setColumnWidth(C0, 300); sh.setColumnWidth(C0 + 1, 300); sh.setRowHeight(r0, 110); sh.setRowHeight(r0 + 1, 110);
  sh.getRange(r0 + 2, C0).setValue('Tra parentesi: MTTR relativo · interventi per turno. Linee del grafico: 1 = come il reparto.').setFontStyle('italic').setFontColor('#5f6368');
}
function aggiornaKpiRisorse() {
  try { completaFermiCarta_(); } catch (e) { Logger.log(e); }
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = krSheet_(ss);
  const f = sh.getRange('B3:H4').getValues();
  const filtri = { periodo: String(f[0][0] || 'Ultimi 90 giorni'), area: String(f[0][3] || 'Tutte'), min: Number(f[0][6]) > 0 ? Number(f[0][6]) : 5, dal: f[1][0], al: f[1][3] };
  const K = kpiRisorseCalcola_(ss, filtri);
  sh.getRange(KR_T - 1, 1, Math.max(sh.getMaxRows() - KR_T + 2, 1), Math.max(sh.getMaxColumns(), 16)).clear();
  sh.getRange(KR_T - 1, 1).setValue('Periodo ' + Utilities.formatDate(K.per.from, TZ, 'dd/MM/yyyy') + ' – ' + Utilities.formatDate(K.per.to, TZ, 'dd/MM/yyyy') + ' · ' + nFermi_(K.nFermi) + ' chiusi · aggiornato ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + ' · script ' + VERSIONE_SCRIPT).setFontColor('#5f6368');
  let r = KR_T;
  sh.getRange(r, 1, 1, KR_H.length).setValues([KR_H]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true).setVerticalAlignment('bottom');
  const pc = x => x === null ? '—' : x, nr = K.rows.length;
  if (nr) {
    const vals = K.rows.map(x => [x.n, x.c, Math.round(x.ore * 10) / 10, pc(x.turni), x.xt === null ? '—' : Math.round(x.xt * 100) / 100, x.rel === null ? '—' : Math.round(x.rel * 100) / 100, x.nrel,
      x.ric, pc(x.tempo), pc(x.desc), x.ano, x.sol, x.comp, x.q]);
    const rng = sh.getRange(r + 1, 1, nr, KR_H.length);
    rng.setValues(vals);
    sh.getRange(r + 1, 8, nr, 3).setNumberFormat('0%');
    const bg = K.rows.map(x => { const row = new Array(KR_H.length).fill(null);
      if (x.rel !== null) row[5] = x.rel <= 0.85 ? '#dff1e7' : x.rel >= 1.2 ? '#fde9a9' : null;
      if (x.tempo !== null) row[8] = x.tempo >= 0.8 ? '#dff1e7' : x.tempo < 0.6 ? '#fde9a9' : null;
      if (x.desc !== null) row[9] = x.desc >= 0.9 ? '#dff1e7' : x.desc < 0.7 ? '#fde9a9' : null;
      return row; });
    rng.setBackgrounds(bg);
    sh.getRange(r + 1, 1, nr, 1).setFontWeight('bold');
  } else sh.getRange(r + 1, 1).setValue('Nessun fermo chiuso nel periodo.');
  r += nr + 2;
  sh.getRange(r, 1).setValue('MTTR relativo: verde ≤ 0,85 (più veloce del reparto sugli stessi guasti), giallo ≥ 1,20; "—" = meno di ' + filtri.min + ' interventi confrontabili. Interventi per turno: solo nei giorni presenti nel foglio Turni. Quadrante: MTTR relativo rispetto a 1,00 e interventi per turno rispetto alla mediana del reparto.').setFontStyle('italic').setFontColor('#5f6368');
  // matrice
  r += 2;
  sh.getRange(r, 1).setValue('MATRICE COMPETENZE · interventi per linea').setFontWeight('bold').setFontSize(12);
  r++;
  sh.getRange(r, 1, 1, K.linee.length + 1).setValues([['Tecnico'].concat(K.linee)]).setFontWeight('bold').setBackground('#dfe7ec');
  if (nr) {
    const mxv = Math.max(1, ...K.mat.map(x => Math.max(...x)));
    const mrng = sh.getRange(r + 1, 1, nr, K.linee.length + 1);
    mrng.setValues(K.rows.map((x, i) => [x.n].concat(K.mat[i].map(v => v || ''))));
    mrng.setBackgrounds(K.mat.map(x => [null].concat(x.map(v => { if (!v) return null; const t = 0.12 + 0.75 * v / mxv; const c = c0 => Math.round(255 - (255 - c0) * t); return '#' + [36, 90, 122].map(c).map(z => ('0' + z.toString(16)).slice(-2)).join(''); }))));
    sh.getRange(r + 1, 2, nr, K.linee.length).setHorizontalAlignment('center');
  }
  r += nr + 1;
  sh.getRange(r, 1, 1, K.linee.length + 1).setValues([['Esperti (≥5 interventi)'].concat(K.esperti.map(e => e <= 1 ? e + ' ⚠' : e))]).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(r, 2, 1, K.linee.length).setBackgrounds([K.esperti.map(e => e <= 1 ? '#fde9a9' : null)]);
  r += 1;
  sh.getRange(r, 1).setValue('⚠ = una sola persona (o nessuna) conosce bene la linea: candidata ad affiancamenti.').setFontStyle('italic').setFontColor('#5f6368');
  // ricorrenza per guasto
  r += 2;
  sh.getRange(r, 1).setValue('RICORRENZA PER TIPO DI GUASTO · stesso guasto sulla stessa macchina entro 24 ore').setFontWeight('bold').setFontSize(12);
  r++;
  sh.getRange(r, 1, 1, 3).setValues([['Tipo guasto', 'Interventi', 'Ricorrenza 24h']]).setFontWeight('bold').setBackground('#dfe7ec');
  if (K.guasti.length) { sh.getRange(r + 1, 1, K.guasti.length, 3).setValues(K.guasti.map(g => [g.g, g.n, g.r])); sh.getRange(r + 1, 3, K.guasti.length, 1).setNumberFormat('0%'); }
  sh.setColumnWidth(1, 220); for (let c = 2; c <= KR_H.length; c++) sh.setColumnWidth(c, c === KR_H.length ? 210 : 105);
  try { krGrafico_(ss, sh, K); } catch (e) { Logger.log(e); }
  sh.setRowHeight(KR_T, 54);
  return K;
}

/* =================== Scadenziario MF (prima foglio Convalide) ===================
   Si incolla lo scadenziario così com'è (Codice processo · Oggetto · Scadenza · Item). Lo script riconosce:
   la persona dal testo dell'Oggetto ("Convalida Sterile Marco Gialli…" → GIALLI Marco, confrontando con Manutentori),
   la categoria dall'Item ("Categoria G - Intervento Meccanico di regolazione") e la data (Date, 2026/10/22 o 22/10/2026).
   Se la stessa persona ha più righe per la stessa categoria vale la scadenza più lontana. */
const SH_CONV = 'Scadenziario MF', SH_CONV_OLD = 'Convalide', CONV_H = ['Codice processo (link)', 'Oggetto', 'Scadenza', 'Item'];
function convSheet_(ss) {
  rinominaMF_(ss);
  let sh = ss.getSheetByName(SH_CONV);
  if (sh) return sh;
  sh = ss.insertSheet(SH_CONV);
  sh.getRange(1, 1, 1, CONV_H.length).setValues([CONV_H]).setFontWeight('bold').setBackground('#dfe7ec');
  sh.getRange(1, 6).setValue('Incolla qui lo scadenziario MF completo (solo valori): basta che ci siano Oggetto, Scadenza e Item. Da qui lo script riempie Cat. G e Cat. H nel foglio Convalide MF.');
  sh.setFrozenRows(1); sh.setColumnWidth(2, 380); sh.setColumnWidth(4, 380);
  return sh;
}
const normNome_ = x => String(x || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z' ]+/g, ' ').replace(/'/g, ' ').trim().split(/\s+/).filter(Boolean);
function dataConv_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/); if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/); if (m) return new Date(+m[3] < 100 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
  return null;
}
let CONV_MEMO = null;
function convalideTutte_(ss) { return CONV_MEMO || (CONV_MEMO = convalideLeggi_(ss)); }
function convalideLeggi_(ss) {
  const sh = convSheet_(ss);
  if (!sh || sh.getLastRow() < 1) return { list: [], ignote: [] };
  const man = tecniciTutti_(ss).map(n => ({ n: n, t: normNome_(n) }));
  const best = {}, ignote = [];
  sh.getDataRange().getValues().forEach(r => {
    const txt = r.map(x => x instanceof Date ? '' : String(x));
    let iItem = txt.findIndex(x => /categoria\s+[A-Z]\b/i.test(x)), cat, desc;
    if (iItem >= 0) { cat = txt[iItem].match(/categoria\s+([A-Z])\b/i)[1].toUpperCase(); desc = (txt[iItem].split(/\s+-\s+/)[1] || '').replace(/\s+/g, ' ').trim(); }
    else if ((iItem = txt.findIndex(x => /\bMD[\s_-]?20\b/i.test(x))) >= 0) { cat = 'MD_20'; desc = 'Corso MD_20'; } // corso MD_20 nello scadenziario
    else if ((iItem = txt.findIndex(x => /media\s*fill|partecipazione\s+(al\s+)?mf\b/i.test(x))) >= 0) { cat = 'MF'; desc = 'Partecipazione MF'; } // partecipazione Media Fill
    else return;
    let d = null; for (const x of r) { if ((d = dataConv_(x))) break; } if (!d) return;
    const ogg = txt.filter((x, i) => i !== iItem).sort((a, b) => b.length - a.length)[0] || '';
    const tok = normNome_(ogg);
    const chi = man.filter(m => m.t.length && m.t.every(w => tok.indexOf(w) >= 0)).sort((a, b) => b.t.length - a.t.length)[0];
    if (!chi) { const k = ogg.replace(/\s+/g, ' ').trim(); if (k && ignote.indexOf(k) < 0) ignote.push(k); return; }
    const key = chi.n + '|' + cat;
    if (!best[key] || best[key].scad < d.getTime()) best[key] = { man: chi.n, cat: cat, desc: desc, scad: d.getTime() };
  });
  return { list: Object.keys(best).map(k => best[k]).sort((a, b) => a.man.localeCompare(b.man) || a.cat.localeCompare(b.cat)), ignote: ignote };
}
// Convalide effettive (scadenziario + data "fatto il" dal foglio Convalide MF)
function convalideEff_(ss) { let c = null; try { c = vestSetup_(ss); } catch (e) { Logger.log(e); } return c && c.conv ? c.conv : convalideTutte_(ss).list; }
function convalide_(ss) { const x = convalideTutte_(ss), list = convalideEff_(ss); return { list: list, ignote: x.ignote.slice(0, 30), cats: [...new Set(list.map(c => c.cat).concat(CONV_FISSE))].sort() }; }
function convMancanti_(ss, list) {
  const no = esonerati_(ss).map(x => x.toUpperCase()), cats = [...new Set(list.map(x => x.cat).concat(CONV_FISSE))].sort();
  if (!cats.length) return [];
  const nomi = tecniciTutti_(ss).filter(n => no.indexOf(n.toUpperCase()) < 0);
  const out = []; nomi.forEach(n => { const m = cats.filter(k => !list.some(x => x.man.toUpperCase() === n.toUpperCase() && x.cat === k)); if (m.length) out.push({ man: n, cats: m }); });
  return out;
}
function rigaConvalide_() {
  let v, manca = [];
  try { const ss = SpreadsheetApp.getActiveSpreadsheet(), no = esonerati_(ss).map(x => x.toUpperCase()); v = convalideEff_(ss).filter(x => no.indexOf(x.man.toUpperCase()) < 0); manca = convMancanti_(ss, v); }
  catch (e) { Logger.log(e); return ''; }
  const oggi = dateOnly_(new Date()).getTime(), gg = x => Math.round((x.scad - oggi) / 864e5);
  const corto = x => cognome_(x.man) + ' ' + x.cat + ' ' + gm_(new Date(x.scad));
  const cfg = (() => { try { return vestSetup_(SpreadsheetApp.getActiveSpreadsheet()) || {}; } catch (e) { return {}; } })(), avv = cfg.avvC || VEST_AVVISO, tl = cfg.tolC || 0;
  const inScad = v.filter(x => gg(x) >= 0 && gg(x) <= avv).sort((a, b) => a.scad - b.scad), inTol = v.filter(x => gg(x) < 0 && -gg(x) <= tl).sort((a, b) => a.scad - b.scad),
    scadute = v.filter(x => -gg(x) > tl).sort((a, b) => a.scad - b.scad);
  if (!inScad.length && !inTol.length && !scadute.length && !manca.length) return '';
  return '🧾 CONVALIDE GMP' + (inScad.length ? ' · in scadenza: ' + inScad.map(corto).join(', ') : '') +
    (inTol.length ? ' · in tolleranza fino al: ' + inTol.map(x => cognome_(x.man) + ' ' + x.cat + ' ' + gm_(piuGiorni_(x.scad, tl))).join(', ') : '') + (scadute.length ? ' · scadute: ' + scadute.map(x => cognome_(x.man) + ' ' + x.cat).join(', ') : '') +
    (manca.length ? ' · mancanti: ' + manca.map(x => cognome_(x.man) + ' ' + x.cats.join(' ')).join(', ') : '');
}

/* =================== Turni (dal foglio Turni) ===================
   Ogni mese si incolla la tabella dei turni (riga con le date + una riga per persona) nel foglio Turni,
   sotto il mese precedente; vale anche un foglio per mese purché il nome inizi con "Turni".
   Celle: 1 / 2 / 3 / G = turno, R = riposo, F = ferie, A (o altro) = assente; "3 1" = fa il 3, era nell'1 (cambio). */
const SH_TURNI = 'Turni';
function turniSheet_(ss) {
  if (ss.getSheets().some(s => /^turni/i.test(s.getName()))) return;
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000) || ss.getSheets().some(s => /^turni/i.test(s.getName()))) return;
    const sh = ss.insertSheet(SH_TURNI);
    sh.getRange('A1').setValue('Incolla qui sotto la tabella dei turni di ogni mese, così com\'è nell\'Excel (riga delle date e una riga per persona, nomi scritti come in Liste). Il mese nuovo sotto il precedente.')
      .setFontStyle('italic').setFontColor('#4d5b66');
  } catch (e) { Logger.log(e); } finally { try { lock.releaseLock(); } catch (e) {} }
}
// Una cella turno scritta "3/1" (fa il 3, era nell'1) Fogli la trasforma in data (3 gennaio): la rileggo come "3 1"
function codCella_(v) {
  if (v instanceof Date) { const g = Number(Utilities.formatDate(v, TZ, 'd')), m = Number(Utilities.formatDate(v, TZ, 'M')); return g + ' ' + m; }
  return String(v).trim().toUpperCase().replace(/[\/\\-]/g, ' ').replace(/\s+/g, ' ');
}
// Tutti i turni letti: { NOME: { 'yyyy-MM-dd': '3 1' } }
function turniTutti_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const out = {}, key = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const data = x => {
    if (x instanceof Date && x.getFullYear() >= 2000) return x;
    const m = String(x || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    return m ? new Date(m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
  };
  ss.getSheets().filter(s => /^turni/i.test(s.getName())).forEach(sh => {
    if (sh.getLastRow() < 2) return;
    let date = null;
    sh.getDataRange().getValues().forEach(r => {
      const dd = r.slice(1).map(data);
      if (dd.filter(Boolean).length >= 7) { date = dd; return; } // riga delle date: inizia un mese
      const nome = String(r[0] || '').trim().replace(/\s+/g, ' ');
      if (!date || !nome) return;
      const p = out[nome] = out[nome] || {};
      r.slice(1).forEach((v, i) => { if (date[i] && v !== '' && v != null) p[key(date[i])] = codCella_(v); });
    });
  });
  return out;
}
// Per l'app: da 7 giorni fa a 35 giorni avanti
function turni_(ss) {
  turniSheet_(ss);
  const t = turniTutti_(ss), da = Utilities.formatDate(new Date(Date.now() - 7 * 864e5), TZ, 'yyyy-MM-dd'), a = Utilities.formatDate(new Date(Date.now() + 35 * 864e5), TZ, 'yyyy-MM-dd');
  const out = {};
  Object.keys(t).forEach(n => { const p = {}; Object.keys(t[n]).forEach(k => { if (k >= da && k <= a) p[k] = t[n][k]; }); if (Object.keys(p).length) out[n] = p; });
  return out;
}
const codTurno_ = v => { const p = String(v || '').trim().split(' '); return { c: p[0] || '', o: p[1] && p[1] !== p[0] ? p[1] : '' }; };
// Riga del passaggio di turno: chi era in quel turno e chi doveva esserci ma era assente
function rigaTurno_(a, t) {
  let tt; try { tt = turniTutti_(); } catch (e) { Logger.log(e); return ''; }
  const k = Utilities.formatDate(a, TZ, 'yyyy-MM-dd'), n = t.replace('T', ''), corto = s => cognome_(s);
  const dentro = [], assenti = [];
  Object.keys(tt).sort().forEach(nome => {
    const x = codTurno_(tt[nome][k]); if (!x.c) return;
    if (x.c === n) dentro.push(corto(nome) + (x.o ? ' (cambio)' : ''));
    else if (x.o === n && !/^[123G]$/.test(x.c)) assenti.push(corto(nome) + (x.c === 'C' ? ' (corso)' : ''));
  });
  if (!dentro.length && !assenti.length) return '';
  return '👷 In turno: ' + (dentro.join(', ') || 'nessuno nel foglio Turni') + (assenti.length ? ' · assenti: ' + assenti.join(', ') : '');
}

/* =================== Programma produzione (foglio Programma) ===================
   Si incolla TUTTO il foglio del programma (solo valori): riga con le date, colonna A con le righe del programma.
   Si usano solo le righe di riempimento ("… rip …") e di confezionamento (spl1, spl2), abbinate alle linee di Equipment. */
function lineaProg_(etichetta, linee) {
  const s = String(etichetta || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (!s) return '';
  const diretto = linee.find(l => l.toLowerCase() === s); if (diretto) return diretto;
  if (/\bprep\b/.test(s)) return '';
  const trova = n => linee.find(l => l.toUpperCase() === n) || '';
  if (/spl ?1\b/.test(s)) return trova('SPL1');
  if (/spl ?2\b/.test(s)) return trova('SPL2');
  if (!/\brip\b/.test(s)) return '';
  if (/\borl\b/.test(s)) return trova('ORL');
  const m = s.match(/\bst ?(\d)\b/); if (!m) return '';
  if (m[1] === '1') return /pomat|gel|tgm/.test(s) ? trova('ST1 TGM') : /liquid|mar/.test(s) ? trova('ST1 MAR') : '';
  return trova('ST' + m[1]);
}
/* Il foglio Programma incollato da Excel si porta dietro migliaia di colonne e righe vuote ma formattate
   (es. 16.000 colonne): il file diventa enorme e lento. Qui si tolgono quelle oltre l'ultima con un valore. */
function pulisciProgramma_(sh) {
  try {
    const lc = sh.getLastColumn(), mc = sh.getMaxColumns(), lr = sh.getLastRow(), mr = sh.getMaxRows();
    let fatto = false;
    // a blocchi da 2000 colonne: un'unica cancellazione di 16.000 colonne può andare in timeout
    for (let m = mc; m > lc + 5; m = sh.getMaxColumns()) { const n = Math.min(2000, m - lc - 5); sh.deleteColumns(m - n + 1, n); fatto = true; SpreadsheetApp.flush(); }
    if (mr > lr + 20) { sh.deleteRows(lr + 21, mr - lr - 20); fatto = true; }
    return fatto;
  } catch (e) { Logger.log(e); return false; }
}
function alleggerisciProgramma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheets().find(s => /^programma/i.test(s.getName()));
  if (!sh) { ss.toast('Foglio Programma non trovato.', 'Pharma-Fix', 6); return; }
  const prima = sh.getMaxColumns() + ' colonne × ' + sh.getMaxRows() + ' righe';
  pulisciProgramma_(sh);
  ss.toast('Foglio Programma alleggerito: da ' + prima + ' a ' + sh.getMaxColumns() + ' × ' + sh.getMaxRows() + '.', 'Pharma-Fix', 8);
}
function programmaTutto_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets().find(s => /^programma/i.test(s.getName()));
  if (!sh || sh.getLastRow() < 3) return {};
  // (la pulizia delle colonne vuote NON si fa qui: su un foglio enorme blocca la risposta all'app; solo dal menu)
  const v = sh.getDataRange().getValues(), linee = [...new Set(equipment_(ss).map(e => String(e.linea).trim()))];
  const data = x => x instanceof Date && x.getFullYear() >= 2000 ? x : null;
  let hr = -1, best = 0; // riga delle date: quella con più date (nel file del programma è la 3)
  v.slice(0, 10).forEach((r, i) => { const n = r.filter(data).length; if (n > best) { best = n; hr = i; } });
  if (hr < 0 || best < 7) return {};
  const date = v[hr].map(x => data(x) ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd') : '');
  const out = {};
  v.slice(hr + 1).forEach(r => {
    const l = lineaProg_(r[0], linee); if (!l) return;
    const p = out[l] = out[l] || {};
    r.forEach((x, i) => { if (i && date[i] && String(x).trim()) p[date[i]] = (p[date[i]] ? p[date[i]] + '\n\n' : '') + String(x).trim(); });
  });
  return out;
}
// Per l'app: da 3 giorni fa a 14 giorni avanti
function programma_(ss) {
  const t = programmaTutto_(ss), da = Utilities.formatDate(new Date(Date.now() - 3 * 864e5), TZ, 'yyyy-MM-dd'), a = Utilities.formatDate(new Date(Date.now() + 14 * 864e5), TZ, 'yyyy-MM-dd');
  const out = {};
  Object.keys(t).forEach(l => { const p = {}; Object.keys(t[l]).forEach(k => { if (k >= da && k <= a) p[k] = t[l][k]; }); if (Object.keys(p).length) out[l] = p; });
  return out;
}
// Passaggio di turno: cambi formato del giorno che arriva (oggi se il turno finisce al mattino, domani altrimenti)
function rigaCambiFormato_(b) {
  let t; try { t = programmaTutto_(); } catch (e) { Logger.log(e); return ''; }
  const mattino = Number(Utilities.formatDate(b, TZ, 'H')) < 12, g = new Date(b.getTime() + (mattino ? 0 : 864e5)), k = Utilities.formatDate(g, TZ, 'yyyy-MM-dd');
  const cf = [];
  Object.keys(t).sort().forEach(l => String(t[l][k] || '').split(/\n\s*\n/).forEach(x => {
    x = x.replace(/\s+/g, ' ').trim(); if (/^cf\b/i.test(x)) cf.push(l + (x.replace(/^cf\b\s*/i, '') ? ' ' + x.replace(/^cf\b\s*/i, '').slice(0, 30) : ''));
  }));
  return cf.length ? '📦 ' + (mattino ? 'Oggi' : 'Domani') + ' cambi formato: ' + cf.join(' · ') : '';
}

/* Correzione diretta: la PRIMA correzione di un fermo chiuso la fa chi l'ha registrato (o un coordinatore) dall'app;
   dalla seconda in poi si chiede al responsabile (Chiedi correzione). Ogni modifica resta nel foglio Correzioni. */
// Ultima correzione diretta di ogni fermo: chi, quando, cosa (per far vedere al tecnico cosa ha corretto il coordinatore)
function correzioniDett_(ss) {
  const sh = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(SH_CORR), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(r => { const t = String(r[7] || '');
    if (r[1] && /^Modifica diretta: /.test(t)) out[String(r[1])] = { da: String(r[6] || ''), ts: r[0] instanceof Date ? r[0].getTime() : 0, txt: t.replace(/^Modifica diretta: /, '') }; });
  return out;
}
function correzioniPerFermo_(ss) {
  const sh = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(SH_CORR), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(r => { const k = String(r[0]); if (k) out[k] = (out[k] || 0) + 1; });
  return out;
}
// Coordinatore: "Va bene" su un fermo della lista Da rivedere
function rivistoFermo_(p) {
  if (!isCoordinatore_(p.manutentore)) throw new Error('Solo i coordinatori possono rivedere i fermi');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI), r = findRow_(sh, p.fermo);
  if (!r) throw new Error('Fermo ' + p.fermo + ' non trovato');
  causaCol_(sh);
  sh.getRange(r, C.RIV).setValue(String(p.manutentore).trim() + ' · ' + Utilities.formatDate(new Date(Number(p.ts) || Date.now()), TZ, 'dd/MM/yyyy HH:mm'));
  return { ok: true };
}
function modificaFermo_(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), fsh = ss.getSheetByName(SH_FERMI);
  const r = findRow_(fsh, p.fermo);
  if (!r) throw new Error('Fermo ' + p.fermo + ' non trovato');
  let sh = ss.getSheetByName(SH_CORR);
  if (sh && sh.getLastRow() > 1 && sh.getRange(2, 11, sh.getLastRow() - 1, 1).getValues().some(x => String(x[0]) === String(p.id))) return { dup: true };
  const v = fsh.getRange(r, 1, 1, HEAD.length).getValues()[0];
  if (String(v[C.STATO - 1]) !== 'CHIUSO') throw new Error('Si possono correggere solo fermi chiusi');
  const io = String(p.manutentore || '').trim(), chi = [v[C.MAN - 1], v[C.APERTO_DA - 1]].map(x => String(x || '').trim().toUpperCase());
  if (chi.indexOf(io.toUpperCase()) < 0 && !isCoordinatore_(io)) throw new Error('La correzione la può fare solo chi ha registrato il fermo o un coordinatore');
  if (!isCoordinatore_(io) && (correzioniPerFermo_(ss)[String(p.fermo)] || 0) > 0) throw new Error('Questo fermo è già stato corretto: per altre modifiche usa Chiedi correzione');
  if (p.rivisto && isCoordinatore_(io)) { causaCol_(fsh); fsh.getRange(r, C.RIV).setValue(io + ' · ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm')); }
  if (!String(p.altro || '').trim()) throw new Error('La descrizione è obbligatoria');
  if (String(p.intervento || '').toUpperCase() === 'SOSTITUZIONE' && !String(p.pezzo || '').trim()) throw new Error('Con SOSTITUZIONE il codice pezzo è obbligatorio');
  const campi = [['Categoria', C.CAT, p.categoria], ['Tipo intervento', C.INT, p.intervento], ['Tipo guasto', C.GUASTO, p.guasto], ['Descrizione', C.ALTRO, p.altro],
    ['Lotto', C.LOTTO, p.lotto], ['Codice pezzo', C.PEZZO, String(p.intervento || '').toUpperCase() === 'SOSTITUZIONE' ? p.pezzo : ''], ['Causa', C.CAUSA, p.causa]];
  if (p.causa != null) { causaCol_(fsh); v[C.CAUSA - 1] = fsh.getRange(r, C.CAUSA).getValue(); }
  const diff = [];
  // orari: se cambiati diventano DICHIARATI (fine non nel futuro, dopo l'inizio, al massimo 12 ore)
  if (p.inizio && p.fine) {
    const a = new Date(Number(p.inizio)), b = new Date(Number(p.fine)), a0 = v[C.INIZIO - 1], b0 = v[C.FINE - 1];
    if (b.getTime() > Date.now() + 120000) throw new Error('Orario non valido: nel futuro');
    if (b <= a) throw new Error('Orario non valido: la fine è prima dell\'inizio');
    if (b - a > 12 * 3600e3) throw new Error('Orario non valido: fermo più lungo di 12 ore');
    const hm = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'dd/MM HH:mm') : '—';
    if (!(a0 instanceof Date) || Math.abs(a0 - a) >= 60000 || !(b0 instanceof Date) || Math.abs(b0 - b) >= 60000) {
      fsh.getRange(r, C.DATA, 1, 3).setValues([[dateOnly_(a), a, b]]);
      fsh.getRange(r, C.TURNO).setValue(turno_(a)); fsh.getRange(r, C.ORARI).setValue('DICHIARATI');
      diff.push('Orari: ' + hm(a0) + '–' + hm(b0) + ' → ' + hm(a) + '–' + hm(b));
    }
  }
  campi.forEach(([nome, col, nuovo]) => {
    if (nuovo == null) return;
    const vecchio = String(v[col - 1] == null ? '' : v[col - 1]).trim(), n = String(nuovo).trim();
    if (vecchio === n) return;
    if (col === C.LOTTO) fsh.getRange(r, col).setNumberFormat('@').setValue(lotto_(n)); else fsh.getRange(r, col).setValue(n);
    diff.push(nome + ': ' + (vecchio || '—') + ' → ' + (n || '—'));
  });
  if (!diff.length) return { nulla: true };
  if (!sh) { sh = ss.insertSheet(SH_CORR); sh.getRange(1, 1, 1, CORR_H.length).setValues([CORR_H]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true); sh.setFrozenRows(1); sh.setColumnWidth(8, 380); sh.setColumnWidth(10, 260); sh.hideColumns(11); }
  const giorno = v[C.INIZIO - 1] instanceof Date ? v[C.INIZIO - 1] : v[C.DATA - 1];
  sh.appendRow([new Date(Number(p.ts) || Date.now()), p.fermo, r, v[C.LINEA - 1], v[C.MACC - 1], giorno, io, 'Modifica diretta: ' + diff.join('; '), 'FATTA', 'prima correzione, fatta dall\'app', p.id]);
  const n = sh.getLastRow();
  sh.getRange(n, 1).setNumberFormat('dd/MM/yyyy HH:mm'); sh.getRange(n, 6).setNumberFormat('dd/MM/yyyy HH:mm'); sh.getRange(n, 9).setBackground('#dff1e7');
  fsh.getRange(r, C.ALTRO).setNote('Corretto da ' + io + ' il ' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + ': ' + diff.join('; '));
  return {};
}

/* ---------- Ricodifica dei fermi registrati prima della nuova classificazione ----------
   "prepara" crea il foglio Ricodifica: una riga per fermo chiuso, con la proposta (dal guasto vecchio e dalla descrizione)
   per Cosa si è guastato / Cosa hai fatto / Perché. Il responsabile controlla e corregge; "applica" scrive nel foglio Fermi
   per ID (categoria dal foglio Voci guasto). Una voce scritta a mano che la macchina non ha viene aggiunta al foglio Guasti. */
const SH_RIC = 'Ricodifica';
const RIC_H = ['ID fermo', 'Data', 'Linea', 'Codice', 'Macchina', 'Descrizione', 'Prima: categoria', 'Prima: intervento', 'Prima: guasto', 'Prima: pezzo',
  'Cosa si è guastato', 'Cosa hai fatto', 'Perché', 'Codice pezzo', 'Applicato'];
const listaLi_ = (ss, nome, def) => { const v = listeVal_(ss), c = v[0].findIndex(h => String(h).trim().toUpperCase() === nome.toUpperCase());
  const l = c < 0 ? [] : v.slice(1).map(r => String(r[c]).trim()).filter(Boolean); return l.length ? l : def; };
const normT_ = t => String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
function propostaRicod_(f, voci, interv, cause) {
  const d = normT_(f.altro), g0 = String(f.guasto || '').trim(), G0 = g0.toUpperCase(), i0 = String(f.int || '').trim().toUpperCase();
  // cosa hai fatto
  let int = i0 === 'RIAVVIO' ? 'RIPRISTINO' : interv.indexOf(i0) >= 0 ? i0 : '';
  if (!int || /HOMING/.test(G0)) {
    if (/HOMING|RESET|RIAVVI|SBLOCC|RIARM/.test(d + ' ' + G0)) int = 'RIPRISTINO';
    else if (!int && (String(f.pezzo || '').trim() || /SOSTITU|\bSOST\b|CAMBIAT|MONTAT/.test(d))) int = 'SOSTITUZIONE';
    else if (!int && /PULIT|PULIZI/.test(d)) int = 'PULIZIA';
    else if (!int && /REGOL|SPOSTAT|REGISTRAT|ALLINEAT|FASAT|CENTRAT|PORTAT|ABBASSAT|ALZAT/.test(d)) int = 'REGOLAZIONE';
  }
  if (interv.indexOf(int) < 0) int = '';
  // cosa si è guastato: il guasto vecchio se è una voce della macchina, altrimenti la voce con più parole in comune con la descrizione
  let gua = voci.find(v => v.toUpperCase() === G0 && G0 !== 'ALTRO' && NON_VOCI.indexOf(G0) < 0) || '';
  if (!gua && d) {
    let best = 0;
    voci.forEach(v => { if (v.toUpperCase() === 'ALTRO') return;
      const w = normT_(v).split(/[^A-Z]+/).filter(x => x.length >= 4 && ['CAMBIO', 'REGOLAZIONE', 'SOSTITUZIONE', 'ZONA'].indexOf(x) < 0);
      const n = w.filter(x => d.indexOf(x.slice(0, Math.max(5, x.length - 2))) >= 0).length;
      if (n > best) { best = n; gua = v; } });
  }
  // perché
  let cau = '';
  if (G0 === 'USURA' || /USURAT|CONSUMAT|ROTT|SPACCAT|LOGOR|SCOLLAT/.test(d)) cau = 'USURA';
  else if (G0 === 'SPIKE CORRENTE' || /SPIKE|TENSIONE|CORRENTE|ALIMENTAZ|BLACKOUT/.test(d)) cau = 'ELETTRICA / ALIMENTAZIONE';
  else if (G0 === 'FORMATO' || /CAMBIO FORMATO|FORMATO/.test(d)) cau = 'FORMATO / REGOLAZIONE PERSA';
  else if (/MATERIALE|DIFETTOS|FUORI SPEC|MAL FORMAT|NON CONFORM/.test(d)) cau = 'MATERIALE FUORI SPEC.';
  else if (/OPERATOR/.test(d)) cau = 'ERRORE OPERATORE';
  if (cause.indexOf(cau) < 0) cau = '';
  // codice pezzo: se è una sostituzione senza codice, il primo numero lungo nella descrizione
  let pz = String(f.pezzo || '').trim();
  if (!pz && int === 'SOSTITUZIONE') { const m = String(f.altro || '').match(/\b\d{5,}\b/); if (m) pz = m[0]; }
  return { gua: gua, int: int, cau: cau, pz: pz };
}
function ricodificaPrepara() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { classificazione_(ss); } catch (e) { Logger.log(e); }
  const sh = ss.getSheetByName(SH_FERMI), last = sh.getLastRow(); if (last < 2) return;
  const g = guasti_(ss), eq = equipment_(ss), interv = listaLi_(ss, 'Tipo intervento', INTERVENTI), cause = listaLi_(ss, 'Causa', CAUSE);
  const v = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  const cauOld = sh.getMaxColumns() >= C.CAUSA ? sh.getRange(2, C.CAUSA, last - 1, 1).getValues() : [];
  const righe = [], voceDi = [], stato = [];
  v.forEach((r, i) => {
    if (!r[C.ID - 1] || String(r[C.STATO - 1]) !== 'CHIUSO') return;
    const code = String(r[C.EQ - 1] || '').trim(), e = eq.find(x => x.code === code) || {};
    const voci = [...new Set([...(g[code.toUpperCase()] || []), ...(g[String(r[C.LINEA - 1]).toUpperCase()] || []), ...(g.TUTTE || [])])].filter(x => NON_VOCI.indexOf(String(x).toUpperCase()) < 0);
    const f = { altro: r[C.ALTRO - 1], guasto: r[C.GUASTO - 1], int: r[C.INT - 1], pezzo: r[C.PEZZO - 1] };
    const pr = propostaRicod_(f, voci, interv, cause);
    if (String((cauOld[i] || [])[0] || '').trim()) pr.cau = String(cauOld[i][0]).trim();
    righe.push([r[C.ID - 1], r[C.INIZIO - 1] instanceof Date ? r[C.INIZIO - 1] : r[C.DATA - 1], r[C.LINEA - 1], code, e.macchina || r[C.MACC - 1], String(r[C.ALTRO - 1] || ''),
      r[C.CAT - 1], r[C.INT - 1], r[C.GUASTO - 1], r[C.PEZZO - 1], pr.gua, pr.int, pr.cau, pr.pz, '']);
    voceDi.push(voci.length ? voci : ['ALTRO']);
    stato.push([pr.gua ? (String(pr.gua).toUpperCase() === String(f.guasto).trim().toUpperCase() ? 'ok' : 'new') : 'no', pr.int ? (pr.int === String(f.int).toUpperCase() ? 'ok' : 'new') : 'no', pr.cau ? 'new' : '']);
  });
  let r = ss.getSheetByName(SH_RIC); if (r) r.clear(); else r = ss.insertSheet(SH_RIC);
  r.getRange(1, 1, 1, RIC_H.length).setValues([RIC_H]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
  r.getRange(1, 11, 1, 4).setBackground('#cfe2f3');
  r.setFrozenRows(1); r.setColumnWidth(6, 360); [11, 12, 13].forEach(c => r.setColumnWidth(c, 190));
  if (!righe.length) return;
  r.getRange(2, 1, righe.length, RIC_H.length).setValues(righe);
  r.getRange(2, 2, righe.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  r.getRange(2, 14, righe.length, 1).setNumberFormat('@');
  const col = { ok: '#ffffff', new: '#fff2cc', no: '#f4cccc', '': '#ffffff' };
  r.getRange(2, 11, righe.length, 3).setBackgrounds(stato.map(x => x.map(k => col[k])));
  r.getRange(2, 11, righe.length, 1).setDataValidations(voceDi.map(l => [SpreadsheetApp.newDataValidation().requireValueInList(l.slice(0, 500), true).setAllowInvalid(true).setHelpText('Voci di questa macchina. Puoi scriverne una nuova: verrà aggiunta al foglio Guasti.').build()]));
  r.getRange(2, 12, righe.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(interv, true).setAllowInvalid(false).build());
  r.getRange(2, 13, righe.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cause, true).setAllowInvalid(false).build());
  r.getRange(1, 17).setValue('Bianco = invariato · giallo = proposta da controllare · rosso = da compilare. Poi menu Pharma-Fix → Ricodifica fermi vecchi: applica. Le celle vuote non cambiano il foglio Fermi.').setWrap(true).setFontStyle('italic');
  r.setColumnWidth(17, 360);
  ss.setActiveSheet(r);
  const n = k => stato.filter(x => x.indexOf(k) >= 0).length;
  ss.toast(righe.length + ' fermi · ' + n('new') + ' con proposte da controllare (giallo) · ' + n('no') + ' da compilare (rosso).', 'Pharma-Fix', 10);
}
function ricodificaApplica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), r = ss.getSheetByName(SH_RIC);
  if (!r || r.getLastRow() < 2) { ss.toast('Prima: menu Pharma-Fix → Ricodifica fermi vecchi: prepara', 'Pharma-Fix', 8); return; }
  const v = r.getRange(2, 1, r.getLastRow() - 1, RIC_H.length).getValues(), fsh = ss.getSheetByName(SH_FERMI);
  // voci nuove scritte a mano → foglio Guasti, riga della macchina
  const gsh = ss.getSheetByName(SH_GUASTI), g = guasti_(ss);
  v.forEach(x => { const code = String(x[3]).trim().toUpperCase(), voce = String(x[10]).trim(); if (!voce || !code || !gsh) return;
    const tutte = [...(g[code] || []), ...(g[String(x[2]).toUpperCase()] || []), ...(g.TUTTE || [])].map(y => String(y).toUpperCase());
    if (tutte.indexOf(voce.toUpperCase()) >= 0) return;
    const gv = gsh.getDataRange().getValues(); let i = gv.findIndex(y => String(y[0]).trim().toUpperCase() === code);
    if (i < 0) { gsh.appendRow([code, String(x[4] || '')]); i = gsh.getLastRow() - 1; }
    const riga = gsh.getRange(i + 1, 1, 1, gsh.getMaxColumns()).getValues()[0]; let c = riga.length; while (c > 2 && !String(riga[c - 1]).trim()) c--;
    if (c + 1 > gsh.getMaxColumns()) gsh.insertColumnsAfter(gsh.getMaxColumns(), 5);
    gsh.getRange(i + 1, c + 1).setValue(voce); (g[code] = g[code] || []).push(voce);
  });
  VOCI_MEMO = null; vociSync_(ss); const vc = vociCat_(ss);
  causaCol_(fsh);
  const quando = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');
  let n = 0;
  v.forEach((x, k) => {
    if (String(x[14]).trim()) return; // già applicato
    const voce = String(x[10]).trim(), int = String(x[11]).trim(), cau = String(x[12]).trim(), pz = String(x[13]).trim();
    if (!voce && !int && !cau && !pz) return;
    const row = findRow_(fsh, x[0]); if (!row) return;
    const cur = fsh.getRange(row, C.CAT, 1, 3).getValues()[0], diff = [];
    const set = (col, nome, val, prima) => { if (val && String(prima).trim() !== val) { fsh.getRange(row, col).setValue(val); diff.push(nome + ': ' + (String(prima).trim() || '—') + ' → ' + val); } };
    set(C.GUASTO, 'guasto', voce, cur[2]);
    if (voce && vc[voce.toUpperCase()]) set(C.CAT, 'categoria', vc[voce.toUpperCase()], cur[0]);
    set(C.INT, 'intervento', int, cur[1]);
    set(C.CAUSA, 'causa', cau, fsh.getRange(row, C.CAUSA).getValue());
    if (pz && !String(fsh.getRange(row, C.PEZZO).getValue()).trim()) set(C.PEZZO, 'pezzo', pz, '');
    if (diff.length) { fsh.getRange(row, C.GUASTO).setNote('Ricodificato il ' + quando + ': ' + diff.join('; ')); n++; }
    r.getRange(k + 2, 15).setValue(diff.length ? quando : 'nessuna modifica');
  });
  bumpVer_(); try { aggiornaValoriKpi(); } catch (e) { Logger.log(e); }
  ss.toast(n + ' fermi ricodificati nel foglio Fermi (nota sulla cella del guasto). Colonna Applicato compilata.', 'Pharma-Fix', 10);
}

/* ---------- Formati e cambi formato (solo coordinatori) ----------
   I dati restano nel file "Base Gestione Formati Packaging" convertito in Fogli Google, nel Drive del responsabile.
   Lo script ne conosce solo l'ID (Proprietà script FORMATI_ID, mai su GitHub) e legge i valori già calcolati dai suoi fogli
   PRODOTTI, MACCHINE e LINEE. L'app li chiede con action=formati: risponde solo a un coordinatore dal suo telefono. */
const FMT_CAR = ['Contenitore', 'Tappo', 'Contagocce', 'Dosaggio', 'Etichetta', 'Astuccio', 'Istruzione', 'Pezzi/astuccio', 'Cartone', 'Fardello', 'Prodotto'];
const FMT_LINEE_X = ['ST1MAR', 'ST1TGM', 'ST2', 'ST3', 'ST4', 'ST6', 'SPL1', 'SPL2']; // colonne D–K di PRODOTTI (lette dall'intestazione)
function collegaFormati() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Formati', 'Incolla l\'indirizzo (o l\'ID) del file "Base Gestione Formati" aperto con Fogli Google.\nVuoto = scollega.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const t = r.getResponseText().trim(), id = (t.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || [])[1] || (/^[a-zA-Z0-9_-]{20,}$/.test(t) ? t : '');
  const props = PropertiesService.getScriptProperties();
  if (!t) { props.deleteProperty('FORMATI_ID'); ui.alert('Formati scollegati.'); return; }
  if (!id) { ui.alert('Indirizzo non riconosciuto: copia l\'indirizzo del file dalla barra del browser.'); return; }
  props.setProperty('FORMATI_ID', id);
  try { CacheService.getScriptCache().remove('fmt'); const d = formatiDati_(true); ui.alert('Collegato: ' + d.prod.length + ' codici, ' + d.macc.length + ' macchine, ' + d.linee.length + ' linee.' + (d.avvisi.length ? '\n\nAttenzione:\n' + d.avvisi.join('\n') : '')); }
  catch (e) { ui.alert('Non riesco a leggere il file: ' + e.message); }
}
function formatiJson_() {
  const c = CacheService.getScriptCache(); let t = null;
  try { t = cacheGet_(c, 'fmt'); } catch (e) { Logger.log(e); }
  if (!t) { t = JSON.stringify(Object.assign({ ok: true }, formatiDati_())); try { cachePut_(c, 'fmt', t); } catch (e) { Logger.log(e); } }
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JSON);
}
// Legge il file dei formati e restituisce i dati compatti per l'app
function formatiDati_(nuovo) {
  const id = PropertiesService.getScriptProperties().getProperty('FORMATI_ID');
  if (!id) throw new Error('File dei formati non collegato (menu Pharma-Fix → Formati: collega il file)');
  const ss = SpreadsheetApp.openById(id), avvisi = [];
  const foglio = n => { const sh = ss.getSheets().find(x => x.getName().trim().toUpperCase() === n); if (!sh) throw new Error('Nel file manca il foglio ' + n); return sh.getDataRange().getValues(); };
  const S = v => String(v == null ? '' : v).trim();
  // PRODOTTI: intestazioni alla riga 2, dati dalla 3
  const P = foglio('PRODOTTI'), h = P[1].map(x => S(x).replace(/\s+/g, ' ').toUpperCase());
  const col = (re, def) => { const i = h.findIndex(x => re.test(x)); if (i < 0 && def == null) avvisi.push('PRODOTTI: colonna non trovata ' + re); return i < 0 ? def : i; };
  const cCod = col(/^CODICE SAP/), cDes = col(/^DESCRIZIONE/), cPrim = col(/^LINEA PRIMARIO/);
  const cX = h.map((x, i) => [x, i]).filter(([x, i]) => i > cPrim && i < cPrim + 12 && /^(ST\d|SPL\d|ST1MAR|ST1TGM|ORL)/.test(x));
  const cAtt = ['CONTENITORE', 'TAPPO', 'CONTAGOCCE', 'DOSAGGIO', 'ETICHETTA', 'FUSTELLA ASTUCCIO', 'ISTRUZIONE', 'PEZZI PER ASTUCCIO', 'FUSTELLA CARTONE', 'FARDELLO'].map(n => col(new RegExp('^' + n)));
  const cFase = col(/^FASE$/), cFor = col(/^CODICE FORMATO/), cFF = col(/^FORMATO PRIMARIO \(FF\)/, -1), cEsito = col(/^ESITO$/, -1), cStato = col(/^STATO FORMATO/, -1), cRep = col(/^REPARTO/, -1);
  const cRicP = col(/^RICETTA PRIMARIO/, -1), cRicS = col(/^RICETTA SECONDARIO/, -1);
  const prod = P.slice(2).filter(r => S(r[cCod])).map(r => ({ c: S(r[cCod]), d: S(r[cDes]), lp: S(r[cPrim]).toUpperCase().replace(/\s+/g, ''),
    ls: cX.filter(([x, i]) => S(r[i]).toUpperCase() === 'X').map(([x]) => x.replace(/\s+/g, '')), a: cAtt.map(i => S(r[i])),
    f: S(r[cFase]), fmt: S(r[cFor]), ff: cFF >= 0 ? S(r[cFF]) : '', es: cEsito >= 0 ? S(r[cEsito]) : '', st: cStato >= 0 ? S(r[cStato]) : '', rep: cRep >= 0 ? S(r[cRep]) : '',
    rp: cRicP >= 0 ? S(r[cRicP]) : '', rs: cRicS >= 0 ? S(r[cRicS]) : '' }));
  // MACCHINE: intestazioni alla riga 1
  const M = foglio('MACCHINE'), hm = M[0].map(x => S(x).replace(/\s+/g, ' ').toUpperCase());
  const cm = n => { const e = hm.indexOf(n); return e >= 0 ? e : hm.findIndex(x => x.indexOf(n) === 0); };
  const iC = cm('CONTENITORE'), car = FMT_CAR.map((n, k) => iC + k);
  const macc = M.slice(1).filter(r => S(r[cm('MACCHINA_ID')])).map(r => ({ id: S(r[cm('MACCHINA_ID')]), l: S(r[cm('LINEA_ID')]).toUpperCase().replace(/\s+/g, ''), pos: Number(r[cm('POS')]) || 0,
    n: S(r[cm('MACCHINA')]), fn: S(r[cm('FUNZIONE')]), fase: S(r[cm('FASE')]).toUpperCase(), q: S(r[cm('LAVORA QUANDO')]).toUpperCase() || 'SEMPRE',
    m: car.map(i => S(r[i]).toUpperCase().slice(0, 1)), ver: S(r[cm('VERIFICATA')]).toUpperCase() !== 'NO' }));
  if (iC < 0) avvisi.push('MACCHINE: colonna Contenitore non trovata');
  const L = foglio('LINEE'), hl = L[0].map(x => S(x).toUpperCase());
  const linee = L.slice(1).filter(r => S(r[0])).map(r => ({ id: S(r[0]).toUpperCase().replace(/\s+/g, ''), rep: S(r[hl.indexOf('REPARTO')]), pro: S(r[hl.findIndex(x => /^PROSEGUE/.test(x))]) }));
  return { prod: prod, macc: macc, linee: linee, car: FMT_CAR, avvisi: avvisi, letto: Date.now() };
}

