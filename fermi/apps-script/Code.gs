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
// Quelli di ST3 vengono dallo storico del template "fermi meccanici"; le altre macchine si compilano nel foglio Guasti.
const DEF_GUASTI = {
  'TUTTE': ['FORMATO', 'SENSORE', 'USURA', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO'],
  'ST3-02': ['H7, Sblocco Homing', 'H7 GUANTONE, Sblocco Homing', 'H10, Sostituzione e regolazione fotocellula della macchina',
    'H32 CPS, rimessa in dima pinze cps, Sostituzione silicone pinze ferma flaconi', 'H32 CTG, rimessa in dima pinze ctg',
    'H19, Rimessa in fase della Stella', 'G25 Regolazione della coclea flaconi', 'SEA VISION'],
  'ST3-03': ['PARAM VIDEOJET'],
  'ST3-04': ['MAG. ISTRUZIONI', 'CHIUSURE AST', 'MAG. ASTUCCI', 'STAMPA DATI VARIABILI', 'CAM. OMRON', 'GUIDE', 'INCASTRO',
    'POS. BOLLINO', 'POS. TAMPER', 'REGOLAZIONE DA PANNELLO'],
  'ST3-06': ['NASTRATORI', 'GUIDE', 'INCASTRO']
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
  DUR: 11, IG: 12, LINEA: 13, TURNO: 14, PEZZO: 15, FOTO_A: 16, FOTO_C: 17, STATO: 18, ID: 19, EQ: 20, APERTO_DA: 21, ORARI: 22, REG: 23, TEL: 24 };

const DEF_LISTE = {
  'Categoria': ['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO'],
  'Tipo intervento': ['SOSTITUZIONE', 'REGOLAZIONE', 'PULIZIA', 'RIAVVIO', 'ALTRO'],
  'Guasto Sterile': ['H7, Sblocco Homing', 'H7 GUANTONE, Sblocco Homing', 'H32 CTG, rimessa in dima pinze ctg',
    'H32 CPS, rimessa in dima pinze cps, Sostituzione silicone pinze ferma flaconi',
    'H10, Sostituzione e regolazione fotocellula della macchina', 'H19, Rimessa in fase della Stella',
    'G25 Regolazione della coclea flaconi', 'SEA VISION', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'FORMATO', 'ALTRO'],
  'Guasto Secondario': ['PARAM VIDEOJET', '"0" MACCHINA', 'GUIDE', 'CHIUSURE AST', 'STAMPA DATI VARIABILI',
    'SPIKE CORRENTE', 'POS. TAMPER', 'POS. BOLLINO', 'INCASTRO', 'ARIA COMPRESSA', 'MAG. ISTRUZIONI',
    'MAG. ASTUCCI', 'NASTRATORI', 'VOMERE', 'CAM. OMRON', 'CAM. ANTARES', 'FORMATO',
    'REGOLAZIONE DA PANNELLO', 'USURA', 'SENSORE', 'PULIZIA', 'SOSTITUZIONE', 'RIAVVIO', 'ALTRO'],
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
    .addSeparator()
    .addItem('Controlla registrazioni', 'controllaRegistrazioni')
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
    .addToUi();
}

function onEdit(e) {
  try {
    if (e.range.getSheet().getName() !== SH_KPI) return;
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

  sh.getRange('A9:H9').setValues([['N° fermi', 'Tempo di fermo', 'MTTR', 'N° lotti', 'Occurrancy (fermi/lotti)',
    'Ore funzionamento (inserire)', 'MTBF (h)', 'Guasto n°1 · RFR']])
    .setFontWeight('bold').setBackground(H).setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A10:H10').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F10').setBackground(Y);
  sh.getRange('F11').setValue('ore della linea/macchina nel periodo scelto (dai dati di produzione)').setFontColor(G).setFontSize(9).setWrap(true);
  sh.getRange('H11').setValue('RFR = n° fermi di quel guasto / totale fermi').setFontColor(G).setFontSize(9);

  // grafici in alto (righe 12-29), tabelle sotto (da riga KPI_T)
  const t = KPI_T;
  sh.getRange(t - 1, 1).setValue('Pareto per FREQUENZA — quante volte succede').setFontWeight('bold');
  sh.getRange(t, 1, 1, 5).setValues([['Tipo guasto', 'N° fermi', 'Tempo', '% cumulata', 'RFR']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, 7).setValue('Pareto per DURATA — quanto ferma la linea').setFontWeight('bold');
  sh.getRange(t, 7, 1, 5).setValues([['Tipo guasto', 'Ore di fermo', 'N° fermi', '% cumulata', '% del tempo']]).setFontWeight('bold').setBackground(H);
  sh.getRange(t - 1, 13).setValue('Fermi per macchina').setFontWeight('bold');
  sh.getRange(t, 13, 1, 6).setValues([['Linea', 'Macchina', 'N° fermi', 'Tempo', 'MTTR', 'Guasto n°1 · RFR']]).setFontWeight('bold').setBackground(H);

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

// Calcolo puro (testabile): righe del foglio Fermi + filtri → numeri e tabelle
const kpiChiave_ = (linea, macc) => (String(linea) + '|' + String(macc)).toUpperCase().replace(/\s+/g, ' ').replace(/ ?\| ?/, '|').trim();
function calcolaKpi_(vals, f) {
  const rows = [];
  const senzaArea = [];
  vals.forEach(r => {
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
      guasto: String(r[C.GUASTO - 1]).trim() || '(vuoto)', dur: dur });
  });
  const n = rows.length, tot = rows.reduce((a, x) => a + x.dur, 0);
  const lotti = new Set(rows.map(x => x.lotto).filter(Boolean)).size;
  const grp = (key) => { const m = {}; rows.forEach(x => { const k = key(x); m[k] = m[k] || { n: 0, t: 0 }; m[k].n++; m[k].t += x.dur; }); return m; };
  const g = grp(x => x.guasto);
  let cum = 0;
  const pareto = Object.keys(g).map(k => [k, g[k].n, g[k].t]).sort((a, b) => b[1] - a[1] || b[2] - a[2])
    .map(r => { cum += r[1]; return [r[0], r[1], r[2], cum / n, r[1] / n]; });
  let cumT = 0;
  const paretoT = Object.keys(g).map(k => [k, g[k].t * 24, g[k].n]).sort((a, b) => b[1] - a[1] || b[2] - a[2])
    .map(r => { cumT += r[1]; return [r[0], r[1], r[2], tot ? cumT / (tot * 24) : 0, tot ? r[1] / (tot * 24) : 0]; });
  const mm = {};
  rows.forEach(x => {
    const k = x.linea + '\u0001' + x.macc;
    const m = mm[k] = mm[k] || { n: 0, t: 0, g: {} };
    m.n++; m.t += x.dur; m.g[x.guasto] = (m.g[x.guasto] || 0) + 1;
  });
  const top1 = (gg, tn) => { const k = Object.keys(gg).sort((a, b) => gg[b] - gg[a])[0]; return k ? k + ' · ' + Math.round(gg[k] / tn * 100) + '%' : ''; };
  const macchine = Object.keys(mm).map(k => { const p = k.split('\u0001'), m = mm[k]; return [p[0], p[1], m.n, m.t, m.t / m.n, top1(m.g, m.n)]; })
    .sort((a, b) => b[3] - a[3]);
  return { n: n, tot: tot, mttr: n ? tot / n : '', lotti: lotti, occ: lotti ? n / lotti : '',
    mtbf: (n && f.ore !== '' && !isNaN(Number(f.ore))) ? Number(f.ore) / n : '',
    top: pareto.length ? pareto[0][0] + ' · ' + Math.round(pareto[0][4] * 100) + '%' : '',
    pareto: pareto, paretoT: paretoT, macchine: macchine, senzaArea: senzaArea };
}

function aggiornaValoriKpi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_KPI);
  if (!sh) return;
  // foglio KPI della versione precedente: lo ricrea tenendo filtri e ore di funzionamento
  if (sh.getRange(KPI_T, 7).getValue() !== 'Tipo guasto' || sh.getRange('A2').getValue() !== 'Area' || sh.getRange(KPI_T, 20).getValue() !== 'Grafico %: guasto') {
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
  const k = calcolaKpi_(vals, { from: p.from, toEnd: p.toEnd, linea: linea, area: area, aree: aree, code: mSel === 'TUTTE' ? '' : mSel.split(' · ')[0],
    ore: sh.getRange('F10').getValue() });

  sh.getRange('A10:E10').setValues([[k.n, k.tot, k.mttr, k.lotti, k.occ]]);
  const ex = k.senzaArea.length;
  sh.getRange('A11').setValue(ex ? '⚠ ' + (ex === 1 ? '1 fermo escluso' : ex + ' fermi esclusi') + ', macchina non trovata in Equipment: ' +
    k.senzaArea.slice(0, 5).join('; ') + (ex > 5 ? '; …' : '') + '. Correggi Linea/Macchina nel foglio Fermi oppure aggiungi la macchina in Equipment.' : '')
    .setFontColor('#a3402d').setFontSize(9).setWrap(false);
  sh.getRange('G10:H10').setValues([[k.mtbf, k.top]]);
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
}

/* =================== SOLUZIONI, DOCUMENTI, ANOMALIE =================== */
// Colonna "Dove": codice macchina (es. ST3-04), nome linea (es. ST3) oppure TUTTE.
const SH_SOL = 'Soluzioni', SH_DOC = 'Documenti', SH_ANO = 'Anomalie', SH_IMP = 'Import_PharmaFix', SH_MAPPA = 'Mappa_import';
const SOL_H = ['ID', 'Data', 'Dove (codice / linea / TUTTE)', 'Linea', 'Macchina', 'Tipo guasto', 'Problema', 'Soluzione', 'Codici ricambio', 'Allegato', 'Autore'];
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
const linkDi_ = (sh, r, c) => { const rt = sh.getRange(r, c).getRichTextValue(); return (rt && rt.getLinkUrl()) || String(sh.getRange(r, c).getValue() || ''); };

function archivio_(ss) {
  const sol = righe_(ss, SH_SOL, SOL_H.length).map(r => ({ id: String(r[0]), ts: ts_(r[1]), dove: String(r[2]).trim(), linea: String(r[3]), macchina: String(r[4]),
    guasto: String(r[5]), problema: String(r[6]), soluzione: String(r[7]), ricambi: String(r[8]), link: String(r[9]), autore: String(r[10]) }));
  const doc = righe_(ss, SH_DOC, DOC_H.length).map(r => ({ id: String(r[0]), ts: ts_(r[1]), dove: String(r[2]).trim(), tipo: String(r[3]).trim().toUpperCase(),
    titolo: String(r[4]), descr: String(r[5]), link: String(r[6]) }));
  const from = Date.now() - 30 * 864e5;
  const ano = righe_(ss, SH_ANO, ANO_H.length).map(r => ({ id: String(r[0]), ts: ts_(r[1]), code: String(r[2]), linea: String(r[3]), macchina: String(r[4]),
    urgenza: Number(r[5]) || 3, descr: String(r[6]), serve: String(r[7]), foto: String(r[8]), da: String(r[9]), stato: String(r[10] || 'APERTA').toUpperCase(),
    gestita: String(r[11]), chiusa: ts_(r[12]), note: String(r[13]) }))
    .filter(a => a.stato !== 'CHIUSA' || (a.chiusa || a.ts || 0) >= from);
  return { soluzioni: sol, documenti: doc, anomalie: ano };
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
  sh.appendRow([p.id, d, p.code || p.linea || 'TUTTE', eq.linea, eq.macchina, p.guasto || '', p.problema, p.soluzione, p.ricambi || '', url, p.manutentore || '']);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('dd/MM/yyyy HH:mm');
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
function controllaDichiarati_(p, inizio) {
  const reg = Number(p.reg) || Date.now(), ts = Number(p.ts);
  const avvio = inizioTurno_(new Date(reg)).getTime();
  if (ts > reg + 120000) throw new Error('Orario non valido: nel futuro');
  if (ts < avvio - 60000) throw new Error('Orario non valido: fuori dal turno in corso');
  if (inizio instanceof Date && ts <= inizio.getTime()) throw new Error('Orario non valido: la fine è prima dell\'inizio');
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
    v.filter(x => x.orari === 'DICHIARATI' && x.reg instanceof Date && x.reg - x.b > 4 * 3600e3)
      .forEach(x => out.push(['Orari dichiarati registrati tardi', x.man + ': registrato ' + Math.round((x.reg - x.b) / 3600e3) + ' ore dopo la fine', 'Fermi riga ' + x.r]));
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
    const corpo = righe.map(([k, v]) => v === undefined ? String(k).toUpperCase() + '\n' : k + ': ' + testo_(v)).join('\n'); // niente link al foglio: il filtro aziendale lo rende illeggibile
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
  const testo = icona + 'PHARMA-FIX · ' + oggetto + '\n\n' + righe.filter(r => r[1] !== undefined).map(([k, v]) => k + ': ' + txt(v)).join('\n');
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
  else if (String(sh.getRange(REP_RIGHE.length, 5).getValue()) !== REP_RIGHE[REP_RIGHE.length - 1][0]) { // righe aggiunte in seguito
    sh.getRange(REP_RIGHE.length, 5, 1, 2).setValues([REP_RIGHE[REP_RIGHE.length - 1]]); nuovo = true;
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
  return sh.getRange(2, 1, last - 1, HEAD.length).getValues().filter(r => r[C.INIZIO - 1] instanceof Date).map(r => ({
    a: r[C.INIZIO - 1], b: r[C.FINE - 1] instanceof Date ? r[C.FINE - 1] : null, stato: r[C.STATO - 1], linea: String(r[C.LINEA - 1] || ''),
    macc: String(r[C.MACC - 1] || ''), guasto: String(r[C.GUASTO - 1] || ''), man: String(r[C.MAN - 1] || r[C.APERTO_DA - 1] || ''),
    altro: String(r[C.ALTRO - 1] || ''), lotto: r[C.LOTTO - 1], orari: r[C.ORARI - 1] }));
}
const minDi_ = f => f.b ? (f.b - f.a) / 60000 : 0;
function anomalieTutte_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return righe_(ss, SH_ANO, ANO_H.length).map(r => ({ ts: r[1] instanceof Date ? r[1] : null, linea: r[3], macc: r[4], urg: Number(r[5]) || 3,
    descr: String(r[6] || ''), da: r[9], stato: String(r[10] || 'APERTA').toUpperCase(), chiusa: r[12] instanceof Date ? r[12] : null }));
}

/* Riepilogo della settimana precedente (lun–dom) */
function riepilogoSettimanale_() { if (repCfg_().settimanale) riepilogoSettimanale(); }
function riepilogoSettimanale() {
  const oggi = dateOnly_(new Date()), dow = Number(Utilities.formatDate(oggi, TZ, 'u'));
  const fine = new Date(oggi.getTime() - (dow - 1) * 864e5), inizio = new Date(fine.getTime() - 7 * 864e5), prima = new Date(inizio.getTime() - 7 * 864e5);
  const tutti = fermiTutti_().filter(f => f.stato === 'CHIUSO' && f.b);
  const sett = tutti.filter(f => f.a >= inizio && f.a < fine), prec = tutti.filter(f => f.a >= prima && f.a < inizio);
  const tot = sett.reduce((x, f) => x + minDi_(f), 0), totP = prec.reduce((x, f) => x + minDi_(f), 0);
  const perLinea = {}; sett.forEach(f => { const l = perLinea[f.linea] = perLinea[f.linea] || { n: 0, m: 0 }; l.n++; l.m += minDi_(f); });
  const perGuasto = {}; sett.forEach(f => { if (!f.guasto) return; const k = f.guasto + ' (' + f.linea + ' ' + nice_(f.macc) + ')'; const g = perGuasto[k] = perGuasto[k] || { n: 0, m: 0 }; g.n++; g.m += minDi_(f); });
  const an = anomalieTutte_(), aperte = an.filter(a => a.stato !== 'CHIUSA');
  const aperti = fermiTutti_().filter(f => f.stato === 'APERTO');
  const freccia = (a, b) => b ? (a > b ? ' ▲ ' : a < b ? ' ▼ ' : ' = ') + 'sett. prima ' + b : '';
  mail_('Riepilogo settimanale', 'RIEPILOGO SETTIMANA ' + gm_(inizio) + '–' + gm_(new Date(fine.getTime() - 864e5)), [
    ['Settimana dal ' + gm_(inizio) + ' al ' + gm_(new Date(fine.getTime() - 864e5))],
    ['Fermi', sett.length + freccia(sett.length, prec.length)],
    ['Tempo di fermo', durTxt_(tot) + (totP ? ' (sett. prima ' + durTxt_(totP) + ')' : '')],
    ['MTTR', sett.length ? durTxt_(tot / sett.length) : '—'],
    ['Per linea', Object.keys(perLinea).sort((a, b) => perLinea[b].m - perLinea[a].m).map(l => l + ': ' + perLinea[l].n + (perLinea[l].n === 1 ? ' fermo · ' : ' fermi · ') + durTxt_(perLinea[l].m) + ' · MTTR ' + durTxt_(perLinea[l].m / perLinea[l].n)).join('<br>') || '—'],
    ['Guasti più frequenti', Object.keys(perGuasto).sort((a, b) => perGuasto[b].n - perGuasto[a].n || perGuasto[b].m - perGuasto[a].m).slice(0, 5).map(k => k + ' ×' + perGuasto[k].n + ' · ' + durTxt_(perGuasto[k].m)).join('<br>') || '—'],
    ['Fermi più lunghi', sett.slice().sort((a, b) => minDi_(b) - minDi_(a)).slice(0, 3).map(f => gm_(f.a) + ' ' + f.linea + ' ' + nice_(f.macc) + ' · ' + durTxt_(minDi_(f)) + (f.guasto ? ' · ' + f.guasto : '')).join('<br>') || '—'],
    ['Anomalie', 'da gestire ' + aperte.length + ' (critiche ' + aperte.filter(a => a.urg === 1).length + ', urgenti ' + aperte.filter(a => a.urg === 2).length + ') · nuove ' + an.filter(a => a.ts && a.ts >= inizio && a.ts < fine).length + ' · chiuse ' + an.filter(a => a.chiusa && a.chiusa >= inizio && a.chiusa < fine).length],
    ['Fermi ancora aperti', aperti.length ? aperti.map(f => f.linea + ' ' + nice_(f.macc) + ' dal ' + gm_(f.a) + ' ' + hm_(f.a)).join('<br>') : 'nessuno']]);
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
  const t = turno_(a), giorno = ['', 'lunedì', 'martedì', 'mercoledì', 'giovedì', 'venerdì', 'sabato', 'domenica'][Number(Utilities.formatDate(a, TZ, 'u'))] + ' ' + gm_(a);
  const tutti = fermiTutti_().filter(linea);
  const nel = tutti.filter(f => f.a >= a && f.a < b).sort((x, y) => x.a - y.a);
  const aperti = tutti.filter(f => f.stato === 'APERTO');
  const an = anomalieTutte_().filter(linea), nuove = an.filter(x => x.ts && x.ts >= a && x.ts < b), chiuse = an.filter(x => x.chiusa && x.chiusa >= a && x.chiusa < b);
  const aperte = an.filter(x => x.stato !== 'CHIUSA');
  const tot = nel.reduce((x, f) => x + minDi_(f), 0);
  const riga = f => hm_(f.a) + (f.b ? '–' + hm_(f.b) : '') + ' ' + f.linea + ' ' + nice_(f.macc) + ' · ' + (f.guasto || 'guasto non indicato') + ' · ' + (f.b ? durTxt_(minDi_(f)) : 'APERTO') +
    ' · ' + (f.man || '?') + (f.altro ? ' · ' + h_(f.altro.slice(0, 70)) : ' · (senza descrizione)') + (f.orari === 'DICHIARATI' ? ' [dichiarato]' : '');
  const solo = soloLinee ? ' · solo ' + soloLinee.join(', ') : '';
  mail_('Passaggio di turno', 'PASSAGGIO TURNO ' + t + ' · ' + giorno + solo, [
    ['Turno ' + t + ' di ' + giorno + ' (' + hm_(a) + '–' + hm_(b) + ')' + (soloLinee ? ' · ciclo continuo: ' + soloLinee.join(', ') : '')],
    ['Fermi nel turno', nel.length ? nel.length + ' · tempo di fermo ' + durTxt_(tot) : 'nessun fermo registrato'],
    ['Dettaglio', nel.length ? nel.slice(0, 20).map(riga).join('<br>') + (nel.length > 20 ? '<br>… e altri ' + (nel.length - 20) : '') : '—'],
    ['Fermi ancora aperti', aperti.length ? aperti.map(f => f.linea + ' ' + nice_(f.macc) + ' dalle ' + hm_(f.a) + ' (' + durTxt_((Date.now() - f.a) / 60000) + ') · ' + (f.man || '?')).join('<br>') : 'nessuno'],
    ['Anomalie nuove', nuove.length ? nuove.map(x => 'urg. ' + x.urg + ' · ' + x.linea + ' ' + nice_(x.macc) + ' · ' + h_(x.descr.slice(0, 70)) + ' · ' + (x.da || '')).join('<br>') : 'nessuna'],
    ['Anomalie chiuse', chiuse.length || 'nessuna'],
    ['Anomalie da gestire', aperte.length + (aperte.length ? ' (critiche ' + aperte.filter(x => x.urg === 1).length + ', urgenti ' + aperte.filter(x => x.urg === 2).length + ')' : '')],
    ['Nota', 'riepilogo automatico dai dati registrati nell\'app']], { noMail: !conMail });
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
    if (action === 'init') return json_(init_());
    if (action === 'ping') return json_({ ok: true });
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
    if (['open', 'close', 'soluzione', 'anomalia', 'anomaliaStato'].indexOf(p.action) >= 0) { checkTecnico_(p.manutentore); checkTelefono_(p.manutentore, p.dev, p.devName); }
    let res;
    if (p.action === 'open') res = openFermo_(p);
    else if (p.action === 'close') res = closeFermo_(p);
    else if (p.action === 'cancel') res = cancelFermo_(p);
    else if (p.action === 'soluzione') res = addSoluzione_(p);
    else if (p.action === 'anomalia') res = addAnomalia_(p);
    else if (p.action === 'anomaliaStato') res = statoAnomalia_(p);
    else throw new Error('Azione sconosciuta');
    SpreadsheetApp.flush();
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
  const lv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LISTE).getDataRange().getValues();
  const col = lv[0].indexOf('Manutentori');
  const ok = col >= 0 && lv.slice(1).some(r => String(r[col]).trim() !== '' && String(r[col]).trim() === String(nome || '').trim());
  if (!ok) throw new Error('Manutentore non abilitato: ' + nome);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function init_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const equipment = equipment_(ss);
  // Liste
  const lv = ss.getSheetByName(SH_LISTE).getDataRange().getValues();
  const liste = {};
  lv[0].forEach((h, i) => { if (h) liste[h] = lv.slice(1).map(r => String(r[i]).trim()).filter(Boolean); });
  // Fermi: aperti + chiusi ultimi 90 giorni
  const sh = ss.getSheetByName(SH_FERMI);
  const last = sh.getLastRow();
  const fermi = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
    const fotoC = sh.getRange(2, C.FOTO_C, last - 1, 1).getFormulas(); // link della foto di chiusura
    const from = Date.now() - 90 * 864e5;
    vals.forEach((r, i) => {
      const start = r[C.INIZIO - 1] instanceof Date ? r[C.INIZIO - 1].getTime() : null;
      const stato = r[C.STATO - 1];
      if (!r[C.ID - 1] || !start) return;
      if (stato !== 'APERTO' && start < from) return;
      const end = r[C.FINE - 1] instanceof Date ? r[C.FINE - 1].getTime() : null;
      fermi.push({ id: r[C.ID - 1], code: r[C.EQ - 1], linea: r[C.LINEA - 1], macchina: r[C.MACC - 1], lotto: String(r[C.LOTTO - 1]),
        start: start, end: end, stato: stato, cat: r[C.CAT - 1], int: r[C.INT - 1], guasto: r[C.GUASTO - 1],
        altro: r[C.ALTRO - 1], man: r[C.MAN - 1], apertoDa: r[C.APERTO_DA - 1],
        pezzo: String(r[C.PEZZO - 1] || ''), orari: String(r[C.ORARI - 1] || ''), fotoC: ((fotoC[i][0] || '').match(/"(https?:[^"]+)"/) || [])[1] || '' });
    });
  }
  return Object.assign({ ok: true, equipment: equipment, liste: liste, guasti: guasti_(ss), fermi: fermi, serverTime: Date.now() }, archivio_(ss));
}

function equipment_(ss) {
  const eqv = ss.getSheetByName(SH_EQ).getDataRange().getValues().slice(1);
  return eqv.filter(r => r[0] && String(r[4]).toUpperCase() !== 'NO')
    .map(r => ({ code: String(r[0]).trim(), linea: String(r[1]), area: String(r[2]), macchina: String(r[3]) }));
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
  sh.getRange(r, C.DATA).setNumberFormat('dd/MM/yyyy');
  sh.getRange(r, C.INIZIO, 1, 2).setNumberFormat('HH:mm');
  sh.getRange(r, C.REG).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(r, C.DUR).setFormula(F_('=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))')).setNumberFormat('[h]:mm');
  sh.getRange(r, C.IG).setFormula(F_('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')'));
  if (url) sh.getRange(r, C.FOTO_A).setFormula(F_(linkFormula_(url)));
  sh.getRange(r, 1, 1, HEAD.length).setBackground('#fff3d6');
  return { row: r };
}

// Lotto obbligatorio: solo cifre, da 4 a 10
function lotto_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!/^\d{4,10}$/.test(s)) throw new Error('Lotto obbligatorio: inserisci il numero di lotto');
  return Number(s);
}

function closeFermo_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const r = findRow_(sh, p.id);
  if (!r) throw new Error('Fermo ' + p.id + ' non trovato');
  if (sh.getRange(r, C.STATO).getValue() === 'CHIUSO') return { dup: true };
  lotto_(p.lotto);
  const d = new Date(Number(p.ts));
  if (p.dichiarato) controllaDichiarati_(p, sh.getRange(r, C.INIZIO).getValue());
  const code = sh.getRange(r, C.EQ).getValue();
  const url = savePhoto_(p.foto, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + code + '_chiusura');
  sh.getRange(r, C.FINE).setValue(d).setNumberFormat('HH:mm');
  if (p.manutentore) sh.getRange(r, C.MAN).setValue(p.manutentore);
  sh.getRange(r, C.CAT, 1, 4).setValues([[p.categoria || '', p.intervento || '', p.guasto || '', p.altro || '']]);
  sh.getRange(r, C.PEZZO).setValue(p.pezzo || '');
  if (url) sh.getRange(r, C.FOTO_C).setFormula(F_(linkFormula_(url)));
  sh.getRange(r, C.STATO).setValue('CHIUSO');
  sh.getRange(r, C.REG).setValue(new Date(Number(p.reg) || Date.now())).setNumberFormat('dd/MM/yyyy HH:mm');
  const tel = String(sh.getRange(r, C.TEL).getValue() || '');
  if (p.dev && tel.indexOf(p.dev) < 0) sh.getRange(r, C.TEL).setValue(tel ? tel + ' / ' + p.dev : p.dev);
  sh.getRange(r, C.LOTTO).setValue(lotto_(p.lotto));
  sh.getRange(r, 1, 1, HEAD.length).setBackground(null);
  try { SpreadsheetApp.flush(); aggiornaValoriKpi(); } catch (e) { Logger.log(e); }
  avvisoRipetuto_(sh, code, p.guasto);
  return { row: r };
}

function cancelFermo_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const r = findRow_(sh, p.id);
  if (!r) return { dup: true };
  if (sh.getRange(r, C.STATO).getValue() !== 'APERTO') throw new Error('Si possono annullare solo fermi aperti');
  sh.deleteRow(r);
  return {};
}
