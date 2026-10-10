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
    .addToUi();
}

function onEdit(e) {
  try {
    if (e.range.getSheet().getName() !== SH_KPI) return;
    if (['B3', 'B4', 'B5', 'D6', 'D7', 'F10'].indexOf(e.range.getA1Notation()) >= 0) aggiornaValoriKpi();
  } catch (x) { Logger.log(x); }
}

// Ricrea da zero il foglio KPI (impaginazione, menu, grafico) e lo calcola
function aggiornaKpi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['KPI_dati', SH_KPI].forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  buildKpi_(ss);
}

function buildKpi_(ss) {
  const sh = ss.getSheetByName(SH_KPI) || ss.insertSheet(SH_KPI, 1);
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getCharts().forEach(c => sh.removeChart(c));
  const H = '#dfe7ec', Y = '#fff3d6', G = '#5a6873';
  sh.getRange('A1').setValue('KPI manutenzione — fermi chiusi').setFontWeight('bold').setFontSize(14);
  sh.getRange('C1').setValue('Cambia le celle gialle: i valori si ricalcolano da soli (oppure menu Pharma-Fix → Aggiorna KPI)').setFontColor(G);
  sh.getRange('A3:A7').setValues([['Linea'], ['Macchina'], ['Periodo'], ['Dal'], ['Al']]).setFontWeight('bold');
  sh.getRange('B3:B5').setValues([['TUTTE'], ['TUTTE'], ['Ultimi 30 giorni']]).setBackground(Y);
  sh.getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PERIODI, true).setAllowInvalid(false).build());
  sh.getRange('B6:B7').setNumberFormat('dd/MM/yyyy');
  sh.getRange('C6:C7').setValues([['Personalizzato: dal →'], ['Personalizzato: al →']]).setFontColor(G).setHorizontalAlignment('right');
  sh.getRange('D6:D7').setBackground(Y).setNumberFormat('dd/MM/yyyy');

  sh.getRange('A9:H9').setValues([['N° fermi', 'Tempo di fermo', 'MTTR', 'N° lotti', 'Occurrancy (fermi/lotti)',
    'Ore funzionamento (inserire)', 'MTBF (h)', 'Guasto più frequente']])
    .setFontWeight('bold').setBackground(H).setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A10:H10').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('F10').setBackground(Y);
  sh.getRange('F11').setValue('ore della linea/macchina nel periodo scelto').setFontColor(G).setFontSize(9);

  sh.getRange('A12').setValue('Pareto guasti — RFR = n° di quel guasto / totale fermi').setFontWeight('bold');
  sh.getRange('A13:E13').setValues([['Tipo guasto', 'N° fermi', 'Tempo', '% cumulata', 'RFR']]).setFontWeight('bold').setBackground(H);
  sh.getRange('G12').setValue('Fermi per macchina').setFontWeight('bold');
  sh.getRange('G13:K13').setValues([['Linea', 'Macchina', 'N° fermi', 'Tempo', 'MTTR']]).setFontWeight('bold').setBackground(H);

  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 230); sh.setColumnWidth(3, 170); sh.setColumnWidth(8, 300);
  sh.setFrozenRows(10);

  try {
    const chart = sh.newChart().asComboChart()
      .addRange(sh.getRange('A13:B33')).addRange(sh.getRange('D13:D33'))
      .setNumHeaders(1)
      .setOption('title', 'Pareto guasti')
      .setOption('legend', { position: 'bottom' })
      .setOption('series', { 0: { type: 'bars', targetAxisIndex: 0, color: '#245a7a' }, 1: { type: 'line', targetAxisIndex: 1, color: '#c23b2a', pointSize: 5 } })
      .setOption('vAxes', { 0: { title: 'N° fermi', minValue: 0 }, 1: { title: '% cumulata', format: 'percent', minValue: 0, maxValue: 1 } })
      .setPosition(2, 13, 0, 0).setOption('width', 720).setOption('height', 380)
      .build();
    sh.insertChart(chart);
  } catch (e) { Logger.log('Grafico non creato: ' + e); }

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
function calcolaKpi_(vals, f) {
  const rows = [];
  vals.forEach(r => {
    if (r[C.STATO - 1] !== 'CHIUSO') return;
    const g0 = dataOra_(r[C.DATA - 1], null);
    const ini = dataOra_(r[C.INIZIO - 1], g0), fin = dataOra_(r[C.FINE - 1], g0);
    if (!ini || !fin) return;
    const giorno = g0 || ini;
    if (giorno < f.from || giorno >= f.toEnd) return;
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
  const mm = grp(x => x.linea + '\u0001' + x.macc);
  const macchine = Object.keys(mm).map(k => { const p = k.split('\u0001'); return [p[0], p[1], mm[k].n, mm[k].t, mm[k].t / mm[k].n]; })
    .sort((a, b) => b[3] - a[3]);
  return { n: n, tot: tot, mttr: n ? tot / n : '', lotti: lotti, occ: lotti ? n / lotti : '',
    mtbf: (n && f.ore !== '' && !isNaN(Number(f.ore))) ? Number(f.ore) / n : '', top: pareto.length ? pareto[0][0] : '',
    pareto: pareto, macchine: macchine };
}

function aggiornaValoriKpi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SH_KPI);
  if (!sh) return;
  // menu Linea e Macchina dal foglio Equipment
  const eq = equipment_(ss);
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
  const k = calcolaKpi_(vals, { from: p.from, toEnd: p.toEnd, linea: linea, code: mSel === 'TUTTE' ? '' : mSel.split(' · ')[0],
    ore: sh.getRange('F10').getValue() });

  sh.getRange('A10:E10').setValues([[k.n, k.tot, k.mttr, k.lotti, k.occ]]);
  sh.getRange('G10:H10').setValues([[k.mtbf, k.top]]);
  sh.getRange('B10:C10').setNumberFormat('[h]:mm'); sh.getRange('E10').setNumberFormat('0.00'); sh.getRange('G10').setNumberFormat('0.0');

  sh.getRange('A14:E300').clearContent(); sh.getRange('G14:K300').clearContent();
  if (k.pareto.length) {
    sh.getRange(14, 1, k.pareto.length, 5).setValues(k.pareto);
    sh.getRange(14, 3, k.pareto.length, 1).setNumberFormat('[h]:mm');
    sh.getRange(14, 4, k.pareto.length, 2).setNumberFormat('0%');
  } else sh.getRange('A14').setValue('Nessun fermo chiuso con questi filtri');
  if (k.macchine.length) {
    sh.getRange(14, 7, k.macchine.length, 5).setValues(k.macchine);
    sh.getRange(14, 10, k.macchine.length, 2).setNumberFormat('[h]:mm');
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
    const from = Date.now() - 90 * 864e5;
    vals.forEach(r => {
      const start = r[C.INIZIO - 1] instanceof Date ? r[C.INIZIO - 1].getTime() : null;
      const stato = r[C.STATO - 1];
      if (!r[C.ID - 1] || !start) return;
      if (stato !== 'APERTO' && start < from) return;
      const end = r[C.FINE - 1] instanceof Date ? r[C.FINE - 1].getTime() : null;
      fermi.push({ id: r[C.ID - 1], code: r[C.EQ - 1], linea: r[C.LINEA - 1], macchina: r[C.MACC - 1], lotto: String(r[C.LOTTO - 1]),
        start: start, end: end, stato: stato, cat: r[C.CAT - 1], int: r[C.INT - 1], guasto: r[C.GUASTO - 1],
        altro: r[C.ALTRO - 1], man: r[C.MAN - 1], apertoDa: r[C.APERTO_DA - 1] });
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
