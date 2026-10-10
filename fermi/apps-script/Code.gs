/**
 * Pharma-Fix Fermi — backend Google Apps Script
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

const SH_FERMI = 'Fermi', SH_EQ = 'Equipment', SH_LISTE = 'Liste', SH_KPI = 'KPI';
const TZ = 'Europe/Rome';

// Colonne del foglio Fermi (A–L = template originale)
const HEAD = [
  'Lotto', 'Data', 'Ora Inizio', 'Ora Fine', 'Manutentore ', 'Macchina', 'Categoria',
  'Tipo intervento', 'Tipo guasto', 'Altro (descrizione breve)', 'Durata intervento MTTR',
  'Intervento-Guasto',
  'Linea', 'Turno', 'Codice pezzo', 'Foto apertura', 'Foto chiusura', 'Stato', 'ID fermo',
  'Equipment', 'Aperto da'
];
const C = { LOTTO: 1, DATA: 2, INIZIO: 3, FINE: 4, MAN: 5, MACC: 6, CAT: 7, INT: 8, GUASTO: 9, ALTRO: 10,
  DUR: 11, IG: 12, LINEA: 13, TURNO: 14, PEZZO: 15, FOTO_A: 16, FOTO_C: 17, STATO: 18, ID: 19, EQ: 20, APERTO_DA: 21 };

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
  'Manutentori': ['ARACHIDE Simone', 'CACIOPPO Riccardo', 'CRISPO Mattia', 'DI RUSCIO Christian',
    'DI STEFANO Alessio', 'GRECO Nicolò', 'GRITTI Alessandro', 'GUIDONE Andrea', 'LAMANUZZI Vincenzo',
    'MANGIAROTTI Marco', 'MORABITO Giulio', 'ORZI Matteo', 'PALUMBO Antonino', 'PARISI Mirko',
    'PESTRICHELLA Mattia', 'RONCONI Fabio', 'SALVINI Alvaro', 'TATA Giuseppe']
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
      if (f) DEF_MACC.Sterile.forEach(([m, n]) => rows.push([cod + '-' + m, nome, 'Sterile', n, 'SI']));
      if (p) DEF_MACC.Secondario.forEach(([m, n]) => rows.push([cod + '-' + m, nome, 'Secondario', n, 'SI']));
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

  buildKpi_(ss);

  // Cartella foto
  if (!props.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('Pharma-Fix Foto Fermi');
    props.setProperty('FOLDER_ID', f.getId());
  }
  // Chiave condivisa con l'app
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 12));

  const s1 = ss.getSheetByName('Foglio1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);

  Logger.log('FATTO. Chiave da copiare in config.js → KEY: ' + props.getProperty('KEY'));
}

function buildKpi_(ss) {
  const sh = ss.getSheetByName(SH_KPI) || ss.insertSheet(SH_KPI);
  sh.clear();
  sh.getRange('A1').setValue('KPI manutenzione — calcolati dal foglio Fermi (solo fermi chiusi)').setFontWeight('bold').setFontSize(13);
  sh.getRange('A3:H3').setValues([['Linea', 'N° fermi', 'Tempo di fermo', 'MTTR', 'N° lotti', 'Occurrancy (fermi/lotti)',
    'Ore funzionamento (inserire a mano)', 'MTBF (h)']]).setFontWeight('bold').setBackground('#dfe7ec').setWrap(true);
  sh.getRange('A4').setFormula('=SORT(UNIQUE(FILTER(Equipment!B2:B,Equipment!B2:B<>"")))');
  sh.getRange('B4').setFormula('=MAP(A4:A,LAMBDA(l,IF(l="","",COUNTIFS(Fermi!M:M,l,Fermi!R:R,"CHIUSO"))))');
  sh.getRange('C4').setFormula('=MAP(A4:A,LAMBDA(l,IF(l="","",SUMIFS(Fermi!K:K,Fermi!M:M,l,Fermi!R:R,"CHIUSO"))))');
  sh.getRange('D4').setFormula('=MAP(B4:B,C4:C,LAMBDA(n,t,IF(OR(n="",n=0),"",t/n)))');
  sh.getRange('E4').setFormula('=MAP(A4:A,LAMBDA(l,IF(l="","",COUNTUNIQUEIFS(Fermi!A:A,Fermi!M:M,l,Fermi!A:A,"<>"))))');
  sh.getRange('F4').setFormula('=MAP(B4:B,E4:E,LAMBDA(n,lo,IF(OR(n="",lo="",lo=0),"",n/lo)))');
  sh.getRange('H4').setFormula('=MAP(B4:B,G4:G,LAMBDA(n,g,IF(OR(n="",n=0,g=""),"",g/n)))');
  sh.getRange('C4:D40').setNumberFormat('[h]:mm');
  sh.getRange('F4:F40').setNumberFormat('0.00');
  sh.getRange('H4:H40').setNumberFormat('0.0');
  sh.getRange('G4:G40').setBackground('#fff3d6');

  sh.getRange('J3').setValue('Pareto guasti (tutte le linee) — RFR = n° guasto / totale fermi').setFontWeight('bold');
  sh.getRange('J4').setFormula('=QUERY(Fermi!A1:U,"select I, count(S), sum(K) where R = \'CHIUSO\' group by I order by count(S) desc label I \'Tipo guasto\', count(S) \'N° fermi\', sum(K) \'Tempo\'",1)');
  sh.getRange('M4').setValue('RFR').setFontWeight('bold');
  sh.getRange('M5').setFormula('=MAP(K5:K,LAMBDA(n,IF(n="","",n/SUM(K5:K))))');
  sh.getRange('L5:L200').setNumberFormat('[h]:mm');
  sh.getRange('M5:M200').setNumberFormat('0.0%');

  sh.getRange('O3').setValue('Fermi per macchina e linea').setFontWeight('bold');
  sh.getRange('O4').setFormula('=QUERY(Fermi!A1:U,"select M, F, count(S), sum(K) where R = \'CHIUSO\' group by M, F order by sum(K) desc label M \'Linea\', F \'Macchina\', count(S) \'N° fermi\', sum(K) \'Tempo\'",1)');
  sh.getRange('R5:R200').setNumberFormat('[h]:mm');
  sh.setColumnWidth(10, 260);
}

/* =================== API =================== */
function doGet(e) {
  try {
    checkKey_(e.parameter.key);
    const action = e.parameter.action || 'init';
    if (action === 'init') return json_(init_());
    if (action === 'ping') return json_({ ok: true });
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
    if (p.action === 'open' || p.action === 'close') checkTecnico_(p.manutentore);
    let res;
    if (p.action === 'open') res = openFermo_(p);
    else if (p.action === 'close') res = closeFermo_(p);
    else if (p.action === 'cancel') res = cancelFermo_(p);
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
  if (!key || k !== key) throw new Error('Chiave non valida: controlla KEY in config.js');
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
  return { ok: true, equipment: equipment, liste: liste, fermi: fermi, serverTime: Date.now() };
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
  const url = savePhoto_(p.foto, stamp + '_' + p.code + '_apertura');
  const lotto = /^\d+$/.test(String(p.lotto || '')) ? Number(p.lotto) : (p.lotto || '');
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
  sh.appendRow(row);
  const r = sh.getLastRow();
  sh.getRange(r, C.DATA).setNumberFormat('dd/MM/yyyy');
  sh.getRange(r, C.INIZIO, 1, 2).setNumberFormat('HH:mm');
  sh.getRange(r, C.DUR).setFormula('=IF(D' + r + '="","",D' + r + '-C' + r + ')').setNumberFormat('[h]:mm');
  sh.getRange(r, C.IG).setFormula('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')');
  if (url) sh.getRange(r, C.FOTO_A).setFormula(linkFormula_(url));
  sh.getRange(r, 1, 1, HEAD.length).setBackground('#fff3d6');
  return { row: r };
}

function closeFermo_(p) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_FERMI);
  const r = findRow_(sh, p.id);
  if (!r) throw new Error('Fermo ' + p.id + ' non trovato');
  if (sh.getRange(r, C.STATO).getValue() === 'CHIUSO') return { dup: true };
  const d = new Date(Number(p.ts));
  const code = sh.getRange(r, C.EQ).getValue();
  const url = savePhoto_(p.foto, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + code + '_chiusura');
  sh.getRange(r, C.FINE).setValue(d).setNumberFormat('HH:mm');
  if (p.manutentore) sh.getRange(r, C.MAN).setValue(p.manutentore);
  sh.getRange(r, C.CAT, 1, 4).setValues([[p.categoria || '', p.intervento || '', p.guasto || '', p.altro || '']]);
  sh.getRange(r, C.PEZZO).setValue(p.pezzo || '');
  if (url) sh.getRange(r, C.FOTO_C).setFormula(linkFormula_(url));
  sh.getRange(r, C.STATO).setValue('CHIUSO');
  if (p.lotto !== undefined && p.lotto !== '') {
    sh.getRange(r, C.LOTTO).setValue(/^\d+$/.test(String(p.lotto)) ? Number(p.lotto) : p.lotto);
  }
  sh.getRange(r, 1, 1, HEAD.length).setBackground(null);
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
