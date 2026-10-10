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

  if (!ss.getSheetByName(SH_KPI)) buildKpi_(ss); // il foglio KPI non viene più ricostruito se esiste già
  if (!ss.getSheetByName(SH_GUASTI)) buildGuasti_(ss);
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
  sh.getRange(2, C.DUR, last - 1, 1).setFormulas(f).setNumberFormat('[h]:mm');
}

// ---------- KPI con filtri (linea, macchina, periodo) ----------
// Per ricostruire il foglio KPI (es. dopo un aggiornamento): esegui aggiornaKpi() dall'editor.
function aggiornaKpi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SH_KPI, SH_KPI_DATI].forEach(n => { const s = ss.getSheetByName(n); if (s) ss.deleteSheet(s); });
  buildKpi_(ss);
}

const SH_KPI_DATI = 'KPI_dati';
const PERIODI = ['Ultimi 7 giorni', 'Ultimi 30 giorni', 'Ultimi 90 giorni', 'Mese corrente', 'Mese scorso', 'Anno corrente', 'Tutto', 'Personalizzato'];

function buildKpi_(ss) {
  // I due fogli si citano a vicenda: si creano entrambi prima di scrivere le formule
  const sh = ss.getSheetByName(SH_KPI) || ss.insertSheet(SH_KPI, 1);
  // Foglio nascosto con i fermi filtrati: tutte le formule KPI leggono da qui
  const dt = ss.getSheetByName(SH_KPI_DATI) || ss.insertSheet(SH_KPI_DATI);
  dt.clear();
  dt.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  dt.getRange('A2').setFormula('=IFERROR(FILTER(Fermi!A2:U, Fermi!R2:R="CHIUSO", ' +
    'IF(KPI!$B$3="TUTTE", Fermi!R2:R<>"", Fermi!M2:M=KPI!$B$3), ' +
    'IF(KPI!$B$4="TUTTE", Fermi!R2:R<>"", Fermi!T2:T=REGEXEXTRACT(KPI!$B$4,"^\\S+")), ' +
    'Fermi!B2:B>=KPI!$B$6, Fermi!B2:B<=KPI!$B$7),)');
  // elenchi per i menu a tendina
  dt.getRange('W1').setValue('TUTTE');
  dt.getRange('W2').setFormula('=SORT(UNIQUE(FILTER(Equipment!B2:B,Equipment!B2:B<>"")))');
  dt.getRange('X1').setValue('TUTTE');
  dt.getRange('X2').setFormula('=IFERROR(FILTER(Equipment!A2:A&" · "&Equipment!D2:D, Equipment!A2:A<>"", ' +
    'IF(KPI!$B$3="TUTTE", Equipment!A2:A<>"", Equipment!B2:B=KPI!$B$3)),)');
  PERIODI.forEach((p, i) => dt.getRange(i + 1, 25).setValue(p)); // colonna Y
  dt.hideSheet();

  sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  const H = '#dfe7ec', Y = '#fff3d6';
  sh.getRange('A1').setValue('KPI manutenzione — fermi chiusi').setFontWeight('bold').setFontSize(14);

  // Filtri
  sh.getRange('A3:A7').setValues([['Linea'], ['Macchina'], ['Periodo'], ['Dal'], ['Al']]).setFontWeight('bold');
  sh.getRange('B3').setValue('TUTTE');
  sh.getRange('B4').setValue('TUTTE');
  sh.getRange('B5').setValue('Ultimi 30 giorni');
  sh.getRange('B3:B5').setBackground(Y);
  const dv = r => SpreadsheetApp.newDataValidation().requireValueInRange(r, true).setAllowInvalid(false).build();
  sh.getRange('B3').setDataValidation(dv(dt.getRange('W1:W60')));
  sh.getRange('B4').setDataValidation(dv(dt.getRange('X1:X300')));
  sh.getRange('B5').setDataValidation(dv(dt.getRange('Y1:Y8')));
  sh.getRange('B6').setFormula('=SWITCH(B5,"Ultimi 7 giorni",TODAY()-6,"Ultimi 30 giorni",TODAY()-29,"Ultimi 90 giorni",TODAY()-89,' +
    '"Mese corrente",EOMONTH(TODAY(),-1)+1,"Mese scorso",EOMONTH(TODAY(),-2)+1,"Anno corrente",DATE(YEAR(TODAY()),1,1),' +
    '"Personalizzato",IF(D6="",TODAY()-29,D6),DATE(2000,1,1))');
  sh.getRange('B7').setFormula('=SWITCH(B5,"Mese scorso",EOMONTH(TODAY(),-1),"Personalizzato",IF(D7="",TODAY(),D7),TODAY())');
  sh.getRange('B6:B7').setNumberFormat('dd/MM/yyyy');
  sh.getRange('C6:C7').setValues([['Personalizzato: dal →'], ['Personalizzato: al →']]).setFontColor('#5a6873').setHorizontalAlignment('right');
  sh.getRange('D6:D7').setBackground(Y).setNumberFormat('dd/MM/yyyy');
  sh.getRange('C4').setFormula('=IF(OR(B4="TUTTE",COUNTIF(KPI_dati!X:X,B4)>0),"","⚠ macchina non della linea scelta: rimetti TUTTE")').setFontColor('#c23b2a');

  // Indicatori
  sh.getRange('A9:H9').setValues([['N° fermi', 'Tempo di fermo', 'MTTR', 'N° lotti', 'Occurrancy (fermi/lotti)',
    'Ore funzionamento (inserire)', 'MTBF (h)', 'Guasto più frequente']])
    .setFontWeight('bold').setBackground(H).setWrap(true).setVerticalAlignment('middle');
  sh.getRange('A10').setFormula('=COUNTIF(KPI_dati!S2:S,"?*")');
  sh.getRange('B10').setFormula('=SUM(KPI_dati!K2:K)');
  sh.getRange('C10').setFormula('=IF(A10=0,"",B10/A10)');
  sh.getRange('D10').setFormula('=IFERROR(COUNTUNIQUE(FILTER(KPI_dati!A2:A,KPI_dati!A2:A<>"")),0)');
  sh.getRange('E10').setFormula('=IF(D10=0,"",A10/D10)');
  sh.getRange('F10').setBackground(Y);
  sh.getRange('G10').setFormula('=IF(OR(A10=0,F10=""),"",F10/A10)');
  sh.getRange('H10').setFormula('=IFERROR(INDEX(A14:A,1),"")');
  sh.getRange('A10:H10').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('B10:C10').setNumberFormat('[h]:mm');
  sh.getRange('E10').setNumberFormat('0.00');
  sh.getRange('G10').setNumberFormat('0.0');
  sh.getRange('F11').setValue('ore della linea/macchina nel periodo scelto').setFontColor('#5a6873').setFontSize(9);

  // Pareto
  sh.getRange('A12').setValue('Pareto guasti — RFR = n° di quel guasto / totale fermi').setFontWeight('bold');
  sh.getRange('A13').setFormula('=IFERROR(QUERY(KPI_dati!A1:U,"select I, count(S), sum(K) where S is not null group by I order by count(S) desc ' +
    'label I \'Tipo guasto\', count(S) \'N° fermi\', sum(K) \'Tempo\'",1),"Nessun fermo chiuso con questi filtri")');
  sh.getRange('D13:E13').setValues([['% cumulata', 'RFR']]);
  sh.getRange('D14').setFormula('=ARRAYFORMULA(IF(B14:B200="","",SCAN(0,B14:B200,LAMBDA(a,b,a+N(b)))/SUM(B14:B200)))');
  sh.getRange('E14').setFormula('=ARRAYFORMULA(IF(B14:B200="","",B14:B200/SUM(B14:B200)))');
  sh.getRange('A13:E13').setFontWeight('bold').setBackground(H);
  sh.getRange('C14:C200').setNumberFormat('[h]:mm');
  sh.getRange('D14:E200').setNumberFormat('0%');

  // Per macchina
  sh.getRange('G12').setValue('Fermi per macchina').setFontWeight('bold');
  sh.getRange('G13').setFormula('=IFERROR(QUERY(KPI_dati!A1:U,"select M, F, count(S), sum(K) where S is not null group by M, F order by sum(K) desc ' +
    'label M \'Linea\', F \'Macchina\', count(S) \'N° fermi\', sum(K) \'Tempo\'",1),"")');
  sh.getRange('K13').setValue('MTTR');
  sh.getRange('K14').setFormula('=ARRAYFORMULA(IF(I14:I200="","",J14:J200/I14:I200))');
  sh.getRange('G13:K13').setFontWeight('bold').setBackground(H);
  sh.getRange('J14:K200').setNumberFormat('[h]:mm');

  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 230); sh.setColumnWidth(3, 170); sh.setColumnWidth(8, 260);
  sh.setFrozenRows(10);

  // Grafico di Pareto: barre = n° fermi, linea = % cumulata
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
  return { ok: true, equipment: equipment, liste: liste, guasti: guasti_(ss), fermi: fermi, serverTime: Date.now() };
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
  sh.appendRow(row);
  const r = sh.getLastRow();
  sh.getRange(r, C.DATA).setNumberFormat('dd/MM/yyyy');
  sh.getRange(r, C.INIZIO, 1, 2).setNumberFormat('HH:mm');
  sh.getRange(r, C.DUR).setFormula('=IF(OR(C' + r + '="",D' + r + '=""),"",MOD(D' + r + '-C' + r + ',1))').setNumberFormat('[h]:mm');
  sh.getRange(r, C.IG).setFormula('=IF(H' + r + '="","",LEFT(H' + r + ',3)&"-"&I' + r + ')');
  if (url) sh.getRange(r, C.FOTO_A).setFormula(linkFormula_(url));
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
  const code = sh.getRange(r, C.EQ).getValue();
  const url = savePhoto_(p.foto, Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmm') + '_' + code + '_chiusura');
  sh.getRange(r, C.FINE).setValue(d).setNumberFormat('HH:mm');
  if (p.manutentore) sh.getRange(r, C.MAN).setValue(p.manutentore);
  sh.getRange(r, C.CAT, 1, 4).setValues([[p.categoria || '', p.intervento || '', p.guasto || '', p.altro || '']]);
  sh.getRange(r, C.PEZZO).setValue(p.pezzo || '');
  if (url) sh.getRange(r, C.FOTO_C).setFormula(linkFormula_(url));
  sh.getRange(r, C.STATO).setValue('CHIUSO');
  sh.getRange(r, C.LOTTO).setValue(lotto_(p.lotto));
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
