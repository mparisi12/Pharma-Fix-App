// Test dello script (Code.gs) con un finto foglio. Lancia: node server.test.js
const { Sheet, carica, ok, fine } = require('./fake-sheets');
const nuovo = () => {
  const sheets = {
    Liste: new Sheet('Liste', [['Manutentori', 'Coordinatori'], ['ROSSI Mario', 'VERDI Anna'], ['BIANCHI Luca', ''], ['VERDI Anna', '']]),
    Equipment: new Sheet('Equipment', [['Codice (QR)', 'Linea', 'Area', 'Macchina', 'Attivo'], ['ST3-02', 'ST3', 'Filling', 'RIEMPITRICE', 'SI'], ['ST3-04', 'ST3', 'Packaging', 'ASTUCCIATRICE', 'SI'],
      ['ST1MAR-06', 'ST1 MAR', 'Packaging', 'BILANCIA', 'SI'], ['ST9-01', 'ST9', 'Packaging', 'DISMESSA', 'NO']]),
    Documenti: new Sheet('Documenti', [['ID']])
  };
  return { sheets, ...carica(sheets) };
};

console.log('Versione dei dati e copia pronta');
{
  const { ctx } = nuovo(); let calcoli = 0;
  ctx.init_ = () => { calcoli++; return { ok: true, testo: 'àèì ✓', n: calcoli, pieno: require('crypto').randomBytes(150000).toString('hex') }; };
  const chiedi = v => JSON.parse(ctx.initCache_(v).t);
  const a = chiedi('');
  ok(a.ver && a.testo === 'àèì ✓', 'prima richiesta: dati completi con versione (accenti intatti)');
  ok(chiedi(a.ver).same === true && calcoli === 1, 'stessa versione: risponde solo "uguale", senza rileggere il foglio');
  ok(chiedi('').n === 1 && calcoli === 1, 'altro telefono senza versione: dati dalla copia pronta (anche spezzata in più pezzi)');
  ctx.bumpVer_();
  const b = chiedi(a.ver);
  ok(!b.same && b.ver !== a.ver && calcoli === 2, 'dopo una modifica: dati nuovi e versione nuova');
  ok(!chiedi('vecchissima').same, 'versione sconosciuta: dati completi');
}
{
  const { ctx, sheets } = nuovo(); const v0 = ctx.dataVer_();
  ctx.onEdit({ range: { getSheet: () => sheets.Liste, getA1Notation: () => 'A2' } });
  ok(ctx.dataVer_() !== v0, 'modifica a mano nel foglio (onEdit): cambia la versione');
}

console.log('\nDocumenti');
{
  const { ctx, sheets } = nuovo();
  ok(JSON.stringify(ctx.docDove_(ctx.SpreadsheetApp.getActiveSpreadsheet())) === JSON.stringify(['TUTTE', 'ST1 MAR', 'ST3', 'ST1MAR-06', 'ST3-02', 'ST3-04']), 'elenco "Dove": TUTTE + linee + codici attivi (senza dismesse)');
  ok(ctx.docElenchi_(ctx.SpreadsheetApp.getActiveSpreadsheet()) === true && sheets.Elenco_Dove.rows.length === 7, 'foglio Elenco_Dove creato');
  const prova = p => { try { return ctx.addDocumento_(p); } catch (e) { return { err: e.message }; } };
  ok(/coordinatori/.test(prova({ manutentore: 'ROSSI Mario', dove: 'ST3', tipo: 'MANUALE', titolo: 'x', link: 'https://a' }).err || ''), 'un tecnico non può aggiungere documenti');
  ok(prova({ id: 'D1', manutentore: 'VERDI Anna', dove: 'st3-04', tipo: 'SCHEDA FORMATO', titolo: 'Scheda 10 ml', link: 'https://drive/x' }).ok, 'il coordinatore aggiunge (codice scritto minuscolo corretto)');
  ok(prova({ id: 'D1', manutentore: 'VERDI Anna', dove: 'ST3-04', tipo: 'SCHEDA FORMATO', titolo: 'Scheda 10 ml', link: 'https://drive/x' }).dup, 'invio ripetuto: nessun doppione');
  ok(/Equipment/.test(prova({ id: 'D2', manutentore: 'VERDI Anna', dove: 'ST9-01', tipo: 'MANUALE', titolo: 'x', link: 'https://d' }).err || ''), 'macchina dismessa rifiutata');
  ok(/link/.test(prova({ id: 'D3', manutentore: 'VERDI Anna', dove: '', tipo: 'ALTRO', titolo: 'x', link: 'drive' }).err || ''), 'link non valido rifiutato');
  sheets.Documenti.rows.push(['', '', '', '', 'Manuale senza Dove', '', 'https://m', '']);
  const docs = ctx.archivio_(ctx.SpreadsheetApp.getActiveSpreadsheet()).documenti;
  ok(docs.some(d => d.titolo === 'Manuale senza Dove' && d.dove === 'TUTTE'), 'riga senza ID e senza Dove: si vede e vale per TUTTE');
}

console.log('\nCoordinatori e manutentori');
{
  const { ctx } = nuovo();
  ok(ctx.isCoordinatore_('verdi anna') && !ctx.isCoordinatore_('ROSSI Mario'), 'riconosce i coordinatori (maiuscole/minuscole)');
  let err = ''; try { ctx.checkTecnico_('NERI Ugo'); } catch (e) { err = e.message; }
  ok(/non abilitato/.test(err), 'nome non in elenco Manutentori: rifiutato');
}
console.log('\nNomi negli avvisi');
{
  const { ctx } = nuovo(); const c = ctx.cognome_;
  ok(c('DI STEFANO Marco') === 'DI STEFANO', 'DI STEFANO Marco → DI STEFANO');
  ok(c('Di Stefano Marco') === 'Di Stefano', 'Di Stefano Marco → Di Stefano');
  ok(c('DI Ruscio Paolo') === 'DI Ruscio', 'DI Ruscio Paolo → DI Ruscio');
  ok(c('DE LUCA Anna') === 'DE LUCA' && c('ROSSI Mario') === 'ROSSI' && c('Rossi Mario') === 'Rossi', 'cognomi normali invariati');
  ok(c('DI') === 'DI' && c('') === '?', 'casi limite');
}
console.log('\nMenu a tendina del foglio Fermi');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows[0] = ['Categoria', 'Tipo intervento', 'Manutentori', 'Coordinatori'];
  sheets.Fermi = new Sheet('Fermi', [['Lotto']]);
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'FORMATO', 'SENSORE'], ['ST3-04', '', 'GUIDE', 'formato']]);
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ok(ctx.fermiConvalida_(ss) === true, 'prima volta: tendine create');
  const col = c => (sheets.Fermi.dv || []).find(d => d.c === c);
  ok([5, 6, 7, 8, 9].every(c => col(c) && col(c).r === 2), 'colonne E, F, G, H, I con la tendina');
  ok(col(5).v.setAllowInvalid[0] === false && col(9).v.setAllowInvalid[0] === true, 'E rigida (solo elenco), I solo segnalata');
  ok(JSON.stringify(sheets.Elenchi_Fermi.rows.slice(1).map(r => r[1]).filter(Boolean)) === JSON.stringify(['FORMATO', 'GUIDE', 'SENSORE']), 'tipi di guasto da Guasti, senza doppioni');
  ok(sheets.Elenchi_Fermi.rows.slice(1).map(r => r[0]).includes('ASTUCCIATRICE') && !sheets.Elenchi_Fermi.rows.slice(1).map(r => r[0]).includes('DISMESSA'), 'macchine attive da Equipment');
  ok(ctx.fermiConvalida_(ss) === false, 'se non cambia niente non rifà il lavoro');
}
fine('Script');
