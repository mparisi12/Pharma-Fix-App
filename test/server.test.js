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
console.log('\nConvalide GMP');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows = [['Manutentori', 'Coordinatori'], ['ORZI Matteo', 'VERDI Anna'], ['DI STEFANO Alessio', ''], ['GRECO Nicolò', ''], ['VERDI Anna', '']];
  sheets.Convalide = new Sheet('Convalide', [[], ['Codice processo (link)', 'Oggetto', 'Scadenza', 'Item'],
    ['SC 1', 'Convalida Sterile Matteo Orzi Meccanico Servizi Tecnici', '2026/10/22', 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 1', 'Convalida Sterile Matteo Orzi Meccanico Servizi Tecnici', '2026/11/04', 'Categoria H - Intervento Meccanico di sostituzione componente \t'],
    ['SC 1', 'Convalida Sterile Matteo Orzi Meccanico', '2025/01/01', 'Categoria H - vecchia'],
    ['SC 2', 'Convalida Sterile Alessio Di Stefano Meccanico Servizi Tecnici', new ctx.DD(2027, 2, 31), 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 3', 'Convalida Sterile Nicolo Greco Meccanico', '22/10/2026', 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 4', 'Convalida Mario Sconosciuto Meccanico', '2027/01/01', 'Categoria G - x']]);
  const r = ctx.convalide_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  const f = (m, c) => r.list.find(x => x.man === m && x.cat === c);
  ok(f('ORZI Matteo', 'G') && new Date(f('ORZI Matteo', 'G').scad).getDate() === 22, 'nome invertito riconosciuto (Matteo Orzi → ORZI Matteo), data 2026/10/22');
  ok(new Date(f('ORZI Matteo', 'H').scad).getMonth() === 10, 'più righe stessa categoria: vale la scadenza più lontana');
  ok(f('ORZI Matteo', 'H').desc === 'Intervento Meccanico di sostituzione componente', 'descrizione dalla colonna Item');
  ok(!!f('DI STEFANO Alessio', 'G'), 'cognome con particella e data come Date');
  ok(!!f('GRECO Nicolò', 'G'), 'accenti: Nicolo = Nicolò, data 22/10/2026');
  ok(r.ignote.length === 1 && /Sconosciuto/.test(r.ignote[0]), 'nomi non trovati in Manutentori segnalati');
}
console.log('\nArea sterile: esonerati e convalide mancanti');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows = [['Manutentori'], ['ORZI Matteo'], ['CACIOPPO Riccardo'], ['RONCONI Luca']];
  sheets.Vestizioni = new Sheet('Vestizioni', [['Tecnico', 'Data vestizione', 'Durata (mesi)', 'Scadenza', 'Giorni mancanti', 'Note'], ['ORZI Matteo', new ctx.DD(2026, 0, 10), 12], ['CACIOPPO Riccardo'], ['RONCONI Luca', '', '', '', '', '', 'NO']]);
  sheets.Convalide = new Sheet('Convalide', [['x', 'Convalida Sterile Matteo Orzi', '2027/01/10', 'Categoria G - regolazione'], ['x', 'Convalida Sterile Matteo Orzi', '2027/01/10', 'Categoria H - sostituzione']]);
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ok(JSON.stringify(ctx.esonerati_(ss)) === '["RONCONI Luca"]', 'colonna Area sterile: NO = esonerato');
  ok(/sterile/i.test(sheets.Vestizioni.rows[0][6]), 'colonna Area sterile presente nel foglio Vestizioni');
  const m = ctx.convMancanti_(ss, ctx.convalideTutte_(ss).list);
  ok(m.length === 1 && m[0].man === 'CACIOPPO Riccardo' && m[0].cats.join() === 'G,H', 'chi è in sterile senza convalide: mancanti G e H (gli esonerati no)');
  ok(/mancanti: CACIOPPO G H/.test(ctx.rigaConvalide_()), 'riepilogo: riga con i mancanti');
}
console.log('\nFoglio Vestizioni: impostazioni e convalide');
{
  const { Sheet: Sh, carica: car } = require('./fake-sheets');
  const sheets = {}; let { ctx } = car(sheets);
  const D = (y, m, d) => new ctx.DD(y, m - 1, d), isD = x => Object.prototype.toString.call(x) === '[object Date]';
  Object.assign(sheets, {
    Liste: new Sh('Liste', [['Manutentori'], ['ORZI Matteo'], ['CACIOPPO Riccardo'], ['RONCONI Fabio']]),
    Vestizioni: new Sh('Vestizioni', [['Tecnico', 'Data vestizione', 'Durata (mesi)', 'Scadenza', 'Giorni mancanti', 'Note'],
      ['ORZI Matteo', D(2025, 10, 28), 13], ['CACIOPPO Riccardo', D(2026, 5, 28), 13], ['RONCONI Fabio', D(2026, 1, 10), '']]),
    Convalide: new Sh('Convalide', [['x', 'Convalida Sterile Matteo Orzi', '2026/10/22', 'Categoria G - Intervento Meccanico di regolazione'],
      ['x', 'Convalida Sterile Matteo Orzi', '2026/11/04', 'Categoria H - Intervento Meccanico di sostituzione componente']]),
    Avvisi: new Sh('Avvisi', [['Avvisi', 'Valore', '', '', 'Report automatici', 'Valore'], ['', '', '', '', 'Preavviso scadenza vestizione sterile (giorni prima)', 45], ['', '', '', '', 'Preavviso scadenza convalide GMP (giorni prima)', 90]])
  });
  sheets.Vestizioni.rows[3][6] = 'NO'; // RONCONI esonerato (colonna G già scritta a mano, senza intestazione)
  sheets.Vestizioni.rows[0][6] = 'Area sterile (SI/NO)';
  const c = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet()), V = sheets.Vestizioni.rows;
  ok(c.cSter === 7 && c.cCat.G === 8 && c.cCat.H === 9 && c.cSet === 11, 'colonne: G area sterile, H/I convalide, K–L impostazioni');
  ok(c.durata === 13 && c.avvV === 45 && c.avvC === 90, 'impostazioni: durata 13 (la più usata), preavvisi presi dal foglio Avvisi');
  ok(!sheets.Avvisi.rows.some(r => /Preavviso/.test(r[4] || '')), 'righe di preavviso tolte dal foglio Avvisi');
  ok(isD(V[1][7]) && V[1][7].getDate() === 22 && V[2][7] === 'MANCA' && V[3][7] === 'esonerato', 'Cat. G: data / MANCA / esonerato');
  ok(/\$L\$2/.test(sheets.Vestizioni.f['2,4'] || ''), 'formula Scadenza usa la durata standard ($L$2)');
  ok((sheets.Vestizioni.cf || []).length === 5, 'colori: esonerati, vestizione scaduta/in preavviso, convalide scadute/in preavviso');
  V[3][2] = ''; const r = ctx.vestizioni_(ctx.SpreadsheetApp.getActiveSpreadsheet()).find(x => x.man === 'RONCONI Fabio');
  ok(r.mesi === 13, 'durata vuota: vale lo standard');
  // nuova categoria nello scadenziario: colonna nuova prima del riquadro Impostazioni
  sheets.Convalide.rows.push(['x', 'Convalida Sterile Matteo Orzi', '2027/01/01', 'Categoria F - pulizia']);
  ({ ctx } = car(sheets)); const c2 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  ok(c2.cCat.F === 10 && c2.cSet === 12 && V[0][11] === 'Impostazioni' && c2.durata === 13, 'categoria nuova: colonna aggiunta, Impostazioni spostate a destra');
}
fine('Script');
