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
  sheets_coord: {
    const { ctx: c2, sheets: sh2 } = nuovo(); sh2.Liste.rows = [['Manutentori', 'Coordinatori'], ['ROSSI Mario', 'PARISI Mirko']];
    let e2 = ''; try { c2.checkTecnico_('PARISI Mirko'); } catch (e) { e2 = e.message; }
    ok(!e2, 'coordinatore non presente tra i Manutentori: può entrare');
    ok(c2.tecniciTutti_(c2.SpreadsheetApp.getActiveSpreadsheet()).join() === 'ROSSI Mario,PARISI Mirko', 'i coordinatori contano anche come tecnici (compiti, vestizioni, convalide)');
  }
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
  ok(c.cSter === 7 && c.cCat.G === 8 && c.cCat.H === 9 && c.cUlt.G === 10 && c.cUlt.H === 11 && c.cSet === 13, 'colonne: G area sterile, H/I scadenze, J/K ultima convalida, M–N impostazioni');
  ok(c.durata === 13 && c.avvV === 45 && c.avvC === 90 && c.durC === 12, 'impostazioni: durata 13 (la più usata), preavvisi presi dal foglio Avvisi');
  ok(!sheets.Avvisi.rows.some(r => /Preavviso/.test(r[4] || '')), 'righe di preavviso tolte dal foglio Avvisi');
  ok(isD(V[1][7]) && V[1][7].getDate() === 22 && V[2][7] === 'MANCA' && V[3][7] === 'esonerato', 'Cat. G: data / MANCA / esonerato');
  ok(/\$N\$2/.test(sheets.Vestizioni.f['2,4'] || ''), 'formula Scadenza usa la durata standard ($N$2)');
  ok((sheets.Vestizioni.cf || []).length === 5, 'colori: esonerati, vestizione scaduta/in preavviso, convalide scadute/in preavviso');
  V[3][2] = ''; const r = ctx.vestizioni_(ctx.SpreadsheetApp.getActiveSpreadsheet()).find(x => x.man === 'RONCONI Fabio');
  ok(r.mesi === 13, 'durata vuota: vale lo standard');
  // nuova categoria nello scadenziario: colonna nuova prima del riquadro Impostazioni
  sheets.Convalide.rows.push(['x', 'Convalida Sterile Matteo Orzi', '2027/01/01', 'Categoria F - pulizia']);
  ({ ctx } = car(sheets)); const c2 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  ok(c2.cCat.F === 12 && c2.cUlt.F === 13 && c2.cSet === 15 && V[0][14] === 'Impostazioni' && c2.durata === 13, 'categoria nuova: colonne aggiunte, Impostazioni spostate a destra');
  // ultima convalida + durata standard
  ({ ctx } = car(sheets)); V[1][c2.cUlt.G - 1] = new ctx.DD(2026, 9, 1); V[2][c2.cUlt.H - 1] = new ctx.DD(2026, 5, 1);
  const c3 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  const g = c3.conv.find(x => x.man === 'ORZI Matteo' && x.cat === 'G'), h = c3.conv.find(x => x.man === 'CACIOPPO Riccardo' && x.cat === 'H');
  ok(g && new Date(g.scad).getFullYear() === 2027 && new Date(g.scad).getMonth() === 9, 'ultima convalida 01/10/2026 + 12 mesi = 01/10/2027 (più lontana dello scadenziario)');
  ok(h && isD(V[2][c3.cCat.H - 1]), 'chi era MANCA con l\'ultima convalida scritta diventa valido');
  const ug = V[2][c3.cUlt.G - 1]; // CACIOPPO non ha G nello scadenziario: resta vuota
  const uo = V[1][c3.cUlt.H - 1]; // ORZI H nello scadenziario 04/11/2026 → ultima = 04/11/2025
  ok(!ug && isD(uo) && uo.getFullYear() === 2025 && uo.getMonth() === 10 && uo.getDate() === 4, 'ultima convalida scritta dallo script = scadenza − 12 mesi (04/11/2025)');
  // foglio impostato con la versione precedente (3 impostazioni + nota): la riga nuova va nella prima riga libera
  const VS = sheets.Vestizioni.rows, cs = c3.cSet - 1; VS[4][cs] = ''; VS[4][cs + 1] = '';
  ({ ctx } = car(sheets)); const c4 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  ok(/durata standard convalide/i.test(VS[4][cs]) && c4.durC === 12 && /Viola/.test(VS[6][cs]), 'foglio già impostato: "Durata standard convalide" aggiunta senza toccare la nota');
}
console.log('\nPassaggio turno e riepilogo: fermi per reparto');
{
  const { ctx } = nuovo(); const msg = [];
  ctx.mail_ = (tipo, titolo, testo) => { msg.push(titolo + '\n' + testo); return true; };
  const now = Date.now(), F = (linea, area, min, extra) => Object.assign({ linea, area, macc: 'X', a: new ctx.DD(now - 90 * 60000), b: new ctx.DD(now - 90 * 60000 + min * 60000), stato: 'CHIUSO', guasto: 'FORMATO', man: 'ROSSI Mario' }, extra || {});
  const fermi = [F('ST3', 'FILLING', 50), F('ST3', 'FILLING', 20), F('ST3', 'PACKAGING', 7), F('SPL2', 'PACKAGING', 220), F('ST2', 'PACKAGING', 2)];
  ctx.fermiTutti_ = () => fermi.map(f => Object.assign({}, f)); ctx.areaDi_ = () => f => f.area; ctx.areaStd_ = a => a; ctx.minDi_ = f => (f.b - f.a) / 60000;
  ctx.macchinaBreve_ = () => () => 'Macchina'; ctx.anomalieTutte_ = () => []; ctx.rigaTurno_ = () => ''; ctx.rigaCambiFormato_ = () => '';
  ctx.inizioTurno_ = () => new ctx.DD(now - 4 * 3600e3);
  ctx.passaggioTurno(false);
  const m = msg[0] || '';
  ok(!/\d+ FERMI ·/.test(m) && !/PACKAGING ·/.test(m), 'niente totali sommati tra reparti né per area');
  ok(/ST3 · Filling 1h 10′ \(2\) · Packaging 7′ \(1\)/.test(m), 'ST3 · Filling 1h 10′ (2) · Packaging 7′ (1)');
  ok(m.indexOf('SPL2 ·') < m.indexOf('ST3 ·') && /SPL2 · Packaging 3h 40′ \(1\)/.test(m), 'reparti dal più fermo; linee con solo Packaging mostrano solo quello');
  ok(/Senza fermi: ST1 MAR/.test(m), 'linee senza fermi in una riga');
}
{
  const { ctx } = nuovo(); const msg = [];
  ctx.mail_ = (tipo, titolo, testo) => { msg.push(titolo + '\n' + testo); return true; };
  const o = new ctx.DD(); o.setHours(0, 0, 0, 0); const dow = (o.getDay() + 6) % 7; const fine = o.getTime() - dow * 864e5, t1 = fine - 3 * 864e5, t0 = fine - 10 * 864e5;
  const F = (t, linea, area, min, extra) => Object.assign({ linea, area, macc: 'X', a: new ctx.DD(t), b: new ctx.DD(t + min * 60000), stato: 'CHIUSO', guasto: 'FORMATO', man: 'ROSSI Mario' }, extra || {});
  const fermi = [F(t1, 'ST3', 'FILLING', 190), F(t1, 'ST3', 'PACKAGING', 125, { orari: 'DICHIARATI' }), F(t1, 'SPL2', 'PACKAGING', 60), F(t0, 'ST3', 'FILLING', 130), F(t0, 'ST3', 'PACKAGING', 145)];
  ctx.fermiTutti_ = () => fermi.map(f => Object.assign({}, f)); ctx.areaDi_ = () => f => f.area; ctx.areaStd_ = a => a; ctx.minDi_ = f => (f.b - f.a) / 60000;
  ctx.macchinaBreve_ = () => () => 'Fardellatrice'; ctx.anomalieTutte_ = () => []; ctx.completaFermiCarta_ = () => ({}); ctx.compitiTutti_ = () => []; ctx.rigaVestizioni_ = () => ''; ctx.rigaConvalide_ = () => '';
  ctx.riepilogoSettimanale();
  const m = msg[0] || ''; if (process.env.MOSTRA) console.log(m);
  ok(/ST3 · Filling 3h 10′ \(1\) ▲\+1h · Packaging 2h 5′ \(1\) ▼−20′/.test(m), 'riepilogo: ST3 · Filling ▲+1h · Packaging ▼−20′ rispetto alla settimana prima');
  ok(/SPL2 · Packaging 1h \(1\)$/m.test(m) && /Packaging registrati dopo: 1 su 2/.test(m), 'senza confronto se la settimana prima non c\'erano fermi; registrati dopo solo Packaging');
}
console.log('\nKPI Risorse');
{
  const { ctx, sheets } = nuovo();
  const t = Date.now() - 12 * 864e5, F = (man, g, min, h, extra) => Object.assign({ a: new ctx.DD(t + h * 3600e3), b: new ctx.DD(t + h * 3600e3 + min * 60000), stato: 'CHIUSO', linea: 'ST3', macc: 'ASTUCCIATRICE', guasto: g, man, altro: 'ok', orari: '', eq: 'ST3-04' }, extra || {});
  const fermi = [];
  for (let i = 0; i < 6; i++) { fermi.push(F('ROSSI Mario', 'FORMATO', 10, i * 30)); fermi.push(F('VERDI Anna', 'FORMATO', 20, i * 30 + 5)); fermi.push(F('BIANCHI Luca', 'FORMATO', 40, i * 30 + 10, { orari: 'DICHIARATI', altro: '' })); }
  fermi.push(F('BIANCHI Luca', 'SENSORE', 15, 200, { linea: 'ST2', eq: 'ST3-02' })); fermi.push(F('BIANCHI Luca', 'SENSORE', 15, 210, { linea: 'ST2', eq: 'ST3-02' }));
  ctx.fermiTutti_ = () => fermi; ctx.completaFermiCarta_ = () => ({}); ctx.anomalieTutte_ = () => [{ ts: new ctx.DD(t), da: 'ROSSI Mario' }]; ctx.compitiTutti_ = () => []; ctx.turniTutti_ = () => ({});
  const K = ctx.aggiornaKpiRisorse(), R = n => K.rows.find(x => x.n === n);
  ok(Math.abs(R('ROSSI Mario').rel - 0.5) < 1e-9 && Math.abs(R('BIANCHI Luca').rel - 2) < 1e-9, 'MTTR relativo rispetto alla mediana dello stesso guasto sulla stessa macchina (anche con gli interventi dei coordinatori)');
  ok(R('BIANCHI Luca').tempo === 0 && R('ROSSI Mario').desc === 1 && R('ROSSI Mario').ano === 1, 'registrati in tempo, descrizione, anomalie segnalate');
  ok(R('BIANCHI Luca').ric > 0, 'ricorrenza: lo stesso guasto torna entro 24h');
  ok(K.esperti[K.linee.indexOf('ST3')] === 2, 'matrice: 2 esperti su ST3 (coordinatrice esclusa)');
  ok(sheets['KPI Risorse'] && sheets['KPI Risorse'].rows[6][0] === 'Tecnico', 'foglio KPI Risorse creato con la tabella');
  fermi.push(F('VERDI Anna', 'FORMATO', 20, 300)); // VERDI Anna è coordinatrice nel foglio Liste
  const K2 = ctx.aggiornaKpiRisorse();
  ok(!K2.rows.some(x => x.n === 'VERDI Anna'), 'i coordinatori non compaiono nella scheda');
  // grafico a quadranti: con turni per avere interventi per turno
  const g0 = new ctx.DD(t); const T = {}; ['ROSSI Mario', 'BIANCHI Luca'].forEach(n => { T[n] = {}; for (let i = 0; i < 10; i++) { const d = new ctx.DD(t + i * 864e5); T[n][ctx.Utilities.formatDate(d, '', 'yyyy-MM-dd')] = '1'; } });
  ctx.turniTutti_ = () => T; ctx.aggiornaKpiRisorse();
  const ch = (sheets['KPI Risorse'].charts || [])[0], dati = (sheets.KPI_Risorse_dati || { rows: [[]] }).rows;
  ok(ch && /Riferimenti/.test(ch.opt.title) && ch.opt.hAxis.logScale && Math.abs(ch.opt.hAxis.viewWindow.min * ch.opt.hAxis.viewWindow.max - 1) < 1e-9 && dati[0].slice().sort().join('|') === ['MTTR relativo', 'ROSSI', 'BIANCHI', 'linea 1', 'linea 2'].sort().join('|'), 'grafico a quadranti: una serie per tecnico + linee di riferimento');
  ctx.aggiornaKpiRisorse(); ok((sheets['KPI Risorse'].charts || []).length === 1, 'ricalcolando non si duplicano i grafici');
  const celle = sheets['KPI Risorse'].rows.flat().filter(v => typeof v === 'string' && /^(↖|↗|↙|↘) /.test(v));
  ok(celle.length === 4 && celle.some(v => /ROSSI/.test(v)), 'i 4 quadranti in celle colorate con i nomi');
}
console.log('\nFoglio Programma pesante');
{
  const { ctx } = nuovo(); const fatto = [];
  const sh = { getLastColumn: () => 305, getMaxColumns: () => 16370, getLastRow: () => 972, getMaxRows: () => 1000, deleteColumns: (a, n) => fatto.push(['col', a, n]), deleteRows: (a, n) => fatto.push(['row', a, n]) };
  ctx.pulisciProgramma_(sh);
  ok(JSON.stringify(fatto) === JSON.stringify([['col', 311, 16060], ['row', 993, 8]]), 'colonne e righe vuote oltre i dati eliminate (16.370 → 310 colonne)');
}
fine('Script');
