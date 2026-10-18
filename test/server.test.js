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
    const { ctx: c2, sheets: sh2 } = nuovo(); sh2.Liste.rows = [['Manutentori', 'Coordinatori'], ['ROSSI Mario', 'NERI Paolo']];
    let e2 = ''; try { c2.checkTecnico_('NERI Paolo'); } catch (e) { e2 = e.message; }
    ok(!e2, 'coordinatore non presente tra i Manutentori: può entrare');
    ok(c2.tecniciTutti_(c2.SpreadsheetApp.getActiveSpreadsheet()).join() === 'ROSSI Mario,NERI Paolo', 'i coordinatori contano anche come tecnici (compiti, vestizioni, convalide)');
  }
}
console.log('\nNomi negli avvisi');
{
  const { ctx } = nuovo(); const c = ctx.cognome_;
  ok(c('DI BELLO Marco') === 'DI BELLO', 'DI BELLO Marco → DI BELLO');
  ok(c('Di Bello Marco') === 'Di Bello', 'Di Bello Marco → Di Bello');
  ok(c('DI Monte Paolo') === 'DI Monte', 'DI Monte Paolo → DI Monte');
  ok(c('DE LUCA Anna') === 'DE LUCA' && c('ROSSI Mario') === 'ROSSI' && c('Rossi Mario') === 'Rossi', 'cognomi normali invariati');
  ok(c('DI') === 'DI' && c('') === '?', 'casi limite');
}
console.log('\nMenu a tendina del foglio Fermi');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows = [['Categoria', 'Tipo intervento', 'Manutentori', 'Coordinatori'], ['MECCANICO', 'REGOLAZIONE', 'ROSSI Mario', 'VERDI Anna'], ['ELETTRICO', 'SOSTITUZIONE', 'BIANCHI Luca', '']];
  sheets.Fermi = new Sheet('Fermi', [['Lotto']]);
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'FORMATO', 'SENSORE'], ['ST3-04', '', 'GUIDE', 'formato']]);
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ok(ctx.fermiConvalida_(ss) === true, 'prima volta: tendine create');
  const col = c => (sheets.Fermi.dv || []).find(d => d.c === c);
  ok([5, 6, 7, 8, 9].every(c => col(c) && col(c).r === 2), 'colonne E, F, G, H, I con la tendina');
  ok(col(5).v.setAllowInvalid[0] === false && col(9).v.setAllowInvalid[0] === true, 'E rigida (solo elenco), I solo segnalata');
  ok(JSON.stringify(sheets.Elenchi_Fermi.rows.slice(1).map(r => r[2]).filter(Boolean)) === JSON.stringify(['BIANCHI Luca', 'ROSSI Mario', 'VERDI Anna']), 'colonna E: elenco tecnici = Manutentori + Coordinatori');
  ok(JSON.stringify(sheets.Elenchi_Fermi.rows.slice(1).map(r => r[1]).filter(Boolean)) === JSON.stringify(['FORMATO', 'GUIDE', 'SENSORE']), 'tipi di guasto da Guasti, senza doppioni');
  ok(sheets.Elenchi_Fermi.rows.slice(1).map(r => r[0]).includes('ASTUCCIATRICE') && !sheets.Elenchi_Fermi.rows.slice(1).map(r => r[0]).includes('DISMESSA'), 'macchine attive da Equipment');
  ok(ctx.fermiConvalida_(ss) === false, 'se non cambia niente non rifà il lavoro');
}
console.log('\nConvalide GMP');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows = [['Manutentori', 'Coordinatori'], ['GIALLI Marco', 'VERDI Anna'], ['DI BELLO Andrea', ''], ['CORSI Nicolò', ''], ['VERDI Anna', '']];
  sheets.Convalide = new Sheet('Convalide', [[], ['Codice processo (link)', 'Oggetto', 'Scadenza', 'Item'],
    ['SC 1', 'Convalida Sterile Marco Gialli Meccanico Servizi Tecnici', '2026/10/22', 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 1', 'Convalida Sterile Marco Gialli Meccanico Servizi Tecnici', '2026/11/04', 'Categoria H - Intervento Meccanico di sostituzione componente \t'],
    ['SC 1', 'Convalida Sterile Marco Gialli Meccanico', '2025/01/01', 'Categoria H - vecchia'],
    ['SC 2', 'Convalida Sterile Andrea Di Bello Meccanico Servizi Tecnici', new ctx.DD(2027, 2, 31), 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 3', 'Convalida Sterile Nicolo Corsi Meccanico', '22/10/2026', 'Categoria G - Intervento Meccanico di regolazione'],
    ['SC 4', 'Convalida Mario Sconosciuto Meccanico', '2027/01/01', 'Categoria G - x']]);
  const r = ctx.convalide_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  const f = (m, c) => r.list.find(x => x.man === m && x.cat === c);
  ok(f('GIALLI Marco', 'G') && new Date(f('GIALLI Marco', 'G').scad).getDate() === 22, 'nome invertito riconosciuto (Marco Gialli → GIALLI Marco), data 2026/10/22');
  ok(new Date(f('GIALLI Marco', 'H').scad).getMonth() === 10, 'più righe stessa categoria: vale la scadenza più lontana');
  ok(f('GIALLI Marco', 'H').desc === 'Intervento Meccanico di sostituzione componente', 'descrizione dalla colonna Item');
  ok(!!f('DI BELLO Andrea', 'G'), 'cognome con particella e data come Date');
  ok(!!f('CORSI Nicolò', 'G'), 'accenti: Nicolo = Nicolò, data 22/10/2026');
  ok(r.ignote.length === 1 && /Sconosciuto/.test(r.ignote[0]), 'nomi non trovati in Manutentori segnalati');
}
console.log('\nArea sterile: esonerati e convalide mancanti');
{
  const { ctx, sheets } = nuovo();
  sheets.Liste.rows = [['Manutentori'], ['GIALLI Marco'], ['MORO Riccardo'], ['FERRI Luca']];
  sheets.Vestizioni = new Sheet('Vestizioni', [['Tecnico', 'Data vestizione', 'Durata (mesi)', 'Scadenza', 'Giorni mancanti', 'Note'], ['GIALLI Marco', new ctx.DD(2026, 0, 10), 12], ['MORO Riccardo'], ['FERRI Luca', '', '', '', '', '', 'NO']]);
  sheets.Convalide = new Sheet('Convalide', [['x', 'Convalida Sterile Marco Gialli', '2027/01/10', 'Categoria G - regolazione'], ['x', 'Convalida Sterile Marco Gialli', '2027/01/10', 'Categoria H - sostituzione']]);
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ok(JSON.stringify(ctx.esonerati_(ss)) === '["FERRI Luca"]', 'colonna Area sterile: NO = esonerato');
  ok(/sterile/i.test(sheets.Vestizioni.rows[0][6]), 'colonna Area sterile presente nel foglio Vestizioni');
  const m = ctx.convMancanti_(ss, ctx.convalideTutte_(ss).list);
  ok(m.length === 1 && m[0].man === 'MORO Riccardo' && m[0].cats.join() === 'G,H', 'chi è in sterile senza convalide: mancanti G e H (gli esonerati no)');
  ok(/mancanti: MORO G H/.test(ctx.rigaConvalide_()), 'riepilogo: riga con i mancanti');
}
console.log('\nFoglio Vestizioni: impostazioni e convalide');
{
  const { Sheet: Sh, carica: car } = require('./fake-sheets');
  const sheets = {}; let { ctx } = car(sheets);
  const D = (y, m, d) => new ctx.DD(y, m - 1, d), isD = x => Object.prototype.toString.call(x) === '[object Date]';
  Object.assign(sheets, {
    Liste: new Sh('Liste', [['Manutentori'], ['GIALLI Marco'], ['MORO Riccardo'], ['FERRI Fabio']]),
    Vestizioni: new Sh('Vestizioni', [['Tecnico', 'Data vestizione', 'Durata (mesi)', 'Scadenza', 'Giorni mancanti', 'Note'],
      ['GIALLI Marco', D(2025, 10, 28), 13], ['MORO Riccardo', D(2026, 5, 28), 13], ['FERRI Fabio', D(2026, 1, 10), '']]),
    Convalide: new Sh('Convalide', [['x', 'Convalida Sterile Marco Gialli', '2026/10/22', 'Categoria G - Intervento Meccanico di regolazione'],
      ['x', 'Convalida Sterile Marco Gialli', '2026/11/04', 'Categoria H - Intervento Meccanico di sostituzione componente']]),
    Avvisi: new Sh('Avvisi', [['Avvisi', 'Valore', '', '', 'Report automatici', 'Valore'], ['', '', '', '', 'Preavviso scadenza vestizione sterile (giorni prima)', 45], ['', '', '', '', 'Preavviso scadenza convalide GMP (giorni prima)', 90]])
  });
  sheets.Vestizioni.rows[3][6] = 'NO'; // FERRI esonerato (colonna G già scritta a mano, senza intestazione)
  sheets.Vestizioni.rows[0][6] = 'Area sterile (SI/NO)';
  const c = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet()), V = sheets.Vestizioni.rows;
  ok(c.cSter === 7 && c.cCat.G === 8 && c.cCat.H === 9 && c.cUlt.G === 10 && c.cUlt.H === 11 && c.cSet === 13, 'colonne: G area sterile, H/I scadenze, J/K ultima convalida, M–N impostazioni');
  ok(c.durata === 13 && c.avvV === 45 && c.avvC === 90 && c.durC === 12, 'impostazioni: durata 13 (la più usata), preavvisi presi dal foglio Avvisi');
  ok(!sheets.Avvisi.rows.some(r => /Preavviso/.test(r[4] || '')), 'righe di preavviso tolte dal foglio Avvisi');
  ok(isD(V[1][7]) && V[1][7].getDate() === 22 && V[2][7] === 'MANCA' && V[3][7] === 'esonerato', 'Cat. G: data / MANCA / esonerato');
  ok(/\$N\$2/.test(sheets.Vestizioni.f['2,4'] || ''), 'formula Scadenza usa la durata standard ($N$2)');
  ok((sheets.Vestizioni.cf || []).length === 5, 'colori: esonerati, vestizione scaduta/in preavviso, convalide scadute/in preavviso');
  V[3][2] = ''; const r = ctx.vestizioni_(ctx.SpreadsheetApp.getActiveSpreadsheet()).find(x => x.man === 'FERRI Fabio');
  ok(r.mesi === 13, 'durata vuota: vale lo standard');
  // nuova categoria nello scadenziario: colonna nuova prima del riquadro Impostazioni
  sheets.Convalide.rows.push(['x', 'Convalida Sterile Marco Gialli', '2027/01/01', 'Categoria F - pulizia']);
  ({ ctx } = car(sheets)); const c2 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  ok(c2.cCat.F === 12 && c2.cUlt.F === 13 && c2.cSet === 15 && V[0][14] === 'Impostazioni' && c2.durata === 13, 'categoria nuova: colonne aggiunte, Impostazioni spostate a destra');
  // ultima convalida + durata standard
  ({ ctx } = car(sheets)); V[1][c2.cUlt.G - 1] = new ctx.DD(2026, 9, 1); V[2][c2.cUlt.H - 1] = new ctx.DD(2026, 5, 1);
  const c3 = ctx.vestSetup_(ctx.SpreadsheetApp.getActiveSpreadsheet());
  const g = c3.conv.find(x => x.man === 'GIALLI Marco' && x.cat === 'G'), h = c3.conv.find(x => x.man === 'MORO Riccardo' && x.cat === 'H');
  ok(g && new Date(g.scad).getFullYear() === 2027 && new Date(g.scad).getMonth() === 9, 'ultima convalida 01/10/2026 + 12 mesi = 01/10/2027 (più lontana dello scadenziario)');
  ok(h && isD(V[2][c3.cCat.H - 1]), 'chi era MANCA con l\'ultima convalida scritta diventa valido');
  const ug = V[2][c3.cUlt.G - 1]; // MORO non ha G nello scadenziario: resta vuota
  const uo = V[1][c3.cUlt.H - 1]; // GIALLI H nello scadenziario 04/11/2026 → ultima = 04/11/2025
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
  ok(ch && /Riferimenti/.test(ch.opt.title) && ch.opt.hAxis.logScale && Math.abs(ch.opt.hAxis.viewWindow.min * ch.opt.hAxis.viewWindow.max - 1) < 1e-9 && dati[0].slice().sort().join('|') === ['MTTR relativo', 'ROSSI', 'BIANCHI', '┃ MTTR come il reparto', '━ interventi come la mediana'].sort().join('|'), 'grafico a quadranti: una serie per tecnico + linee di riferimento');
  ctx.aggiornaKpiRisorse(); ok((sheets['KPI Risorse'].charts || []).length === 1, 'ricalcolando non si duplicano i grafici');
  const celle = sheets['KPI Risorse'].rows.flat().filter(v => typeof v === 'string' && /^(↖|↗|↙|↘) /.test(v));
  ok(celle.length === 4 && celle.some(v => /ROSSI/.test(v)), 'i 4 quadranti in celle colorate con i nomi');
}
console.log('\nFoglio Programma pesante');
{
  const { ctx } = nuovo(); const fatto = [];
  let mc = 16370; const sh = { getLastColumn: () => 305, getMaxColumns: () => mc, getLastRow: () => 972, getMaxRows: () => 1000, deleteColumns: (a, n) => { fatto.push(['col', a, n]); mc -= n; }, deleteRows: (a, n) => fatto.push(['row', a, n]) };
  ctx.pulisciProgramma_(sh);
  ok(mc === 310 && fatto.filter(x => x[0] === 'col').every(x => x[2] <= 2000) && fatto.some(x => x[0] === 'row'), 'colonne vuote eliminate a blocchi da 2000 (16.370 → 310), righe in eccesso tolte');
}

console.log('\nTurni: cella "3/1" diventata data');
{
  const date = ['OTTOBRE'].concat([9, 10, 11, 12, 13, 14, 15].map(g => 'x/' + g + '/10/26'));
  const sheets = { Liste: new Sheet('Liste', [['Manutentori'], ['ROSSI Mario']]), Equipment: new Sheet('Equipment', [['Codice (QR)']]),
    Turni: new Sheet('Turni', [['Incolla'], date]) };
  const { ctx } = carica(sheets);
  sheets.Turni.rows.push(['ROSSI Mario', new ctx.DD(2026, 0, 3), new ctx.DD(2026, 2, 1), '3/2', 'R', 2, '3 1', 'F']);
  const p = ctx.turniTutti_(ctx.SpreadsheetApp.getActiveSpreadsheet())['ROSSI Mario'];
  ok(p['2026-10-09'] === '3 1' && p['2026-10-10'] === '1 3' && p['2026-10-11'] === '3 2', 'data 3 gennaio → "3 1", 1 marzo → "1 3", "3/2" → "3 2"');
  ok(p['2026-10-12'] === 'R' && p['2026-10-13'] === '2' && p['2026-10-14'] === '3 1', 'celle normali invariate');
}


console.log('\nClassificazione: voci guasto, interventi a 4, causa');
{
  const { ctx, sheets, props } = nuovo();
  const C = ctx.C, riga = (int, gua, cat) => { const r = new Array(ctx.HEAD.length).fill(''); r[C.ID - 1] = 'F' + Math.random(); r[C.INT - 1] = int; r[C.GUASTO - 1] = gua; r[C.CAT - 1] = cat; r[C.STATO - 1] = 'CHIUSO'; return r; };
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.slice(), riga('RIAVVIO', 'CAM. OMRON', 'AUTOMATION'), riga('REGOLAZIONE', 'CAM. OMRON', 'AUTOMATION'), riga('REGOLAZIONE', 'CAM. OMRON', 'MECCANICO'), riga('ALTRO', 'GUIDE', 'MECCANICO')]);
  sheets.Liste.rows = [['Categoria', 'Tipo intervento', 'Manutentori'], ['MECCANICO', 'SOSTITUZIONE', 'ROSSI Mario'], ['ELETTRICO', 'REGOLAZIONE', ''], ['AUTOMATION', 'PULIZIA', ''], ['ALTRO', 'RIAVVIO', ''], ['', 'ALTRO', '']];
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'FORMATO', 'SENSORE', 'USURA', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO', 'RIAVVIO', 'RIPARAZIONE'], ['ST3-04', '', 'CAM. OMRON', 'GUIDE', 'PARAM VIDEOJET']]);
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  ctx.classificazione_(ss);
  ok(sheets.Liste.rows.slice(1).map(r => r[1]).filter(Boolean).join() === 'REGOLAZIONE,SOSTITUZIONE,RIPRISTINO,PULIZIA', 'Liste: Tipo intervento a 4 voci (RIAVVIO, ALTRO tolti)');
  ok(sheets.Fermi.rows[1][C.INT - 1] === 'RIPRISTINO' && sheets.Fermi.rows[4][C.INT - 1] === 'ALTRO', 'storico: RIAVVIO → RIPRISTINO, il resto invariato');
  ok(sheets.Fermi.rows[0][C.CAUSA - 1] === 'Causa', 'foglio Fermi: colonna Causa aggiunta');
  ok(sheets.Guasti.rows[1].slice(2).filter(Boolean).join() === 'SENSORE,ARIA COMPRESSA,ALTRO', 'TUTTE: tolte cause e interventi (FORMATO, USURA, SPIKE, RIAVVIO, RIPARAZIONE)');
  const v = ctx.vociCat_(ss);
  ok(v['CAM. OMRON'] === 'AUTOMATION' && /storico/.test(sheets['Voci guasto'].rows.find(r => r[0] === 'CAM. OMRON')[2]), 'Voci guasto: CAM. OMRON = AUTOMATION dallo storico (2 su 3)');
  ok(v['SENSORE'] === 'ELETTRICO' && v['PARAM VIDEOJET'] === 'AUTOMATION' && v['GUIDE'] === 'MECCANICO' && !v['ALTRO'], 'voci senza storico: categoria dal nome; ALTRO escluso');
  sheets.Guasti.rows[2].push('NASTRATORI'); ctx.classificazione_(ss);
  const n = sheets['Voci guasto'].rows.length;
  ok(sheets['Voci guasto'].rows.some(r => r[0] === 'NASTRATORI') && sheets.Liste.rows[1][1] === 'REGOLAZIONE', 'voce nuova nel foglio Guasti: aggiunta in Voci guasto (la migrazione non si ripete)');
  const LT = sheets.Liste.rows[0].map(String);
  ok(LT.includes('Tipo intervento · spiegazione') && LT.includes('Causa'), 'Liste: aggiunte le colonne Tipo intervento · spiegazione e Causa');
  const cD = LT.indexOf('Tipo intervento · spiegazione'), cC = LT.indexOf('Causa');
  ok(sheets.Liste.rows[2][1] === 'SOSTITUZIONE' && /codice obbligatorio/.test(sheets.Liste.rows[2][cD]) && sheets.Liste.rows.slice(1).map(r => r[cC]).filter(Boolean).length === 6, 'spiegazione sulla stessa riga dell\'intervento, 6 cause');
  sheets['Voci guasto'].rows.find(r => r[0] === 'GUIDE')[1] = 'ELETTRICO'; ctx.classificazione_(ss);
  ok(sheets['Voci guasto'].rows.length === n && sheets['Voci guasto'].rows.find(r => r[0] === 'GUIDE')[1] === 'ELETTRICO', 'le categorie decise dal responsabile non vengono toccate');
}


console.log('\nRicodifica dei fermi vecchi');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C;
  const riga = (id, int, gua, altro, pz) => { const r = new Array(ctx.HEAD.length).fill(''); Object.assign(r, { [C.ID - 1]: id, [C.INT - 1]: int, [C.GUASTO - 1]: gua, [C.ALTRO - 1]: altro, [C.PEZZO - 1]: pz || '', [C.STATO - 1]: 'CHIUSO', [C.EQ - 1]: 'ST3-04', [C.LINEA - 1]: 'ST3', [C.CAT - 1]: 'MECCANICO' }); return r; };
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.slice(), riga('A', 'SOSTITUZIONE', 'USURA', 'LAMA SBUCCIATRICE 9500085'), riga('B', 'ALTRO', 'ALTRO', 'regolato chiusura astucci lato operatore'),
    riga('C', 'RIAVVIO', 'CAM. OMRON', ''), riga('D', 'REGOLAZIONE', 'H7, Sblocco Homing', '')]);
  sheets.Liste.rows = [['Categoria', 'Tipo intervento', 'Manutentori'], ['MECCANICO', 'SOSTITUZIONE', 'ROSSI Mario'], ['ELETTRICO', 'RIAVVIO', ''], ['AUTOMATION', 'ALTRO', '']];
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'SENSORE', 'ALTRO'], ['ST3-04', 'AST', 'CHIUSURE AST', 'CAM. OMRON', 'H7, Sblocco Homing']]);
  ctx.ricodificaPrepara();
  const R = sheets.Ricodifica.rows, by = id => R.find(r => r[0] === id);
  ok(by('A')[11] === 'SOSTITUZIONE' && by('A')[12] === 'USURA' && by('A')[13] === '9500085' && !by('A')[10], 'lama: SOSTITUZIONE, causa USURA, codice pezzo dalla descrizione, guasto da scegliere');
  ok(by('B')[10] === 'CHIUSURE AST' && by('B')[11] === 'REGOLAZIONE', 'ALTRO + "regolato chiusura astucci": CHIUSURE AST · REGOLAZIONE');
  ok(by('C')[11] === 'RIPRISTINO' && by('D')[11] === 'RIPRISTINO', 'RIAVVIO e Sblocco Homing → RIPRISTINO');
  by('A')[10] = 'LAMA SBUCCIATRICE';
  ctx.ricodificaApplica();
  const F = id => sheets.Fermi.rows.find(r => r[C.ID - 1] === id);
  ok(F('A')[C.GUASTO - 1] === 'LAMA SBUCCIATRICE' && F('A')[C.CAUSA - 1] === 'USURA' && F('A')[C.PEZZO - 1] === '9500085', 'applicato al foglio Fermi per ID (guasto, causa, pezzo)');
  ok(sheets.Guasti.rows[2].includes('LAMA SBUCCIATRICE') && F('A')[C.CAT - 1] === 'MECCANICO', 'voce nuova aggiunta al foglio Guasti della macchina, categoria da Voci guasto');
  ok(F('C')[C.CAT - 1] === 'AUTOMATION' && by('A')[14] && by('A')[14] !== 'nessuna modifica', 'categoria dalla voce (CAM. OMRON → AUTOMATION); colonna Applicato compilata');
}


console.log('\nKPI: per causa, per intervento, micro-fermi, MTBF stimato, guasti cronici');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C, DD = ctx.DD, oggi = new DD(); oggi.setHours(0, 0, 0, 0);
  const riga = (gg, min, gua, int, linea) => { const r = new Array(ctx.HEAD.length).fill(''); const a = new DD(oggi.getTime() - gg * 864e5 + 10 * 3600e3);
    Object.assign(r, { [C.STATO - 1]: 'CHIUSO', [C.DATA - 1]: new DD(a.getFullYear(), a.getMonth(), a.getDate()), [C.INIZIO - 1]: a, [C.FINE - 1]: new DD(a.getTime() + min * 60000),
      [C.GUASTO - 1]: gua, [C.INT - 1]: int, [C.LINEA - 1]: linea || 'ST3', [C.MACC - 1]: 'ASTUCCIATRICE', [C.EQ - 1]: 'ST3-04', [C.LOTTO - 1]: '601095' }); return r; };
  const vals = [riga(1, 60, 'CHIUSURE AST', 'SOSTITUZIONE'), riga(2, 30, 'CHIUSURE AST', 'REGOLAZIONE'), riga(3, 1, 'GUIDE', 'RIAVVIO'), riga(3, 2, 'GUIDE', 'REGOLAZIONE')];
  const f = { from: new DD(oggi.getTime() - 29 * 864e5), toEnd: new DD(oggi.getTime() + 864e5), linea: 'TUTTE', area: '', aree: {}, code: '', ore: 100, cause: ['USURA', '', 'USURA', 'MATERIALE FUORI SPEC.'] };
  const k = ctx.calcolaKpi_(vals, f);
  ok(k.n === 4 && k.micro === 2 && Math.abs(k.mttr * 1440 - 45) < 0.01, 'micro-fermi (< 3 min) contati a parte: MTTR = media dei fermi veri (45 min)');
  ok(k.perCausa[0][0] === 'USURA' && k.perCausa[0][1] === 2 && k.perCausa[k.perCausa.length - 1][0] === '(non indicata)', 'tabella per causa: USURA in cima, "non indicata" in fondo');
  ok(k.perInt.some(r => r[0] === 'RIPRISTINO' && r[1] === 1) && k.mtbf === 25, 'per intervento (RIAVVIO conta come RIPRISTINO); MTBF = ore / n° fermi');
  // stima ore dal Programma
  const d = n => new DD(oggi.getTime() - n * 864e5);
  sheets.Equipment.rows = [['Codice', 'Linea', 'Area', 'Macchina', 'Attivo'], ['ST3-04', 'ST3', 'Packaging', 'AST', 'SI'], ['SPL1-01', 'SPL1', 'Packaging', 'X', 'SI']];
  sheets.Programma = new Sheet('Programma', [['x'], ['y'], ['', d(9), d(8), d(7), d(6), d(5), d(4), d(3), d(2), d(1), oggi, d(-1)], ['st3 rip', '601095', '', '601096', '', '', '', '', '', '601097', '', '601098'], ['spl1', '', '601001']]);
  const st = ctx.oreStimate_(ctx.SpreadsheetApp.getActiveSpreadsheet(), { from: d(29), toEnd: new DD(oggi.getTime() + 864e5) }, 'PACKAGING', 'TUTTE');
  ok(st.giorni === 4 && st.ore === 96 && st.linee === 2, 'ore stimate: giorni con lotti fino a oggi (domani escluso) × 24 h, su tutte le linee dell\'area');
  // guasti cronici
  const f2 = (gg, min, gua) => ({ a: new DD(oggi.getTime() - gg * 864e5), b: new DD(oggi.getTime() - gg * 864e5 + min * 60000), guasto: gua, linea: 'ST3', macc: 'AST' });
  const tutti = [f2(3, 60, 'H7'), f2(10, 60, 'H7'), f2(40, 20, 'H7'), f2(5, 30, 'GUIDE'), f2(6, 30, 'GUIDE'), f2(35, 90, 'GUIDE'), f2(4, 200, 'ALTRO'), f2(4, 200, 'ALTRO')];
  const rr = ctx.croniciRighe_(tutti, oggi, (l, m) => m);
  ok(rr.length === 4 && /H7 · ST3 AST ×2 · 2h ▲ \+1h 40/.test(rr[2]) && /GUIDE.*×2 · 1h ▼ −30/.test(rr[3]) && !rr.join().includes('ALTRO'), 'guasti cronici: ore delle ultime 4 settimane con tendenza (▲/▼), ALTRO escluso');
}


console.log('\nRaccolta uniforme: definizioni, rivisto, correzioni');
{
  const { ctx, sheets, props } = nuovo(); const C = ctx.C, ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  props.CLASS_V1 = '1'; props.CLASS_V2 = '1';
  sheets['Voci guasto'] = new Sheet('Voci guasto', [['Voce (come nel foglio Guasti)', 'Categoria', 'Da dove viene (giallo = da controllare)', '', 'Una riga per voce del foglio Guasti: …'], ['INCASTRO', 'MECCANICO', 'dal nome'], ['GUIDE', 'MECCANICO', 'dal nome']]);
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'INCASTRO', 'GUIDE']]);
  ctx.classificazione_(ss);
  const V = sheets['Voci guasto'].rows;
  ok(V[0][3] === 'Quando si usa (la legge il tecnico)' && /bloccato nel trasporto/.test(V[1][3]) && !V[2][3] && !V[0][4] && /^Una riga/.test(V[0][5]), 'Voci guasto: colonna "Quando si usa" (definizioni di partenza), nota spostata');
  ok(/bloccato/.test(ctx.vociDesc_(ss).INCASTRO), 'definizioni lette per l\'app');
  const r = new Array(ctx.HEAD.length).fill(''); r[C.ID - 1] = 'F9'; r[C.STATO - 1] = 'CHIUSO'; r[C.MAN - 1] = 'ROSSI Mario';
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.slice(), r]);
  let err = ''; try { ctx.rivistoFermo_({ fermo: 'F9', manutentore: 'ROSSI Mario' }); } catch (e) { err = e.message; }
  ok(/coordinatori/.test(err), 'solo un coordinatore può segnare "rivisto"');
  ctx.rivistoFermo_({ fermo: 'F9', manutentore: 'VERDI Anna', ts: Date.now() });
  ok(/^VERDI Anna · /.test(sheets.Fermi.rows[1][C.RIV - 1]) && sheets.Fermi.rows[0][C.PROP - 1] === 'Voce proposta (con ALTRO)', 'rivisto: nome e data nella colonna Rivisto; colonne in fondo create');
  sheets.Correzioni = new Sheet('Correzioni', [['h'], [new ctx.DD(), 'F9', 2, 'ST3', 'X', '', 'VERDI Anna', 'Modifica diretta: Tipo guasto: A → B', 'FATTA']]);
  const d = ctx.correzioniDett_(ss);
  ok(d.F9 && d.F9.da === 'VERDI Anna' && d.F9.txt === 'Tipo guasto: A → B', 'correzioni: chi e cosa, per il messaggio al tecnico');
}

console.log('\nFermo con più interventi (giri)');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C;
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.slice()]);
  const t0 = Date.now() - 50 * 60000, t1 = t0 + 20 * 60000, t2 = t1 + 25 * 60000;
  const base = { code: 'ST3-04', lotto: '012345', manutentore: 'ROSSI Mario' };
  ctx.openFermo_({ ...base, id: 'F1', ts: t0 });
  ctx.closeFermo_({ ...base, id: 'F1', ts: t1, guasto: 'GUIDE', categoria: 'MECCANICO', intervento: 'REGOLAZIONE', altro: 'guida regolata', gruppo: 'F1', giro: 1 });
  ctx.openFermo_({ ...base, id: 'F2', ts: t1, gruppo: 'F1', giro: 2 });
  const R = sheets.Fermi.rows;
  ok(R[1][C.GRUPPO - 1] === 'F1' && R[1][C.GIRO - 1] === 1 && R[1][C.STATO - 1] === 'CHIUSO', 'intervento 1: chiuso, Fermo n. F1, Intervento n. 1');
  ok(R[2][C.GRUPPO - 1] === 'F1' && R[2][C.GIRO - 1] === 2 && R[2][C.STATO - 1] === 'APERTO' && R[2][C.INIZIO - 1].getTime() === t1, 'intervento 2: aperto dallo stesso istante, stesso Fermo n.');
  ok(R[0][C.GRUPPO - 1] === 'Fermo n.' && R[0][C.GIRO - 1] === 'Intervento n.' && R[0][C.MULTI - 1] === 'Più interventi?', 'intestazioni delle colonne nuove');
  ctx.closeFermo_({ ...base, id: 'F2', ts: t2, guasto: 'INCASTRO', categoria: 'MECCANICO', intervento: 'PULIZIA', altro: 'pulito e poi sostituito sensore', multi: 'Descrizione con più interventi: confermato "uno solo" dal tecnico' });
  ok(R[2][C.GIRO - 1] === 2 && /uno solo/.test(R[2][C.MULTI - 1]), 'chiusura: il numero d\'intervento resta 2, avviso "più interventi" registrato');
  const f = ctx.init_().fermi;
  ok(f.length === 2 && f.every(x => x.gruppo === 'F1') && f.find(x => x.id === 'F2').giro === 2 && /uno solo/.test(f.find(x => x.id === 'F2').multi), 'l\'app riceve Fermo n., Intervento n. e avviso');
  ctx.openFermo_({ ...base, id: 'F3', ts: t2 + 60000 });
  ok(!R[3][C.GRUPPO - 1] && !ctx.init_().fermi.find(x => x.id === 'F3').gruppo, 'fermo normale: nessun Fermo n.');
  ctx.controllaRegistrazioni();
  const ctl = sheets.Controlli.rows;
  {
    const { ctx: c3, sheets: s3 } = nuovo(); const h = c3.HEAD.slice(); h[C.GRUPPO - 1] = 'Note mie';
    s3.Fermi = new Sheet('Fermi', [h]);
    c3.openFermo_({ ...base, id: 'X1', ts: t0, gruppo: 'X1', giro: 2 });
    ok(s3.Fermi.rows[0][C.GRUPPO - 1] === 'Note mie' && !s3.Fermi.rows[1][C.GRUPPO - 1] && s3.Fermi.rows[0][C.CAUSA - 1] === 'Causa', 'colonna AB già usata per altro: non viene sovrascritta');
  }
  ok(ctl.some(r => r[0] === 'Più interventi in un fermo?' && /ROSSI/.test(r[1]) && r[2] === 'Fermi riga 3') && !ctl.some(r => r[0] === 'Fermi sovrapposti'), 'Controlli: segnala l\'avviso ignorato; interventi consecutivi non sono "sovrapposti"');
}


console.log('\nFermi (macchina ferma) e interventi separati');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C, DD = ctx.DD, oggi = new DD(); oggi.setHours(0, 0, 0, 0);
  const riga = (gg, ini, min, gua) => { const r = new Array(ctx.HEAD.length).fill(''); const a = new DD(oggi.getTime() - gg * 864e5 + ini * 60000);
    Object.assign(r, { [C.STATO - 1]: 'CHIUSO', [C.DATA - 1]: new DD(a.getFullYear(), a.getMonth(), a.getDate()), [C.INIZIO - 1]: a, [C.FINE - 1]: new DD(a.getTime() + min * 60000),
      [C.GUASTO - 1]: gua, [C.INT - 1]: 'REGOLAZIONE', [C.LINEA - 1]: 'ST3', [C.MACC - 1]: 'AST', [C.EQ - 1]: 'ST3-04', [C.LOTTO - 1]: '601095' }); return r; };
  // fermo 1 = 3 interventi consecutivi (gruppo G1), fermo 2 = un intervento solo
  const vals = [riga(1, 600, 20, 'GUIDE'), riga(1, 620, 10, 'SENSORE'), riga(1, 630, 30, 'GUIDE'), riga(2, 600, 15, 'INCASTRO')];
  const k = ctx.calcolaKpi_(vals, { from: new DD(oggi.getTime() - 29 * 864e5), toEnd: new DD(oggi.getTime() + 864e5), linea: 'TUTTE', area: '', aree: {}, code: '', ore: 100, cause: [], gruppi: ['G1', 'G1', 'G1', ''] });
  ok(k.n === 2 && k.nInt === 4, 'N° fermi = 2 (macchina ferma), N° interventi = 4');
  ok(k.mtbf === 50 && Math.abs(k.durFermo * 1440 - 37.5) < 0.01 && Math.abs(k.mttr * 1440 - 18.75) < 0.01, 'MTBF sui fermi (100 h / 2), durata media fermo 37,5 min, MTTR sugli interventi 18,75 min');
  ok(k.macchine[0][2] === 2 && k.pareto.find(r => r[0] === 'GUIDE')[1] === 2, 'fermi per macchina: 2 fermi; Pareto: GUIDE 2 interventi');
  const f = (min0, min, gua, g, id) => ({ id: id, gruppo: g, a: new DD(oggi.getTime() + min0 * 60000), b: new DD(oggi.getTime() + (min0 + min) * 60000), guasto: gua, linea: 'ST3', macc: 'AST', area: 'PACKAGING', orari: '' });
  const L = [f(600, 20, 'GUIDE', 'G1', 'a'), f(620, 10, 'SENSORE', 'G1', 'b'), f(630, 30, 'GUIDE', 'G1', 'c'), f(700, 45, 'INCASTRO', '', 'd')];
  const rr = ctx.righeReparti_(L, ['ST3']).join('|');
  ok(/\(2 fermi · 4 int\.\)/.test(rr) && /1h 45/.test(rr), 'passaggio turno / riepilogo: "(2 fermi · 4 int.)", ore totali uguali');
  const pl = ctx.piuLunghi_(L, (l, m) => m, false).join('|');
  ok(/GUIDE \+ SENSORE · 1h \(3 interventi\)/.test(pl) && pl.indexOf('GUIDE + SENSORE') < pl.indexOf('INCASTRO'), 'più lunghi: il fermo con 3 interventi conta intero (1h, guasti in fila)');
  ok(ctx.righeReparti_([L[3]], ['ST3']).join().includes('(1)'), 'fermo con un solo intervento: resta "(1)" come prima');
}


console.log('\nPiù operazioni nella descrizione (antifurbi)');
{
  const { ctx } = nuovo();
  ok(ctx.piuOperazioni_('regolato posizione sensore e regolato lama sbucciatrice') && ctx.piuOperazioni_('rimesso in dima pinze ctg e regolato sensore ingrese'), 'i fermi inventati di prova: riconosciuti');
  ok(ctx.piuOperazioni_('Pulizia rulli e pulegge della videojet e sostituita testina di stampa'), 'tipi diversi (pulizia + sostituzione)');
  ['regolato distanz acart', 'leva sensore S19 trovata spostata, rimessa in posizione', 'Regolato larghezza tramoggia alzate da 1200 a 1215', 'Sostituito contrasto in gomma etichettatrice', 'Regolato guide e sensore ingresso', 'Eseguito sblocco con Guantone.']
    .forEach(x => ok(!ctx.piuOperazioni_(x), 'un\'operazione sola: non segnalato — "' + x + '"'));
}


console.log('\nFormati (solo coordinatori, file collegato)');
{
  const { ctx, props } = nuovo();
  const P = [['CONFEZIONATO SU'], ['Codice SAP', 'Descrizione', 'Linea PRIMARIO\n(filling)', 'ST3', 'SPL1', 'Contenitore', 'Tappo', 'Contagocce', 'Dosaggio\n(ml / g)', 'Etichetta', 'Fustella astuccio', 'Istruzione', 'Pezzi per astuccio', 'Fustella cartone', 'Fardello', 'Fase', 'Codice formato\ndel codice', 'ESITO'],
    ['6000001', 'PROVA A 5ML', 'ST3', 'X', '', 'FL.X', 'T1', 'N/A', '5 ml', 'E1', 'A1', 'P1', '1', 'C1', '', 'SECONDARIO', 'FP-A', 'OK'],
    ['5000002', 'PROVA B', 'ST3', '', '', 'FL.X', 'T1', 'N/A', '10 ml', '', '', '', '', '', '', 'PRIMARIO', 'FF-X', 'OK']];
  const M = [['Macchina_ID', 'Linea_ID', 'Reparto', 'Pos.', 'Macchina', 'Funzione', 'Fase', 'Lavora quando', 'Contenitore', 'Tappo', 'Contagocce', 'Dosaggio (ml)', 'Etichetta', 'Astuccio', 'Istruzione', 'Pezzi/ astuccio', 'Cartone', 'Fardello', 'Prodotto', 'Verificata'],
    ['ST3-02', 'ST3', 'S3', 2, 'RIEMPITRICE', 'x', 'PRIMARIO', 'SEMPRE', 'P', 'P', '', 'R', '', '', '', '', '', '', '', 'SÌ'], ['ST3-04', 'ST3', 'S3', 4, 'ASTUCCIATRICE', 'x', 'SECONDARIO', 'SEMPRE', 'P', '', '', '', '', 'P', 'R', '', '', '', '', 'NO']];
  const Ln = [['Linea_ID', 'Reparto', 'Sigla linea', 'Tipo linea', 'Prosegue su (secondario)'], ['ST3', 'STERILE 3', 'IMA', '', 'ST3, SPL1']];
  const fs = { PRODOTTI: new Sheet('PRODOTTI', P), MACCHINE: new Sheet('MACCHINE', M), LINEE: new Sheet('LINEE', Ln) };
  ctx.SpreadsheetApp.openById = () => ({ getSheets: () => Object.values(fs) });
  let err = ''; try { ctx.formatiDati_(); } catch (e) { err = e.message; }
  ok(/non collegato/.test(err), 'senza file collegato: messaggio chiaro');
  props.FORMATI_ID = 'abcdefghijklmnopqrstuvwxyz';
  const d = ctx.formatiDati_();
  ok(d.prod.length === 2 && d.prod[0].ls.join() === 'ST3' && d.prod[0].a[3] === '5 ml' && d.prod[0].fmt === 'FP-A' && d.prod[1].f === 'PRIMARIO', 'PRODOTTI: codice, linee con X, caratteristiche, formato, fase');
  ok(d.macc.length === 2 && d.macc[0].m.slice(0, 4).join() === 'P,P,,R' && d.macc[1].ver === false && d.macc[0].n === 'RIEMPITRICE', 'MACCHINE: matrice P/R per caratteristica, "verificata"');
  ok(d.linee[0].pro === 'ST3, SPL1' && !d.avvisi.length, 'LINEE lette, nessun avviso');
  const prova = man => { const r = ctx.doGet({ parameter: { key: 'k', action: 'formati', man: man, dev: 'T-1' } }); return JSON.parse(r.t || r.getContent && r.getContent() || '{}'); };
  ctx.checkKey_ = () => {}; ctx.checkTelefono_ = () => {};
  ok(/coordinatori/.test(prova('ROSSI Mario').error || ''), 'un tecnico chiede i formati: rifiutato dallo script');
  const j = prova('VERDI Anna');
  ok(j.ok && j.prod.length === 2, 'un coordinatore li riceve');
}

fine('Script');
