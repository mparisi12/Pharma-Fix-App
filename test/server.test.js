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
console.log('\nConvalide MF (ex Vestizioni) e Scadenziario MF (ex Convalide)');
{
  const { Sheet: Sh, carica: car } = require('./fake-sheets');
  const sheets = {}; let { ctx } = car(sheets);
  const isD = x => x && typeof x.getTime === 'function' && !isNaN(x.getTime()), ymd = x => isD(x) ? x.getFullYear() + '-' + (x.getMonth() + 1) + '-' + x.getDate() : String(x);
  const D = (y, m, d) => new ctx.DD(y, m - 1, d);
  // foglio com'era (formato "Vestizioni" con le colonne aggiunte nel tempo e il riquadro Impostazioni)
  const OLD = ['Tecnico', 'Data vestizione', 'Durata (mesi)', 'Scadenza', 'Giorni mancanti', 'Note', 'Area sterile (SI/NO)', 'Scadenza Cat. G\nregolazione', 'Scadenza Cat. H\nsostituzione',
    'Ultima convalida Cat. G', 'Ultima convalida Cat. H', 'Scadenza Cat. MD_20\ncorso', 'Ultima convalida Cat. MD_20', '', 'Impostazioni', 'Valore'];
  const set = ['Durata standard vestizione (mesi)', 'Preavviso vestizione (giorni prima)', 'Preavviso convalide GMP (giorni prima)', 'Durata standard convalide interventi (mesi)',
    'Tolleranza vestizione (giorni dopo la scadenza)', 'Tolleranza convalide GMP (giorni dopo la scadenza)', 'Durata corso MD_20 (mesi)'], val = [12, 30, 45, 12, 30, 20, 12];
  const riga = (nome, vest, note, ster, md) => { const r = new Array(16).fill(''); r[0] = nome; r[1] = vest || ''; r[2] = vest ? 12 : ''; r[5] = note || ''; r[6] = ster || ''; r[12] = md || ''; return r; };
  const V0 = [OLD, riga('MORO Riccardo'), riga('GIALLI Marco', D(2026, 1, 10), 'nota di Gialli', '', D(2025, 5, 22)), riga('FERRI Luca', '', '', 'NO'), riga('', ''), riga('', ''), riga('', ''), riga('', ''), riga('', '')];
  set.forEach((x, i) => { V0[1 + i][14] = x; V0[1 + i][15] = val[i]; });
  Object.assign(sheets, {
    Liste: new Sh('Liste', [['Manutentori'], ['GIALLI Marco'], ['MORO Riccardo'], ['FERRI Luca']]),
    Vestizioni: new Sh('Vestizioni', V0),
    Convalide: new Sh('Convalide', [['Codice processo (link)', 'Oggetto', 'Scadenza', 'Item'], ['SC MF.1', 'Convalida Sterile Marco Gialli', '2027/01/10', 'Categoria G - Intervento Meccanico di regolazione'],
      ['SC MF.1', 'Convalida Sterile Marco Gialli', '2027/03/05', 'Categoria H - Intervento Meccanico di sostituzione componente'], ['SC MF.9', 'Formazione Riccardo Moro', '2026/12/01', 'Corso MD 20']])
  });
  let ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  const c = ctx.vestSetup_(ss), V = sheets.Vestizioni.rows;
  ok(ss.getSheetByName('Convalide MF') === sheets.Vestizioni && ss.getSheetByName('Scadenziario MF') === sheets.Convalide && !ss.getSheetByName('Vestizioni') && !ss.getSheetByName('Convalide'),
    'fogli rinominati: Vestizioni → Convalide MF, Convalide → Scadenziario MF (dati tenuti)');
  ok(V[0].slice(0, 13).map(x => String(x).split('\n')[0]).join('|') === 'Tecnico|Area sterile (SI/NO)|MD_20|Vestizione|Partecipazione MF|Cat. G regolazione|Cat. H sostituzione|Scadenza|Scadenza|Scadenza|Scadenza|Scadenza|Note' &&
    V[0].slice(7, 12).map(x => String(x).split('\n')[1]).join('|') === 'MD_20|Vestizione|Partecipazione MF|Cat. G|Cat. H', 'colonne in ordine: MD_20, Vestizione, Partecipazione MF, Cat. G, Cat. H, poi le scadenze nello stesso ordine, poi Note');
  const G = V.find(r => r[0] === 'GIALLI Marco'), M = V.find(r => r[0] === 'MORO Riccardo'), F = V.find(r => r[0] === 'FERRI Luca');
  ok(ymd(G[2]) === '2025-5-22' && ymd(G[3]) === '2026-1-10' && G[4] === '' && G[12] === 'nota di Gialli', 'dati spostati: MD_20 e vestizione "fatto il", note tenute; colonne Durata e Giorni mancanti tolte');
  ok(ymd(G[7]) === '2026-5-22' && ymd(G[8]) === '2027-1-10' && G[9] === 'MANCA' && ymd(G[10]) === '2027-1-10' && ymd(G[11]) === '2027-3-5', 'scadenze: MD_20 e vestizione = data + 12 mesi; partecipazione MF MANCA; Cat. G e H dallo scadenziario');
  ok(ymd(G[5]) === '2026-1-10' && ymd(G[6]) === '2026-3-5', 'Cat. G e H "ultima convalida" scritte dallo script (scadenza dello scadenziario − 12 mesi)');
  ok(ymd(M[7]) === '2026-12-1' && M[8] === 'MANCA' && M[10] === 'MANCA', 'MD_20 letto dallo scadenziario ("Corso MD 20"); il resto MANCA');
  ok(F.slice(7, 12).every(x => x === 'esonerato') && JSON.stringify(ctx.esonerati_(ss)) === '["FERRI Luca"]', 'Area sterile NO: esonerato da tutto');
  ok(c.avvV === 30 && c.avvC === 45 && c.tolV === 30 && c.tolC === 20 && c.durMF === 12 && c.durMD === 12 && c.durata === 12, 'impostazioni vecchie tenute, partecipazione MF 12 mesi di base');
  const cs = V[0].indexOf('Impostazioni'); ok(cs === 14 && V.slice(1, 9).map(r => r[cs]).some(x => /^Durata partecipazione MF/.test(x)) && V.slice(1, 12).some(r => /^Giallo chiaro/.test(r[cs] || '')), 'riquadro Impostazioni in O, con la durata della partecipazione MF e la legenda');
  ok((sheets.Vestizioni.cf || []).length === 7, 'colori: esonerati + (vestizione) + (corsi e convalide) oltre tolleranza / in tolleranza / in preavviso');
  const conv = ctx.convalide_(ss);
  ok(conv.cats.join() === 'G,H,MD_20,MF' && conv.list.some(x => x.man === 'GIALLI Marco' && x.cat === 'MD_20'), 'app: categorie G, H, MD_20, MF; MD_20 di Gialli dalla data scritta');
  const m = ctx.convMancanti_(ss, conv.list);
  ok(m.find(x => x.man === 'GIALLI Marco').cats.join() === 'MF' && m.find(x => x.man === 'MORO Riccardo').cats.join() === 'G,H,MF' && !m.some(x => x.man === 'FERRI Luca'), 'mancanti: Gialli la partecipazione MF, Moro G, H e MF; esonerati no');
  ok(/mancanti: .*MORO G H MF/.test(ctx.rigaConvalide_()), 'riepilogo del lunedì: riga con i mancanti');
  ok((ctx.vestizioni_(ss).find(x => x.man === 'GIALLI Marco') || {}).scad === D(2027, 1, 10).getTime(), 'app: vestizione di Gialli scade il 10/01/2027');
  // durata partecipazione MF 6 mesi, data scritta a mano
  ({ ctx } = car(sheets)); ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  V.find(r => /^Durata partecipazione MF/.test(r[cs] || ''))[cs + 1] = 6; V.find(r => r[0] === 'MORO Riccardo')[4] = new ctx.DD(2026, 2, 1);
  const c2 = ctx.vestSetup_(ss), M2 = V.find(r => r[0] === 'MORO Riccardo');
  ok(c2.durMF === 6 && ymd(M2[9]) === '2026-9-1' && c2.conv.some(x => x.man === 'MORO Riccardo' && x.cat === 'MF'), 'partecipazione MF: 01/03/2026 + 6 mesi = 01/09/2026');
  ok(V.find(r => r[0] === 'GIALLI Marco')[12] === 'nota di Gialli' && V[0][14] === 'Impostazioni', 'rilanciato: il foglio non si riscrive, note e impostazioni restano');
  // categoria nuova nello scadenziario: le sue due colonne in fondo, dati tenuti
  sheets.Convalide.rows.push(['SC MF.1', 'Convalida Sterile Marco Gialli', '2027/06/01', 'Categoria F - pulizia']);
  ({ ctx } = car(sheets)); ss = ctx.SpreadsheetApp.getActiveSpreadsheet(); ctx.vestSetup_(ss);
  const h = V[0].map(x => String(x).replace(/\n/g, ' '));
  ok(h.indexOf('Cat. F ultima convalida') === 7 && h.indexOf('Scadenza Cat. F') === 13 && h[14] === 'Note' && h[16] === 'Impostazioni' && V.find(r => r[0] === 'GIALLI Marco')[14] === 'nota di Gialli' && ymd(V.find(r => r[0] === 'GIALLI Marco')[13]) === '2027-6-1',
    'categoria F nuova: colonne aggiunte dopo Cat. H e dopo Scadenza Cat. H, note e impostazioni spostate');
  // riepilogo: vestizione scaduta da 10 giorni con tolleranza 45 = in tolleranza; da 60 = scaduta
  ({ ctx } = car(sheets)); ss = ctx.SpreadsheetApp.getActiveSpreadsheet(); const oggi = new ctx.DD(); oggi.setHours(0, 0, 0, 0);
  V.find(r => /^Tolleranza vestizione/.test(String(r[16] || '')))[17] = 45;
  V.find(r => r[0] === 'GIALLI Marco')[3] = new ctx.DD(oggi.getFullYear() - 1, oggi.getMonth(), oggi.getDate() - 10);
  V.find(r => r[0] === 'MORO Riccardo')[3] = new ctx.DD(oggi.getFullYear() - 1, oggi.getMonth(), oggi.getDate() - 60);
  const rv = ctx.rigaVestizioni_();
  ok(/in tolleranza fino al: GIALLI/.test(rv) && /scadute: MORO/.test(rv), 'riepilogo vestizioni: GIALLI in tolleranza, MORO scaduto: ' + rv);
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
  ok(V[0][3] === 'Quando si usa (la legge il tecnico)' && /bloccato nel trasporto/.test(V[1][3]) && !V[2][3] && V[0][4] === 'Unisci in (voce che resta)' && /^Una riga/.test(V[0][5]), 'Voci guasto: colonna "Quando si usa" (definizioni di partenza), nota spostata, colonna E "Unisci in"');
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
  ok(k.macchine[0][2] === 2 && k.pareto.find(r => r[0] === 'REG · GUIDE')[1] === 2, 'fermi per macchina: 2 fermi; Pareto: REG · GUIDE 2 interventi');
  const v2 = vals.concat([Object.assign(riga(3, 600, 25, 'GUIDE'), { [C.INT - 1]: 'SOSTITUZIONE' })]);
  const k2 = ctx.calcolaKpi_(v2, { from: new DD(oggi.getTime() - 29 * 864e5), toEnd: new DD(oggi.getTime() + 864e5), linea: 'TUTTE', area: '', aree: {}, code: '', ore: 100, cause: [], gruppi: ['G1', 'G1', 'G1', '', ''] });
  ok(k2.pareto.find(r => r[0] === 'REG · GUIDE')[1] === 2 && k2.pareto.find(r => r[0] === 'SOS · GUIDE')[1] === 1 && /^REG · GUIDE · 40%/.test(k2.top) && /^REG · GUIDE/.test(k2.macchine[0][5]), 'regolazione e sostituzione della stessa voce contate separatamente (Pareto, guasto n°1, per macchina)');
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


console.log('\nVoci guasto: tendina per macchina nel foglio Fermi e unione dei doppioni');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C, ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  sheets.Guasti = new Sheet('Guasti', [['Codice', 'Macchina', 'Tipi'], ['TUTTE', '', 'SENSORE', 'ALTRO'], ['ST3', '', 'GUIDE'], ['ST3-04', 'ASTUCCIATRICE', 'MAG. ASTUCCI', 'Magazzino astucci', 'CAM. OMRON'], ['ST3-02', '', 'H7, Sblocco Homing', 'H7 GUANTONE, Sblocco Homing']]);
  const riga = (eq, lin, mac, g) => { const r = new Array(30).fill(''); r[C.EQ - 1] = eq; r[C.LINEA - 1] = lin; r[C.MACC - 1] = mac; r[C.GUASTO - 1] = g; return r; };
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.concat(ctx.EXTRA_H), riga('ST3-04', 'ST3', 'ASTUCCIATRICE', 'Magazzino astucci'), riga('', 'ST3', 'RIEMPITRICE', 'H7 GUANTONE, Sblocco Homing'), riga('', 'ST7', 'BOH', 'X'), riga('ST3-04', 'ST3', 'ASTUCCIATRICE', 'CAM. OMRON'), riga('ST3-04', 'ST3', 'ASTUCCIATRICE', 'mag. astucci ')]);
  sheets.Soluzioni = new Sheet('Soluzioni', [ctx.SOL_H, ['s1', '', 'ST3-04', 'ST3', 'ASTUCCIATRICE', 'MAG. ASTUCCI', 'p', 's']]);
  sheets['Elenchi_Fermi'] = new Sheet('Elenchi_Fermi', [['M', 'T'], ['', 'SENSORE']]);
  ctx.fermiGuastoDV_(ss, sheets.Fermi, 2, 4);
  const dv = sheets.Fermi.dvr, l = r => (dv[r + ',' + C.GUASTO].requireValueInList || [[]])[0];
  ok(l(2).join('|') === 'MAG. ASTUCCI|Magazzino astucci|CAM. OMRON|GUIDE|SENSORE|ALTRO', 'riga ST3-04: solo le voci dell\'astucciatrice + linea ST3 + TUTTE, ALTRO in fondo');
  ok(l(3).indexOf('H7 GUANTONE, Sblocco Homing') >= 0 && l(3).indexOf('MAG. ASTUCCI') < 0, 'riga scritta a mano senza codice: macchina trovata da linea + nome (riempitrice ST3)');
  ok(!!dv['4,' + C.GUASTO].requireValueInRange, 'macchina sconosciuta: tutte le voci');
  ok(dv['2,' + C.GUASTO] === dv['5,' + C.GUASTO], 'stessa macchina: stessa regola (una sola per codice)');
  // onEdit: cambia la macchina di una riga → nuova tendina
  sheets.Fermi.rows[3][C.EQ - 1] = 'ST3-02'; sheets.Fermi.dvr = {};
  ctx.onEdit({ range: { getSheet: () => sheets.Fermi, getRow: () => 4, getColumn: () => C.EQ, getLastColumn: () => C.EQ, getNumRows: () => 1, getA1Notation: () => 'T4' }, source: ss });
  ok(sheets.Fermi.dvr && (sheets.Fermi.dvr['4,' + C.GUASTO].requireValueInList || [[]])[0].indexOf('H7, Sblocco Homing') === 0, 'modifica a mano del codice macchina: la tendina Tipo guasto si aggiorna');
  // Unione: Magazzino astucci → MAG. ASTUCCI, H7 GUANTONE → H7, catena CAM. OMRON → CAMERA → CAMERA OMRON (rinomina)
  sheets['Voci guasto'] = new Sheet('Voci guasto', [['Voce (come nel foglio Guasti)', 'Categoria', 'Da dove viene', 'Quando si usa', 'Unisci in (voce che resta)'],
    ['MAG. ASTUCCI', 'MECCANICO'], ['Magazzino astucci', 'MECCANICO', '', '', 'MAG. ASTUCCI'], ['H7, Sblocco Homing', 'AUTOMATION'], ['H7 GUANTONE, Sblocco Homing', 'AUTOMATION', '', '', 'H7, Sblocco Homing'], ['CAM. OMRON', 'AUTOMATION', '', '', 'CAMERA OMRON'], ['H7, Sblocco Homing', 'AUTOMATION']]);
  const log = ctx.unioniApplica_(ss), G = sheets.Guasti.rows, F = sheets.Fermi.rows, VG = sheets['Voci guasto'].rows;
  ok(G[3].slice(2).filter(Boolean).join('|') === 'MAG. ASTUCCI|CAMERA OMRON' && G[4].slice(2).filter(Boolean).join('|') === 'H7, Sblocco Homing', 'Guasti: voci unite senza doppioni nella riga, rinomina applicata');
  ok(F[1][C.GUASTO - 1] === 'MAG. ASTUCCI' && F[2][C.GUASTO - 1] === 'H7, Sblocco Homing' && F[4][C.GUASTO - 1] === 'CAMERA OMRON' && F[3][C.GUASTO - 1] === 'X', 'storico Fermi aggiornato (le altre voci restano)');
  ok(VG.length === 4 && VG.map(r => r[0]).join('|') === 'Voce (come nel foglio Guasti)|MAG. ASTUCCI|H7, Sblocco Homing|CAMERA OMRON' && !VG[3][4], 'Voci guasto: righe doppie tolte, rinomina con categoria tenuta');
  ok(F[5][C.GUASTO - 1] === 'MAG. ASTUCCI', 'storico: stessa voce scritta in minuscolo/con spazi → come nel foglio Voci guasto');
  ok(sheets.Soluzioni.rows[1][5] === 'MAG. ASTUCCI' || sheets.Soluzioni.rows[1][5] === 'MAG. ASTUCCI', 'Soluzioni: tipo guasto unito');
  ok(log.length === 3 && sheets['Unioni voci'].rows.length === 4 && sheets['Unioni voci'].rows.some(r => r[1] === 'MAGAZZINO ASTUCCI' && r[3] === 1), 'registro nel foglio "Unioni voci" con le righe cambiate');
  ok(ctx.unioniApplica_(ss) === null, 'rilanciato: niente da fare');
}

console.log('\nFoglio Fermi: righe copiate a mano e macchina corretta a mano');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C, ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  const riga = (id, eq, lin, mac) => { const r = new Array(30).fill(''); r[C.ID - 1] = id; r[C.EQ - 1] = eq; r[C.LINEA - 1] = lin; r[C.MACC - 1] = mac; return r; };
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.concat(ctx.EXTRA_H), riga('F1', 'ST3-04', 'ST3', 'ASTUCCIATRICE'), riga('F1', 'ST3-02', 'ST1', 'ASTUCCIATRICE'), riga('F1', 'ST3-04', 'ST3', 'ASTUCCIATRICE'), riga('F2', 'ST3-02', 'ST3', 'RIEMPITRICE')]);
  const F = sheets.Fermi.rows;
  ok(ctx.idDoppi_(sheets.Fermi) === 2 && F[1][C.ID - 1] === 'F1' && F[2][C.ID - 1] === 'F1-2' && F[3][C.ID - 1] === 'F1-3' && F[4][C.ID - 1] === 'F2', 'ID doppi (righe copiate): la copia riceve F1-2, F1-3; la prima resta F1');
  ok(F[1][C.GRUPPO - 1] === 'F1' && F[2][C.GRUPPO - 1] === 'F1' && F[3][C.GRUPPO - 1] === 'F1' && !F[4][C.GRUPPO - 1], '"Fermo n." lega le righe dello stesso fermo');
  ok(ctx.idDoppi_(sheets.Fermi) === 0, 'rilanciato: niente da fare');
  ctx.eqCoerente_(ss, sheets.Fermi, 3, C.MACC);
  ok(F[2][C.EQ - 1] === 'ST3-04' && F[2][C.LINEA - 1] === 'ST3', 'cambiata la Macchina (nome unico): codice e linea si sistemano');
  F[4][C.EQ - 1] = 'ST1MAR-06'; ctx.eqCoerente_(ss, sheets.Fermi, 5, C.EQ);
  ok(F[4][C.MACC - 1] === 'BILANCIA' && F[4][C.LINEA - 1] === 'ST1 MAR', 'cambiato il Codice: macchina e linea dal foglio Equipment');
  F[4][C.LINEA - 1] = 'ST9'; ctx.eqCoerente_(ss, sheets.Fermi, 5, C.LINEA);
  ok(F[4][C.EQ - 1] === 'ST1MAR-06' && F[4][C.LINEA - 1] === 'ST1 MAR', 'linea sbagliata: il nome unico della macchina rimette linea giusta');
}

console.log('\nSoluzioni: scarto con motivo');
{
  const { ctx, sheets } = nuovo(); const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  sheets.Soluzioni = new Sheet('Soluzioni', [ctx.SOL_H.slice(0, 13), ['S1', new ctx.DD(), 'ST3-04', 'ST3', 'ASTUCCIATRICE', 'INCASTRO', 'flaconi incastrati', 'regolato', '', '', 'ROSSI Mario', 'DA APPROVARE', '']]);
  ctx.solColonne_(sheets.Soluzioni);
  ok(sheets.Soluzioni.rows[0][13] === 'Motivo (se scartata)' && sheets.Soluzioni.rows[0][12] === 'Approvata da', 'colonna "Motivo (se scartata)" aggiunta dopo "Approvata da"');
  let err = ''; try { ctx.approvaSoluzione_({ id: 'S1', esito: 'NO', motivo: '', manutentore: 'VERDI Anna' }); } catch (e) { err = e.message; }
  ok(/motivo/.test(err) && sheets.Soluzioni.rows[1][11] === 'DA APPROVARE', 'scarto senza motivo rifiutato');
  ctx.approvaSoluzione_({ id: 'S1', esito: 'NO', motivo: 'Troppo vaga: manca dove e come', manutentore: 'VERDI Anna' });
  ok(sheets.Soluzioni.rows[1][11] === 'SCARTATA' && sheets.Soluzioni.rows[1][13] === 'Troppo vaga: manca dove e come', 'scartata: stato e motivo scritti');
  const a = ctx.archivio_(ss);
  ok(!a.soluzioni.length && a.solScartate.length === 1 && a.solScartate[0].motivo === 'Troppo vaga: manca dove e come' && a.solScartate[0].autore === 'ROSSI Mario' && a.solScartate[0].da === 'VERDI Anna', 'l\'app riceve la scartata con motivo (non tra le soluzioni)');
}

console.log('\nFoglio Guasti: voci per macchina');
{
  const { ctx, sheets } = nuovo(); const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  sheets.Guasti = new Sheet('Guasti', [['Codice (QR) / Linea / TUTTE', 'Macchina (promemoria)', 'Tipi di guasto'], ['TUTTE', 'vale per tutte', 'SENSORE', ' ', 'INCASTRO'], ['st3-04', 'ST3 · ASTUCCIATRICE', 'MAG. ASTUCCI', '', 'CHIUSURE AST  '], ['ST3-05', 'ST3 · BILANCIA']]);
  const g = ctx.guasti_(ss);
  ok(g['ST3-04'].join('|') === 'MAG. ASTUCCI|CHIUSURE AST' && g.TUTTE.join('|') === 'SENSORE|INCASTRO', 'righe del foglio Guasti: codice in maiuscolo, colonna B solo promemoria, celle vuote e spazi ignorati');
  ok(Array.isArray(g['ST3-05']) && g['ST3-05'].length === 0, 'macchina senza voci: elenco vuoto (l\'app mostrerà linea + TUTTE + ALTRO)');
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.slice()]); sheets.Liste.rows = [['Manutentori'], ['ROSSI Mario']];
  ok(JSON.stringify(ctx.init_().guasti) === JSON.stringify(g), 'l\'app riceve esattamente le voci del foglio Guasti');
}

console.log('\nPiù foto in chiusura');
{
  const { ctx, sheets } = nuovo(); const C = ctx.C;
  const r = new Array(31).fill(''); r[C.ID - 1] = 'FX1'; r[C.EQ - 1] = 'ST3-04';
  sheets.Fermi = new Sheet('Fermi', [ctx.HEAD.concat(ctx.EXTRA_H.slice(0, 6)), r]);
  const files = {}; ctx.PropertiesService.getScriptProperties().setProperty('FOLDER_ID', 'F');
  ctx.DriveApp = { getFolderById: () => ({ getFilesByName: n => { let done = !files[n]; return { hasNext: () => !done, next: () => { done = true; return { getUrl: () => files[n] }; } }; },
    createFile: b => { const n = b.getName ? b.getName() : 'x'; files[n] = 'https://drive/' + n; return { getUrl: () => files[n] }; } }) };
  const nb = ctx.Utilities.newBlob; ctx.Utilities.newBlob = (d, t, n) => Object.assign(nb(d), { getName: () => n });
  const a = ctx.fotoFermo_({ fermo: 'FX1', id: 'P000001', n: 2, foto: 'AAAA', ts: Date.now() });
  ctx.fotoFermo_({ fermo: 'FX1', id: 'P000001', n: 2, foto: 'AAAA', ts: Date.now() }); // invio ripetuto
  ctx.fotoFermo_({ fermo: 'FX1', id: 'P000002', n: 3, foto: 'BBBB', ts: Date.now() });
  const cell = String(sheets.Fermi.rows[1][C.FOTO_X - 1]), righe = cell.split('\n');
  ok(sheets.Fermi.rows[0][C.FOTO_X - 1] === 'Altre foto chiusura' && righe.length === 2 && /_chiusura_2_/.test(righe[0]) && /_chiusura_3_/.test(righe[1]), 'foto 2 e 3: un link per riga in "Altre foto chiusura", nome con il numero della foto');
  ok(Object.keys(files).length === 2, 'invio ripetuto della stessa foto: non la duplica');
  let err = ''; try { ctx.fotoFermo_({ fermo: 'NON-ESISTE', id: 'P3', n: 2, foto: 'A' }); } catch (e) { err = e.message; }
  ok(/non trovato/.test(err), 'foto di un fermo che non c\'è: rifiutata');
}
fine('Script');
