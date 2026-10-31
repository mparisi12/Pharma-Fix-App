// Test dell'app nel browser (Playwright) con un finto Google Sheet e dati inventati.
// Lancia: npm install && npx playwright install chromium && node app.test.js
const http = require('http'), fs = require('fs'), path = require('path');
const { chromium } = require('playwright');
const { ok, fine } = require('./fake-sheets');
const ROOT = path.join(__dirname, '..'), PORT = 8779;
const MIME = { '.html': 'text/html', '.js': 'text/javascript', '.json': 'application/json', '.webmanifest': 'application/manifest+json', '.png': 'image/png' };
const srv = http.createServer((q, r) => { const f = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]).replace(/\/$/, '/index.html'));
  if (!f.startsWith(ROOT) || !fs.existsSync(f)) { r.writeHead(404); return r.end(); } r.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(r); });

const now = Date.now(), g0 = new Date(); g0.setHours(0, 0, 0, 0);
// Macchine finte: nell'app vera arrivano solo dal foglio Equipment (non più da defaults.js)
const D = { equipment: [['ST3-01', 'Filling', 'RIORDINATORE'], ['ST3-02', 'Filling', 'RIEMPITRICE'], ['ST3-03', 'Packaging', 'ETICHETTATRICE'],
    ['ST3-04', 'Packaging', 'ASTUCCIATRICE'], ['ST3-05', 'Packaging', 'BILANCIA'], ['ST3-06', 'Packaging', 'CARTONATRICE']].map(([code, area, macchina]) => ({ code, linea: 'ST3', area, macchina })),
  liste: { Manutentori: ['ROSSI Mario', 'BIANCHI Luca'], Coordinatori: ['VERDI Anna'] }, guasti: {}, soluzioni: [], anomalie: [], turni: {}, programma: {}, continuo: ['ST3'],
  fermi: [{ id: 'F0', code: 'ST3-02', linea: 'ST3', macchina: 'RIEMPITRICE', start: now - 20 * 60000, stato: 'APERTO', lotto: '601090', man: 'BIANCHI Luca', apertoDa: 'BIANCHI Luca' }],
  compiti: [], documenti: [], vestizioni: [{ man: 'ROSSI Mario', scad: g0.getTime() + 40 * 864e5 }],
  convalide: { list: [{ man: 'ROSSI Mario', cat: 'G', desc: 'Intervento Meccanico di regolazione', scad: g0.getTime() - 3 * 864e5 }, { man: 'BIANCHI Luca', cat: 'H', desc: 'Intervento Meccanico di sostituzione componente', scad: g0.getTime() + 200 * 864e5 }], ignote: [] },
  sterileNo: ['VERDI Anna'] };
let ver = 1, full = 0, same = 0; const posts = [];

(async () => {
  await new Promise(r => srv.listen(PORT, r));
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 390, height: 844 }, serviceWorkers: 'block' }); const p = await ctx.newPage();
  const errori = []; p.on('pageerror', e => errori.push(e.message));
  await ctx.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ contentType: 'text/css', body: '' }));
  await ctx.route(/script\.google\.com/, r => {
    const u = r.request().url();
    if (r.request().method() === 'POST') {
      const o = JSON.parse(r.request().postData()); posts.push(o); ver++;
      if (o.action === 'open') D.fermi.unshift({ id: o.id, code: o.code, linea: 'ST3', macchina: 'X', start: o.ts, stato: 'APERTO', lotto: o.lotto, man: o.manutentore, apertoDa: o.manutentore });
      if (o.action === 'compito') D.compiti.push({ id: o.id, ts: o.ts, da: o.manutentore, a: o.a, code: o.code || '', linea: '', titolo: o.titolo, descr: o.descr, pr: o.pr, entro: o.entroTs, stato: 'DA FARE', storia: '' });
      if (o.action === 'compitoStato') Object.assign(D.compiti.find(k => k.id === o.id), { stato: o.stato, presoDa: o.manutentore, preso: o.ts });
      if (o.action === 'documento') D.documenti.push({ id: o.id, ts: o.ts, dove: o.dove, tipo: o.tipo, titolo: o.titolo, descr: o.descr, link: o.link });
      return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' });
    }
    if (/action=login/.test(u)) return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' });
    const v = (u.match(/[?&]v=([^&]+)/) || [])[1];
    if (v === 'V' + ver) { same++; return r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, same: true, ver: 'V' + ver }) }); }
    full++; r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, ...D, ver: 'V' + ver }) });
  });
  const attesa = ms => p.waitForTimeout(ms);
  async function entra(nome, cognome) {
    await p.goto(`http://localhost:${PORT}/fermi/`); await attesa(500);
    if (await p.isVisible('#lpin')) await p.fill('#lpin', 'prova');
    await p.fill('#lnome', nome); await p.fill('#lcogn', cognome); await p.click('#lgo'); await attesa(1800);
  }
  async function esci() { await p.evaluate(() => closeSh()); await p.click('#manbtn'); await p.click('#logout'); await attesa(300); }

  console.log('Accesso e dati');
  await entra('Anna', 'Verdi');
  ok(/VERDI Anna/.test(await p.innerText('#manbtn')), 'accesso del coordinatore (solo nella colonna Coordinatori)');
  ok(/RIEMPITRICE/.test(await p.innerText('#view')), 'la home mostra il fermo aperto');
  const f0 = full; for (let i = 0; i < 3; i++) { await p.evaluate(() => refresh()); await attesa(150); }
  ok(full === f0 && same >= 3, 'aggiornamenti senza modifiche: nessun riscaricamento completo');

  console.log('\nApri fermo');
  await p.click('#manual'); await attesa(300); await p.click('#pl [data-v="ST3"]'); await attesa(200); await p.click('#pm [data-v="ST3-04"]'); await attesa(500);
  ok(await p.isVisible('#lotto'), 'popup centrale con il lotto');
  await p.fill('#lotto', '601095'); await p.click('#sheet #go'); await attesa(2500);
  ok(posts.some(o => o.action === 'open' && o.code === 'ST3-04' && o.lotto === '601095'), 'fermo inviato al foglio');
  ok(full > f0, 'dopo l\'invio i dati nuovi si riscaricano');
  // lotto corretto a mano: al fermo successivo sulla stessa linea viene riproposto, anche se c'è il programma
  const lt = await p.evaluate(() => { const eq = eqBy('ST3-05') || { code: 'ST3-05', linea: 'ST3', macchina: 'BILANCIA', area: 'Packaging' }; const old = S.programma;
    S.programma = { ST3: { [dKey(giornoProd())]: '601090' } }; closeSh(); openSheet(eq);
    const r = { v: document.querySelector('#lotto').value, sel: (document.querySelector('#lch [aria-pressed="true"]') || {}).textContent || '', hint: document.querySelector('#sheet .hint').textContent };
    closeSh(); S.programma = old; return r; });
  ok(lt.v === '601095' && /Ultimo usato · 601095/.test(lt.sel) && /Ultimo lotto usato su ST3/.test(lt.hint), 'apri fermo: riproposto l\'ultimo lotto scritto a mano (601095), non il primo del programma (601090)');

  console.log('\nDocumenti');
  await p.evaluate(() => go({ name: 'doc', code: 'ST3-04', tab: 'sch' })); await attesa(300);
  await p.click('#adddoc'); await attesa(300); await p.click('#dgo');
  ok(/titolo/i.test(await p.innerText('#derr')), 'senza titolo: chiede il titolo');
  await p.fill('#dti', 'Scheda formato prova'); await p.fill('#dli', 'https://drive.google.com/file/d/abc/view'); await p.click('#dgo'); await attesa(2000);
  ok(posts.some(o => o.action === 'documento' && o.tipo === 'SCHEDA FORMATO' && o.dove === 'ST3-04'), 'documento inviato');
  ok(/Scheda formato prova/.test(await p.innerText('#view')), 'documento in elenco');

  console.log('\nCompiti');
  await p.evaluate(() => go({ name: 'home' })); await p.click('#manbtn'); await p.click('#rep'); await attesa(300); await p.click('#kass'); await attesa(300);
  ok(!!(await p.$('#ka [data-v="VERDI Anna"]')), 'si possono assegnare compiti anche ai coordinatori');
  await p.click('#ka [data-v="ROSSI Mario"]'); await p.fill('#kt', 'Controllo cinghia prova'); await p.click('#kok'); await attesa(2000);
  ok(posts.some(o => o.action === 'compito' && o.a === 'ROSSI Mario'), 'compito assegnato');
  ok(!(await p.isVisible('#kvedi')), 'chi assegna non vede il popup');
  await esci(); await entra('Mario', 'Rossi');
  ok(await p.isVisible('#kvedi') && /nuovo compito/i.test(await p.innerText('#sheet')), 'il tecnico vede il popup "nuovo compito"');
  ok(!!(await p.$('#manbtn .kpill')), 'pallino sul nome (compito + vestizione)');
  await p.click('#kvedi'); await attesa(300); await p.click('#ktake'); await attesa(2000);
  ok(posts.some(o => o.action === 'compitoStato' && o.stato === 'IN CORSO' && o.manutentore === 'ROSSI Mario'), 'preso in carico');
  await p.reload(); await attesa(1800);
  ok(!(await p.isVisible('#kvedi')), 'compito già visto: il popup non torna');
  await p.click('#manbtn'); await attesa(200);
  ok(!!(await p.$('#sheet .vbox')), 'avviso vestizione nel menu');
  ok(/Cat\. G SCADUTA/.test(await p.innerText('#sheet')), 'avviso convalida GMP scaduta nel menu');
  await p.click('#mabil'); await attesa(200);
  ok((await p.$$('#sheet .abil')).length === 3 && /Mancante/.test(await p.innerText('#sheet')), 'le mie abilitazioni: vestizione, Cat. G scaduta, Cat. H mancante');
  await esci(); await entra('Anna', 'Verdi'); await p.click('#manbtn'); await attesa(200);
  ok(/Scadenze del reparto/.test(await p.innerText('#vrep')), 'coordinatore: voce Scadenze del reparto');
  await p.click('#vrep'); await attesa(200); await p.click('#sfil [data-v="scad"]'); await attesa(200);
  ok((await p.$$eval('#sheet .sgrid .n', x => x.map(e => e.textContent))).join() === 'BIANCHI Luca,ROSSI Mario', 'filtro Scadute: scaduti + chi non ha la convalida (BIANCHI senza Cat. G); esonerati fuori');
  ok(/MANCA/.test(await p.innerText('#sheet .sgrid')), 'convalida mancante segnata MANCA');
  // coordinatore: scartare una soluzione chiede il motivo
  await p.evaluate(() => { closeSh(); S.soluzioni.push({ id: 'SP1', ts: Date.now(), dove: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', guasto: 'INCASTRO', problema: 'I flaconi si incastravano', soluzione: 'Regolato sponde e vomere', autore: 'ROSSI Mario', stato: 'DA APPROVARE' }); go({ name: 'sol', code: 'ST3-04' }); }); await attesa(300);
  await p.click('#daapp [data-app="NO"]'); await attesa(150);
  ok(await p.isVisible('#daapp .scmot'), 'Scarta: compare la scelta del motivo');
  await p.click('#daapp [data-app="NOK"]'); await attesa(150);
  ok(!posts.some(o => o.action === 'approvaSol' && o.id === 'SP1'), 'senza motivo non scarta');
  await p.click('#daapp .scmot [data-m="Troppo vaga: manca dove e come"]'); await attesa(100);
  if (process.env.SHOT_SOL) { await p.evaluate(() => document.querySelector('#daapp').scrollIntoView({ block: 'start' })); await p.screenshot({ path: process.env.SHOT_SOL.replace('.png', '2.png') }); }
  await p.click('#daapp [data-app="NOK"]'); await attesa(800);
  ok((posts.filter(o => o.action === 'approvaSol' && o.id === 'SP1').pop() || {}).motivo === 'Troppo vaga: manca dove e come', 'scartata con il motivo (arriva a chi l\'ha proposta)');
  await p.evaluate(() => { S.soluzioni = S.soluzioni.filter(x => x.id !== 'SP1'); go('home'); }); await attesa(200);
  // tolleranza 30 giorni: la Cat. G di ROSSI scaduta da 3 giorni è arancio, ancora valida
  await p.evaluate(() => { S.avviso = Object.assign({}, S.avviso, { tolV: 30, tolC: 30 }); shScadRep('tutti'); }); await attesa(200);
  ok(await p.$$eval('#sheet .sgrid .tag.o', x => x.length) === 1 && /In tolleranza · 1/.test(await p.innerText('#sfil')) && /arancio/.test(await p.innerText('#sheet')), 'tolleranza: cella arancio con l\'ultimo giorno valido, filtro "In tolleranza"');
  if (process.env.SHOT_TOL) await p.screenshot({ path: process.env.SHOT_TOL });
  ok(!/SCAD\./.test(await p.innerText('#sheet .sgrid')), 'in tolleranza non è segnata SCAD.');
  const tb = await p.evaluate(() => boxConv('ROSSI Mario'));
  ok(/vbox tol/.test(tb) && /valida ancora fino al/.test(tb) && /ancora 27 giorni/.test(tb), 'tecnico: box arancio "scaduta · in tolleranza", valida fino al… (27 giorni rimasti)');
  if (process.env.SHOT_TOL) { await p.evaluate(() => { closeSh(); shAbil('ROSSI Mario', () => {}); }); await attesa(200); await p.screenshot({ path: process.env.SHOT_TOL.replace('.png', '2.png') }); }
  await p.evaluate(() => { S.avviso = Object.assign({}, S.avviso, { tolV: 0, tolC: 0 }); closeSh(); });
  // corso MD_20: stessa logica delle convalide, testi da corso
  const md = await p.evaluate(() => { const old = S.convalide; S.convalide = { list: old.list.concat([{ man: 'ROSSI Mario', cat: 'MD_20', desc: 'Corso MD_20', scad: Date.now() + 10 * 864e5 }, { man: 'ROSSI Mario', cat: 'MF', desc: 'Partecipazione MF', scad: Date.now() - 5 * 864e5 }]), ignote: [], cats: ['G', 'H', 'MD_20', 'MF'] };
    const b = boxConv('ROSSI Mario'), mb = boxConv('BIANCHI Luca'); shScadRep('tutti'); const t = document.querySelector('#sheet').innerText, hs = [...document.querySelectorAll('#sheet .sgrid .h')].map(e => e.innerText.replace(/\s+/g, ' '));
    closeSh(); shAbil('ROSSI Mario', () => {}); const ab = [...document.querySelectorAll('#sheet .abil .k')].map(e => e.textContent); closeSh(); S.convalide = old; return { b, mb, t, hs, ab }; });
  if (process.env.SHOT_MF) { await p.evaluate(() => { const old = S.convalide; S.convalide = { list: old.list.concat([{ man: 'ROSSI Mario', cat: 'MD_20', desc: 'Corso MD_20', scad: Date.now() + 10 * 864e5 }, { man: 'ROSSI Mario', cat: 'MF', desc: 'Partecipazione MF', scad: Date.now() - 5 * 864e5 }]), ignote: [], cats: ['G', 'H', 'MD_20', 'MF'] }; shScadRep('tutti'); window.__oldConv = old; }); await attesa(200); await p.screenshot({ path: process.env.SHOT_MF }); await p.evaluate(() => { closeSh(); S.convalide = window.__oldConv; }); }
  ok(/Partecipazione MF SCADUTA/.test(md.b) && /Partecipazione MF: data non registrata/.test(md.mb), 'partecipazione MF: avviso al femminile (scaduta), chi non ce l\'ha "data non registrata"');
  ok(md.hs.join('|') === 'CORSO MD_20|VESTIZ.|PART. MF|CAT. G REGOL.|CAT. H SOSTIT.' && md.ab.join('') === 'MDVMFGH', 'ordine come nel foglio Convalide MF: MD_20, vestizione, partecipazione MF, Cat. G, Cat. H: ' + md.hs.join('|') + ' / ' + md.ab.join(''));
  ok(/Corso MD_20 in scadenza/.test(md.b) && /Prenota il corso/.test(md.b) && /Corso MD_20: data non registrata/.test(md.mb), 'corso MD_20: avviso in scadenza; chi non ce l\'ha: "data non registrata"');
  ok(/CORSO\s*MD_20/.test(md.t), 'Scadenze del reparto: colonna CORSO MD_20');

  console.log('\nAccesso che resta (iPhone)');
  await esci(); await entra('Mario', 'Rossi');
  await p.reload(); await attesa(1500);
  ok(/ROSSI Mario/.test(await p.innerText('#manbtn')) && await p.isHidden('#login'), 'riaprendo l\'app: resta dentro, niente codice né nome');
  await p.evaluate(() => { const s = JSON.parse(localStorage.getItem('pf_sess')); s.exp = Date.now() + 3600e3; localStorage.setItem('pf_sess', JSON.stringify(s)); sessionStorage.clear(); });
  await p.reload(); await attesa(1500);
  ok((await p.evaluate(() => JSON.parse(localStorage.getItem('pf_sess')).exp - Date.now())) > 25 * 864e5, 'ogni apertura rinnova l\'accesso (30 giorni), non scade a fine turno');
  await p.evaluate(() => { localStorage.clear(); sessionStorage.clear(); });
  await p.reload(); await attesa(2000);
  ok(await p.isHidden('#login') && /ROSSI Mario/.test(await p.innerText('#manbtn')) && await p.evaluate(() => !!localStorage.getItem('pf_pin')), 'memoria del browser svuotata: codice e accesso ripresi dalla copia nell\'archivio');
  const ios = await b.newContext({ viewport: { width: 390, height: 844 }, serviceWorkers: 'block', userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1' });
  await ios.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ contentType: 'text/css', body: '' })); await ios.route(/script\.google\.com/, r => r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, ...D }) }));
  const q = await ios.newPage(); await q.goto(`http://localhost:${PORT}/fermi/?eq=ST3-04`); await q.waitForTimeout(800);
  ok(await q.isVisible('#iosban'), 'iPhone in Safari (es. QR con la fotocamera): avviso "apri l\'app dalla schermata Home"');
  await q.addInitScript(() => Object.defineProperty(navigator, 'standalone', { get: () => true })); await q.reload(); await q.waitForTimeout(800);
  ok(!(await q.isVisible('#iosban')), 'iPhone dall\'icona sulla schermata Home: nessun avviso');
  await ios.close();

  console.log('\nRete che va e viene');
  let giu = 0; await ctx.unroute(/script\.google\.com/);
  await ctx.route(/script\.google\.com/, r => { if (giu > 0) { giu--; return r.abort('internetdisconnected'); } r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, same: true, ver: 'V' + ver }) }); });
  await p.evaluate(() => closeSh());
  giu = 1; await p.evaluate(() => refresh()); await attesa(300);
  ok(!/Non connesso/.test(await p.innerText('#sync')), 'un solo tentativo fallito: niente pallino rosso');
  await attesa(5500);
  ok(/Sincronizzato/.test(await p.innerText('#sync')), 'riprova da solo dopo 5 secondi');
  giu = 2; await p.evaluate(() => refresh()); await attesa(5800);
  const tx = await p.innerText('#sync');
  ok(/Non connesso/.test(tx) && /dati delle \d\d:\d\d/.test(tx), 'due tentativi falliti: rosso con l\'ora degli ultimi dati');
  await p.evaluate(() => window.dispatchEvent(new Event('online'))); await attesa(800);
  ok(/Sincronizzato/.test(await p.innerText('#sync')), 'rete tornata: si riconnette subito');

  console.log('\nTocco sulla scritta in alto: aggiorna subito');
  await ctx.unroute(/script\.google\.com/); let get = 0;
  D.fermi.unshift({ id: 'F-ALTRO', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 60000, stato: 'APERTO', lotto: '123456', man: 'BIANCHI Luca', apertoDa: 'BIANCHI Luca' });
  await ctx.route(/script\.google\.com/, r => { get++; r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, ...D, ver: 'NUOVA' }) }); });
  await p.evaluate(() => go('home')); await attesa(200);
  await p.click('#sync'); await attesa(800);
  ok(get >= 1 && /fermi? apert/.test(await p.innerText('body')), 'tocco sul pallino: dati richiesti subito e conteggio dei fermi aperti');
  ok(/BIANCHI/.test(await p.innerText('#view')), 'il fermo aperto dal collega compare subito nella home');

  console.log('\nChiusura fermo: cosa si è guastato → cosa hai fatto → perché');
  await ctx.unroute(/script\.google\.com/);
  const storia = k => ({ id: 'H' + k, code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - (k + 2) * 3600e3, end: Date.now() - (k + 1) * 3600e3, stato: 'CHIUSO', guasto: 'CAM. OMRON', int: k < 2 ? 'REGOLAZIONE' : 'RIAVVIO', cat: 'AUTOMATION', man: 'ROSSI Mario', lotto: '601095' });
  D.fermi = D.fermi.concat([0, 1, 2].map(storia), [{ id: 'FC', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 60000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' }]);
  D.guasti = { 'ST3-04': ['CAM. OMRON', 'MAG. ISTRUZIONI'], TUTTE: ['SENSORE', 'ALTRO'] }; D.vociCat = { 'CAM. OMRON': 'AUTOMATION', 'MAG. ISTRUZIONI': 'MECCANICO', SENSORE: 'ELETTRICO' };
  await ctx.route(/script\.google\.com/, r => { if (r.request().method() === 'POST') { posts.push(JSON.parse(r.request().postData())); return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' }); }
    r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, ...D, ver: 'CL' }) }); });
  await p.evaluate(() => refresh()); await attesa(500);
  await p.evaluate(() => closeSheet(S.fermi.find(f => f.id === 'FC'))); await attesa(300);
  ok(await p.isVisible('#st2.locked') && !(await p.isVisible('#st1 .autotag')), 'all\'inizio solo "cosa si è guastato" è attivo, niente categoria da scegliere');
  await p.click('#gua [data-v="CAM. OMRON"]'); await attesa(200);
  ok(/AUTOMATION/.test(await p.innerText('#catbox')) && !(await p.isVisible('#catbox .chips')), 'CAM. OMRON: categoria AUTOMATION automatica');
  ok(await p.getAttribute('#int [data-v="REGOLAZIONE"]', 'aria-pressed') === 'true' && /proposto/.test(await p.innerText('#iprop')), 'intervento proposto: il più usato per quel guasto su quella macchina (RIAVVIO conta come RIPRISTINO)');
  await p.click('#int [data-v="SOSTITUZIONE"]'); await attesa(150);
  ok(await p.isVisible('#pezzo') && !(await p.innerText('#iprop')), 'SOSTITUZIONE: compare subito il codice pezzo');
  await p.click('#gua [data-v="ALTRO"]'); await attesa(150);
  ok(await p.isVisible('#catbox .chips') && await p.getAttribute('#int [data-v="SOSTITUZIONE"]', 'aria-pressed') === 'true', 'ALTRO: la categoria la sceglie il tecnico; la sua scelta di intervento resta');
  ok(await p.$$eval('#sheet input:not([type]), #sheet textarea', l => l.length > 0 && l.every(e => e.getAttribute('autocomplete') === 'off')), 'campi di testo senza suggerimenti di indirizzi/contatti (autocomplete off)');
  ok(await p.isVisible('#phgrp .camtile') && await p.isVisible('#phgrp .galbtn') && (await p.getAttribute('#gal', 'capture')) === null, 'foto chiusura: si può scattare o scegliere dalla galleria (input senza capture)');
  const big = await p.evaluate(() => { const c = document.createElement('canvas'); c.width = 4000; c.height = 3000; const x = c.getContext('2d'); x.fillStyle = '#3a6'; x.fillRect(0, 0, 4000, 3000); return c.toDataURL('image/jpeg', .9).split(',')[1]; });
  const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('#phgrp .galbtn')]);
  await fc.setFiles({ name: 'galleria.jpg', mimeType: 'image/jpeg', buffer: Buffer.from(big, 'base64') }); await attesa(400);
  ok(await p.isVisible('#phgrp canvas') && await p.isVisible('#phgrp .lgal'), 'foto dalla galleria caricata: si può cerchiare e cambiare');
  await p.click('#gua [data-v="CAM. OMRON"]'); await p.click('#int [data-v="REGOLAZIONE"]');
  await p.click('#cau [data-v="USURA"]'); await p.click('#cau [data-v="USURA"]'); await p.click('#cau [data-v="MATERIALE FUORI SPEC."]');
  await p.fill('#altro', 'regolata messa a fuoco telecamera'); await p.click('#go'); await attesa(300);
  ok(/è giusto\?/.test(await p.innerText('#err')) && !posts.some(o => o.action === 'close' && o.id === 'FC'), 'fermo più corto di 3 minuti: chiede conferma e non chiude');
  await p.click('#go'); await attesa(800);
  const cl = posts.filter(o => o.action === 'close' && o.id === 'FC').pop() || {};
  const dim = await p.evaluate(f => new Promise(r => { const im = new Image(); im.onload = () => r([im.naturalWidth, im.naturalHeight]); im.src = 'data:image/jpeg;base64,' + f; }), cl.foto || '');
  ok(dim[0] === 1024 && dim[1] === 768, 'foto dalla galleria 4000×3000 ridimensionata a 1024×768 prima dell\'invio');
  await p.evaluate(() => { closeSh(); go({ name: 'anoform', code: 'ST3-04' }); }); await attesa(300);
  ok(await p.isVisible('#aph .camtile') && await p.isVisible('#aph .galbtn'), 'anomalie: foto anche dalla galleria');
  await p.evaluate(() => go({ name: 'home' })); await attesa(200);
  ok(cl.categoria === 'AUTOMATION' && cl.intervento === 'REGOLAZIONE' && cl.guasto === 'CAM. OMRON' && cl.causa === 'MATERIALE FUORI SPEC.', 'inviato: categoria automatica, intervento, guasto e causa (secondo tocco toglie la scelta)');

  D.liste['Tipo intervento'] = ['REGOLAZIONE', 'SOSTITUZIONE', 'RIPRISTINO', 'PULIZIA']; D.intDesc = { RIPRISTINO: 'reset allarme, homing' }; D.liste.Causa = ['USURA', 'MANUTENZIONE NON ESEGUITA'];
  D.fermi.push({ id: 'FC2', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 300000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' });
  await p.evaluate(() => { S.ver = 'x'; return refresh(); }); await attesa(400);
  await p.evaluate(() => { closeSh(); closeSheet(S.fermi.find(f => f.id === 'FC2')); }); await attesa(300);
  ok(/reset allarme, homing/.test(await p.innerText('#int')) && /MANUTENZIONE NON ESEGUITA/.test(await p.innerText('#cau')), 'interventi, spiegazioni e cause presi dal foglio Liste');
  await p.evaluate(() => closeSh());

  console.log('\nRaccolta uniforme: definizioni, voce mancante con ALTRO, Da rivedere, correzioni al tecnico');
  D.vociDesc = { 'CAM. OMRON': 'telecamera Omron che scarta o non legge' };
  D.fermi.push({ id: 'FC3', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 600000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' },
    { id: 'RV1', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 5 * 3600e3, end: Date.now() - 4.5 * 3600e3, stato: 'CHIUSO', guasto: 'CAM. OMRON', int: 'REGOLAZIONE', altro: 'sostituito cavo telecamera', man: 'ROSSI Mario', lotto: '601095', orari: 'TEMPO REALE' },
    { id: 'CR1', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 8 * 3600e3, end: Date.now() - 7.5 * 3600e3, stato: 'CHIUSO', guasto: 'MAG. ISTRUZIONI', int: 'REGOLAZIONE', altro: 'x', man: 'ROSSI Mario', lotto: '601095', rivisto: 'VERDI Anna', corr: { da: 'VERDI Anna', ts: 123, txt: 'Tipo guasto: INCASTRO → MAG. ISTRUZIONI' } });
  await p.evaluate(() => { S.ver = 'y'; return refresh(); }); await attesa(400);
  await p.evaluate(() => { closeSh(); closeSheet(S.fermi.find(f => f.id === 'FC3')); }); await attesa(300);
  await p.click('#gua [data-v="CAM. OMRON"]'); await attesa(150);
  ok(/telecamera Omron che scarta/.test(await p.innerText('#catbox')), 'sotto la voce scelta compare "quando si usa" (foglio Voci guasto)');
  await p.click('#gua [data-v="ALTRO"]'); await p.click('#catbox [data-v="MECCANICO"]').catch(() => {}); await p.click('#int [data-v="REGOLAZIONE"]');
  await p.fill('#altro', 'lama tagliata'); await p.click('#go'); await attesa(200);
  ok(await p.isVisible('#gprop') && /che voce manca/i.test(await p.innerText('#err')), 'con ALTRO: obbligatorio scrivere la voce che manca');
  await p.fill('#gprop', 'lama sbucciatrice'); await p.click('#go'); await attesa(200); await p.click('#go').catch(() => {}); await attesa(600);
  ok((posts.filter(o => o.action === 'close' && o.id === 'FC3').pop() || {}).proposta === 'lama sbucciatrice', 'la voce proposta viene inviata con il fermo');
  // Proponi come soluzione: campi guidati e controlli prima dell'invio
  D.fermi.push({ id: 'FC4', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 900000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' });
  await p.evaluate(() => { S.ver = 'x'; return refresh(); }); await attesa(400);
  await p.evaluate(() => { closeSh(); closeSheet(S.fermi.find(f => f.id === 'FC4')); }); await attesa(300);
  // categoria: proposta e modificabile; "A SCELTA" la sceglie sempre il tecnico
  await p.evaluate(() => { S.vociCat = Object.assign({}, S.vociCat, { SENSORE: 'A SCELTA' }); });
  await p.click('#gua [data-v="SENSORE"]'); await attesa(150);
  ok(await p.isVisible('#catbox .catch') && !(await p.$('#catbox .catch [aria-pressed="true"]')) && /dipende da cosa hai fatto/i.test(await p.innerText('#catbox')), 'voce A SCELTA (SENSORE): il tecnico sceglie la categoria, niente preselezionato');
  if (process.env.SHOT_CAT) { await p.evaluate(() => document.querySelector('#catbox').scrollIntoView({ block: 'center' })); await p.screenshot({ path: process.env.SHOT_CAT }); }
  await p.click('#catbox .catch [data-v="ELETTRICO"]'); await attesa(100);
  ok(await p.evaluate(() => document.querySelector('#catbox .catch [data-v="ELETTRICO"]').getAttribute('aria-pressed')) === 'true' && /sensibilità, taratura/.test(await p.innerText('#catbox')), 'SENSORE: il tecnico sceglie ELETTRICO (con l\'aiuto "sensibilità, taratura")');
  await p.click('#int [data-v="REGOLAZIONE"]'); await attesa(100);
  await p.fill('#altro', 'regolata lama'); await p.click('#psol'); await attesa(100);
  ok(await p.isVisible('#pcau') && await p.isVisible('#pcome') && (await p.inputValue('#pcome')) === 'regolata lama' && /Proponi come soluzione per i colleghi/.test(await p.innerText('#st4')), 'soluzione: domanda chiara, tre campi guidati, "come" parte dalla descrizione');
  const prova = async (sin, cau, come, okk) => { await p.fill('#pprob', sin); await p.fill('#pcau', cau); await p.fill('#pcome', come); if ((await p.isChecked('#pok')) !== okk) await p.click('#pok'); await p.click('#go'); await attesa(200); return p.innerText('#err'); };
  ok(/Troppo poco/.test(await prova('astucci segnati in uscita', 'lama storta', 'regolata lama', true)), 'soluzione troppo corta: "troppo poco per insegnare"');
  ok(/non ha risolto/i.test(await prova('astucci segnati in uscita', 'lama storta', 'provato a regolare la lama e la guida più volte ma senza risultato, da rivedere', true)), 'racconta tentativi falliti: non è una soluzione');
  ok(/Conferma che ha risolto/.test(await prova('astucci segnati in uscita', 'lama storta', 'allentate le 2 viti lato OP, riallineata la lama alla guida, stretto e provato 20 astucci', false)), 'serve la spunta "Ha risolto"');
  await p.fill('#pcau', 'lama inserimento storta'); await p.click('#err').catch(() => {});
  if (process.env.SHOT_SOL) { await p.evaluate(() => { document.querySelector('#err').hidden = true; document.querySelector('#pswrap').scrollIntoView({ block: 'center' }); }); await p.screenshot({ path: process.env.SHOT_SOL }); }
  if (!(await p.isChecked('#pok'))) await p.click('#pok');
  await p.click('#go'); await attesa(300); await p.click('#go').catch(() => {}); await attesa(700);
  const so = posts.filter(o => o.action === 'soluzione').pop() || {};
  ok((posts.filter(o => o.action === 'close' && o.id === 'FC4').pop() || {}).categoria === 'ELETTRICO', 'inviata la categoria scelta dal tecnico');
  ok(so.problema === 'astucci segnati in uscita' && /^Causa: lama inserimento storta\n allentate|^Causa: lama inserimento storta\nallentate/.test(so.soluzione), 'inviata: problema = cosa si vedeva, soluzione = causa + come');
  // più foto in chiusura (fino a 3): la 2ª e la 3ª partono dopo la chiusura come operazioni a sé
  D.fermi.push({ id: 'FC5', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 900000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' });
  await p.evaluate(() => { S.ver = 'x5'; return refresh(); }); await attesa(400);
  await p.evaluate(() => { closeSh(); closeSheet(S.fermi.find(f => f.id === 'FC5')); }); await attesa(300);
  ok(!(await p.isVisible('#phx [data-add]')), 'prima della foto principale non si aggiungono altre foto');
  const img = await p.evaluate(() => { const c = document.createElement('canvas'); c.width = 1600; c.height = 1200; const x = c.getContext('2d'); x.fillStyle = '#c63'; x.fillRect(0, 0, 1600, 1200); return c.toDataURL('image/jpeg', .9).split(',')[1]; });
  const scegli = async sel => { const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click(sel)]); await fc.setFiles({ name: 'f.jpg', mimeType: 'image/jpeg', buffer: Buffer.from(img, 'base64') }); await attesa(400); };
  await scegli('#phgrp .galbtn');
  ok(/ancora 2/.test(await p.innerText('#phx')), 'dopo la prima foto: "puoi aggiungerne ancora 2"');
  await scegli('#phx [data-add="gal"]');
  if (process.env.SHOT_FOTO) { await p.evaluate(() => document.querySelector('#phx').scrollIntoView({ block: 'center' })); await p.screenshot({ path: process.env.SHOT_FOTO.replace('.png', 'b.png') }); }
  await scegli('#phx [data-add="gal"]');
  ok((await p.$$('#phx .phxi canvas')).length === 2 && !(await p.isVisible('#phx [data-add]')), '3 foto in tutto: poi i pulsanti spariscono');
  if (process.env.SHOT_FOTO) { await p.evaluate(() => document.querySelector('#phgrp').scrollIntoView({ block: 'start' })); await p.screenshot({ path: process.env.SHOT_FOTO }); }
  await p.click('#phx [data-rm="1"]'); await attesa(150);
  ok((await p.$$('#phx .phxi')).length === 1 && /ancora 1/.test(await p.innerText('#phx')), '"Togli": la foto 3 si toglie e se ne può aggiungere ancora 1');
  await scegli('#phx [data-add="cam"]').catch(() => {});
  await p.click('#gua [data-v="CAM. OMRON"]'); await p.click('#int [data-v="REGOLAZIONE"]'); await attesa(100);
  await p.fill('#altro', 'regolata telecamera e messa a fuoco'); await p.click('#go'); await attesa(300); await p.click('#go').catch(() => {}); await attesa(900);
  const fx = posts.filter(o => o.action === 'foto' && o.fermo === 'FC5'), cl5 = posts.filter(o => o.action === 'close' && o.id === 'FC5').pop() || {};
  ok(cl5.foto && fx.length === 2 && fx.map(o => o.n).join() === '2,3' && fx.every(o => o.foto && o.foto.length < 400000), 'chiusura con la foto 1, poi foto 2 e 3 inviate a parte (ridimensionate)');
  ok(posts.indexOf(cl5) < posts.indexOf(fx[0]), 'prima parte la chiusura, poi le foto');
  const det = await p.evaluate(() => { const f = Object.assign({}, S.fermi.find(x => x.id === 'FC5') || { id: 'FC5', start: Date.now() - 6e5, end: Date.now(), linea: 'ST3', macchina: 'AST' }, { stato: 'CHIUSO', end: Date.now(), fotoC: 'https://drive/1', fotoX: ['https://drive/2', 'https://drive/3'] });
    intSheet(f, eqBy('ST3-04')); const t = [...document.querySelectorAll('#sheet .fotorow a')].map(a => a.textContent.trim()); closeSh(); return t; });
  ok(det.join('|') === 'Foto 1|Foto 2|Foto 3', 'dettaglio dell\'intervento: Foto 1 · Foto 2 · Foto 3');
  // tecnico: soluzione scartata con il motivo
  await p.evaluate(() => { closeSh(); S.solScartate = [{ id: 'SX1', ts: Date.now(), macchina: 'ASTUCCIATRICE', guasto: 'INCASTRO', problema: 'regolato sponde', autore: 'ROSSI Mario', motivo: 'Troppo vaga: manca dove e come', da: 'VERDI Anna' }]; go('home'); }); await attesa(300);
  ok(await p.isVisible('#scartban'), 'tecnico: avviso "una tua soluzione è stata scartata"');
  await p.click('#scartban'); await attesa(200);
  ok(/Troppo vaga: manca dove e come/.test(await p.innerText('#sheet')), 'il tecnico legge perché è stata scartata');
  await p.evaluate(() => { closeSh(); go('home'); }); await attesa(300);
  ok(!(await p.isVisible('#scartban')), 'letta: l\'avviso sparisce');
  await p.evaluate(() => { closeSh(); go('home'); }); await attesa(300);
  ok(await p.isVisible('#corrban') && !(await p.isVisible('#revban')), 'tecnico: avviso "un tuo fermo è stato corretto"; nessuna lista Da rivedere');
  await p.click('#corrban'); await attesa(200);
  ok(/INCASTRO → MAG. ISTRUZIONI/.test(await p.innerText('#sheet')), 'il tecnico vede cosa ha cambiato il coordinatore');
  await p.click('#corrok'); await attesa(300);
  ok(!(await p.isVisible('#corrban')), 'dopo "Ho capito" l\'avviso sparisce');
  await esci(); await entra('Anna', 'Verdi'); await p.evaluate(() => { S.ver = 'z'; return refresh(); }); await attesa(400); await p.evaluate(() => { closeSh(); go('home'); }); await attesa(300);
  ok(await p.isVisible('#revban'), 'coordinatore: in home "N fermi da rivedere"');
  await p.click('#revban'); await attesa(200);
  ok(/parla di sostituzione/.test(await p.innerText('#sheet')), 'motivo mostrato: descrizione "sostituito" ma intervento REGOLAZIONE');
  await p.click('[data-ok="RV1"]'); await attesa(600);
  ok(posts.some(o => o.action === 'rivisto' && o.fermo === 'RV1' && o.manutentore === 'VERDI Anna'), '"Va bene": il fermo viene segnato come rivisto');
  await p.evaluate(() => { closeSh(); S.ver = 'zz'; return refresh(); }); await attesa(400); await p.evaluate(() => go('home')); await attesa(200);
  ok(await p.evaluate(() => S.fermi.some(f => f.id === 'RV1' && !f.rivisto)) && !(await p.isVisible('#revban')), '"Va bene": l\'avviso sparisce subito, anche se il foglio non ha ancora aggiornato i dati');
  await p.evaluate(() => { rivSegna('RV1', false); go('home'); }); await attesa(200);
  ok(await p.isVisible('#revban'), 'se il foglio rifiuta la revisione, il fermo torna in "Da rivedere"');
  await p.evaluate(() => { rivSegna('RV1', true); go('home'); }); await attesa(100);

  console.log('\nTelefono nuovo aperto dal QR (macchine solo dal foglio)');
  const tel = await b.newContext({ viewport: { width: 390, height: 844 }, serviceWorkers: 'block' });
  await tel.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ contentType: 'text/css', body: '' }));
  await tel.route(/script\.google\.com/, r => r.fulfill({ contentType: 'application/json',
    body: r.request().method() === 'POST' || /action=login/.test(r.request().url()) ? '{"ok":true}' : JSON.stringify({ ok: true, ...D, ver: 'V' + ver }) }));
  const t = await tel.newPage(); t.on('pageerror', e => errori.push(e.message));
  await t.goto(`http://localhost:${PORT}/fermi/?eq=ST3-05`); await t.waitForTimeout(400);
  ok(await t.evaluate(() => S.equipment.length === 0), 'prima dell\'accesso nessuna macchina in memoria: nel codice pubblico non c\'è l\'elenco reale');
  await t.fill('#lpin', 'prova'); await t.fill('#lnome', 'Mario'); await t.fill('#lcogn', 'Rossi'); await t.click('#lgo'); await t.waitForTimeout(3500);
  const sh = await t.innerText('#sheet');
  ok(/Apri fermo/i.test(sh) && /ST3-05/.test(sh) && /BILANCIA/i.test(sh), 'dopo l\'accesso le macchine arrivano dal foglio e si apre il fermo della macchina del QR');
  await tel.close();

  console.log('\nFermo con più interventi: giro, divisione alla chiusura, controllo antifurbi');
  const D2 = { ...D, guasti: { TUTTE: ['SENSORE', 'GUIDE', 'ALTRO'] }, vociCat: { SENSORE: 'ELETTRICO', GUIDE: 'MECCANICO' },
    fermi: [{ id: 'GA', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 30 * 60000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' }] };
  const P2 = []; let v2 = 1;
  const gc = await b.newContext({ viewport: { width: 390, height: 844 }, serviceWorkers: 'block' });
  await gc.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ contentType: 'text/css', body: '' }));
  await gc.route(/script\.google\.com/, r => {
    if (r.request().method() === 'POST') { // come il foglio vero: apre e chiude le righe, con Fermo n. e Intervento n.
      const o = JSON.parse(r.request().postData()); P2.push(o); v2++;
      if (o.action === 'open') D2.fermi.unshift({ id: o.id, code: o.code, linea: 'ST3', macchina: 'ASTUCCIATRICE', start: o.ts, stato: 'APERTO', lotto: o.lotto, man: o.manutentore, apertoDa: o.manutentore, gruppo: o.gruppo || '', giro: o.giro || 0 });
      if (o.action === 'close') { const f = D2.fermi.find(x => x.id === o.id); Object.assign(f, { stato: 'CHIUSO', end: o.ts, guasto: o.guasto, int: o.intervento, altro: o.altro, multi: o.multi || '' }, o.gruppo ? { gruppo: o.gruppo, giro: o.giro } : {}); }
      return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' });
    }
    r.fulfill({ contentType: 'application/json', body: /action=login/.test(r.request().url()) ? '{"ok":true}' : JSON.stringify({ ok: true, ...D2, ver: 'G' + v2 }) });
  });
  const g = await gc.newPage(); g.on('pageerror', e => errori.push(e.message));
  await g.goto(`http://localhost:${PORT}/fermi/`); await g.waitForTimeout(400);
  await g.fill('#lpin', 'prova'); await g.fill('#lnome', 'Mario'); await g.fill('#lcogn', 'Rossi'); await g.click('#lgo'); await g.waitForTimeout(1800);
  await g.evaluate(() => { closeSh(); closeSheet(allFermi().find(f => f.id === 'GA')); }); await g.waitForTimeout(300);
  ok(await g.isVisible('#giro') && /Intervento fatto, continuo/i.test(await g.innerText('#giro')), 'Packaging: pulsante "Intervento fatto, continuo" accanto a Chiudi fermo');
  await g.click('#gua [data-v="SENSORE"]'); await g.click('#int [data-v="REGOLAZIONE"]'); await g.fill('#altro', 'regolato sensore ingresso');
  await g.click('#giro'); await g.waitForTimeout(2500);
  const c1 = P2.find(o => o.action === 'close' && o.id === 'GA') || {}, o2 = P2.find(o => o.action === 'open' && o.gruppo === 'GA') || {};
  ok(c1.gruppo === 'GA' && c1.giro === 1 && !c1.dichiarato, 'giro: intervento 1 chiuso adesso (tempo reale), Fermo n. GA');
  ok(o2.giro === 2 && o2.ts === c1.ts && o2.code === 'ST3-04' && o2.lotto === '601095', 'giro: intervento 2 aperto dallo stesso istante, stessa macchina e stesso lotto');
  await g.evaluate(() => go({ name: 'equip', code: 'ST3-04' })); await g.waitForTimeout(300);
  ok(/intervento 2/.test(await g.innerText('#view')) && /^(0:)?3\d:/.test((await g.innerText('#view .hero .timer')).trim()), 'scheda macchina: "intervento 2" e cronometro del fermo intero (dall\'inizio del primo)');
  await g.click('[data-close]'); await g.waitForTimeout(300);
  ok(/intervento 2/i.test(await g.innerText('#sheet h3')) && /1\.\s*SENSORE/.test(await g.innerText('#sheet .giri')), 'chiusura intervento 2: titolo e interventi già registrati');

  await g.click('#gua [data-v="GUIDE"]'); await g.click('#int [data-v="SOSTITUZIONE"]'); await g.fill('#pezzo', 'X-100'); await g.fill('#altro', 'sostituita guida e poi regolato sensore');
  await g.click('#go'); await g.waitForTimeout(200); await g.click('#go'); await g.waitForTimeout(300); // 1° tocco: conferma micro-fermo
  ok(await g.isVisible('#err.mwarn') && /Sembrano più interventi/.test(await g.innerText('#err')) && !P2.some(o => o.action === 'close' && o.id === o2.id), 'antifurbi: descrizione con più interventi → avviso, non chiude');
  await g.click('#mdiv'); await g.waitForTimeout(200);
  ok(await g.isChecked('#multi') && await g.isVisible('#mfin') && /Salva intervento 2 e continua/i.test(await g.innerText('#go')) && await g.isHidden('#giro'), '"Dividi in più interventi": chiede la fine di questo intervento');
  await g.uncheck('#multi'); await g.dispatchEvent('#multi', 'change'); await g.click('#go'); await g.waitForTimeout(300);
  ok(await g.isVisible('#err.mwarn'), 'chiudendo comunque: l\'avviso torna');
  await g.click('#muno'); await g.waitForTimeout(2000);
  const c2 = P2.find(o => o.action === 'close' && o.id === o2.id) || {};
  ok(c2.gruppo === 'GA' && c2.giro === 2 && /confermato: un intervento solo/.test(c2.multi || ''), '"No, è un intervento solo": chiude e registra l\'avviso ignorato');

  // divisione alla chiusura: fermo già finito registrato con gli orari (come per il filling)
  const hm = t => { const d = new Date(t); return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); };
  const t1 = Date.now() - 40 * 60000, t2 = Date.now() - 25 * 60000, t3 = Date.now() - 10 * 60000;
  await g.evaluate(() => { closeSh(); closeSheet({ id: 'RG', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', lotto: '601095' }, { reg: true }); }); await g.waitForTimeout(300);
  ok(await g.isHidden('#giro'), 'fermo registrato con gli orari: niente giro, solo divisione alla chiusura');
  await g.fill('#tini', hm(t1)); await g.fill('#tfin', hm(t2)); await g.check('#multi'); await g.dispatchEvent('#multi', 'change');
  ok(/Fine intervento 1/.test(await g.innerText('#tfinl')), 'con "più interventi" la Fine diventa la fine dell\'intervento 1');
  await g.click('#gua [data-v="SENSORE"]'); await g.click('#int [data-v="PULIZIA"]'); await g.fill('#altro', 'pulita fotocellula');
  await g.click('#go'); await g.waitForTimeout(500);
  ok(/intervento 2/i.test(await g.innerText('#sheet h3')) && await g.inputValue('#tini') === hm(t2), 'dopo l\'intervento 1 si apre l\'intervento 2 con inizio = fine dell\'1');
  await g.fill('#tfin', hm(t3)); await g.click('#gua [data-v="GUIDE"]'); await g.click('#int [data-v="REGOLAZIONE"]'); await g.fill('#altro', 'guida regolata');
  await g.click('#go'); await g.waitForTimeout(2500);
  const rg = P2.filter(o => o.gruppo === 'RG');
  ok(rg.length === 4 && rg.filter(o => o.action === 'close').map(o => o.giro).join() === '1,2' && rg.every(o => o.action !== 'open' || o.dichiarato), 'registrati 2 interventi, Fermo n. RG, orari dichiarati');
  const ra = rg.find(o => o.action === 'close' && o.giro === 1), rb = rg.find(o => o.action === 'open' && o.giro === 2);
  ok(ra && rb && Math.abs(ra.ts - rb.ts) < 60000, 'l\'intervento 2 inizia dove finisce l\'1');
  ok(await g.evaluate(() => /Più interventi/.test(motiviRev({ id: 'x', code: 'ST3-04', guasto: 'GUIDE', int: 'SOSTITUZIONE', altro: 'sostituita cinghia, regolato nonio', stato: 'CHIUSO' }, []).join()) && !piuInterventi('regolato sensore ingresso') && !!piuInterventi('pulito + resettato')),
    'coordinatore: "Più interventi in un fermo?" tra i motivi da rivedere');
  // giro con "Scegli orario": l'intervento finisce all'ora scritta e il successivo parte da lì (dichiarati)
  D2.fermi.unshift({ id: 'GB', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 40 * 60000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' });
  await g.evaluate(() => { S.ver = 'gb'; return refresh(); }); await g.waitForTimeout(300);
  await g.evaluate(() => { closeSh(); closeSheet(allFermi().find(f => f.id === 'GB')); }); await g.waitForTimeout(300);
  const tg = Date.now() - 20 * 60000;
  await g.click('#gua [data-v="GUIDE"]'); await g.click('#int [data-v="REGOLAZIONE"]'); await g.fill('#altro', 'regolate guide');
  await g.click('#cwhen [data-v="ora"]'); await g.fill('#cfin', hm(tg)); await g.click('#giro'); await g.waitForTimeout(1500);
  const gbc = P2.find(o => o.action === 'close' && o.id === 'GB') || {}, gbo = P2.find(o => o.action === 'open' && o.gruppo === 'GB') || {};
  ok(hm(gbc.ts) === hm(tg) && gbc.dichiarato && gbo.ts === gbc.ts && gbo.dichiarato, 'giro con "Scegli orario": intervento chiuso all\'ora scritta, il successivo parte da lì (dichiarati)');
  await g.evaluate(id => { closeSh(); closeSheet(allFermi().find(f => f.id === id)); }, gbo.id); await g.waitForTimeout(300);
  await g.click('#canc');
  ok(/annullare l'intervento 2/.test(await g.innerText('#canc')), 'annulla intervento 2: la conferma parla dell\'intervento, non del fermo');
  await g.click('#canc'); await g.waitForTimeout(1500);
  ok(P2.some(o => o.action === 'cancel' && o.id === gbo.id) && !P2.some(o => o.action === 'cancel' && o.id === 'GB'), 'annulla intervento 2: cancellato solo quello, l\'intervento 1 resta');
  ok(await g.evaluate(() => !!piuInterventi('regolato posizione sensore e regolato lama sbucciatrice') && !!piuInterventi('rimesso in dima pinze ctg e regolato sensore ingrese') && !piuInterventi('Regolato guide e sensore ingresso') && !piuInterventi('leva trovata spostata, rimessa in posizione')),
    'antifurbi nell\'app: due operazioni unite da "e" riconosciute, una sola no');
  await gc.close();

  console.log('\nFormati e cambi formato (solo coordinatori)');
  const FX = { ok: true, car: ['Contenitore', 'Tappo', 'Contagocce', 'Dosaggio', 'Etichetta', 'Astuccio', 'Istruzione', 'Pezzi/astuccio', 'Cartone', 'Fardello', 'Prodotto'], letto: Date.now(), avvisi: [],
    linee: [{ id: 'SPL2', rep: 'SPL2', pro: '' }],
    prod: [{ c: '6000001', d: 'PROVA A', lp: 'ST6', ls: ['SPL2'], a: ['FL.A', '', '', '5 ml', 'E1', 'A1', '', '1', 'C1', 'BUND'], f: 'SECONDARIO', fmt: 'FP-A', ff: 'FF-A', es: 'OK' },
           { c: '6000002', d: 'PROVA B', lp: 'ST5', ls: ['SPL1', 'SPL2'], a: ['FL.B', '', '', '10 ml', 'E2', 'A1', '', '1', 'C2', ''], f: 'SECONDARIO', fmt: 'FP-B', ff: 'FF-B', es: 'OK' },
           { c: '6000003', d: 'PROVA A2', lp: 'ST6', ls: ['SPL2'], a: ['FL.A', '', '', '5 ml', 'E9', 'A1', '', '1', 'C1', 'BUND'], f: 'SECONDARIO', fmt: 'FP-A', ff: 'FF-A', es: 'OK' }],
    macc: [{ id: 'SPL2-01', l: 'SPL2', pos: 1, n: 'ETICHETTATRICE', fase: 'SECONDARIO', q: 'SEMPRE', m: ['P', '', '', '', 'R', '', '', '', '', '', ''], ver: true },
           { id: 'SPL2-02', l: 'SPL2', pos: 2, n: 'CARTONATRICE PAK', fase: 'SECONDARIO', q: 'SOLO SENZA FARDELLO', m: ['', '', '', '', '', '', '', '', 'P', '', ''], ver: true },
           { id: 'SPL2-03', l: 'SPL2', pos: 3, n: 'FARDELLATRICE', fase: 'SECONDARIO', q: 'SOLO CON FARDELLO', m: ['', '', '', '', '', '', '', '', '', 'P', ''], ver: true }] };
  let fmtReq = 0; await ctx.unroute(/script\.google\.com/);
  await ctx.route(/script\.google\.com/, r => { const u = decodeURIComponent(r.request().url()); if (/action=formati/.test(u)) { fmtReq++; return r.fulfill({ contentType: 'application/json', body: /VERDI/.test(u) ? JSON.stringify(FX) : '{"ok":false,"error":"Sezione riservata ai coordinatori"}' }); }
    if (r.request().method() === 'POST') return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' });
    r.fulfill({ contentType: 'application/json', body: /action=login/.test(u) ? '{"ok":true}' : JSON.stringify({ ok: true, ...D, ver: 'F' }) }); });
  await esci(); await entra('Mario', 'Rossi'); await p.evaluate(() => closeSh()); await p.click('#manbtn'); await attesa(200);
  ok(!(await p.isVisible('#mfmt')) && fmtReq === 0, 'tecnico: nessuna voce Formati e nessun dato dei formati scaricato');
  await esci(); await entra('Anna', 'Verdi'); await p.evaluate(() => closeSh()); await p.click('#manbtn'); await attesa(200);
  ok(await p.isVisible('#mfmt'), 'coordinatore: voce "Formati e cambi formato" nel menu del nome');
  await p.click('#mfmt'); await attesa(500);
  await p.fill('#fa', '6000001'); await p.fill('#fb', '6000002'); await p.dispatchEvent('#fb', 'input'); await attesa(200);
  const ftx = await p.innerText('#fres');
  ok(/CAMBIO FORMATO/.test(ftx) && /FP-A → FP-B/.test(ftx) && /ETICHETTATRICE\s*CAMBIO PEZZI/.test(ftx) && /Contenitore: FL\.A → FL\.B/.test(ftx), 'cambio formato: banner, macchina con cambio pezzi (da → a)');
  ok(/CARTONATRICE PAK\s*DA AVVIARE/.test(ftx) && /FARDELLATRICE\s*SI FERMA/.test(ftx), 'regola fardello: cartonatrice da avviare, fardellatrice si ferma');
  await p.fill('#fb', '6000003'); await p.dispatchEvent('#fb', 'input'); await attesa(200);
  ok(/STESSO FORMATO – SOLO RICETTE/.test(await p.innerText('#fres')) && /Etichetta: E1 → E9/.test(await p.innerText('#fres')), 'stesso formato, cambia solo l\'etichetta: solo ricetta');
  ok(await p.evaluate(() => !Object.keys(localStorage).some(k => /FP-A|6000002/.test(localStorage.getItem(k) || ''))), 'i dati dei formati non restano salvati sul telefono');
  await p.evaluate(() => closeSh());

  console.log('\nVoci guasto: solo quelle della macchina (foglio Guasti)');
  const vg = await p.evaluate(() => {
    const old = { g: S.guasti, f: S.fermi };
    S.guasti = { 'ST3-04': ['MAG. ASTUCCI', 'CHIUSURE AST', 'SENSORE'], 'ST3-03': ['PARAM VIDEOJET'], ST3: ['GUIDE'], 'ST1MAR-05': ['MAG. ISTRUZIONI'], TUTTE: ['SENSORE', 'ARIA COMPRESSA', 'INCASTRO'] };
    const t = Date.now(), f = (id, code, g) => ({ id, code, linea: 'ST3', macchina: 'X', start: t - 3600e3, end: t - 3000e3, stato: 'CHIUSO', guasto: g });
    S.fermi = old.f.concat([f('V1', 'ST3-04', 'CHIUSURE AST'), f('V2', 'ST3-04', 'CHIUSURE AST'), f('V3', 'ST3-04', 'INCASTRO'), f('V4', 'ST3-03', 'CHIUSURE AST')]);
    const eq = { code: 'ST3-04', linea: 'ST3', area: 'Packaging', macchina: 'ASTUCCIATRICE' };
    const lista = guastiFor(eq), altra = guastiFor({ code: 'ST3-05', linea: 'ST3', area: 'Packaging', macchina: 'BILANCIA' }), mar = guastiFor({ code: 'ST1MAR-05', linea: 'ST1 MAR', area: 'Packaging', macchina: 'ASTUCCIATRICE' });
    const chipsTxt = html => [...new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html').querySelectorAll('[data-v]')].map(b => b.dataset.v);
    const ordine = chipsTxt(guaChips(eq, null));
    closeSh(); closeSheet({ id: 'VX', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: t - 600000, stato: 'APERTO', lotto: '601095', man: S.man, apertoDa: S.man });
    const chiusura = [...document.querySelectorAll('#gua [data-v]')].map(b => b.dataset.v); closeSh();
    S.guasti = old.g; S.fermi = old.f;
    return { lista, altra, mar, ordine, chiusura };
  });
  ok(vg.lista.join('|') === 'MAG. ASTUCCI|CHIUSURE AST|SENSORE|GUIDE|ARIA COMPRESSA|INCASTRO|ALTRO', 'ST3-04: voci della macchina + linea ST3 + TUTTE, senza doppioni (SENSORE una volta), ALTRO in fondo: ' + vg.lista.join('|'));
  ok(!vg.lista.includes('PARAM VIDEOJET') && !vg.lista.includes('MAG. ISTRUZIONI'), 'niente voci di altre macchine (PARAM VIDEOJET dell\'etichettatrice, MAG. ISTRUZIONI di ST1 MAR)');
  ok(vg.altra.join('|') === 'GUIDE|SENSORE|ARIA COMPRESSA|INCASTRO|ALTRO', 'macchina senza riga propria: solo linea + TUTTE + ALTRO');
  ok(vg.mar.join('|') === 'MAG. ISTRUZIONI|SENSORE|ARIA COMPRESSA|INCASTRO|ALTRO', 'altra linea: la riga di linea ST3 non c\'entra');
  ok(vg.ordine.slice(0, 2).join('|') === 'CHIUSURE AST|INCASTRO' && vg.ordine.at(-1) === 'ALTRO', 'in ordine di frequenza su QUESTA macchina (il fermo dell\'etichettatrice non conta), ALTRO sempre in fondo: ' + vg.ordine.join('|'));
  ok(vg.chiusura.join('|') === vg.ordine.join('|'), 'chiusura fermo: le stesse voci, nello stesso ordine');

  console.log('\nFermi per linea');
  const SP = process.env.SHOT_LIN || '';
  await p.evaluate(() => { const t = Date.now(), h = 3600e3, mk = (id, code, mac, ore, min, g, int, man, alt) => ({ id, code, linea: 'ST3', macchina: mac, start: t - ore * h, end: t - ore * h + min * 60000, stato: 'CHIUSO', guasto: g, int, man, altro: alt || '', lotto: '601095' });
    S.equipment = S.equipment.concat([{ code: 'ST3-03', linea: 'ST3', area: 'Packaging', macchina: 'ETICHETTATRICE' }]);
    S.fermi = S.fermi.filter(f => f.linea !== 'ST3' || f.stato === 'APERTO').concat([mk('L1', 'ST3-04', 'ASTUCCIATRICE', 2, 25, 'MAG. ASTUCCI', 'REGOLAZIONE', 'ROSSI Mario', 'regolate guide magazzino'), mk('L2', 'ST3-04', 'ASTUCCIATRICE', 30, 40, 'MAG. ASTUCCI', 'REGOLAZIONE', 'BIANCHI Luca'),
      mk('L3', 'ST3-03', 'ETICHETTATRICE', 5, 15, 'PARAM VIDEOJET', 'REGOLAZIONE', 'ROSSI Mario'), mk('L4', 'ST3-02', 'RIEMPITRICE', 8, 50, 'H7, Sblocco Homing', 'RIPRISTINO', 'BIANCHI Luca'), mk('L5', 'ST3-04', 'ASTUCCIATRICE', 60, 20, 'CHIUSURE AST', 'SOSTITUZIONE', 'ROSSI Mario'), mk('L6', 'ST3-02', 'RIEMPITRICE', 200, 30, 'H7, Sblocco Homing', 'RIPRISTINO', 'ROSSI Mario')]); go('home'); });
  await attesa(300);
  ok(await p.isVisible('.lchips2 [data-linea="ST3"]'), 'home: sotto "Chiusi oggi" le linee con il numero di interventi');
  if (SP) await p.evaluate(() => document.querySelector('.lchips2').scrollIntoView({ block: 'center' })), await p.screenshot({ path: SP.replace('.png', '_home.png') });
  await p.click('.lchips2 [data-linea="ST3"]'); await attesa(300);
  const lv = await p.innerText('#view');
  ok(/linea st3/i.test(lv) && /filling/i.test(lv) && /packaging/i.test(lv) && /astucciatrice/i.test(lv) && /riempitrice/i.test(lv), 'vista linea: Filling e Packaging separati, per macchina');
  ok(/linea · 5/i.test(lv) && /1 fermo · 1 intervento · 50′/.test(lv), 'ultimi 7 giorni: 5 interventi (quello di 8 giorni fa escluso); Filling: 1 fermo, 1 intervento, 50 min');
  if (SP) await p.screenshot({ path: SP, fullPage: true });
  await p.click('#lgg [data-v="30"]'); await attesa(200);
  ok(/linea · 6/i.test(await p.innerText('#view')), '30 giorni: compare anche il fermo di 8 giorni fa');
  await p.click('.mrow[data-eq="ST3-04"]'); await attesa(300);
  ok(await p.isVisible('.lnkline[data-linea="ST3"]'), 'scheda macchina: link "Fermi di tutta la linea ST3"');
  if (SP) { await p.evaluate(() => document.querySelector('.lnkline').scrollIntoView({ block: 'center' })); await p.screenshot({ path: SP.replace('.png', '_mac.png') }); }
  await p.click('.lnkline'); await attesa(300);
  ok(/linea st3/i.test(await p.innerText('#view')), 'dal link si apre la vista della linea');
  await p.evaluate(() => go('home'));

  console.log('\nTelefono in attesa di approvazione: schermata che si sblocca da sola, ID in più punti, approvazione dal coordinatore');
  {
    let approvato = false; const P3 = []; const daApp = [{ tecnico: 'ROSSI Mario', dev: 'T-NUOVO1', modello: 'Android K', quando: Date.now() - 4 * 60000, vecchi: [{ modello: 'Android K', ultimo: Date.now() - 10 * 864e5 }], altro: '' },
      { tecnico: 'BIANCHI Luca', dev: 'T-ALTRO9', modello: 'iOS iPhone', quando: Date.now() - 60 * 60000, vecchi: [], altro: 'NERI Paolo' }];
    const tw = await b.newContext({ viewport: { width: 390, height: 844 }, serviceWorkers: 'block' });
    await tw.route(/fonts\.(googleapis|gstatic)\.com/, r => r.fulfill({ contentType: 'text/css', body: '' }));
    await tw.route(/script\.google\.com/, r => { const u = decodeURIComponent(r.request().url());
      if (r.request().method() === 'POST') { P3.push(JSON.parse(r.request().postData())); return r.fulfill({ contentType: 'application/json', body: '{"ok":true}' }); }
      if (/action=login/.test(u)) return r.fulfill({ contentType: 'application/json', body: approvato || /VERDI/.test(u) ? '{"ok":true}' : JSON.stringify({ ok: false, error: 'Telefono non autorizzato per ROSSI Mario: telefono nuovo. Chiedi al responsabile di approvarlo' }) });
      if (/action=telefoni/.test(u)) return r.fulfill({ contentType: 'application/json', body: JSON.stringify(/VERDI/.test(u) ? { ok: true, telefoni: daApp } : { ok: false, error: 'Sezione riservata ai coordinatori' }) });
      r.fulfill({ contentType: 'application/json', body: JSON.stringify({ ok: true, ...D, ver: 'TW' }) }); });
    const w = await tw.newPage(); w.on('pageerror', e => errori.push(e.message));
    await w.goto(`http://localhost:${PORT}/fermi/`); await w.waitForTimeout(600);
    const dev0 = await w.evaluate(() => DEV);
    await w.fill('#lpin', 'prova'); await w.fill('#lnome', 'Mario'); await w.fill('#lcogn', 'Rossi'); await w.click('#lgo'); await w.waitForTimeout(1500);
    ok(await w.isVisible('#lwait') && await w.isHidden('#lform') && /in attesa/i.test(await w.innerText('#lwait')) && /ROSSI Mario/.test(await w.innerText('#lwtxt')), 'telefono in attesa: schermata di attesa al posto dell\'errore');
    if (SP) await w.screenshot({ path: SP.replace('.png', '_telwait.png') });
    approvato = true; await w.click('#lwgo'); await w.waitForTimeout(1200);
    ok(await w.isHidden('#login') && /ROSSI Mario/.test(await w.innerText('#manbtn')), 'approvato: entra senza riscrivere nome e codice');
    const copie = await w.evaluate(async () => ({ c: devCookie(), i: await media('get', 'cfg:dev'), k: await (await (await caches.open(DEV_CACHE)).match(DEV_URL)).text() }));
    ok(/^T-[A-Z0-9]{4,12}$/.test(dev0) && copie.c === dev0 && copie.i === dev0 && copie.k === dev0, 'ID del telefono salvato anche in cookie, archivio e cache');
    await w.evaluate(() => localStorage.removeItem('pf_dev')); await w.reload(); await w.waitForTimeout(1200);
    ok(await w.evaluate(() => DEV) === dev0, 'memoria piccola svuotata: stesso ID (dal cookie)');
    await w.evaluate(() => { localStorage.removeItem('pf_dev'); document.cookie = 'pf_dev=; max-age=0'; }); await w.reload(); await w.waitForTimeout(1200);
    ok(await w.evaluate(() => DEV) === dev0, 'memoria piccola e cookie svuotati: stesso ID (dall\'archivio)');
    await w.evaluate(async () => { localStorage.removeItem('pf_dev'); document.cookie = 'pf_dev=; max-age=0'; await media('del', 'cfg:dev'); }); await w.reload(); await w.waitForTimeout(1200);
    ok(await w.evaluate(() => DEV) === dev0, 'resta solo la cache: stesso ID');
    ok(!(await w.isVisible('#telban')), 'un tecnico non vede i telefoni da approvare');
    // coordinatore
    await w.evaluate(() => { salvaSess(null); showLogin(); }); await w.fill('#lnome', 'Anna'); await w.fill('#lcogn', 'Verdi'); await w.click('#lgo'); await w.waitForTimeout(1500);
    ok(await w.isVisible('#telban') && /2 telefoni da approvare/.test(await w.innerText('#telban')), 'coordinatore: banner "2 telefoni da approvare" in home');
    await w.click('#telban'); await w.waitForTimeout(300);
    const sh = await w.innerText('#sheet');
    ok(/telefono nuovo/i.test(sh) && /Stesso modello/.test(sh) && /telefono di un altro/i.test(sh) && /NERI Paolo/.test(sh), 'elenco: caso "stesso modello" e caso "telefono di un altro" spiegati');
    if (SP) await w.screenshot({ path: SP.replace('.png', '_telapp.png') });
    await w.click('[data-tel="0"][data-es="OK"]'); await w.waitForTimeout(1500);
    const op = P3.find(o => o.action === 'telefono') || {};
    ok(op.tecnico === 'ROSSI Mario' && op.devT === 'T-NUOVO1' && op.esito === 'OK' && op.manutentore === 'VERDI Anna', 'Approva: inviato al foglio');
    ok(/1 telefono da approvare/i.test(await w.innerText('#sheet h3').catch(() => '')) || /BIANCHI Luca/.test(await w.innerText('#sheet')) && !/ROSSI Mario/.test(await w.innerText('#sheet')), 'approvato: sparisce subito dall\'elenco');
    await tw.close();
  }

  ok(!errori.length, 'nessun errore JavaScript nella pagina' + (errori.length ? ': ' + errori.join(' | ') : ''));
  await b.close(); srv.close(); fine('App');
})().catch(e => { console.error(e); process.exit(1); });
