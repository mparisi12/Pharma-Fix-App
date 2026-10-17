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
const D = { liste: { Manutentori: ['ROSSI Mario', 'BIANCHI Luca'], Coordinatori: ['VERDI Anna'] }, guasti: {}, soluzioni: [], anomalie: [], turni: {}, programma: {}, continuo: ['ST3'],
  fermi: [{ id: 'F0', code: 'ST3-02', linea: 'ST3', macchina: 'RIEMPITRICE + TAPPATRICE FLEXFILL F940', start: now - 20 * 60000, stato: 'APERTO', lotto: '601090', man: 'BIANCHI Luca', apertoDa: 'BIANCHI Luca' }],
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
  ok(/FLEXFILL/.test(await p.innerText('#view')), 'la home mostra il fermo aperto');
  const f0 = full; for (let i = 0; i < 3; i++) { await p.evaluate(() => refresh()); await attesa(150); }
  ok(full === f0 && same >= 3, 'aggiornamenti senza modifiche: nessun riscaricamento completo');

  console.log('\nApri fermo');
  await p.click('#manual'); await attesa(300); await p.click('#pl [data-v="ST3"]'); await attesa(200); await p.click('#pm [data-v="ST3-04"]'); await attesa(500);
  ok(await p.isVisible('#lotto'), 'popup centrale con il lotto');
  await p.fill('#lotto', '601095'); await p.click('#sheet #go'); await attesa(2500);
  ok(posts.some(o => o.action === 'open' && o.code === 'ST3-04' && o.lotto === '601095'), 'fermo inviato al foglio');
  ok(full > f0, 'dopo l\'invio i dati nuovi si riscaricano');

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
  D.fermi = D.fermi.concat([0, 1, 2].map(storia), [{ id: 'FC', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 600000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' }]);
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
  await p.click('#gua [data-v="CAM. OMRON"]'); await p.click('#int [data-v="REGOLAZIONE"]');
  await p.click('#cau [data-v="USURA"]'); await p.click('#cau [data-v="USURA"]'); await p.click('#cau [data-v="MATERIALE FUORI SPEC."]');
  await p.fill('#altro', 'regolata messa a fuoco telecamera'); await p.click('#go'); await attesa(800);
  const cl = posts.filter(o => o.action === 'close' && o.id === 'FC').pop() || {};
  ok(cl.categoria === 'AUTOMATION' && cl.intervento === 'REGOLAZIONE' && cl.guasto === 'CAM. OMRON' && cl.causa === 'MATERIALE FUORI SPEC.', 'inviato: categoria automatica, intervento, guasto e causa (secondo tocco toglie la scelta)');

  D.liste['Tipo intervento'] = ['REGOLAZIONE', 'SOSTITUZIONE', 'RIPRISTINO', 'PULIZIA']; D.intDesc = { RIPRISTINO: 'reset allarme, homing' }; D.liste.Causa = ['USURA', 'MANUTENZIONE NON ESEGUITA'];
  D.fermi.push({ id: 'FC2', code: 'ST3-04', linea: 'ST3', macchina: 'ASTUCCIATRICE', start: Date.now() - 300000, stato: 'APERTO', lotto: '601095', man: 'ROSSI Mario', apertoDa: 'ROSSI Mario' });
  await p.evaluate(() => { S.ver = 'x'; return refresh(); }); await attesa(400);
  await p.evaluate(() => { closeSh(); closeSheet(S.fermi.find(f => f.id === 'FC2')); }); await attesa(300);
  ok(/reset allarme, homing/.test(await p.innerText('#int')) && /MANUTENZIONE NON ESEGUITA/.test(await p.innerText('#cau')), 'interventi, spiegazioni e cause presi dal foglio Liste');
  await p.evaluate(() => closeSh());

  ok(!errori.length, 'nessun errore JavaScript nella pagina' + (errori.length ? ': ' + errori.join(' | ') : ''));
  await b.close(); srv.close(); fine('App');
})().catch(e => { console.error(e); process.exit(1); });
