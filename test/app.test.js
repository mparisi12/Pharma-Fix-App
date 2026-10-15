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
const D = { liste: { Manutentori: ['ROSSI Mario', 'BIANCHI Luca', 'VERDI Anna'], Coordinatori: ['VERDI Anna'] }, guasti: {}, soluzioni: [], anomalie: [], turni: {}, programma: {}, continuo: ['ST3'],
  fermi: [{ id: 'F0', code: 'ST3-02', linea: 'ST3', macchina: 'RIEMPITRICE + TAPPATRICE FLEXFILL F940', start: now - 20 * 60000, stato: 'APERTO', lotto: '601090', man: 'BIANCHI Luca', apertoDa: 'BIANCHI Luca' }],
  compiti: [], documenti: [], vestizioni: [{ man: 'ROSSI Mario', scad: g0.getTime() + 40 * 864e5 }] };
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
  ok(/VERDI Anna/.test(await p.innerText('#manbtn')), 'accesso del coordinatore');
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

  ok(!errori.length, 'nessun errore JavaScript nella pagina' + (errori.length ? ': ' + errori.join(' | ') : ''));
  await b.close(); srv.close(); fine('App');
})().catch(e => { console.error(e); process.exit(1); });
