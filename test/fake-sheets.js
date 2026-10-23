// Finto Google Sheet per provare Code.gs con Node, senza Google. Solo quello che serve ai test.
const fs = require('fs'), path = require('path'), vm = require('vm'), zlib = require('zlib');
const any = () => new Proxy(function () {}, { get: (t, k) => k === 'build' ? () => ({}) : () => any(), apply: () => any() });
function a1(s) {
  const m = s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/); if (!m) return null;
  const c = x => x.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0); const r1 = +m[2], c1 = c(m[1]);
  return m[3] ? [r1, c1, +m[4] - r1 + 1, c(m[3]) - c1 + 1] : [r1, c1, 1, 1];
}
class Sheet {
  constructor(n, rows = []) { this.n = n; this.rows = rows; }
  getName() { return this.n; } getLastRow() { return this.rows.length; } getMaxRows() { return 1000; } getMaxColumns() { return 30; } getLastColumn() { return Math.max(0, ...Array.from(this.rows, r => r ? r.length : 0)); }
  appendRow(r) { this.rows.push(r); } clearContents() { this.rows.length = 0; } clear() { this.rows.length = 0; }
  getDataRange() { return this.getRange(1, 1, Math.max(this.rows.length, 1), Math.max(1, this.getLastColumn())); }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === 'string') { const x = a1(r); if (!x) return any(); [r, c, nr, nc] = x; }
    const sh = this;
    const R = new Proxy({}, { get: (t, k) => {
      if (k === 'getValues' || k === 'getDisplayValues') return () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (sh.rows[r - 1 + i] || [])[c - 1 + j] ?? ''));
      if (k === 'getValue') return () => R.getValues()[0][0];
      if (k === 'setValues') return v => { v.forEach((row, i) => row.forEach((x, j) => { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; sh.rows[r - 1 + i][c - 1 + j] = x; })); return R; };
      if (k === 'setValue') return x => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; sh.rows[r - 1 + i][c - 1 + j] = x; } return R; };
      if (k === 'createTextFinder') return id => ({ matchEntireCell: () => ({ findNext: () => { for (let i = 0; i < nr; i++) if (String((sh.rows[r - 1 + i] || [])[c - 1]) === String(id)) return { getRow: () => r + i }; return null; } }) });
      if (k === 'clearContent') return () => { for (let i = 0; i < nr; i++) if (sh.rows[r - 1 + i]) for (let j = 0; j < nc; j++) sh.rows[r - 1 + i][c - 1 + j] = ''; return R; };
      if (k === 'getA1Notation') return () => { const L = n => { let x = ''; while (n > 0) { x = String.fromCharCode(65 + (n - 1) % 26) + x; n = Math.floor((n - 1) / 26); } return x; }; return L(c) + r; };
      if (k === 'getFormula') return () => (sh.f || {})[r + ',' + c] || '';
      if (k === 'setFormula') return f => { (sh.f = sh.f || {})[r + ',' + c] = f; return R; };
      if (k === 'setDataValidation') return v => { (sh.dv = sh.dv || []).push({ r, c, nr, nc, v }); return R; };
      if (k === 'setDataValidations') return v => { v.forEach((row, i) => (sh.dvr = sh.dvr || {})[(r + i) + ',' + c] = row[0]); return R; };
      return () => R;
    } });
    return R;
  }
}
Sheet.prototype.insertRowsAfter = function () { return this; };
Sheet.prototype.setRowHeight = function () { return this; };
Sheet.prototype.deleteRow = function (r) { this.rows.splice(r - 1, 1); return this; };
Sheet.prototype.getCharts = function () { return this.charts || []; };
Sheet.prototype.removeChart = function (c) { this.charts = (this.charts || []).filter(x => x !== c); };
Sheet.prototype.insertChart = function (c) { (this.charts = this.charts || []).push(c); };
Sheet.prototype.newChart = function () { const o = { opt: {} }; const b = new Proxy({}, { get: (t, k) => k === 'build' ? () => o : k === 'setOption' ? (a, v) => { o.opt[a] = v; return b; } : (...a) => { o[k] = a; return b; } }); return b; };
Sheet.prototype.insertColumnBefore = function (c) { this.rows.forEach(r => { while (r.length < c - 1) r.push(''); r.splice(c - 1, 0, ''); }); return this; };
Sheet.prototype.setConditionalFormatRules = function (r) { this.cf = r; };
Sheet.prototype.setName = function (n) { this.n = n; return this; };
for (const m of ['setFrozenColumns', 'hideSheet', 'setFrozenRows', 'setColumnWidth', 'hideColumns', 'insertColumnsAfter', 'clearConditionalFormatRules']) Sheet.prototype[m] = function () { return this; };

function carica(sheets) {
  const src = fs.readFileSync(path.join(__dirname, '../fermi/apps-script/Code.gs'), 'utf8');
  const props = {}, cache = {};
  const ss = { getSheetByName: n => Object.values(sheets).find(x => x && x.n === n) || null, insertSheet: n => (sheets[n] = new Sheet(n)), deleteSheet: s => delete sheets[s.n], getSheets: () => Object.values(sheets), toast() {}, setActiveSheet() {} };
  const blob = b => ({ getBytes: () => b, getDataAsString: () => Buffer.from(b).toString('utf8') });
  const p2 = n => String(n).padStart(2, '0');
  const ctx = {
    console, Logger: { log() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, newConditionalFormatRule: () => { const o = {}; const b = new Proxy({}, { get: (t, k) => k === 'build' ? () => o : (...a) => { o[k] = a; return b; } }); return b; }, newDataValidation: () => { const o = {}; const b = new Proxy({}, { get: (t, k) => k === 'build' ? () => o : (...a) => { o[k] = a; return b; } }); return b; }, flush() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, getAll: ks => Object.fromEntries(ks.map(k => [k, cache[k]])), putAll: o => Object.assign(cache, o) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) },
    Utilities: {
      newBlob: d => blob(typeof d === 'string' ? [...Buffer.from(d, 'utf8')] : d),
      gzip: b => blob([...zlib.gzipSync(Buffer.from(b.getBytes()))]), ungzip: b => blob([...zlib.gunzipSync(Buffer.from(b.getBytes()))]),
      base64Encode: b => Buffer.from(b).toString('base64'), base64Decode: s => [...Buffer.from(s, 'base64')],
      formatDate: (d, tz, f) => f.replace(/yyyy|MM|M|dd|d|HH|H|mm|u/g, t => ({ yyyy: d.getFullYear(), MM: p2(d.getMonth() + 1), M: d.getMonth() + 1, dd: p2(d.getDate()), d: d.getDate(), HH: p2(d.getHours()), H: d.getHours(), mm: p2(d.getMinutes()), u: (d.getDay() + 6) % 7 + 1 })[t])
    }
  };
  vm.createContext(ctx);
  vm.runInContext(src + ';this.C=C;this.HEAD=HEAD;this.DOC_H=DOC_H;this.SOL_H=SOL_H;this.EXTRA_H=EXTRA_H;this.DD=Date;this.isCoordinatore_=isCoordinatore_;this.PARTICELLE_=PARTICELLE_;this.C=C;', ctx);
  return { ctx, ss, props, cache };
}

// Mini-asserzioni: stampano OK/NO e alla fine escono con errore se qualcosa non va
let falliti = 0;
function ok(cond, msg) { console.log((cond ? '  ok  ' : '  NO  ') + msg); if (!cond) falliti++; }
function fine(nome) { console.log(falliti ? `\n${nome}: ${falliti} test FALLITI` : `\n${nome}: tutto ok`); process.exit(falliti ? 1 : 0); }
module.exports = { Sheet, carica, ok, fine };
