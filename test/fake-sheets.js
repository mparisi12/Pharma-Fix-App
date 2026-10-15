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
  getName() { return this.n; } getLastRow() { return this.rows.length; } getMaxRows() { return 1000; } getMaxColumns() { return 30; }
  appendRow(r) { this.rows.push(r); } clearContents() { this.rows = []; } clear() { this.rows = []; }
  getDataRange() { return this.getRange(1, 1, Math.max(this.rows.length, 1), Math.max(1, ...this.rows.map(r => r.length))); }
  getRange(r, c, nr = 1, nc = 1) {
    if (typeof r === 'string') { const x = a1(r); if (!x) return any(); [r, c, nr, nc] = x; }
    const sh = this;
    const R = new Proxy({}, { get: (t, k) => {
      if (k === 'getValues' || k === 'getDisplayValues') return () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (sh.rows[r - 1 + i] || [])[c - 1 + j] ?? ''));
      if (k === 'getValue') return () => R.getValues()[0][0];
      if (k === 'setValues') return v => { v.forEach((row, i) => row.forEach((x, j) => { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; sh.rows[r - 1 + i][c - 1 + j] = x; })); return R; };
      if (k === 'setValue') return x => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || []; sh.rows[r - 1 + i][c - 1 + j] = x; } return R; };
      if (k === 'createTextFinder') return id => ({ matchEntireCell: () => ({ findNext: () => { for (let i = 0; i < nr; i++) if (String((sh.rows[r - 1 + i] || [])[c - 1]) === String(id)) return { getRow: () => r + i }; return null; } }) });
      return () => R;
    } });
    return R;
  }
}
for (const m of ['hideSheet', 'setFrozenRows', 'setColumnWidth', 'hideColumns', 'insertColumnsAfter', 'clearConditionalFormatRules']) Sheet.prototype[m] = function () { return this; };

function carica(sheets) {
  const src = fs.readFileSync(path.join(__dirname, '../fermi/apps-script/Code.gs'), 'utf8');
  const props = {}, cache = {};
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)), deleteSheet: s => delete sheets[s.n], getSheets: () => Object.values(sheets), toast() {} };
  const blob = b => ({ getBytes: () => b, getDataAsString: () => Buffer.from(b).toString('utf8') });
  const p2 = n => String(n).padStart(2, '0');
  const ctx = {
    console, Logger: { log() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, newDataValidation: () => any(), flush() {} },
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
  vm.runInContext(src + ';this.C=C;this.HEAD=HEAD;this.DOC_H=DOC_H;this.DD=Date;this.isCoordinatore_=isCoordinatore_;this.PARTICELLE_=PARTICELLE_;', ctx);
  return { ctx, ss, props, cache };
}

// Mini-asserzioni: stampano OK/NO e alla fine escono con errore se qualcosa non va
let falliti = 0;
function ok(cond, msg) { console.log((cond ? '  ok  ' : '  NO  ') + msg); if (!cond) falliti++; }
function fine(nome) { console.log(falliti ? `\n${nome}: ${falliti} test FALLITI` : `\n${nome}: tutto ok`); process.exit(falliti ? 1 : 0); }
module.exports = { Sheet, carica, ok, fine };
