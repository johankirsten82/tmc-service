'use strict';
/**
 * MOCK of the full candidate Apps Script (photos + Check machine + Food stock).
 * This does not call Google. The live v5 project is not modified.
 * Import code below is a sample string, not a deployment secret.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const assert = require('assert');

const CANDIDATE = path.join(__dirname, 'candidate-Code.gs');
const ARTIFACT = '/opt/cursor/artifacts/TMC-Service-Code-stock.gs';
const LIVE_V5 = '/opt/cursor/artifacts/TMC-Service-Code.gs';
const code = fs.readFileSync(CANDIDATE, 'utf8');
const PIN = 'ab'.repeat(32);
const IMPORT = 'sample-import-code';
const LINK = 'mock-link-secret-0001';

function linkLine(s) { const m = s.match(/var LINK_SECRET_SETUP = '([^']*)', LINK_URL_SETUP = '([^']*)';/); if (!m) throw new Error('link line missing'); return m; }

const results = [];
function t(name, fn) {
  try { fn(); results.push('ok  ' + name); }
  catch (e) { results.push('FAIL ' + name + ' — ' + (e && e.message ? e.message : e)); }
}

t('MOCK files: candidate is sanitized, artifact matches it except the live link line', function () {
  const c = linkLine(code);
  assert.strictEqual(c[1], '__LINK_SECRET__');
  assert.strictEqual(c[2], '__LINK_URL__');
  assert.ok(code.includes('function photo_') && code.includes('function linkApply_') && code.includes('function stockImportEntry_'));
  assert.ok(code.includes("recs_(store, 'flc')"));
  assert.ok(!/script\.google\.com|AKfyc/.test(code));
  if (!fs.existsSync(ARTIFACT) || !fs.existsSync(LIVE_V5)) return;
  const art = fs.readFileSync(ARTIFACT, 'utf8');
  const live = fs.readFileSync(LIVE_V5, 'utf8');
  const a = linkLine(art), v = linkLine(live);
  if (a[1] !== v[1] || a[2] !== v[2]) throw new Error('artifact link line does not match live v5');
  if (a[1] === '__LINK_SECRET__') throw new Error('artifact lost the live link line');
  const strip = function (s) { return s.replace(/var LINK_SECRET_SETUP = '[^']*', LINK_URL_SETUP = '[^']*';/, 'LINK_LINE'); };
  const A = strip(art).split('\n'), B = strip(code).split('\n'), bad = [];
  for (let i = 0; i < Math.max(A.length, B.length); i++) if (A[i] !== B[i]) bad.push(i + 1);
  if (bad.length) throw new Error('artifact and candidate differ beyond the link line: ' + bad.join(','));
  console.log('    sha256 live-v5 ' + crypto.createHash('sha256').update(live).digest('hex'));
  console.log('    sha256 artifact ' + crypto.createHash('sha256').update(art).digest('hex'));
  console.log('    sha256 candidate ' + crypto.createHash('sha256').update(code).digest('hex'));
});

function load(src) {
  const bag = { FUEL_PULLED: String(Date.now()), STOCK_IMPORT_CODE: IMPORT, LINK_SECRET: LINK, LINK_URL: 'https://fuel-link.example/exec' };
  const triggers = [];
  let crashAfter = null, sets = 0;
  function sheet() {
    const sh = { rows: [] };
    sh.getLastRow = function () {
      for (let i = sh.rows.length - 1; i >= 0; i--) if ((sh.rows[i] || []).some(function (c) { return c !== '' && c != null; })) return i + 1;
      return 0;
    };
    sh.appendRow = function (v) { sh.rows.push(v.slice()); };
    sh.setFrozenRows = function () {};
    sh.getRange = function (r, c, nr, nc) {
      if (arguments.length === 2) { nr = 1; nc = 1; }
      return {
        setNumberFormat: function () { return this; },
        getDisplayValues: function () { return read().map(function (row) { return row.map(function (v) { return v == null ? '' : String(v); }); }); },
        getValues: function () { return read(); },
        setValues: function (data) {
          data.forEach(function (line, i) {
            const ri = r - 1 + i;
            if (!sh.rows[ri]) sh.rows[ri] = [];
            line.forEach(function (cell, j) { sh.rows[ri][c - 1 + j] = cell; });
          });
          sets++;
          if (crashAfter != null && sets >= crashAfter) throw new Error('interrupted');
          return this;
        },
        clearContent: function () { return this; }
      };
      function read() {
        const o = [];
        for (let i = 0; i < nr; i++) {
          const row = sh.rows[r - 1 + i] || [], line = [];
          for (let j = 0; j < nc; j++) line.push(row[c - 1 + j] == null ? '' : row[c - 1 + j]);
          o.push(line);
        }
        return o;
      }
    };
    sh.getDataRange = function () { return sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 8); };
    return sh;
  }
  const sheets = {};
  const ss = {
    name: 'TMC Service Sheet',
    getName: function () { return ss.name; },
    rename: function (n) { ss.name = n; },
    getSheetByName: function (n) { return sheets[n] || null; },
    insertSheet: function (n) { sheets[n] = sheet(); return sheets[n]; }
  };
  const sandbox = {
    console: console, Date: Date, Math: Math, JSON: JSON, Object: Object, Array: Array, String: String, Number: Number,
    RegExp: RegExp, Error: Error, isFinite: isFinite, parseInt: parseInt,
    PropertiesService: { getScriptProperties: function () { return {
      getProperty: function (k) { return Object.prototype.hasOwnProperty.call(bag, k) ? bag[k] : null; },
      setProperty: function (k, v) { bag[k] = String(v); },
      deleteProperty: function (k) { delete bag[k]; },
      getProperties: function () { return Object.assign({}, bag); }
    }; } },
    LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, waitLock: function () {}, releaseLock: function () {} }; } },
    Utilities: {
      getUuid: function () { return crypto.randomUUID(); },
      base64Encode: function (b) { return Buffer.from(b).toString('base64'); },
      base64Decode: function (s) { return Buffer.from(String(s), 'base64'); },
      newBlob: function (d) { const b = Buffer.from(d); return { getBytes: function () { return Buffer.from(b); } }; }
    },
    DriveApp: { createFolder: function () { throw new Error('drive'); }, getFolderById: function () { throw new Error('drive'); }, getFileById: function () { throw new Error('drive'); } },
    SpreadsheetApp: { getActive: function () { return ss; }, getActiveSpreadsheet: function () { return ss; }, openById: function () { sandbox._fuel++; return { getSheetByName: function () { return null; } }; } },
    ContentService: { MimeType: { JSON: 'application/json' }, createTextOutput: function (s) { const o = { getContent: function () { return s; }, setMimeType: function () { return o; } }; return o; } },
    ScriptApp: { getProjectTriggers: function () { return triggers.slice(); }, deleteTrigger: function () {}, newTrigger: function () { return { timeBased: function () { const api = { everyDays: function () { return api; }, atHour: function () { return api; }, nearMinute: function () { return api; }, everyMinutes: function () { return api; }, create: function () { return {}; } }; return api; } }; } },
    UrlFetchApp: { fetch: function () { throw new Error('network'); } },
    _bag: bag, _fuel: 0, _ss: ss,
    setCrash: function (n) { crashAfter = n; sets = 0; }
  };
  vm.runInNewContext(src, sandbox, { filename: 'candidate-Code.gs' });
  return sandbox;
}

function seed(ctx, kind, id, data) {
  const sh = ctx.sheet_(ctx.REC, ctx.REC_HEAD);
  const u = Date.now();
  sh.appendRow([kind + ':' + id, kind, id, u, 0, u, 'seed', JSON.stringify(data)]);
}
function user() { return { id: 'uStaff', name: 'Sam Staff', status: 'active', admin: 0, perms: ['view', 'add'], salt: 's', pinH: PIN, dev: 'ph', at: '2026-11-08T00:00:00.000Z' }; }
function oilTransfer(qty) {
  return { action: 'importStockTransfer', code: IMPORT, transfer: {
    id: 'trsf1001', v: 1, sourceApp: 'fi', sourceIssueId: 'trsf1001', destinationApp: 'tmcservice',
    at: '2026-11-08T11:31:00.000Z', by: 'SAMPLE Staff', qtyScale: 3,
    lines: [{ id: 'lineOil001', sourceItemId: 'foodoil88', name: 'SAMPLE Hydraulic oil', unit: 'L', qty: qty }]
  } };
}
function phoneStock(records, pid) {
  const R = {};
  (records || []).forEach(function (r) { let d = r.data; if (typeof d === 'string') { try { d = JSON.parse(d); } catch (e) { d = null; } } R[r.kind + ':' + r.id] = { kind: r.kind, del: r.del, data: d }; });
  const all = function (k) { return Object.keys(R).map(function (x) { return R[x]; }).filter(function (r) { return r.kind === k && !r.del && r.data; }).map(function (r) { return r.data; }); };
  const p = all('pt').find(function (x) { return String(x.id) === String(pid); });
  let q = p ? (+p.qty0 || 0) : 0;
  all('sc').forEach(function (c) { if (String(c.pid) === String(pid)) q += +c.delta || 0; });
  all('flc').forEach(function (l) { if (l && String(l.pid) === String(pid)) q += (+l.qtyTh || 0) / 1000; });
  all('jc').forEach(function (j) { (j.parts || []).forEach(function (l) { if (String(l.pid) === String(pid)) q -= +l.qty || 0; }); });
  return Math.round(q * 1000) / 1000;
}

let ctx;
function boot() { ctx = load(code); seed(ctx, 'u', 'uStaff', user()); seed(ctx, 'pt', 'p100', { id: 'p100', pn: 'P-100', name: 'SAMPLE Hydraulic oil', unit: 'L', qty0: 4, min: 1 }); }

t('MOCK candidate: 4 L + Food 2 L allows a 1 L job and both sides show 5 L', function () {
  boot();
  const imp = ctx.handle_(oilTransfer('2'));
  assert.strictEqual(imp.result, 'accepted');
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
  assert.strictEqual(ctx.handle_(oilTransfer('2')).result, 'duplicate');
  assert.strictEqual(ctx.handle_(oilTransfer('9')).error, 'bad_request');
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
  assert.strictEqual(ctx._bag.STOCK_IMPORT_CODE ? ctx.handle_({ action: 'sync', code: IMPORT, changes: [{ kind: 'u', id: 'uStaff', u: Date.now(), data: { id: 'uStaff', name: 'Hacker' } }] }).error : '', 'bad_request');
  const login = ctx.handle_({ action: 'login', id: 'uStaff', pinH: PIN, device: 'ph' });
  assert.strictEqual(login.ok, true);
  const job = { id: 'job1', fleet: 'TRK1', parts: [{ pid: 'p100', pn: 'P-100', qty: 1 }], at: '2026-11-08T12:00:00.000Z' };
  const sync = ctx.handle_({ token: login.token, device: 'ph', since: 0, changes: [{ kind: 'jc', id: 'job1', u: 5000, del: 0, data: job }] });
  assert.strictEqual(sync.ok, true);
  assert.ok(!sync.refusedStock);
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 5);
  assert.strictEqual(phoneStock(sync.records, 'p100'), 5);
  const again = ctx.handle_({ token: login.token, device: 'ph', since: 0, changes: [{ kind: 'jc', id: 'job1', u: 6000, del: 0, data: job }] });
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 5);
  assert.strictEqual(phoneStock(again.records, 'p100'), 5);
  const too = ctx.handle_({ token: login.token, device: 'ph', since: 0, changes: [{ kind: 'jc', id: 'jobBig', u: 7000, del: 0, data: { id: 'jobBig', fleet: 'TRK1', parts: [{ pid: 'p100', qty: 6 }] } }] });
  assert.ok(too.refusedStock && too.refusedStock.indexOf('jc:jobBig') >= 0);
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 5);
  assert.strictEqual(ctx._fuel, 0);
  const fak = ctx.handle_({ token: login.token, device: 'ph', since: 0, changes: [{ kind: 'flc', id: 'lineFake1', u: 8000, del: 0, data: { id: 'lineFake1', pid: 'p100', qtyTh: 9000 } }] });
  assert.strictEqual(fak.accepted, 0);
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 5);
});

t('MOCK candidate: interrupted Food import recovers without a second 2 L', function () {
  boot();
  ctx.setCrash(3);
  let threw = false;
  try { ctx.handle_(oilTransfer('2')); } catch (e) { threw = e.message === 'interrupted'; }
  assert.strictEqual(threw, true);
  ctx.setCrash(null);
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
  assert.strictEqual(ctx.handle_(oilTransfer('2')).result, 'accepted');
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
  assert.strictEqual(ctx.handle_(oilTransfer('2')).result, 'duplicate');
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
  const pt = JSON.parse(ctx.readAll_().map['pt:p100'].data);
  assert.strictEqual(pt.qty0, 4);
  assert.strictEqual(pt.pn, 'P-100');
});

t('MOCK candidate: Check machine and a unit change are unchanged', function () {
  boot();
  const bad = ctx.handle_({ action: 'linkCall', secret: 'nope-nope-nope-nope', call: { fid: 'fill1', vid: 'TRK1' } });
  assert.strictEqual(bad.error, 'bad_link');
  const ok = ctx.handle_({ action: 'linkCall', secret: LINK, call: { fid: 'fill1', vid: 'TRK1', name: 'Truck', date: '2026-11-08', rate: 12.2, normal: 8 } });
  assert.strictEqual(ok.ok, true);
  assert.ok(ctx.readAll_().map['cm:fill1']);
  const login = ctx.handle_({ action: 'login', id: 'uStaff', pinH: PIN, device: 'ph' });
  const phone = ctx.handle_({ token: login.token, device: 'ph', since: 0, changes: [{ kind: 'cm', id: 'fill9', u: 9000, del: 0, data: { id: 'fill9' } }] });
  assert.strictEqual(phone.accepted, 0);
  assert.ok(!ctx.readAll_().map['cm:fill9']);
  assert.strictEqual(ctx.handle_(oilTransfer('2')).result, 'accepted');
  const unit = oilTransfer('1');
  unit.transfer = JSON.parse(JSON.stringify(unit.transfer));
  unit.transfer.id = 'trsf1002';
  unit.transfer.sourceIssueId = 'trsf1002';
  unit.transfer.lines[0].id = 'lineOil002';
  unit.transfer.lines[0].unit = 'bag';
  assert.strictEqual(ctx.handle_(unit).error, 'bad_request');
  assert.strictEqual(ctx.stockMap_(ctx.readAll_()).p100, 6);
});

const failed = results.filter(function (r) { return r.indexOf('FAIL') === 0; });
results.forEach(function (r) { console.log(r); });
console.log(failed.length ? failed.length + ' failed' : results.filter(function (r) { return r.indexOf('ok') === 0; }).length + ' passed');
process.exit(failed.length ? 1 : 0);
