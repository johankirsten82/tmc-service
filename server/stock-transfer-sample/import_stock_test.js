'use strict';
/**
 * Mock tests for the Service stock-transfer sample.
 * Two receivers, separate stores. No live Sheet, Drive, or URL.
 * Import codes here are sample strings, not a deployment secret.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const gs = fs.readFileSync(path.join(__dirname, 'importStockTransfer.gs'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '../../index.html'), 'utf8');
const sandbox = { module: { exports: {} }, console: console };
vm.runInNewContext(gs, sandbox, { filename: 'importStockTransfer.gs' });
const api = sandbox.module.exports;
const SERVICE = 'sample-service-import';
const WORKSHOP = 'sample-workshop-import';
const WORKSHOP_APP = 'workshop-sample';

function env(opt) {
  opt = opt || {};
  const rows = {};
  const writes = [];
  let locked = false;
  let crashAt = opt.crashAt == null ? null : opt.crashAt;
  const parts = JSON.parse(JSON.stringify(opt.parts || []));
  return {
    appId: opt.appId || api.STOCK_APP_ID,
    code: function () { return opt.code == null ? SERVICE : opt.code; },
    parts: function () { return parts.map(function (p) { return JSON.parse(JSON.stringify(p)); }); },
    tryLock: function () { if (locked) return false; locked = true; return true; },
    unlock: function () { locked = false; },
    locked: function () { return locked; },
    get: function (kind, id) {
      const r = rows[kind + ':' + id];
      return r ? JSON.parse(JSON.stringify(r)) : null;
    },
    list: function (kind) {
      return Object.keys(rows).filter(function (k) { return k.indexOf(kind + ':') === 0; }).map(function (k) { return JSON.parse(JSON.stringify(rows[k])); });
    },
    write: function (kind, id, data) {
      rows[kind + ':' + id] = JSON.parse(JSON.stringify(data));
      writes.push(kind + ':' + id);
      if (crashAt != null && writes.length >= crashAt) throw new Error('interrupted');
    },
    writes: writes,
    rows: rows,
    partsLive: parts,
    setCrash: function (n) { crashAt = n; }
  };
}
function line(id, item, name, unit, qty) {
  return { id: id, sourceItemId: item, name: name, unit: unit, qty: qty };
}
function transfer(over) {
  const t = {
    id: 'trsf1001', v: 1, sourceApp: 'fi', sourceIssueId: 'trsf1001', destinationApp: 'tmcservice',
    at: '2026-11-08T11:31:00.000Z', by: 'SAMPLE Staff', qtyScale: 3,
    lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1.25')]
  };
  return Object.assign(t, over || {});
}
function req(t, code) {
  return { action: 'importStockTransfer', code: code || SERVICE, transfer: t || transfer() };
}
function qty(e, stockId) {
  return api.stockSumTh_(e.list('flc'), stockId);
}
function kinds(e) {
  return Object.keys(e.rows).map(function (k) { return k.split(':')[0]; }).filter(function (k, i, a) { return a.indexOf(k) === i; }).sort();
}

const results = [];
function t(name, fn) {
  try { fn(); results.push('ok  ' + name); }
  catch (err) { results.push('FAIL ' + name + ' — ' + (err && err.stack ? err.stack.split('\n')[0] + ' ' + err.message : err)); }
}

t('source file has no live endpoint or import secret', function () {
  assert.ok(!/script\.google\.com|AKfyc/i.test(gs));
  assert.ok(!gs.includes(SERVICE) && !gs.includes(WORKSHOP));
  assert.ok(gs.includes("var STOCK_APP_ID = 'tmcservice'"));
  assert.ok(!/script\.google\.com|AKfyc|STOCK_IMPORT_CODE|sample-service-import|workshop-sample/.test(html.slice(html.indexOf('SAMPLE food stock'))));
});

t('decimal quantities and part-number preservation', function () {
  const oil = { id: 'p100', pn: 'P-100', name: 'SAMPLE Hydraulic oil', unit: 'L', qty0: 4 };
  const cafe = { id: 'p200', pn: 'P-200', name: 'SAMPLE Café', unit: 'kg', qty0: 1 };
  const e = env({ parts: [oil, cafe] });
  const before = JSON.stringify(e.partsLive);
  const body = transfer({
    lines: [
      line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1'),
      line('lineRice02', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '0.5'),
      line('lineOil001', 'foodoil88', 'sample hydraulic oil', 'L', '1.25'),
      line('lineCafe01', 'foodcafe01', 'sample cafe', 'kg', '1.250')
    ]
  });
  const a = api.stockImportHandle_(e, req(body));
  assert.strictEqual(a.result, 'accepted');
  assert.strictEqual(qty(e, 'foodrice25'), 1500);
  assert.strictEqual(qty(e, 'part_p100'), 1250);
  assert.strictEqual(qty(e, 'part_p200'), 1250);
  const oilLine = e.get('fsk', 'part_p100');
  assert.strictEqual(oilLine.pn, 'P-100');
  assert.strictEqual(oilLine.partId, 'p100');
  assert.strictEqual(e.get('fsk', 'part_p200').pn, 'P-200');
  assert.strictEqual(e.get('fsk', 'foodrice25').pn, '');
  assert.strictEqual(JSON.stringify(e.partsLive), before);
  assert.deepStrictEqual(kinds(e), ['fbd', 'flc', 'fsk', 'ftr']);
});

t('repeated transfer is duplicate and an altered body does not credit again', function () {
  const e = env();
  const first = req(transfer());
  assert.strictEqual(api.stockImportHandle_(e, first).result, 'accepted');
  const n = e.writes.length;
  const again = transfer();
  again.lines = [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1.25')].reverse();
  assert.strictEqual(api.stockImportHandle_(e, req(again)).result, 'duplicate');
  const altered = transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '9.5')] });
  const d = api.stockImportHandle_(e, req(altered));
  assert.strictEqual(d.result, 'duplicate');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
  assert.strictEqual(e.writes.length, n);
  assert.strictEqual(api.stockImportHandle_(e, { action: 'importStockTransfer', code: 'nope', transfer: altered }).error, 'bad_code');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
});

t('altered payload while the first attempt is open writes nothing further', function () {
  const e = env({ crashAt: 1 });
  let threw = false;
  try { api.stockImportHandle_(e, req(transfer())); } catch (err) { threw = err.message === 'interrupted'; }
  assert.strictEqual(threw, true);
  assert.strictEqual(e.locked(), false);
  assert.strictEqual(e.get('ftr', 'trsf1001').status, 'open');
  assert.strictEqual(e.list('flc').length, 0);
  const altered = api.stockImportHandle_(e, req(transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '9')] })));
  assert.strictEqual(altered.error, 'bad_request');
  assert.strictEqual(e.list('flc').length, 0);
  e.setCrash(null);
  assert.strictEqual(api.stockImportHandle_(e, req(transfer())).result, 'accepted');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
  assert.strictEqual(api.stockImportHandle_(e, req(transfer())).result, 'duplicate');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
});

t('crash after the line credit still resumes once when the lock is gone', function () {
  const e = env({ crashAt: 4 });
  let threw = false;
  try { api.stockImportHandle_(e, req(transfer())); } catch (err) { threw = true; }
  assert.strictEqual(threw, true);
  assert.strictEqual(e.locked(), false);
  assert.strictEqual(e.get('flc', 'lineRice01').qtyTh, 1250);
  assert.strictEqual(e.get('ftr', 'trsf1001').status, 'open');
  e.setCrash(null);
  const writes = e.writes.length;
  assert.strictEqual(api.stockImportHandle_(e, req(transfer())).result, 'accepted');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
  assert.ok(e.writes.length > writes);
  assert.strictEqual(e.get('ftr', 'trsf1001').status, 'applied');
  assert.strictEqual(api.stockImportHandle_(e, req(transfer())).result, 'duplicate');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
});

t('two recipients reject the other app and do not share rows', function () {
  const service = env({ appId: 'tmcservice', code: SERVICE });
  const workshop = env({ appId: WORKSHOP_APP, code: WORKSHOP });
  assert.strictEqual(api.stockImportHandle_(service, req(transfer({ destinationApp: WORKSHOP_APP }))).error, 'wrong_app');
  assert.strictEqual(Object.keys(service.rows).length, 0);
  assert.strictEqual(api.stockImportHandle_(workshop, req(transfer(), WORKSHOP)).error, 'wrong_app');
  assert.strictEqual(Object.keys(workshop.rows).length, 0);
  assert.strictEqual(api.stockImportHandle_(service, req(transfer())).result, 'accepted');
  assert.strictEqual(api.stockImportHandle_(workshop, req(transfer({ destinationApp: WORKSHOP_APP }), WORKSHOP)).result, 'accepted');
  assert.strictEqual(qty(service, 'foodrice25'), 1250);
  assert.strictEqual(qty(workshop, 'foodrice25'), 1250);
  assert.notStrictEqual(service.rows['flc:lineRice01'], workshop.rows['flc:lineRice01']);
  assert.strictEqual(service.get('ftr', 'trsf1001').destinationApp, 'tmcservice');
  assert.strictEqual(workshop.get('ftr', 'trsf1001').destinationApp, WORKSHOP_APP);
});

t('ambiguous same-name items write nothing; a different unit does not bind', function () {
  const a = { id: 'p1', pn: 'A-1', name: 'SAMPLE Rice 25kg', unit: 'kg', qty0: 3 };
  const b = { id: 'p2', pn: 'B-2', name: 'sample rice 25kg', unit: 'kg', qty0: 8 };
  const e = env({ parts: [a, b] });
  const snap = JSON.stringify(e.partsLive);
  const amb = api.stockImportHandle_(e, req(transfer()));
  assert.strictEqual(amb.error, 'ambiguous_item');
  assert.strictEqual(Object.keys(e.rows).length, 0);
  assert.strictEqual(JSON.stringify(e.partsLive), snap);
  const other = env({ parts: [a] });
  const before = JSON.stringify(other.partsLive);
  const own = api.stockImportHandle_(other, req(transfer({ lines: [line('lineBag001', 'foodbag001', 'SAMPLE Rice 25kg', 'bag', '2')] })));
  assert.strictEqual(own.result, 'accepted');
  assert.strictEqual(other.get('fsk', 'foodbag001').partId, '');
  assert.strictEqual(other.get('fsk', 'foodbag001').pn, '');
  assert.strictEqual(JSON.stringify(other.partsLive), before);
  assert.ok(!other.get('fsk', 'part_p1'));
});

t('unit change on a bound Food item is refused and the first credit stays', function () {
  const e = env();
  assert.strictEqual(api.stockImportHandle_(e, req(transfer())).result, 'accepted');
  const next = transfer({
    id: 'trsf1002', sourceIssueId: 'trsf1002',
    lines: [line('lineRice99', 'foodrice25', 'Renamed rice', 'bag', '3')]
  });
  assert.strictEqual(api.stockImportHandle_(e, req(next)).error, 'bad_request');
  assert.strictEqual(qty(e, 'foodrice25'), 1250);
  assert.strictEqual(e.list('fsk').length, 1);
  const same = transfer({
    id: 'trsf1003', sourceIssueId: 'trsf1003',
    lines: [line('lineRice77', 'foodrice25', 'Renamed rice', 'kg', '1')]
  });
  assert.strictEqual(api.stockImportHandle_(e, req(same)).result, 'accepted');
  assert.strictEqual(qty(e, 'foodrice25'), 2250);
  assert.strictEqual(e.get('fsk', 'foodrice25').name, 'SAMPLE Rice 25kg');
  assert.strictEqual(e.get('fbd', 'fi_foodrice25').unit, 'kg');
});

t('bad quantities, source identity, and the import code cannot run another action', function () {
  const e = env({ parts: [{ id: 'p100', pn: 'P-100', name: 'Bolt', unit: 'ea', qty0: 4 }] });
  function err(body) { return api.stockImportHandle_(e, req(body)).error; }
  assert.strictEqual(err(transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1.2500')] })), 'bad_qty');
  assert.strictEqual(err(transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '0')] })), 'bad_qty');
  assert.strictEqual(err(transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1e1')] })), 'bad_qty');
  assert.strictEqual(err(transfer({ lines: [line('lineRice01', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '-1')] })), 'bad_qty');
  assert.strictEqual(err(transfer({ sourceApp: 'food' })), 'bad_request');
  assert.strictEqual(err(transfer({ sourceIssueId: 'other1' })), 'bad_request');
  assert.strictEqual(err(transfer({ lines: [line('no', 'foodrice25', 'SAMPLE Rice 25kg', 'kg', '1')] })), 'bad_request');
  const pin = req(transfer());
  pin.transfer.pinH = 'secret';
  assert.strictEqual(api.stockImportHandle_(e, pin).error, 'bad_request');
  assert.strictEqual(api.stockImportHandle_(e, { action: 'sync', code: SERVICE, changes: [{ kind: 'u', id: 'u1' }] }).error, 'bad_request');
  assert.strictEqual(api.stockImportHandle_(e, { action: 'importStockTransfer', code: '', transfer: transfer() }).error, 'bad_code');
  const busy = env();
  busy.tryLock();
  assert.strictEqual(api.stockImportHandle_(busy, req(transfer())).error, 'busy_try_again');
  assert.strictEqual(Object.keys(e.rows).length, 0);
  assert.strictEqual(e.partsLive[0].qty0, 4);
  assert.strictEqual(e.partsLive[0].pn, 'P-100');
});

t('frontend received-stock sample does not change parts math', function () {
  assert.ok(html.includes('function stStock(pid){const p=stPart(pid);let q=p?(+p.qty0||0):0;all("sc").forEach'));
  assert.ok(html.includes('id="foodStock"'));
  assert.ok(html.includes('SAMPLE Rice 25kg'));
  assert.ok(html.includes('SAMPLE — not live stock') || html.includes('foodSample'));
  assert.ok(html.includes('#sample-food-stock'));
  assert.ok(html.includes('if(!window.TMC.foodStockSample)'));
  assert.ok(!html.includes('importStockTransfer'));
});

const failed = results.filter(function (r) { return r.indexOf('FAIL') === 0; });
results.forEach(function (r) { console.log(r); });
console.log(failed.length ? failed.length + ' failed' : results.filter(function (r) { return r.indexOf('ok') === 0; }).length + ' passed');
if (failed.length) failed.forEach(function (r) { console.log(r); });
process.exit(failed.length ? 1 : 0);
