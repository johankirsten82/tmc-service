/**
 * SAMPLE ONLY. Not deployed. Not part of live TMC Service Apps Script v5
 * (photos + Check machine). Do not paste this over that project.
 *
 * Service receiver for Food stock-transfer contract v1.
 * Own import code (script property STOCK_IMPORT_CODE). Own rows in this
 * app's _Records sheet. No Food users, PINs, team code, or database.
 * The phone never sees the import code or the receiver URL.
 *
 * Contract gaps this sample does not paper over with a new error code:
 * - Same transfer id, different body, before the receipt is applied:
 *   v1 has no code for that. This returns bad_request and writes nothing
 *   further. Resend the original body to finish. It does not invent
 *   payload_mismatch.
 * - A later transfer for a bound Food item whose unit string changed:
 *   v1 says the binding still receives a credit after a rename, and does
 *   not say what a different unit means. This returns bad_request and does
 *   not add the quantity onto a line of another unit.
 * - The credit is Service's own received-stock line (kinds fsk / flc / fbd / ftr).
 *   It does not change part qty0, stock corrections, job cards, service due,
 *   or maintenance. A unique name+unit match stores the existing part id and
 *   copies the part number for display. It does not write the part row.
 *   If parts-store on-hand should increase, the contract needs to say so:
 *   that would change job-card stock checks.
 */
var STOCK_APP_ID = 'tmcservice';
var STOCK_IMPORT_PROP = 'STOCK_IMPORT_CODE';
var STOCK_SOURCE_APP = 'fi';
var STOCK_QTY_SCALE = 3;
var STOCK_ID = /^[A-Za-z0-9_-]{4,64}$/;
var STOCK_QTY = /^(?:0|[1-9]\d{0,8})(?:\.\d{1,3})?$/;
var STOCK_AT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/;
var STOCK_KINDS = { ftr: 1, fbd: 1, fsk: 1, flc: 1 };

function stockFold_(s) {
  return String(s || '').trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}
function stockQtyTh_(s) {
  if (typeof s !== 'string' || !STOCK_QTY.test(s)) return null;
  var p = s.split('.'), whole = parseInt(p[0], 10), frac = ((p[1] || '') + '000').slice(0, 3);
  var th = whole * 1000 + parseInt(frac, 10);
  if (!(th > 0)) return null;
  return th;
}
function stockAllow_(obj, keys) {
  var i, ok = {};
  for (i = 0; i < keys.length; i++) ok[keys[i]] = 1;
  return Object.keys(obj).every(function (k) { return ok[k]; });
}
function stockCanon_(t) {
  var lines = t.lines.map(function (l) {
    return { id: l.id, sourceItemId: l.sourceItemId, name: String(l.name).trim(), unit: String(l.unit).trim(), qty: l.qty };
  });
  lines.sort(function (a, b) { return a.id < b.id ? -1 : a.id > b.id ? 1 : 0; });
  return JSON.stringify({
    v: t.v, sourceApp: t.sourceApp, sourceIssueId: t.sourceIssueId, destinationApp: t.destinationApp,
    at: t.at, by: String(t.by).trim(), qtyScale: t.qtyScale, lines: lines
  });
}
function stockItems_(env) {
  var bound = {}, items = [], parts = env.parts ? env.parts() : [];
  env.list('fsk').forEach(function (s) { if (s && s.partId) bound[String(s.partId)] = 1; });
  (parts || []).forEach(function (p) {
    if (!p || p.id == null || bound[String(p.id)]) return;
    items.push({ stockId: 'part_' + p.id, partId: String(p.id), pn: p.pn == null ? '' : String(p.pn), name: p.name == null ? '' : String(p.name), unit: p.unit == null ? '' : String(p.unit), from: 'part' });
  });
  env.list('fsk').forEach(function (s) {
    if (!s) return;
    items.push({ stockId: String(s.id), partId: s.partId ? String(s.partId) : '', pn: s.pn == null ? '' : String(s.pn), name: s.name == null ? '' : String(s.name), unit: s.unit == null ? '' : String(s.unit), from: 'fsk' });
  });
  return items;
}
function stockPlan_(env, t) {
  var i, plan = [], items = stockItems_(env);
  for (i = 0; i < t.lines.length; i++) {
    var l = t.lines[i], th = stockQtyTh_(l.qty), name = String(l.name).trim(), unit = String(l.unit).trim();
    var b = env.get('fbd', 'fi_' + l.sourceItemId), stock = null, created = null;
    if (b) {
      if (String(b.unit).trim() !== unit) return { ok: false, error: 'bad_request' };
      stock = { stockId: String(b.stockId), partId: b.partId ? String(b.partId) : '', pn: b.pn == null ? '' : String(b.pn), name: b.name, unit: b.unit };
    } else {
      var hit = items.filter(function (it) { return stockFold_(it.name) === stockFold_(name) && String(it.unit).trim() === unit; });
      if (hit.length > 1) return { ok: false, error: 'ambiguous_item' };
      if (hit.length === 1) stock = hit[0];
      else {
        created = { id: l.sourceItemId, name: name, unit: unit, pn: '', partId: '', sourceApp: STOCK_SOURCE_APP, sourceItemId: l.sourceItemId };
        stock = { stockId: created.id, partId: '', pn: '', name: name, unit: unit, from: 'new' };
        items.push(stock);
      }
    }
    var old = env.get('flc', l.id);
    if (old && (String(old.transferId) !== String(t.id) || old.qtyTh !== th)) return { ok: false, error: 'bad_request' };
    plan.push({ line: l, th: th, name: name, unit: unit, stock: stock, created: created, have: !!old, bound: !!b });
  }
  return { ok: true, plan: plan };
}
function stockShape_(t, appId) {
  if (!t || typeof t !== 'object' || Array.isArray(t)) return { ok: false, error: 'bad_request' };
  if (!stockAllow_(t, ['id', 'v', 'sourceApp', 'sourceIssueId', 'destinationApp', 'at', 'by', 'qtyScale', 'lines'])) return { ok: false, error: 'bad_request' };
  if (typeof t.id !== 'string' || !STOCK_ID.test(t.id)) return { ok: false, error: 'bad_request' };
  if (t.v !== 1 || t.sourceApp !== STOCK_SOURCE_APP || t.sourceIssueId !== t.id) return { ok: false, error: 'bad_request' };
  if (typeof t.destinationApp !== 'string' || !t.destinationApp) return { ok: false, error: 'bad_request' };
  if (t.destinationApp !== appId) return { ok: false, error: 'wrong_app' };
  if (typeof t.at !== 'string' || !STOCK_AT.test(t.at)) return { ok: false, error: 'bad_request' };
  if (typeof t.by !== 'string' || !String(t.by).trim() || String(t.by).length > 120) return { ok: false, error: 'bad_request' };
  if (t.qtyScale !== STOCK_QTY_SCALE) return { ok: false, error: 'bad_request' };
  if (!Array.isArray(t.lines) || !t.lines.length || t.lines.length > 40) return { ok: false, error: 'bad_request' };
  var i, seen = {};
  for (i = 0; i < t.lines.length; i++) {
    var l = t.lines[i];
    if (!l || typeof l !== 'object' || Array.isArray(l)) return { ok: false, error: 'bad_request' };
    if (!stockAllow_(l, ['id', 'sourceItemId', 'name', 'unit', 'qty'])) return { ok: false, error: 'bad_request' };
    if (typeof l.id !== 'string' || !STOCK_ID.test(l.id) || seen[l.id]) return { ok: false, error: 'bad_request' };
    if (typeof l.sourceItemId !== 'string' || !STOCK_ID.test(l.sourceItemId)) return { ok: false, error: 'bad_request' };
    if (typeof l.name !== 'string' || l.name.length > 120 || !String(l.name).trim()) return { ok: false, error: 'bad_request' };
    if (typeof l.unit !== 'string' || l.unit.length > 20 || !String(l.unit).trim()) return { ok: false, error: 'bad_request' };
    if (stockQtyTh_(l.qty) == null) return { ok: false, error: 'bad_qty' };
    seen[l.id] = 1;
  }
  return { ok: true };
}
function stockWrite_(env, kind, id, data) { env.write(kind, id, data); }
function credit_(env, req) {
  var t = req && req.transfer;
  var id = t && typeof t.id === 'string' ? t.id : '';
  if (STOCK_ID.test(id)) {
    var prev = env.get('ftr', id);
    if (prev && prev.status === 'applied') return { ok: true, transferId: id, result: 'duplicate' };
  }
  if (Object.keys(req).some(function (k) { return k !== 'action' && k !== 'code' && k !== 'transfer'; })) return { ok: false, error: 'bad_request' };
  var shape = stockShape_(t, env.appId || STOCK_APP_ID);
  if (!shape.ok) return shape;
  var canon = stockCanon_(t);
  var open = env.get('ftr', t.id);
  if (open && open.canon !== canon) return { ok: false, error: 'bad_request' };
  var planned = stockPlan_(env, t);
  if (!planned.ok) return planned;
  if (!open) stockWrite_(env, 'ftr', t.id, {
    id: t.id, canon: canon, status: 'open', lineIds: t.lines.map(function (l) { return l.id; }),
    at: t.at, by: String(t.by).trim(), sourceApp: t.sourceApp, sourceIssueId: t.sourceIssueId, destinationApp: t.destinationApp
  });
  planned.plan.forEach(function (step) {
    var st = step.stock;
    if (!step.bound && !env.get('fbd', 'fi_' + step.line.sourceItemId)) {
      stockWrite_(env, 'fbd', 'fi_' + step.line.sourceItemId, {
        id: 'fi_' + step.line.sourceItemId, sourceApp: STOCK_SOURCE_APP, sourceItemId: step.line.sourceItemId,
        stockId: st.stockId, partId: st.partId || '', pn: st.pn || '', name: st.name, unit: st.unit
      });
    }
    if (step.created && !env.get('fsk', step.created.id)) stockWrite_(env, 'fsk', step.created.id, step.created);
    else if (st.from === 'part' && !env.get('fsk', st.stockId)) {
      stockWrite_(env, 'fsk', st.stockId, {
        id: st.stockId, name: st.name, unit: String(st.unit).trim(), pn: st.pn || '', partId: st.partId,
        sourceApp: STOCK_SOURCE_APP, sourceItemId: step.line.sourceItemId
      });
    }
    if (!step.have) stockWrite_(env, 'flc', step.line.id, {
      id: step.line.id, transferId: t.id, stockId: st.stockId, sourceItemId: step.line.sourceItemId,
      sourceApp: STOCK_SOURCE_APP, qtyTh: step.th, name: step.name, unit: step.unit
    });
  });
  var done = env.get('ftr', t.id) || {};
  done.status = 'applied';
  stockWrite_(env, 'ftr', t.id, done);
  return { ok: true, transferId: t.id, result: 'accepted' };
}
function importStockTransfer_(env, req) {
  if (!req || req.action !== 'importStockTransfer') return { ok: false, error: 'bad_request' };
  var want = String(env && env.code ? env.code() : '');
  if (!want || String(req.code || '') !== want) return { ok: false, error: 'bad_code' };
  if (!env.tryLock()) return { ok: false, error: 'busy_try_again' };
  try { return credit_(env, req); }
  finally { env.unlock(); }
}
/** This code path accepts only importStockTransfer. It cannot sync or edit users. */
function stockImportHandle_(env, req) {
  if (!req || typeof req !== 'object' || Array.isArray(req)) return { ok: false, error: 'bad_request' };
  if (req.action !== 'importStockTransfer') return { ok: false, error: 'bad_request' };
  return importStockTransfer_(env, req);
}
function stockSumTh_(lines, stockId) {
  var n = 0;
  (lines || []).forEach(function (l) { if (l && String(l.stockId) === String(stockId)) n += l.qtyTh || 0; });
  return n;
}

/* Sheet glue for a later approved deploy into THIS app's script. Not wired to live v5. */
function stockSheetEnv_() {
  var REC = '_Records', HEAD = ['Key', 'Kind', 'Id', 'UpdatedAt', 'Deleted', 'ServerTime', 'Device', 'Json'];
  function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
  function sh() {
    var s = ss(), t = s.getSheetByName(REC) || s.insertSheet(REC);
    if (t.getLastRow() === 0) { t.appendRow(HEAD); t.setFrozenRows(1); }
    return t;
  }
  function read() {
    var t = sh(), n = t.getLastRow(), map = {};
    if (n > 1) t.getRange(2, 1, n - 1, HEAD.length).getDisplayValues().forEach(function (r, i) {
      if (String(r[4]) === '1') return;
      map[String(r[0])] = { kind: String(r[1]), id: String(r[2]), data: String(r[7]), row: i + 2 };
    });
    return { t: t, map: map };
  }
  var lock = LockService.getScriptLock();
  return {
    appId: STOCK_APP_ID,
    code: function () { return String(PropertiesService.getScriptProperties().getProperty(STOCK_IMPORT_PROP) || ''); },
    parts: function () {
      var store = read(), out = [];
      Object.keys(store.map).forEach(function (k) {
        var r = store.map[k]; if (r.kind !== 'pt') return;
        try { var d = JSON.parse(r.data); if (d) out.push(d); } catch (e) {}
      });
      return out;
    },
    tryLock: function () { return lock.tryLock(20000); },
    unlock: function () { try { lock.releaseLock(); } catch (e) {} },
    get: function (kind, id) {
      var r = read().map[kind + ':' + id]; if (!r) return null;
      try { return JSON.parse(r.data); } catch (e) { return null; }
    },
    list: function (kind) {
      var store = read(), out = [];
      Object.keys(store.map).forEach(function (k) {
        var r = store.map[k]; if (r.kind !== kind) return;
        try { var d = JSON.parse(r.data); if (d) out.push(d); } catch (e) {}
      });
      return out;
    },
    write: function (kind, id, data) {
      if (!STOCK_KINDS[kind]) return;
      var store = read(), key = kind + ':' + id, now = Date.now();
      var row = [key, kind, id, now, 0, now, 'stock-import', JSON.stringify(data)];
      var old = store.map[key];
      if (old) store.t.getRange(old.row, 1, 1, row.length).setNumberFormat('@').setValues([row]);
      else store.t.getRange(store.t.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    }
  };
}
function importStockTransfer(req) { return stockImportHandle_(stockSheetEnv_(), req); }
if (typeof module !== 'undefined') module.exports = {
  importStockTransfer_: importStockTransfer_, stockImportHandle_: stockImportHandle_,
  stockQtyTh_: stockQtyTh_, stockFold_: stockFold_, stockSumTh_: stockSumTh_, STOCK_APP_ID: STOCK_APP_ID
};
