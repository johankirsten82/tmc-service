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
 * Usable stock is the existing total: opening qty0 + corrections − job-card use,
 * plus Food line credits (flc). A receipt does not change qty0 or any old
 * correction. A new Food item is a real part with opening 0, so job cards can
 * select it. The same flc term has to be added to live stockMap_ when this
 * receiver is deployed; v5 still sums only qty0 + corrections − job cards, so
 * it would refuse a job that spends the received litres until that sum is there.
 * Same transfer id with a different body returns bad_request, including after
 * the receipt was applied. The same body returns duplicate.
 */
var STOCK_APP_ID = 'tmcservice';
var STOCK_IMPORT_PROP = 'STOCK_IMPORT_CODE';
var STOCK_SOURCE_APP = 'fi';
var STOCK_QTY_SCALE = 3;
var STOCK_ID = /^[A-Za-z0-9_-]{4,64}$/;
var STOCK_QTY = /^(?:0|[1-9]\d{0,8})(?:\.\d{1,3})?$/;
var STOCK_AT = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,3})?Z$/;
var STOCK_KINDS = { ftr: 1, fbd: 1, flc: 1, pt: 1 };

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
function stockPart_(env, id) {
  var row = env.get('pt', id);
  if (row) return row;
  var parts = env.parts ? env.parts() : [], i;
  for (i = 0; i < parts.length; i++) if (parts[i] && String(parts[i].id) === String(id)) return parts[i];
  return null;
}
function stockItems_(env) {
  var items = [], seen = {};
  function add(p) {
    if (!p || p.id == null || seen[String(p.id)]) return;
    seen[String(p.id)] = 1;
    items.push({ partId: String(p.id), pn: p.pn == null ? '' : String(p.pn), name: p.name == null ? '' : String(p.name), unit: p.unit == null ? '' : String(p.unit) });
  }
  (env.parts ? env.parts() : []).forEach(add);
  env.list('pt').forEach(add);
  return items;
}
function stockNewId_(env, sourceItemId) {
  if (!stockPart_(env, sourceItemId)) return sourceItemId;
  if (!stockPart_(env, 'fi_' + sourceItemId)) return 'fi_' + sourceItemId;
  return 'fi_' + sourceItemId + '_2';
}
function stockPn_(env, sourceItemId) {
  var taken = {};
  stockItems_(env).forEach(function (p) { if (p.pn) taken[p.pn] = 1; });
  if (!taken[sourceItemId]) return sourceItemId;
  var alt = ('fi-' + sourceItemId).slice(0, 60);
  return taken[alt] ? sourceItemId : alt;
}
function stockPlan_(env, t) {
  var i, plan = [], items = stockItems_(env);
  for (i = 0; i < t.lines.length; i++) {
    var l = t.lines[i], th = stockQtyTh_(l.qty), name = String(l.name).trim(), unit = String(l.unit).trim();
    var b = env.get('fbd', 'fi_' + l.sourceItemId), stock = null, created = null;
    if (b) {
      if (String(b.unit).trim() !== unit) return { ok: false, error: 'bad_request' };
      stock = { partId: String(b.partId), pn: b.pn == null ? '' : String(b.pn), name: b.name, unit: String(b.unit).trim() };
    } else {
      var hit = items.filter(function (it) { return stockFold_(it.name) === stockFold_(name) && String(it.unit).trim() === unit; });
      if (hit.length > 1) return { ok: false, error: 'ambiguous_item' };
      if (hit.length === 1) stock = hit[0];
      else {
        var nid = stockNewId_(env, l.sourceItemId);
        created = { id: nid, pn: stockPn_(env, l.sourceItemId), name: name, unit: unit, qty0: 0, min: 0, sourceApp: STOCK_SOURCE_APP, sourceItemId: l.sourceItemId };
        stock = { partId: nid, pn: created.pn, name: name, unit: unit };
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
  var prev = STOCK_ID.test(id) ? env.get('ftr', id) : null;
  if (Object.keys(req).some(function (k) { return k !== 'action' && k !== 'code' && k !== 'transfer'; })) return { ok: false, error: 'bad_request' };
  var shape = stockShape_(t, env.appId || STOCK_APP_ID);
  if (prev && prev.status === 'applied') {
    if (!shape.ok || stockCanon_(t) !== prev.canon) return { ok: false, error: 'bad_request' };
    return { ok: true, transferId: id, result: 'duplicate' };
  }
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
        partId: st.partId, pn: st.pn || '', name: st.name, unit: st.unit
      });
    }
    if (step.created && !env.get('pt', step.created.id)) stockWrite_(env, 'pt', step.created.id, step.created);
    if (!step.have) stockWrite_(env, 'flc', step.line.id, {
      id: step.line.id, transferId: t.id, pid: st.partId, sourceItemId: step.line.sourceItemId,
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
function stockOnHandTh_(env, partId) {
  var p = stockPart_(env, partId), th = p ? Math.round((Number(p.qty0) || 0) * 1000) : 0;
  env.list('sc').forEach(function (c) { if (c && String(c.pid) === String(partId)) th += Math.round((Number(c.delta) || 0) * 1000); });
  env.list('flc').forEach(function (l) { if (l && String(l.pid) === String(partId)) th += l.qtyTh || 0; });
  env.list('jc').forEach(function (j) { (j && j.parts || []).forEach(function (l) { if (String(l.pid) === String(partId)) th -= Math.round((Number(l.qty) || 0) * 1000); }); });
  return th;
}
/** One job-card use. Same job id does not deduct again. Not an import error code. */
function stockConsume_(env, partId, qty, jcId) {
  var need = stockQtyTh_(String(qty));
  if (need == null) return { ok: false, error: 'bad_qty' };
  if (env.get('jc', jcId)) return { ok: true, same: 1, onHandTh: stockOnHandTh_(env, partId) };
  if (stockOnHandTh_(env, partId) < need) return { ok: false, blocked: 'stock', onHandTh: stockOnHandTh_(env, partId) };
  env.write('jc', jcId, { id: jcId, parts: [{ pid: String(partId), qty: need / 1000 }] });
  return { ok: true, onHandTh: stockOnHandTh_(env, partId) };
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
  stockQtyTh_: stockQtyTh_, stockFold_: stockFold_, stockOnHandTh_: stockOnHandTh_, stockConsume_: stockConsume_, STOCK_APP_ID: STOCK_APP_ID
};
