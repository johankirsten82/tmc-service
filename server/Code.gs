/**
 * TMC BASE — Google Sheet sync (generic: users + records). One Sheet + one script PER APP.  (base VER 4)
 * Setup:
 *  1) Make a new Google Sheet for this app → Extensions → Apps Script → paste this file.
 *  2) OPTIONAL: Project Settings → Script properties → TEAM_CODE = <old-style sync code>.
 *     Not needed any more: phones without a code sign in with name + PIN and get a session token.
 *     Keep it only if older app versions still send a code. (Never write a code in this file.)
 *  2b) EASIEST: run setupOnce() once from the editor (Run → Allow). It names the Sheet, makes the tabs,
 *     adds the first Admin (JA Kirsten, no PIN yet), installs TRIGGERS and, only if TEAM_CODE_SETUP
 *     was replaced, stores TEAM_CODE in Script Properties. Safe to run again (changes nothing twice).
 *  2c) SAFE SIGN-IN ONLY: Script property STRICT_TOKEN = 1 → the old team code may no longer change
 *     users ("u" records); those changes need a signed-in token. Default off (old phones keep working).
 *  3) Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 *     Put the /exec link in the app's CONFIG SYNC_URL (public, safe). Leave SYNC_CODE empty.
 *     After ANY change to this file: Deploy → Manage deployments → Edit → Version: NEW version → Deploy
 *     (the /exec link stays the same; without a new version the old code keeps running).
 * Open actions (no code, no token; script lock + per-name / per-hour caps):
 *   status {name}               → {ok,state:none|pending|active|rejected|blocked|setup,id,name,salt,hasPin}
 *   register {name,device}      → creates a PENDING user (Admin must approve) → status
 *   login {id,pinH,device}      → pinH = SHA256(salt+":"+pin) computed on the phone → {ok,token}; 5 wrong / hour / name
 *   setpin {id|setup,name,salt,pinH,device} → first PIN of an approved user (or very first Admin) → {ok,token}
 * Protected: sync {token | code, device, since, changes} → {ok,now,accepted,records}. Last write (u) wins.
 *   With a token, only an Admin may change other users ("u" records); a staff token may only
 *   change its own user record without touching status/admin.
 * ADMIN FIX RULE: staff may only ADD new records. Changing or deleting an existing record (any kind
 *   except their own "u" record) is refused unless the token belongs to an Admin. Old team-code phones
 *   have no name, so they are treated as staff (add only). Every Admin change is written to the
 *   "Fix log" tab: Time, Admin, Kind, Id, Old value, New value.
 * PERMISSIONS: user record .perms = ticked ids (missing = PERM_DEFAULT). Staff need 'add' to add a kind,
 *   'view' to receive it, 'reports'/'export' for those actions. Only an Admin can change .perms.
 * Kind "u" = users {id,name,status:pending|active|rejected|blocked,admin,salt,pinH,dev,at}.
 * Approval is stored on the user record (by name), NOT bound to a device, so a reinstalled phone
 * signs in again with name + PIN and no new approval.
 */
var ADMIN_NAME = 'JA Kirsten';
var ADMIN_ALIASES = ['jakirsten', 'johankirsten', 'johanakirsten', 'jkirsten'];
var MAX_FAIL = 5, FAIL_WIN = 3600000, MAX_REG_HOUR = 30;
/* PERMISSIONS: keep these the SAME as CONFIG.PERMS / PERM_DEFAULT in index.html.
   PERM_KINDS: record kind → which tick is needed to see it (view) and to add it (add).
   Kinds not listed are open to every approved user. Fix + approving users are always Admin-only. */
var PERMS = ['view', 'add', 'reports', 'export', 'store', 'storeEd'];
var PERM_DEFAULT = ['view', 'add'];
/* a list = any ONE of these ticks. v7: pt = part, sc = Admin stock correction (the log), jc = job card */
var PERM_KINDS = { x: { view: 'view', add: 'add' }, s: { view: 'view', add: 'add' }, fv: { view: 'view', add: 'admin' },
  pt: { view: ['store', 'storeEd', 'add'], add: 'storeEd' }, sc: { view: ['store', 'storeEd', 'add'], add: 'admin' },
  jc: { view: ['view', 'store', 'storeEd', 'add'], add: 'add' },
  m: { view: ['view', 'store', 'storeEd', 'add'], add: 'add' } };
var REC = '_Records';
var REC_HEAD = ['Key', 'Kind', 'Id', 'UpdatedAt', 'Deleted', 'ServerTime', 'Device', 'Json'];

function teamCode_() { return String(PropertiesService.getScriptProperties().getProperty('TEAM_CODE') || '').trim(); }
function strict_() { return String(props_().getProperty('STRICT_TOKEN') || '').trim() === '1'; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}
function readAll_() {
  var sh = sheet_(REC, REC_HEAD), n = sh.getLastRow(), map = {}, rows = [];
  if (n > 1) rows = sh.getRange(2, 1, n - 1, REC_HEAD.length).getDisplayValues();  /* text as typed: never Sheets' Date objects */
  rows.forEach(function (r, i) {
    var key = String(r[0]), kind = String(r[1]), id = key.indexOf(kind + ':') === 0 ? key.slice(kind.length + 1) : String(r[2]);  /* Id from Key: exact even if Sheets once turned '2099-01-01' into a date */
    map[key] = { key: key, kind: kind, id: id, u: Number(r[3]), del: Number(r[4]) ? 1 : 0, s: Number(r[5]), dev: String(r[6]), data: String(r[7]), row: i + 2 };
  });
  return { sh: sh, map: map };
}
function merge_(store, changes, device, now) {
  var acc = 0;
  (changes || []).forEach(function (c) {
    if (!c || !c.kind || c.id === undefined || c.id === null) return;
    /* users are never removed by a sync/restore; only an explicit purge:1 delete (never sent by the app) removes one */
    if (String(c.kind) === 'u' && !(c.del && (c.purge === 1 || c._discard === 1)) && (c.del || c.data == null || typeof c.data !== 'object')) return;  /* v9: _discard set only by discard_() */
    var key = c.kind + ':' + c.id, old = store.map[key], u = Number(c.u) || 0;
    if (old && old.u >= u) return;
    store.map[key] = { key: key, kind: String(c.kind), id: String(c.id), u: u, del: c.del ? 1 : 0, s: now, dev: device || '',
      data: c.del ? (c._discard === 1 ? JSON.stringify({ discarded: 1, at: new Date(now).toISOString() }) : '') : JSON.stringify(c.data == null ? null : c.data), row: old ? old.row : 0, dirty: true };
    acc++;
  });
  return acc;
}
function writeRecs_(store) {
  var sh = store.sh, add = [];
  Object.keys(store.map).forEach(function (k) {
    var r = store.map[k]; if (!r.dirty) return;
    var v = [r.key, r.kind, r.id, r.u, r.del, r.s, r.dev, r.data];
    if (r.row) sh.getRange(r.row, 1, 1, v.length).setNumberFormat('@').setValues([v]); else add.push(v);
    delete r.dirty;
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, REC_HEAD.length).setNumberFormat('@').setValues(add);  /* plain text: Ids like 2099-01-01 stay text */
}
/* NEVER send PIN secrets in a sync: only the signed-in phone's own record keeps them (viewerId); team code gets none */
var SECRET_U = ['pinH', 'salt', 'pin'];
function stripUser_(data, viewerId) {
  try { var d = JSON.parse(data); if (!d || typeof d !== 'object') return data;
    if (viewerId && d.id === viewerId) return data;
    d.hasPin = d.pinH ? 1 : 0; SECRET_U.forEach(function (f) { delete d[f]; }); return JSON.stringify(d);
  } catch (e) { return ''; }
}
function changedSince_(store, since, me, viewerId) {
  return Object.keys(store.map).map(function (k) { return store.map[k]; }).filter(function (r) { return r.s > since && canKind_(me, r.kind, 'view'); })
    .map(function (r) {
      var data = r.data;
      if (r.kind === 'u' && data) data = stripUser_(data, viewerId);
      return { kind: r.kind, id: r.id, u: r.u, del: r.del, data: data };
    });
}
/* token phones don't hold other users' PIN hashes: keep the stored one unless the Admin asked for a PIN reset */
function keepPins_(store, changes) {
  return (changes || []).map(function (c) {
    if (!c || c.kind !== 'u' || c.del || !c.data || c.data.pinH) return c;
    var old = userById_(store, c.id); if (!old || !old.pinH || c.data.pinReset) return c;
    var d = JSON.parse(JSON.stringify(c.data)); d.pinH = old.pinH; d.salt = old.salt; delete d.hasPin;
    return { kind: c.kind, id: c.id, u: c.u, del: c.del, data: d };
  });
}
/** Readable "Users" tab (the app never reads it). Apps add their own tabs in rebuildViews_. */
function rebuildViews_(store) {
  var rows = [];
  Object.keys(store.map).forEach(function (k) {
    var r = store.map[k]; if (r.kind !== 'u' || r.del) return;
    var d = {}; try { d = JSON.parse(r.data) || {}; } catch (e) {}
    rows.push([d.name || '', d.status || '', d.admin ? 'yes' : '', d.pinH ? 'yes' : 'no', d.at || '', d.dev || '', d.admin ? 'ALL (Admin)' : perms_(d).join(', ')]);
  });
  rows.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])); });
  var sh = sheet_('Users', ['Name', 'Status', 'Admin', 'PIN set', 'Requested', 'Device', 'Permissions']);
  if (sh.getRange(1, 7).getValues()[0][0] !== 'Permissions') sh.getRange(1, 1, 1, 7).setValues([['Name', 'Status', 'Admin', 'PIN set', 'Requested', 'Device', 'Permissions']]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 7).setValues(rows);
  /* APP VIEWS GO HERE (e.g. one readable tab per record kind). */
  storeViews_(store);
}
/* ---- v7 PARTS STORE + JOB CARDS (Service's own records only; no link to any other app) ---- */
function recs_(store, kind) { var o = []; Object.keys(store.map).forEach(function (k) { var r = store.map[k]; if (r.kind !== kind || r.del) return; var d = null; try { d = JSON.parse(r.data); } catch (e) {} if (d) o.push(d); }); return o; }
function r3_(n) { return Math.round(n * 1000) / 1000; }
/* stock = qty when added + Admin corrections − parts used on job cards */
function stockMap_(store) {
  var q = {};
  recs_(store, 'pt').forEach(function (p) { q[String(p.id)] = (q[String(p.id)] || 0) + (Number(p.qty0) || 0); });
  recs_(store, 'sc').forEach(function (c) { q[String(c.pid)] = (q[String(c.pid)] || 0) + (Number(c.delta) || 0); });
  recs_(store, 'jc').forEach(function (j) { (j.parts || []).forEach(function (l) { q[String(l.pid)] = (q[String(l.pid)] || 0) - (Number(l.qty) || 0); }); });
  return q;
}
/* stock corrections: Admin token only (never staff, never the old team code), with a reason.
   job cards: a staff/team-code card that would take stock below zero is refused (with its linked service record);
   only an Admin token may save with override. */
/* v8: a part's vehicles = list of Fuel fleet numbers, or ['*'] = All vehicles. Missing/empty (parts made before v8) = All vehicles. */
function ptVeh_(v) { if (!Array.isArray(v)) return undefined; var o = [], seen = {};
  v.forEach(function (x) { x = String(x == null ? '' : x).trim().slice(0, 30); if (x && !seen[x] && o.length < 100) { seen[x] = 1; o.push(x); } });
  return o.indexOf('*') >= 0 ? ['*'] : o; }
function ptVehTxt_(v) { v = ptVeh_(v); return !v || !v.length || v[0] === '*' ? 'All vehicles' : v.join(', '); }
function storeRules_(store, me, changes, refused) {
  var isAdm = !!(me && me.admin), q = stockMap_(store), dropS = {};
  (changes || []).forEach(function (c) { if (c && c.kind === 'pt' && !c.del && c.data && typeof c.data === 'object' && 'veh' in c.data) { var v = ptVeh_(c.data.veh); if (v === undefined) delete c.data.veh; else c.data.veh = v; } });
  changes = (changes || []).filter(function (c) {
    if (!c || c.kind !== 'sc' || store.map['sc:' + c.id]) return true;
    if (!isAdm || c.del || !c.data || !String(c.data.reason || '').trim()) { refused.n++; return false; }
    return true;
  });
  changes.forEach(function (c) {  /* new parts/corrections in this batch count first */
    if (!c || c.del || !c.data || store.map[c.kind + ':' + c.id]) return;
    if (c.kind === 'pt') q[String(c.id)] = (q[String(c.id)] || 0) + (Number(c.data.qty0) || 0);
    if (c.kind === 'sc') q[String(c.data.pid)] = (q[String(c.data.pid)] || 0) + (Number(c.data.delta) || 0);
  });
  changes = changes.filter(function (c) {
    if (!c || c.kind !== 'jc' || c.del || !c.data || store.map['jc:' + c.id]) return true;
    var need = {}; (c.data.parts || []).forEach(function (l) { need[String(l.pid)] = (need[String(l.pid)] || 0) + (Number(l.qty) || 0); });
    var neg = Object.keys(need).some(function (pid) { return r3_((q[pid] || 0) - need[pid]) < 0; });
    if (neg && !isAdm) { refused.stock = refused.stock || []; refused.stock.push('jc:' + c.id); if (c.data.sid) { dropS[String(c.data.sid)] = 1; refused.stock.push('s:' + c.data.sid); } return false; }
    Object.keys(need).forEach(function (pid) { q[pid] = r3_((q[pid] || 0) - need[pid]); });
    return true;
  });
  return changes.filter(function (c) { return !(c && c.kind === 's' && dropS[String(c.id)] && !store.map['s:' + c.id]); });
}
function viewTab_(name, head, rows) {
  var sh = sheet_(name, head), w = head.length;
  sh.getRange(1, 1, 1, w).setValues([head]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, w).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, w).setNumberFormat('@').setValues(rows);
}
function chadDT_(iso) { var d = new Date(iso); return isNaN(d) ? String(iso || '') : new Date(d.getTime() + CHAD_OFFSET_MS_).toISOString().slice(0, 16).replace('T', ' '); }
function storeViews_(store) {
  var q = stockMap_(store), P = recs_(store, 'pt'), S = recs_(store, 'sc'), J = recs_(store, 'jc');
  if (!P.length && !S.length && !J.length) return;
  P.sort(function (a, b) { return String(a.pn).localeCompare(String(b.pn)); });
  viewTab_('Parts', ['Part no', 'Name', 'Vehicles', 'Unit', 'In stock', 'Minimum', 'To order'], P.map(function (p) { var s = r3_(q[String(p.id)] || 0), m = Number(p.min) || 0;
    return [String(p.pn || ''), String(p.name || ''), ptVehTxt_(p.veh), String(p.unit || ''), String(s), String(m), s < m ? String(r3_(m - s)) : '']; }));
  S.sort(function (a, b) { return String(b.at).localeCompare(String(a.at)); });
  viewTab_('Stock log', ['Time (Chad)', 'Part no', 'Name', 'Was', 'Now', 'Change', 'Reason', 'By'], S.map(function (c) {
    return [chadDT_(c.at), String(c.pn || ''), String(c.name || ''), String(c.old), String(c.new), String(c.delta), String(c.reason || ''), String(c.by || '')]; }));
  J.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
  viewTab_('Job cards', ['Job no', 'Date', 'Fleet', 'Name', 'Meter', 'Work done', 'Done by', 'Parts used', 'Next due', 'Admin override'], J.map(function (j) {
    return [String(j.no || j.id), String(j.date || ''), String(j.fleet || ''), String(j.name || ''), j.meter === '' || j.meter == null ? '' : String(j.meter) + ' ' + String(j.mu || ''),
      String(j.work || ''), String(j.by || ''), (j.parts || []).map(function (l) { return l.pn + ' x' + l.qty; }).join(', '), String(j.due == null ? '' : j.due), j.override ? 'yes (' + (j.override.by || '') + ')' : '']; }));
}
function props_() { return PropertiesService.getScriptProperties(); }
function nk_(x) { var n = String(x || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[\s.]+/g, ''); return ADMIN_ALIASES.indexOf(n) >= 0 ? ADMIN_ALIASES[0] : n; }
function clean_(x) { return String(x || '').trim().replace(/\s+/g, ' ').slice(0, 60); }
function users_(store) {
  var o = [];
  Object.keys(store.map).forEach(function (k) { var r = store.map[k]; if (r.kind !== 'u' || r.del) return; var d = null; try { d = JSON.parse(r.data); } catch (e) {} if (d && d.id) o.push(d); });
  return o;
}
function userById_(store, id) { return users_(store).filter(function (u) { return u.id === String(id || ''); })[0] || null; }
function putUser_(store, u, device) { merge_(store, [{ kind: 'u', id: u.id, u: Math.max(Date.now(), ((store.map['u:' + u.id] || {}).u || 0) + 1), del: 0, data: u }], device, Date.now()); writeRecs_(store); try { rebuildViews_(store); } catch (e) {} }
/* same name rules as the app: lowercase, no spaces/dots, Admin aliases; one approved match only */
function status_(store, raw, device) {
  var name = clean_(raw), K = nk_(name); if (!K) return { ok: false, error: 'need_name' };
  var U = users_(store), all = U.filter(function (u) { return nk_(u.name) === K; }), ok = all.filter(function (u) { return u.status === 'active'; });
  if (ok.length === 1) return { ok: true, state: 'active', id: ok[0].id, name: ok[0].name, salt: ok[0].pinH ? ok[0].salt : '', hasPin: !!ok[0].pinH, admin: !!ok[0].admin };
  if (ok.length === 0) {
    if (all.some(function (u) { return u.status === 'blocked'; })) return { ok: true, state: 'blocked' };
    if (all.some(function (u) { return u.status === 'rejected'; })) return { ok: true, state: 'rejected' };
    var hasAdmin = U.some(function (u) { return u.status === 'active' && u.admin; });
    if (!hasAdmin && K === ADMIN_ALIASES[0]) return { ok: true, state: 'setup', name: ADMIN_NAME };
  }
  var own = U.filter(function (u) { return u.status === 'pending' && nk_(u.name) === K && (ok.length < 2 || (u.dev === device && clean_(u.name).toLowerCase() === name.toLowerCase())); })[0];
  return own ? { ok: true, state: 'pending', id: own.id, name: own.name, other: !!(device && own.dev && own.dev !== device) } : { ok: true, state: 'none' };
}
function fails_(K) { var p = props_(), now = Date.now(), a = []; try { a = JSON.parse(p.getProperty('FAIL_' + K) || '[]'); } catch (e) {} return a.filter(function (t) { return now - t < FAIL_WIN; }); }
function newToken_(u, device) {
  var p = props_(), all = p.getProperties(), tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  Object.keys(all).forEach(function (k) { if (k.indexOf('TOK_') !== 0) return; try { var v = JSON.parse(all[k]); if (v.uid === u.id && v.dev === String(device || '')) p.deleteProperty(k); } catch (e) {} });
  p.setProperty('TOK_' + tok, JSON.stringify({ uid: u.id, dev: String(device || ''), at: Date.now() }));
  return tok;
}
function session_(store, tok) {
  if (!tok || !/^[a-f0-9]{32,80}$/i.test(String(tok))) return null;
  var v = props_().getProperty('TOK_' + tok); if (!v) return null;
  var s; try { s = JSON.parse(v); } catch (e) { return null; }
  var u = userById_(store, s.uid);
  return u && u.status === 'active' ? u : { blocked: true };
}
/* ONE request per person per app: a pending or active user with the same name key already exists */
function dupReq_(store, name, exceptId) { var K = nk_(clean_(name)), U = users_(store);
  if (U.filter(function (u) { return nk_(u.name) === K && u.status === 'active'; }).length >= 2) return false;  /* 2+ approved share the name: a new request is how a 3rd person asks */
  return !!K && U.some(function (u) { return String(u.id) !== String(exceptId || '') && nk_(u.name) === K && (u.status === 'pending' || u.status === 'active'); }); }
function noDupReq_(store, changes, refused) {
  return (changes || []).filter(function (c) {
    if (!c || c.kind !== 'u' || c.del || !c.data || userById_(store, c.id) || c.data.status !== 'pending') return true;
    if (dupReq_(store, c.data.name, c.id)) { refused.dup = (refused.dup || 0) + 1; return false; } return true;
  });
}
/* SERVER CHANGE Service v9 (= base VER 7 remove-on-block): REMOVE a user = Admin token only, never self, never the last Admin.
   That user's session tokens (TOK_*) are deleted so the old PIN session stops working. Team code / staff: refused. */
function discard_(store, me, changes, refused) {
  return (changes || []).filter(function (c) {
    if (!c || c.kind !== 'u') return true;
    if (c._discard !== undefined) delete c._discard;  /* phones can never set the internal flag themselves */
    if (!c.del || !c.data || c.data.discarded !== 1) return true;
    var u = userById_(store, c.id); if (!u) return false;
    var why = !me || !me.admin ? 'admin_only_fix' : String(u.id) === String(me.id) ? 'self' :
      (u.admin && !users_(store).some(function (x) { return x.id !== u.id && x.admin && x.status === 'active'; })) ? 'last_admin' : isMain_(u) ? 'main_admin' : '';
    if (why) { refused.disc = refused.disc || []; refused.disc.push(String(c.id)); refused.discWhy = why; return false; }
    var p = props_(), all = p.getProperties();
    Object.keys(all).forEach(function (k) { if (k.indexOf('TOK_') !== 0) return; try { if (JSON.parse(all[k]).uid === u.id) p.deleteProperty(k); } catch (e) {} });
    c._discard = 1; return true;
  });
}
/* v9: the MAIN Admin (name key ADMIN_ALIASES[0] = jakirsten, any alias, with the admin flag) can never be removed, blocked,
   rejected or lose Admin (v10: or be renamed to a name that is not the main name / an alias), by any token or the team code. Refused changes go back as refusedDiscard so the phone re-pulls. */
function isMain_(u) { return !!(u && u.admin && nk_(u.name) === ADMIN_ALIASES[0]); }
function protectMain_(store, changes, refused) {
  return (changes || []).filter(function (c) {
    if (!c || c.kind !== 'u') return true;
    var u = userById_(store, c.id); if (!isMain_(u)) return true;
    var d = c.data && typeof c.data === 'object' ? c.data : null;
    /* v10: also no demote and no rename away from the main name / aliases; Reset PIN (same record, PIN fields cleared) stays allowed */
    var why = (c.del || !d || d.status !== 'active') ? 'main_admin' : !d.admin ? 'main_demote' : nk_(d.name) !== ADMIN_ALIASES[0] ? 'main_rename' : '';
    if (!why) return true;
    refused.disc = refused.disc || []; refused.disc.push(String(c.id)); refused.discWhy = why; return false;
  });
}
/* a staff token may only change non-user records and its own user record (not status/admin) */
function allowed_(store, me, changes) {
  if (me.admin) return changes || [];
  return (changes || []).filter(function (c) {
    if (!c || c.kind !== 'u') return true;
    if (String(c.id) !== me.id || c.del || !c.data) return false;
    return c.data.status === me.status && !!c.data.admin === !!me.admin && sameList_(perms_(c.data), perms_(me));
  });
}
/* ADMIN FIX RULE: non-admins may only add NEW records (never change/delete an existing one). */
var FIX_LOG = 'Fix log', FIX_HEAD = ['Time', 'Admin', 'Kind', 'Id', 'Old value', 'New value'];
function addOnly_(store, changes, refused) {
  return (changes || []).filter(function (c) {
    if (!c || !c.kind || c.kind === 'u') return true;
    var old = store.map[c.kind + ':' + c.id];
    if (!c.del && !old) { if (canKind_(refused.me, c.kind, 'add')) return true; refused.n++; return false; }
    if (old && !c.del && old.data === JSON.stringify(c.data == null ? null : c.data)) return true;  // same add sent again
    refused.n++; return false;
  });
}
function logFixes_(store, me, changes) {
  var rows = [], now = new Date().toISOString();
  (changes || []).forEach(function (c) {
    if (!c || !c.kind || c.kind === 'u') return;
    var old = store.map[c.kind + ':' + c.id]; if (!old || old.u >= (Number(c.u) || 0)) return;
    var nv = c.del ? 'DELETED' : JSON.stringify(c.data == null ? null : c.data);
    if (nv === old.data) return;
    rows.push([now, me ? me.name : 'team code', c.kind, String(c.id), old.del ? 'DELETED' : old.data, nv]);
  });
  if (rows.length) { var sh = sheet_(FIX_LOG, FIX_HEAD); sh.getRange(sh.getLastRow() + 1, 1, rows.length, FIX_HEAD.length).setValues(rows); }
}
function perms_(u) { return Array.isArray(u && u.perms) ? u.perms : PERM_DEFAULT; }
function can_(me, p) { return !me || !!me.admin || perms_(me).indexOf(p) >= 0; }  /* me=null = old team code (full, as before) */
function canKind_(me, kind, what) { var k = PERM_KINDS[kind]; if (!k || !k[what]) return true; var n = k[what];
  return Array.isArray(n) ? n.some(function (p) { return can_(me, p); }) : can_(me, n); }
function sameList_(a, b) { return JSON.stringify((a || []).slice().sort()) === JSON.stringify((b || []).slice().sort()); }
function open_(req, store) {
  var a = req.action, dev = String(req.device || '').slice(0, 40);
  if (a === 'status') return status_(store, req.name, dev);
  if (a === 'register') {
    var st = status_(store, req.name, dev); if (!st.ok || st.state !== 'none') return st;
    if (dupReq_(store, req.name)) return { ok: true, state: 'dup', refusedWhy: 'duplicate_request' };  /* ONE request per name per app */
    var p = props_(), now = Date.now(), regs = []; try { regs = JSON.parse(p.getProperty('REG_HOUR') || '[]'); } catch (e) {}
    regs = regs.filter(function (t) { return now - t < 3600000; });
    if (regs.length >= MAX_REG_HOUR) return { ok: false, error: 'too_many_requests' };
    regs.push(now); p.setProperty('REG_HOUR', JSON.stringify(regs));
    putUser_(store, { id: 'r' + now + Math.floor(Math.random() * 1000), name: clean_(req.name), status: 'pending', dev: dev, at: new Date(now).toISOString() }, dev);
    return status_(store, req.name, dev);
  }
  if (a === 'login') {
    var u = userById_(store, req.id); if (!u) return { ok: false, error: 'no_user' };
    if (u.status !== 'active') return { ok: false, error: 'not_active', state: u.status };
    var K = nk_(u.name), F = fails_(K);
    if (F.length >= MAX_FAIL) return { ok: false, error: 'locked', until: F[0] + FAIL_WIN };
    if (!u.pinH || String(req.pinH || '') !== u.pinH) { F.push(Date.now()); props_().setProperty('FAIL_' + K, JSON.stringify(F)); return { ok: false, error: F.length >= MAX_FAIL ? 'locked' : 'bad_pin', until: F[0] + FAIL_WIN }; }
    props_().deleteProperty('FAIL_' + K);
    return { ok: true, token: newToken_(u, dev), id: u.id };
  }
  if (a === 'setpin') {
    if (!/^[a-f0-9]{64}$/.test(String(req.pinH || '')) || !req.salt) return { ok: false, error: 'bad_pin' };
    var v;
    if (req.setup) {
      var s2 = status_(store, req.name, dev); if (s2.state !== 'setup') return { ok: false, error: 'not_allowed' };
      v = { id: 'u' + Date.now(), name: ADMIN_NAME, status: 'active', admin: 1, dev: dev, at: new Date().toISOString() };
    } else {
      v = userById_(store, req.id); if (!v || v.status !== 'active' || v.pinH) return { ok: false, error: 'not_allowed' };
    }
    v.salt = String(req.salt).slice(0, 40); v.pinH = String(req.pinH); delete v.pinReset; delete v.hasPin; putUser_(store, v, dev);
    return { ok: true, token: newToken_(v, dev), id: v.id };
  }
  return null;
}
/* ===== VER 9 PHOTOS TO DRIVE (Johan 1 Nov 2026) — same as base / Fuel / Purchases =====
   action 'photo' / 'photoget'. Server-only kind 'pl'. Daily photoCleanup. */
var PHOTO_KIND = 'pl', PHOTO_KEEP_DAYS = 60, PHOTO_MAX = 400000, PHOTO_TRIGGER = { fn: 'photoCleanup', hour: 3 };
function photoFolder_() {
  var p = props_(), id = p.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var nf = DriveApp.createFolder(String(APP_SHEET_NAME).replace(/ Sheet$/, '') + ' photos');
  p.setProperty('PHOTO_FOLDER_ID', nf.getId()); return nf;
}
function photoCell_(store, id) {
  id = String(id || ''); if (!id) return '';
  var r = store.map[PHOTO_KIND + ':' + id]; if (!r || r.del) return id;
  try { var d = JSON.parse(r.data || '{}'); return d.removed ? 'photo removed' : (d.url || id); } catch (e) { return id; }
}
function noPhotoRecs_(changes) { return (changes || []).filter(function (c) { return !(c && c.kind === PHOTO_KIND); }); }
function photo_(store, me, req) {
  var id = String(req.id || '');
  if (!/^[A-Za-z0-9_-]{3,60}$/.test(id)) return { ok: false, error: 'bad_id' };
  var key = PHOTO_KIND + ':' + id, old = store.map[key], od = null;
  try { od = old && !old.del ? JSON.parse(old.data || 'null') : null; } catch (e) { od = null; }
  if (req.action === 'photoget') {
    if (!(me && me.admin)) return { ok: false, error: 'admin_only' };
    if (!od) return { ok: false, error: 'no_photo' };
    if (od.removed || !od.fid) return { ok: false, error: 'photo_removed' };
    try { return { ok: true, id: id, d: 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(od.fid).getBlob().getBytes()) }; }
    catch (e) { return { ok: false, error: 'photo_removed' }; }
  }
  if (od) return { ok: true, id: id, url: od.url || '', dup: 1, removed: od.removed ? 1 : 0 };
  var m = /^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/.exec(String(req.d || ''));
  if (!m) return { ok: false, error: 'bad_photo' };
  if (m[1].length > PHOTO_MAX) return { ok: false, error: 'too_big' };
  var at = Date.parse(req.at) || Date.now(), now = Date.now();
  if (now - at > PHOTO_KEEP_DAYS * 864e5) return { ok: true, id: id, skipped: 'old' };
  var f = photoFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', id + '.jpg'));
  var data = { id: id, fid: f.getId(), url: f.getUrl(), at: new Date(at).toISOString(), by: me ? me.name : '', ref: String(req.ref || '').slice(0, 80) };
  merge_(store, [{ kind: PHOTO_KIND, id: id, u: now, del: 0, data: data }], req.device, now); writeRecs_(store);
  try { rebuildViews_(store); } catch (ve) {}
  return { ok: true, id: id, url: data.url };
}
function photoCleanup() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var store = readAll_(), now = Date.now(), cut = now - PHOTO_KEEP_DAYS * 864e5, n = 0, ch = [];
    Object.keys(store.map).forEach(function (k) {
      var r = store.map[k]; if (r.kind !== PHOTO_KIND || r.del) return;
      var d; try { d = JSON.parse(r.data || '{}'); } catch (e) { return; }
      if (d.removed || (Date.parse(d.at) || now) >= cut) return;
      try { if (d.fid) DriveApp.getFileById(d.fid).setTrashed(true); } catch (e) {}
      d.removed = 1; d.url = ''; d.removedAt = new Date(now).toISOString(); ch.push({ kind: PHOTO_KIND, id: r.id, u: Math.max(now, r.u + 1), del: 0, data: d }); n++;
    });
    try { var it = photoFolder_().getFiles(); while (it.hasNext()) { var f = it.next(); if (f.getDateCreated().getTime() < cut) { f.setTrashed(true); } } } catch (e) {}
    if (ch.length) { merge_(store, ch, 'cleanup', now); writeRecs_(store); try { rebuildViews_(store); } catch (ve) {} }
    return n;
  } finally { lock.releaseLock(); }
}

function handle_(req) {
  if (!req) return { ok: false, error: 'bad_request' };
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'busy_try_again' };
  try {
    var store = readAll_();
    var o = open_(req, store); if (o) return o;
    var code = teamCode_(), changes = noPhotoRecs_(req.changes), me = null, refused = { n: 0 };
    var viewerId = '';
    /* v6 ADMIN FIX RULE (server): a valid token wins over the team code, so an Admin token is never treated as staff
       (tests/admin_fix_test_spec.md A2). A phone with a token sends only the token; the code is used only when there is no valid token. */
    var tokMe = req.token ? session_(store, req.token) : null;
    if (tokMe || !(code && req.code && String(req.code).trim() === code)) {
      if (!req.token) return { ok: false, error: req.code ? 'bad_code' : 'need_login' };
      me = tokMe;
      if (!me) return { ok: false, error: 'bad_token' };
      if (me.blocked) return { ok: false, error: 'not_active' };
      changes = keepPins_(store, allowed_(store, me, discard_(store, me, protectMain_(store, changes, refused), refused)));
      refused.me = me; viewerId = me.id;
      if (!me.admin) changes = addOnly_(store, changes, refused); else logFixes_(store, me, changes);
    } else {
      if (strict_()) changes = (changes || []).filter(function (c) { if (c && c.kind === 'u') { refused.strict = (refused.strict || 0) + 1; return false; } return true; });  /* STRICT_TOKEN: users only via token */
      changes = keepPins_(store, addOnly_(store, discard_(store, null, protectMain_(store, changes, refused), refused), refused)); /* old team-code access: add only, no remove */
      var tu = req.token ? session_(store, req.token) : null; if (tu && !tu.blocked) viewerId = tu.id;
    }
    changes = (changes || []).filter(function (c) { return !c || c.kind !== 'fv'; });  /* Fuel copies come only from pullFuel */
    changes = storeRules_(store, me, changes, refused);  /* v7: Admin-only stock corrections, no stock below zero without Admin */
    changes = noDupReq_(store, changes, refused);
    if (req.action === 'photo' || req.action === 'photoget') return photo_(store, me, req);
    if (req.action === 'report' || req.action === 'export') {  /* ticked actions, checked here on the server */
      var need = req.action === 'report' ? 'reports' : 'export';
      if (!can_(me, need)) return { ok: false, error: 'not_allowed', perm: need };
      var cnt = {}, rows = [];
      Object.keys(store.map).forEach(function (k) { var r = store.map[k]; if (r.kind === 'u' || r.del || !canKind_(me, r.kind, 'view')) return;
        cnt[r.kind] = (cnt[r.kind] || 0) + 1; if (req.action === 'export') rows.push({ kind: r.kind, id: r.id, data: JSON.parse(r.data || 'null') }); });
      return req.action === 'report' ? { ok: true, counts: cnt } : { ok: true, rows: rows };
    }
    try { if (Date.now() - Number(props_().getProperty('FUEL_PULLED') || 0) > FUEL_EVERY_MS) pullFuelInto_(store); } catch (fe) {}
    var now = Date.now(), since = Number(req.since) || 0;
    var acc = merge_(store, changes, req.device, now), warn = '';
    if (acc) { writeRecs_(store); try { rebuildViews_(store); } catch (ve) { warn = 'views: ' + String(ve && ve.message || ve).slice(0, 200); } }
    var res = { ok: true, now: now, accepted: acc, records: changedSince_(store, since, me, viewerId) };
    if (refused.disc) { res.refused = refused.disc.length + (refused.n || 0); res.refusedWhy = refused.discWhy; res.refusedDiscard = refused.disc; }  /* SERVER CHANGE v9 */
    else if (refused.n) { res.refused = refused.n; res.refusedWhy = 'admin_only_fix'; }
    else if (refused.strict) { res.refused = refused.strict; res.refusedWhy = 'need_token'; }
    else if (refused.dup) { res.refused = refused.dup; res.refusedWhy = 'duplicate_request'; }
    if (refused.stock && refused.stock.length) res.refusedStock = refused.stock;
    if (warn) res.warn = warn;
    return res;
  } finally { lock.releaseLock(); }
}
function doPost(e) {
  try {
    var body = e && e.postData && e.postData.contents;
    if (!body) return out_({ ok: false, error: 'empty_request' });
    var req; try { req = JSON.parse(body); } catch (pe) { return out_({ ok: false, error: 'bad_json' }); }
    return out_(handle_(req));
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err).slice(0, 300) }); }
}
function doGet() { return out_({ ok: true, app: 'TMC Service sync', hint: 'running' }); }
/* ===== ONE-TIME SETUP (standard in every app): open the editor, pick setupOnce, Run, Allow. =====
   TMC Base replaces the placeholder with the real team code ONLY in the copy pasted into Google
   (never commit a real code). Left as the placeholder → TEAM_CODE is not touched (token-only app). */
var APP_SHEET_NAME = 'TMC Service Sheet';  /* each app: '<app name> Sheet' */
var TEAM_CODE_SETUP = '__TMC_TEAM_CODE__';
/* Time triggers this app needs, e.g. { fn: 'dailyJob', hour: 5 }. Base has none. */
var TRIGGERS = [{ fn: 'pullFuel', hour: 5 }];
/* PHOTO_TRIGGER added in installTriggers_ below */
function renameSheetOnce_() { try { var ss = SpreadsheetApp.getActive(); if (ss.getName() !== APP_SHEET_NAME) { ss.rename(APP_SHEET_NAME); return 'renamed'; } return 'ok'; } catch (e) { return 'skip'; } }
function installTriggers_() {
  var have = {}, out = [];
  ScriptApp.getProjectTriggers().forEach(function (t) { var f = t.getHandlerFunction(); if (have[f]) ScriptApp.deleteTrigger(t); else have[f] = 1; });  /* drop duplicates */
  TRIGGERS.concat([PHOTO_TRIGGER]).forEach(function (t) {
    if (have[t.fn]) { out.push(t.fn + ':kept'); return; }
    ScriptApp.newTrigger(t.fn).timeBased().everyDays(1).atHour(t.hour || 5).nearMinute(t.minute || 0).create(); have[t.fn] = 1; out.push(t.fn + ':added');
  });
  return out;
}
function setupOnce() {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var p = props_(), out = { sheet: renameSheetOnce_() };
    if (TEAM_CODE_SETUP && TEAM_CODE_SETUP.indexOf('__') !== 0) { p.setProperty('TEAM_CODE', String(TEAM_CODE_SETUP).trim()); out.teamCode = 'set'; }
    else out.teamCode = teamCode_() ? 'kept' : 'none (token sign-in only)';
    var store = readAll_(); sheet_(FIX_LOG, FIX_HEAD);
    if (!users_(store).some(function (u) { return u.admin && u.status === 'active'; })) {
      putUser_(store, { id: 'u' + Date.now(), name: ADMIN_NAME, status: 'active', admin: 1, dev: 'setup', at: new Date().toISOString() }, 'setup'); out.admin = 'added';
    } else { rebuildViews_(store); out.admin = 'kept'; }
    out.triggers = installTriggers_();
    out.strict = strict_() ? 'on' : 'off';
    return JSON.stringify(out);
  } finally { lock.releaseLock(); }
}
if (typeof module !== 'undefined') module.exports = { handle_: handle_, merge_: merge_, setupOnce: setupOnce, discard_: discard_, protectMain_: protectMain_, isMain_: isMain_, photoCleanup: photoCleanup };

/* ===== TMC Service: copy vehicles + latest meter from TMC Fuel (ONE-WAY, READ-ONLY) =====
   Opens Fuel's Sheet by ID and only READS (getDataRange().getDisplayValues()). Nothing in this
   block calls any set/append/insert/delete on Fuel. Copies only fleet ID, name, unit, latest reading
   into this app's own kind 'fv' records. */
var FUEL_SHEET_ID = '1mLETc2v9hCSWMN_9cNErMEei-R2elFwXHwHmozQgbz4';
var FUEL_EVERY_MS = 30 * 60 * 1000;
function fuelTab_(ss, name) {
  var sh = ss.getSheetByName(name); if (!sh) return [];
  var v = sh.getDataRange().getDisplayValues(); if (v.length < 2) return [];
  var h = v[0].map(function (x) { return String(x).trim(); });
  return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}
/* v6: the Fuel day is the CHAD local day (Africa/Ndjamena = UTC+1 all year, no daylight saving).
   A plain date (Daily readings 'Day') or a time with no zone is already local → its date part.
   A time with a zone (fill-up 'Date (ISO UTC)', e.g. 2026-10-27T23:30:00Z = 28 Oct 00:30 in Chad) → shifted to Chad, then the date. */
var CHAD_OFFSET_MS_ = 60 * 60 * 1000;
function chadDay_(s) { if (s instanceof Date) return isNaN(s) ? '' : new Date(s.getTime() + CHAD_OFFSET_MS_).toISOString().slice(0, 10);
  s = String(s || '').trim(); if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s) || /^\d{4}-\d{2}-\d{2}[T ]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?$/.test(s)) return s.slice(0, 10);
  var d = new Date(s); return isNaN(d) ? '' : new Date(d.getTime() + CHAD_OFFSET_MS_).toISOString().slice(0, 10); }
function fuelDay_(s) { return chadDay_(s); }
/* v8: Kouri (Fuel id s_Kouri, ID 'Kouri', "Kouri Fuel") is a fuel supplier, not a vehicle; a vehicle Fuel marks
   Stopped ('Working' = stopped, i.e. active:false) is not serviced. Neither is copied, and an old copy is removed. */
function fuelSkip_(v) { var k = function (x) { return String(x == null ? '' : x).trim().toLowerCase().replace(/^s_/, ''); };
  return k(v['ID']) === 'kouri' || k(v['Vehicles']) === 'kouri fuel' || /^(stopped|false)$/.test(k(v['Working'])) || v.active === false; }
function fuelLatest_(veh, daily, fills) {
  var best = {};
  function see(fleet, when, val, unit) { fleet = String(fleet || '').trim(); var n = Number(String(val).replace(/[ ,]/g, '')); when = fuelDay_(when);
    if (!fleet || !when || String(val).trim() === '' || !isFinite(n)) return;  /* v6: compare by CHAD DAY; same day (fill-up + daily reading) → keep the HIGHEST reading */
    var b0 = best[fleet]; if (!b0 || when > b0.w || (when === b0.w && n > b0.r)) best[fleet] = { w: when, r: n, u: unit || (b0 && b0.u) }; }
  daily.forEach(function (d) { see(d['Fleet no'], d['Day'], d['Reading']); });
  fills.forEach(function (f) { see(f['Fleet no'], f['Date'], f['Meter'], f['Unit']); });
  return veh.filter(function (v) { return String(v['ID'] || '').trim() && !fuelSkip_(v); }).map(function (v) {
    var id = String(v['ID']).trim(), b = best[id] || {}, unit = /^h/i.test(String(v['Meter'] || b.u || '')) ? 'h' : 'km';
    return { id: id, fleet: id, name: String(v['Vehicles'] || v['Model'] || '').trim(), unit: unit, reading: b.r === undefined ? null : b.r, day: b.w ? String(b.w).slice(0, 10) : '' };
  });
}
function pullFuelInto_(store) {
  var ss = SpreadsheetApp.openById(FUEL_SHEET_ID);
  var VT = fuelTab_(ss, 'Vehicles'), rows = fuelLatest_(VT, fuelTab_(ss, 'DailyReadings'), fuelTab_(ss, 'FillUps'));
  var now = Date.now(), ch = [], gone = { Kouri: 1, s_Kouri: 1 };
  VT.forEach(function (v) { var id = String(v['ID'] || '').trim(); if (id && fuelSkip_(v)) gone[id] = 1; });
  Object.keys(store.map).forEach(function (key) { var o = store.map[key]; if (o.kind !== 'fv' || o.del) return;  /* only ones Fuel lists as Kouri/Stopped: an empty read never wipes the list */
    var d = {}; try { d = JSON.parse(o.data || '{}') || {}; } catch (e) {}
    if (gone[o.id] || fuelSkip_({ ID: o.id, Vehicles: d.name })) ch.push({ kind: 'fv', id: o.id, u: Math.max(now, o.u + 1), del: 1, data: null }); });
  rows.forEach(function (r) { var old = store.map['fv:' + r.id]; if (old && old.data === JSON.stringify(r)) return;
    ch.push({ kind: 'fv', id: r.id, u: Math.max(now, (old ? old.u : 0) + 1), del: 0, data: r }); });
  if (ch.length) { merge_(store, ch, 'fuel-copy', now); writeRecs_(store); }
  props_().setProperty('FUEL_PULLED', String(now));   /* this app's own script property, not Fuel */
  return rows.length;
}
function pullFuel() { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return pullFuelInto_(readAll_()); } finally { lock.releaseLock(); } }
if (typeof module !== 'undefined') { module.exports.fuelLatest_ = fuelLatest_; module.exports.pullFuelInto_ = pullFuelInto_; module.exports.fuelSkip_ = fuelSkip_; }

/* ===== LINK (Fuel v28 <-> Service v12): "Check machine (high fuel use)" calls. =====
   Only ONE small record type crosses: Fuel -> Service {fid, vid, name, date, rate, normal}; Service -> Fuel {fid} "done".
   Nothing else is sent or read: each app keeps its own Sheet, users, PINs, tokens and caches.
   Shared secret + the other app's /exec URL live in THIS script's properties (LINK_SECRET, LINK_URL), set by setupOnce.
   Outbox = script properties LQ_<id> (kept until the other app says ok → retried every 5 min while offline); LS_<id> = sent (never sent twice). */
var LINK_SECRET_SETUP = '__LINK_SECRET__', LINK_URL_SETUP = '__LINK_URL__';  /* real values only in the copy pushed to Google, never committed */
var LINK_TRIGGER_FN = 'linkFlush';
function linkOn_() { var p = props_(); return !!(p.getProperty('LINK_SECRET') && p.getProperty('LINK_URL')); }
function linkOk_(s) { var k = String(props_().getProperty('LINK_SECRET') || ''); return k.length >= 16 && String(s || '') === k; }
function linkId_(x) { return String(x == null ? '' : x).replace(/[^A-Za-z0-9_.:-]/g, '').slice(0, 60); }
function linkNum_(x) { var n = Number(x); return isFinite(n) ? Math.round(n * 10) / 10 : null; }
function linkTxt_(x) { return String(x == null ? '' : x).replace(/[\u0000-\u001f<>]/g, '').slice(0, 80); }
function linkQueue_(id, body) { var p = props_(); id = linkId_(id); if (!id || p.getProperty('LQ_' + id) || p.getProperty('LS_' + id)) return false; p.setProperty('LQ_' + id, JSON.stringify(body)); return true; }
function linkFlush() {  /* trigger + after each sync: send what is waiting; keep it when offline / refused */
  var p = props_(), all = p.getProperties(), url = p.getProperty('LINK_URL'), key = p.getProperty('LINK_SECRET'), n = 0;
  if (!url || !key) return 'off';
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('LQ_') !== 0) return; var id = k.slice(3), body = {};
    try { body = JSON.parse(all[k]); } catch (e) { p.deleteProperty(k); return; }
    body.secret = key;
    try {
      var r = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain', payload: JSON.stringify(body), muteHttpExceptions: true, followRedirects: true });
      var j = {}; try { j = JSON.parse(r.getContentText()); } catch (e2) {}
      if (r.getResponseCode() === 200 && j && j.ok) { p.setProperty('LS_' + id, String(Date.now())); p.deleteProperty(k); n++; }
    } catch (e3) { /* offline: try again next time */ }
  });
  return n;
}
function linkSetup_() {
  var p = props_();
  if (LINK_SECRET_SETUP.indexOf('__') !== 0) p.setProperty('LINK_SECRET', LINK_SECRET_SETUP);
  if (LINK_URL_SETUP.indexOf('__') !== 0) p.setProperty('LINK_URL', LINK_URL_SETUP);
  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === LINK_TRIGGER_FN; }))
    ScriptApp.newTrigger(LINK_TRIGGER_FN).timeBased().everyMinutes(5).create();
  return linkOn_() ? 'on' : 'off (no secret/url)';
}
var linkSetup0_ = setupOnce;
setupOnce = function () { var o = {}; try { o = JSON.parse(linkSetup0_()); } catch (e) { o = { base: 'error: ' + String(e && e.message || e).slice(0, 120) }; } try { o.link = linkSetup_(); } catch (e2) { o.link = 'error: ' + String(e2 && e2.message || e2).slice(0, 120); } return JSON.stringify(o); };
var linkHandle0_ = handle_;
handle_ = function (req) {
  if (req && (req.action === LINK_IN_ACTION)) return linkIn_(req);
  if (req && Array.isArray(req.changes)) req.changes = req.changes.filter(function (c) { return !(c && LINK_SERVER_KINDS[c.kind]); });  /* phones never write link records */
  return linkHandle0_(req);
};
var linkPost0_ = doPost;
doPost = function (e) { var o = linkPost0_(e); try { if (linkOn_() && Object.keys(props_().getProperties()).some(function (k) { return k.indexOf('LQ_') === 0; })) linkFlush(); } catch (er) {} return o; };
function linkIn_(req) {
  if (!linkOk_(req.secret)) return { ok: false, error: 'bad_link' };
  var lock = LockService.getScriptLock(); if (!lock.tryLock(25000)) return { ok: false, error: 'busy_try_again' };
  try { return linkApply_(readAll_(), req); } finally { lock.releaseLock(); }
}
/* Service side: Fuel's call → job record cm:<fill-up id> "Check machine (high fuel use)" (only this one record type is accepted).
   Marking it done = a phone adds cd:<same id> (add-only, 'add' tick) → "done" goes back to Fuel once. */
var LINK_IN_ACTION = 'linkCall', LINK_SERVER_KINDS = { cm: 1 };
PERM_KINDS.cm = { view: 'view', add: 'admin' }; PERM_KINDS.cd = { view: 'view', add: 'add' };
function linkApply_(store, req) {
  var c = req.call || {}, fid = linkId_(c.fid); if (!fid || !linkId_(c.vid)) return { ok: false, error: 'bad_call' };
  var old = store.map['cm:' + fid]; if (old && !old.del) return { ok: true, dup: true };   /* the same call twice = one job */
  var now = Date.now(), d = { id: fid, fid: fid, vid: linkId_(c.vid), name: linkTxt_(c.name), date: linkTxt_(c.date), rate: linkNum_(c.rate), normal: linkNum_(c.normal), at: new Date(now).toISOString() };
  merge_(store, [{ kind: 'cm', id: fid, u: now, del: 0, data: d }], 'fuel-link', now); writeRecs_(store); try { rebuildViews_(store); } catch (ve) {}
  return { ok: true };
}
var linkMerge0_ = merge_;
merge_ = function (store, changes, device, now) {
  var acc = linkMerge0_(store, changes, device, now);
  try { (changes || []).forEach(function (c) {
    if (!c || c.kind !== 'cd' || c.del) return; var r = store.map['cd:' + c.id], m = store.map['cm:' + c.id];
    if (r && !r.del && m && !m.del) linkQueue_(c.id, { action: 'linkDone', fid: linkId_(c.id) });
  }); } catch (er) {}
  return acc;
};
