'use strict';
/**
 * Mock tests for the Service Apps Script photo update.
 * Loads server/Code.gs (link constants redacted). No live Sheet, Drive, or URL calls.
 * If the deployment artifact and the live backup are both on disk, checks they differ
 * only on the link-constant line and that the artifact still matches the backup's line.
 * Never prints those constants.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const assert = require('assert');

const REPO = path.join(__dirname, 'Code.gs');
const LIVE = '/home/ubuntu/.cursor/projects/workspace/uploads/Code_e55d.gs';
const SAMPLE = '/home/ubuntu/.cursor/projects/workspace/uploads/Code_192f.gs';
const ART = '/opt/cursor/artifacts/TMC-Service-Code.gs';
const PIN = 'ab'.repeat(32);
const PIN2 = 'cd'.repeat(32);
const TEST_SECRET = 'test-link-secret-0123456789abcd';
const TEST_URL = 'https://fuel-link.example/exec';
const REC_HEAD = ['Key', 'Kind', 'Id', 'UpdatedAt', 'Deleted', 'ServerTime', 'Device', 'Json'];

const results = [];
function t(name, fn) {
  try { fn(); results.push('ok  ' + name); }
  catch (e) { results.push('FAIL ' + name + ' — ' + (e && e.message ? e.message : e)); }
}

function linkConsts(src) {
  const m = src.match(/var LINK_SECRET_SETUP = '([^']*)', LINK_URL_SETUP = '([^']*)';/);
  if (!m) throw new Error('link constant line missing or reshaped');
  return { secret: m[1], url: m[2] };
}
function sha256(s) { return crypto.createHash('sha256').update(s).digest('hex'); }

function makeSheet(name, ss) {
  const sh = {
    name: name,
    rows: [],
    getLastRow() {
      for (let i = this.rows.length - 1; i >= 0; i--) {
        const row = this.rows[i] || [];
        if (row.some(function (c) { return c !== '' && c != null; })) return i + 1;
      }
      return 0;
    },
    appendRow(v) { this.rows.push(v.slice()); },
    setFrozenRows() {},
    getRange(r, c, nr, nc) {
      if (arguments.length === 2) { nr = 1; nc = 1; }
      const sheet = this;
      return {
        setNumberFormat() { return this; },
        getValues() { return read(); },
        getDisplayValues() { return read().map(function (row) { return row.map(function (v) { return v == null ? '' : String(v); }); }); },
        setValues(data) {
          data.forEach(function (line, i) {
            const ri = r - 1 + i;
            if (!sheet.rows[ri]) sheet.rows[ri] = [];
            line.forEach(function (cell, j) { sheet.rows[ri][c - 1 + j] = cell; });
          });
          return this;
        },
        clearContent() {
          for (let i = 0; i < nr; i++) {
            const row = sheet.rows[r - 1 + i];
            if (!row) continue;
            for (let j = 0; j < nc; j++) row[c - 1 + j] = '';
          }
          return this;
        }
      };
      function read() {
        const o = [];
        for (let i = 0; i < nr; i++) {
          const row = sheet.rows[r - 1 + i] || [];
          const line = [];
          for (let j = 0; j < nc; j++) line.push(row[c - 1 + j] == null ? '' : row[c - 1 + j]);
          o.push(line);
        }
        return o;
      }
    },
    getDataRange() {
      const n = Math.max(this.getLastRow(), 1);
      let w = 1;
      this.rows.forEach(function (row) { w = Math.max(w, row.length); });
      return this.getRange(1, 1, n, w);
    }
  };
  return sh;
}

function load(code) {
  const bag = {};
  const triggers = [];
  const fetches = [];
  const folderNames = [];
  const files = {};
  const folders = {};
  let seq = 1;
  const fuelWrites = { n: 0 };
  function fileObj(meta) {
    return {
      getId: function () { return meta.id; },
      getUrl: function () { return 'https://drive.example/file/' + meta.id; },
      getBlob: function () { return { getBytes: function () { return Buffer.from(meta.bytes); } }; },
      setTrashed: function (v) { meta.trashed = !!v; },
      isTrashed: function () { return !!meta.trashed; },
      getDateCreated: function () { return meta.created; }
    };
  }
  function folderObj(meta) {
    return {
      getId: function () { return meta.id; },
      isTrashed: function () { return !!meta.trashed; },
      createFile: function (blob) {
        const id = 'file' + (seq++);
        files[id] = { id: id, bytes: Buffer.from(blob.getBytes()), created: new Date(), trashed: false, folder: meta.id };
        return fileObj(files[id]);
      },
      getFiles: function () {
        const list = Object.keys(files).map(function (k) { return files[k]; }).filter(function (f) { return f.folder === meta.id && !f.trashed; });
        let i = 0;
        return { hasNext: function () { return i < list.length; }, next: function () { return fileObj(list[i++]); } };
      }
    };
  }
  function makeSS(label, readonly) {
    const sheets = {};
    const ss = {
      name: label,
      getName: function () { return ss.name; },
      rename: function (n) { if (readonly) fuelWrites.n++; ss.name = n; },
      getSheetByName: function (n) { return sheets[n] || null; },
      insertSheet: function (n) {
        if (readonly) fuelWrites.n++;
        const sh = makeSheet(n, ss);
        if (readonly) {
          const raw = sh.getRange;
          sh.appendRow = function () { fuelWrites.n++; };
          sh.getRange = function () {
            const rg = raw.apply(sh, arguments);
            const set = rg.setValues;
            rg.setValues = function () { fuelWrites.n++; return set.apply(rg, arguments); };
            rg.clearContent = function () { fuelWrites.n++; return rg; };
            return rg;
          };
        }
        sheets[n] = sh;
        return sh;
      }
    };
    return { ss: ss, sheets: sheets };
  }
  const service = makeSS('Untitled', false);
  const fuel = makeSS('Fuel Sheet', true);
  function fuelSheet(name, grid) {
    const sh = fuel.ss.insertSheet(name);
    fuelWrites.n = 0;
    sh.rows = grid.map(function (r) { return r.slice(); });
  }
  fuelSheet('Vehicles', [
    ['ID', 'Vehicles', 'Model', 'Meter', 'Working'],
    ['Kouri', 'Kouri Fuel', '', 'km', 'working'],
    ['STOP1', 'Old truck', '', 'h', 'stopped'],
    ['TRK1', 'Water truck', '', 'h', 'working']
  ]);
  fuelSheet('DailyReadings', [
    ['Fleet no', 'Day', 'Reading'],
    ['TRK1', '2026-11-01', '100'],
    ['Kouri', '2026-11-01', '5']
  ]);
  fuelSheet('FillUps', [
    ['Fleet no', 'Date', 'Meter', 'Unit'],
    ['TRK1', '2026-11-02', '110', 'h']
  ]);
  fuelWrites.n = 0;

  const sandbox = {
    Buffer: Buffer,
    console: console,
    Date: Date,
    Math: Math,
    JSON: JSON,
    Object: Object,
    Array: Array,
    String: String,
    Number: Number,
    RegExp: RegExp,
    Error: Error,
    isFinite: isFinite,
    PropertiesService: { getScriptProperties: function () {
      return {
        getProperty: function (k) { return Object.prototype.hasOwnProperty.call(bag, k) ? bag[k] : null; },
        setProperty: function (k, v) { bag[k] = String(v); },
        deleteProperty: function (k) { delete bag[k]; },
        getProperties: function () { return Object.assign({}, bag); }
      };
    } },
    LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, waitLock: function () {}, releaseLock: function () {} }; } },
    Utilities: {
      getUuid: function () { return crypto.randomUUID(); },
      base64Encode: function (bytes) { return Buffer.from(bytes).toString('base64'); },
      base64Decode: function (s) { return Buffer.from(String(s), 'base64'); },
      newBlob: function (data, mime, name) { const buf = Buffer.from(data); return { getBytes: function () { return Buffer.from(buf); }, mime: mime, name: name }; }
    },
    DriveApp: {
      createFolder: function (name) {
        folderNames.push(String(name));
        const id = 'fld' + (seq++);
        folders[id] = { id: id, name: String(name), trashed: false };
        return folderObj(folders[id]);
      },
      getFolderById: function (id) { if (!folders[id]) throw new Error('no folder'); return folderObj(folders[id]); },
      getFileById: function (id) { if (!files[id]) throw new Error('no file'); return fileObj(files[id]); }
    },
    SpreadsheetApp: {
      getActive: function () { return service.ss; },
      getActiveSpreadsheet: function () { return service.ss; },
      openById: function (id) { sandbox._opened.push(String(id)); return fuel.ss; }
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: function (s) { const o = { getContent: function () { return s; }, setMimeType: function () { return o; } }; return o; }
    },
    ScriptApp: {
      getProjectTriggers: function () { return triggers.slice(); },
      deleteTrigger: function (tr) { const i = triggers.indexOf(tr); if (i >= 0) triggers.splice(i, 1); },
      newTrigger: function (fn) {
        return { timeBased: function () {
          const spec = { fn: fn, getHandlerFunction: function () { return fn; } };
          const api = {
            everyDays: function () { spec.kind = 'day'; return api; },
            atHour: function (h) { spec.hour = h; return api; },
            nearMinute: function (m) { spec.minute = m; return api; },
            everyMinutes: function (n) { spec.kind = 'min'; spec.every = n; return api; },
            create: function () { triggers.push(spec); return spec; }
          };
          return api;
        } };
      }
    },
    UrlFetchApp: { fetch: function (url, opt) { fetches.push({ url: url, payload: opt.payload }); return { getResponseCode: function () { return 200; }, getContentText: function () { return '{"ok":true}'; } }; } },
    _bag: bag, _triggers: triggers, _fetches: fetches, _files: files, _folders: folders, _folderNames: folderNames, _fuelWrites: fuelWrites, _opened: []
  };
  vm.runInNewContext(code, sandbox, { filename: 'Code.gs' });
  sandbox.props_().setProperty('FUEL_PULLED', String(Date.now()));
  return sandbox;
}

function jpeg(text) {
  const b64 = Buffer.from(text).toString('base64');
  return { b64: b64, d: 'data:image/jpeg;base64,' + b64 };
}
function seed(ctx, kind, id, data, opt) {
  opt = opt || {};
  const sh = ctx.sheet_(ctx.REC, ctx.REC_HEAD);
  const u = opt.u || Date.now();
  sh.appendRow([kind + ':' + id, kind, id, u, opt.del ? 1 : 0, opt.s || u, opt.dev || 'seed', JSON.stringify(data)]);
}
function user(id, name, extra) {
  return Object.assign({ id: id, name: name, status: 'active', admin: 0, salt: 'salt', pinH: PIN, dev: 'seed', at: '2026-11-01T00:00:00.000Z' }, extra || {});
}
function login(ctx, id, device, pinH) {
  return ctx.handle_({ action: 'login', id: id, pinH: pinH || PIN, device: device || 'phone' });
}
function recs(ctx) {
  const store = ctx.readAll_();
  const o = {};
  Object.keys(store.map).forEach(function (k) { o[k] = store.map[k]; });
  return o;
}
function headerOf(ctx) {
  const sh = ctx.sheet_(ctx.REC, ctx.REC_HEAD);
  return sh.getRange(1, 1, 1, 8).getDisplayValues()[0];
}

const code = fs.readFileSync(REPO, 'utf8');
let ctx;

t('sources: repo redacted, artifact matches live except link line', function () {
  const repoC = linkConsts(code);
  assert.strictEqual(repoC.secret, '__LINK_SECRET__');
  assert.strictEqual(repoC.url, '__LINK_URL__');
  assert.ok(code.includes('function photoCleanup'));
  assert.ok(code.includes('linkCall'));
  assert.ok(code.includes("m: { view: ['view', 'store', 'storeEd', 'add'], add: 'add' }"));
  assert.ok(code.includes("var REC_HEAD = ['Key', 'Kind', 'Id', 'UpdatedAt', 'Deleted', 'ServerTime', 'Device', 'Json'];"));
  if (!fs.existsSync(LIVE) || !fs.existsSync(ART)) return;
  const live = fs.readFileSync(LIVE, 'utf8');
  const art = fs.readFileSync(ART, 'utf8');
  const sample = fs.existsSync(SAMPLE) ? fs.readFileSync(SAMPLE, 'utf8') : '';
  const liveC = linkConsts(live);
  const artC = linkConsts(art);
  if (artC.secret !== liveC.secret || artC.url !== liveC.url) throw new Error('artifact link constants do not match the live backup');
  if (artC.secret.indexOf('__') === 0) throw new Error('artifact lost the live link constants');
  if (repoC.secret === liveC.secret) throw new Error('redacted server file still contains the live link secret');
  const strip = function (s) { return s.replace(/var LINK_SECRET_SETUP = '[^']*', LINK_URL_SETUP = '[^']*';/, 'LINK_LINE'); };
  const a = strip(art).split('\n');
  const b = strip(code).split('\n');
  const bad = [];
  for (let i = 0; i < Math.max(a.length, b.length); i++) if (a[i] !== b[i]) bad.push(i + 1);
  if (bad.length) throw new Error('artifact and repo Code.gs differ beyond the link line: ' + bad.join(','));
  if (sample && sample.includes('linkCall')) throw new Error('sample unexpectedly contains Check machine');
  if (!live.includes('linkCall') || live.includes('function photoCleanup')) throw new Error('live backup identity check failed');
  results.push('    sha256 live ' + sha256(live));
  results.push('    sha256 sample ' + sha256(sample));
  results.push('    sha256 artifact ' + sha256(art));
  results.push('    sha256 repo ' + sha256(code));
});

function boot() { ctx = load(code); }

t('photo save, dup, get, and refusal', function () {
  boot();
  seed(ctx, 'u', 'uStaff', user('uStaff', 'Sam Staff', { perms: ['view', 'add'] }));
  seed(ctx, 'u', 'uAdmin', user('uAdmin', 'JA Kirsten', { admin: 1 }));
  const staff = login(ctx, 'uStaff', 'ph-s');
  const admin = login(ctx, 'uAdmin', 'ph-a');
  assert.strictEqual(staff.ok, true);
  const img = jpeg('maint-photo-1');
  const id = 'ph' + Date.now() + 'ab12';
  const saved = ctx.handle_({ action: 'photo', token: staff.token, device: 'ph-s', id: id, d: img.d, at: new Date().toISOString(), ref: 'maint TRK1 tyres' });
  assert.strictEqual(saved.ok, true);
  assert.strictEqual(saved.id, id);
  assert.ok(saved.url);
  assert.deepStrictEqual(ctx._folderNames, ['TMC Service photos']);
  const map = recs(ctx);
  const pl = JSON.parse(map['pl:' + id].data);
  assert.strictEqual(pl.ref, 'maint TRK1 tyres');
  assert.strictEqual(pl.by, 'Sam Staff');
  assert.strictEqual(Buffer.from(ctx._files[pl.fid].bytes).toString(), 'maint-photo-1');
  const again = ctx.handle_({ action: 'photo', token: staff.token, id: id, d: img.d, at: new Date().toISOString(), ref: 'maint TRK1 tyres' });
  assert.strictEqual(again.dup, 1);
  assert.strictEqual(Object.keys(ctx._files).length, 1);
  const got = ctx.handle_({ action: 'photoget', token: admin.token, id: id });
  assert.strictEqual(got.ok, true);
  assert.strictEqual(got.d, img.d);
  assert.strictEqual(ctx.handle_({ action: 'photoget', token: staff.token, id: id }).error, 'admin_only');
  assert.strictEqual(ctx.handle_({ action: 'photo', id: id, d: img.d }).error, 'need_login');
  assert.strictEqual(ctx.handle_({ action: 'photo', token: staff.token, id: 'x', d: img.d }).error, 'bad_id');
  assert.strictEqual(ctx.handle_({ action: 'photo', token: staff.token, id: 'phok1', d: 'data:image/png;base64,aaaa' }).error, 'bad_photo');
  assert.strictEqual(ctx.handle_({ action: 'photo', token: staff.token, id: 'phbig1', d: 'data:image/jpeg;base64,' + 'A'.repeat(400001) }).error, 'too_big');
  const old = ctx.handle_({ action: 'photo', token: staff.token, id: 'phold1', d: img.d, at: new Date(Date.now() - 61 * 864e5).toISOString() });
  assert.strictEqual(old.skipped, 'old');
  assert.strictEqual(ctx.handle_({ action: 'photoget', token: admin.token, id: 'missing1' }).error, 'no_photo');
  const sneak = ctx.handle_({ action: 'photo', token: staff.token, id: 'phsneak1', d: jpeg('x').d, changes: [{ kind: 's', id: 'sSneak', u: Date.now(), del: 0, data: { id: 'sSneak', note: 'no' } }] });
  assert.strictEqual(sneak.ok, true);
  assert.ok(!recs(ctx)['s:sSneak']);
  assert.strictEqual(ctx._fetches.length, 0);
  assert.strictEqual(ctx._fuelWrites.n, 0);
});

t('maintenance kind m: add, permission, no service restart', function () {
  boot();
  seed(ctx, 'u', 'uStaff', user('uStaff', 'Sam Staff', { perms: ['view', 'add'] }));
  seed(ctx, 'u', 'uView', user('uView', 'View Only', { perms: ['view'] }));
  seed(ctx, 'u', 'uStore', user('uStore', 'Store Only', { perms: ['storeEd'] }));
  seed(ctx, 'u', 'uRep', user('uRep', 'Reports Only', { perms: ['reports'] }));
  seed(ctx, 'u', 'uAdmin', user('uAdmin', 'JA Kirsten', { admin: 1 }));
  seed(ctx, 'fv', 'TRK1', { id: 'TRK1', fleet: 'TRK1', name: 'Water truck', unit: 'h', reading: 110, day: '2026-11-02' });
  const staff = login(ctx, 'uStaff', 'ph-s').token;
  const admin = login(ctx, 'uAdmin', 'ph-a').token;
  const rec = { id: 'm100', type: 'tyres', fleet: 'TRK1', name: 'Water truck', date: '2026-11-08', meter: 110, mu: 'h', qty: 2, cost: '', notes: 'front', by: 'Sam Staff', user: 'Sam Staff', at: '2026-11-08T08:00:00.000Z', p: 'ph100' };
  const add = ctx.handle_({ token: staff, device: 'ph-s', since: 0, changes: [{ kind: 'm', id: 'm100', u: 1000, del: 0, data: rec }] });
  assert.strictEqual(add.accepted, 1);
  assert.strictEqual(add.records.filter(function (r) { return r.kind === 's'; }).length, 0);
  assert.ok(add.records.some(function (r) { return r.kind === 'm' && r.id === 'm100'; }));
  const edited = Object.assign({}, rec, { notes: 'changed by staff' });
  const deny = ctx.handle_({ token: staff, device: 'ph-s', since: 0, changes: [{ kind: 'm', id: 'm100', u: 2000, del: 0, data: edited }] });
  assert.ok(deny.refused >= 1);
  assert.strictEqual(JSON.parse(recs(ctx)['m:m100'].data).notes, 'front');
  const viewAdd = ctx.handle_({ token: login(ctx, 'uView', 'ph-v').token, device: 'ph-v', since: 0, changes: [{ kind: 'm', id: 'm200', u: 1000, del: 0, data: Object.assign({}, rec, { id: 'm200' }) }] });
  assert.ok(viewAdd.refused >= 1);
  assert.ok(!recs(ctx)['m:m200']);
  const storeSee = ctx.handle_({ token: login(ctx, 'uStore', 'ph-t').token, device: 'ph-t', since: 0, changes: [] });
  assert.ok(storeSee.records.some(function (r) { return r.kind === 'm'; }));
  const rep = ctx.handle_({ token: login(ctx, 'uRep', 'ph-r').token, device: 'ph-r', since: 0, changes: [] });
  assert.ok(!rep.records.some(function (r) { return r.kind === 'm'; }));
  const fix = Object.assign({}, rec, { notes: 'admin fix' });
  const adm = ctx.handle_({ token: admin, device: 'ph-a', since: 0, changes: [{ kind: 'm', id: 'm100', u: 3000, del: 0, data: fix }] });
  assert.strictEqual(adm.accepted, 1);
  assert.strictEqual(JSON.parse(recs(ctx)['m:m100'].data).notes, 'admin fix');
  const log = ctx.sheet_('Fix log', ['Time', 'Admin', 'Kind', 'Id', 'Old value', 'New value']);
  const logRows = log.getRange(2, 1, log.getLastRow() - 1, 6).getDisplayValues();
  assert.ok(logRows.some(function (r) { return r[1] === 'JA Kirsten' && r[2] === 'm' && r[3] === 'm100'; }));
  assert.ok(!recs(ctx)['s:m100']);
  assert.deepStrictEqual(headerOf(ctx), REC_HEAD);
  assert.strictEqual(ctx._fuelWrites.n, 0);
});

t('phones cannot write pl or cm; photo cleanup removes old Service photos only', function () {
  boot();
  seed(ctx, 'u', 'uAdmin', user('uAdmin', 'JA Kirsten', { admin: 1 }));
  const admin = login(ctx, 'uAdmin', 'ph-a');
  const img = jpeg('keep-me');
  const id = 'phkeep01';
  assert.strictEqual(ctx.handle_({ action: 'photo', token: admin.token, device: 'ph-a', id: id, d: img.d, at: new Date().toISOString(), ref: 'maint TRK1 battery' }).ok, true);
  const oldId = 'phold02';
  assert.strictEqual(ctx.handle_({ action: 'photo', token: admin.token, device: 'ph-a', id: oldId, d: jpeg('old-me').d, at: new Date().toISOString(), ref: 'old' }).ok, true);
  const sh = ctx.sheet_(ctx.REC, ctx.REC_HEAD);
  const n = sh.getLastRow();
  const rows = sh.getRange(2, 1, n - 1, 8).getValues();
  rows.forEach(function (r, i) {
    if (r[0] !== 'pl:' + oldId) return;
    const d = JSON.parse(r[7]);
    d.at = new Date(Date.now() - 90 * 864e5).toISOString();
    sh.getRange(i + 2, 8, 1, 1).setValues([[JSON.stringify(d)]]);
  });
  const folder = ctx._folders[ctx.props_().getProperty('PHOTO_FOLDER_ID')];
  ctx._files.orphan = { id: 'orphan', bytes: Buffer.from('x'), created: new Date(Date.now() - 90 * 864e5), trashed: false, folder: folder.id };
  ctx._files.fresh = { id: 'fresh', bytes: Buffer.from('y'), created: new Date(), trashed: false, folder: folder.id };
  const nRemoved = ctx.photoCleanup();
  assert.strictEqual(nRemoved, 1);
  const old = JSON.parse(recs(ctx)['pl:' + oldId].data);
  assert.strictEqual(old.removed, 1);
  assert.strictEqual(old.url, '');
  assert.strictEqual(ctx._files[old.fid].trashed, true);
  assert.strictEqual(ctx.handle_({ action: 'photoget', token: admin.token, id: oldId }).error, 'photo_removed');
  const kept = JSON.parse(recs(ctx)['pl:' + id].data);
  assert.ok(!kept.removed);
  assert.strictEqual(ctx._files[kept.fid].trashed, false);
  assert.strictEqual(ctx._files.orphan.trashed, true);
  assert.strictEqual(ctx._files.fresh.trashed, false);
  const inject = ctx.handle_({ token: admin.token, device: 'ph-a', since: 0, changes: [{ kind: 'pl', id: 'phhack', u: Date.now(), del: 0, data: { id: 'phhack', fid: 'nope' } }] });
  assert.strictEqual(inject.accepted, 0);
  assert.ok(!recs(ctx)['pl:phhack']);
});

t('Check machine call, done queue, and no Fuel write', function () {
  boot();
  ctx.props_().setProperty('LINK_SECRET', TEST_SECRET);
  ctx.props_().setProperty('LINK_URL', TEST_URL);
  seed(ctx, 'u', 'uStaff', user('uStaff', 'Sam Staff', { perms: ['view', 'add'] }));
  const staff = login(ctx, 'uStaff', 'ph-s').token;
  const call = { fid: 'fill42', vid: 'TRK1', name: 'Water truck', date: '2026-11-08', rate: 12.26, normal: 8 };
  const bad = ctx.handle_({ action: 'linkCall', secret: 'nope-nope-nope-nope', call: call });
  assert.strictEqual(bad.error, 'bad_link');
  const ok = ctx.handle_({ action: 'linkCall', secret: TEST_SECRET, call: call });
  assert.strictEqual(ok.ok, true);
  assert.strictEqual(ctx.handle_({ action: 'linkCall', secret: TEST_SECRET, call: call }).dup, true);
  const cm = JSON.parse(recs(ctx)['cm:fill42'].data);
  assert.strictEqual(cm.vid, 'TRK1');
  assert.strictEqual(cm.rate, 12.3);
  assert.strictEqual(cm.normal, 8);
  assert.ok(!recs(ctx)['pl:fill42']);
  assert.ok(!recs(ctx)['m:fill42']);
  const phoneCm = ctx.handle_({ token: staff, device: 'ph-s', since: 0, changes: [{ kind: 'cm', id: 'fill99', u: Date.now(), del: 0, data: { id: 'fill99', vid: 'TRK1' } }] });
  assert.strictEqual(phoneCm.accepted, 0);
  assert.ok(!recs(ctx)['cm:fill99']);
  const cd = { id: 'fill42', by: 'Sam Staff', at: '2026-11-08T09:00:00.000Z' };
  const done = ctx.handle_({ token: staff, device: 'ph-s', since: 0, changes: [{ kind: 'cd', id: 'fill42', u: 5000, del: 0, data: cd }] });
  assert.strictEqual(done.accepted, 1);
  const queued = JSON.parse(ctx.props_().getProperty('LQ_fill42'));
  assert.deepStrictEqual(Object.keys(queued).sort(), ['action', 'fid']);
  assert.strictEqual(queued.action, 'linkDone');
  assert.strictEqual(queued.fid, 'fill42');
  assert.strictEqual(ctx.linkFlush(), 1);
  assert.strictEqual(ctx._fetches.length, 1);
  assert.strictEqual(ctx._fetches[0].url, TEST_URL);
  const sent = JSON.parse(ctx._fetches[0].payload);
  assert.deepStrictEqual(Object.keys(sent).sort(), ['action', 'fid', 'secret']);
  assert.strictEqual(sent.secret, TEST_SECRET);
  assert.strictEqual(sent.fid, 'fill42');
  assert.ok(!ctx.props_().getProperty('LQ_fill42'));
  assert.ok(ctx.props_().getProperty('LS_fill42'));
  assert.strictEqual(ctx.linkFlush(), 0);
  assert.ok(recs(ctx)['cm:fill42']);
  assert.ok(recs(ctx)['cd:fill42']);
  assert.strictEqual(ctx._fuelWrites.n, 0);
  assert.strictEqual(ctx._opened.length, 0);
});

t('admin re-sign-in, main-admin protection, removed staff token stops', function () {
  boot();
  seed(ctx, 'u', 'uAdmin', user('uAdmin', 'JA Kirsten', { admin: 1, dev: 'ph-a' }));
  seed(ctx, 'u', 'u2', user('u2', 'Second Admin', { admin: 1 }));
  seed(ctx, 'u', 'uStaff', user('uStaff', 'Sam Staff', { perms: ['view', 'add'] }));
  const first = login(ctx, 'uAdmin', 'ph-a');
  assert.strictEqual(first.ok, true);
  assert.strictEqual(ctx.handle_({ token: first.token, device: 'ph-a', since: 0, changes: [] }).ok, true);
  const second = login(ctx, 'uAdmin', 'ph-a');
  assert.strictEqual(second.ok, true);
  assert.notStrictEqual(second.token, first.token);
  assert.strictEqual(ctx.handle_({ token: first.token, device: 'ph-a', since: 0, changes: [] }).error, 'bad_token');
  assert.strictEqual(ctx.handle_({ token: second.token, device: 'ph-a', since: 0, changes: [] }).ok, true);
  const baseU = recs(ctx)['u:uAdmin'].u;
  const keep = ctx.handle_({ token: second.token, device: 'ph-a', since: 0, changes: [{ kind: 'u', id: 'uAdmin', u: baseU + 10, del: 0, data: { id: 'uAdmin', name: 'JA Kirsten', status: 'active', admin: 1, dev: 'ph-a', at: '2026-11-08T00:00:00.000Z' } }] });
  assert.strictEqual(keep.accepted, 1);
  assert.strictEqual(JSON.parse(recs(ctx)['u:uAdmin'].data).pinH, PIN);
  const reset = ctx.handle_({ token: second.token, device: 'ph-a', since: 0, changes: [{ kind: 'u', id: 'uAdmin', u: baseU + 20, del: 0, data: { id: 'uAdmin', name: 'JA Kirsten', status: 'active', admin: 1, pinReset: baseU + 20, dev: 'ph-a', at: '2026-11-08T00:00:00.000Z' } }] });
  assert.strictEqual(reset.accepted, 1);
  assert.ok(!JSON.parse(recs(ctx)['u:uAdmin'].data).pinH);
  assert.strictEqual(ctx.handle_({ action: 'status', name: 'JA Kirsten' }).hasPin, false);
  const set = ctx.handle_({ action: 'setpin', id: 'uAdmin', salt: 'newsalt', pinH: PIN2, device: 'ph-a' });
  assert.strictEqual(set.ok, true);
  assert.strictEqual(ctx.handle_({ token: second.token, device: 'ph-a', since: 0, changes: [] }).error, 'bad_token');
  assert.strictEqual(ctx.handle_({ token: set.token, device: 'ph-a', since: 0, changes: [] }).ok, true);
  assert.strictEqual(login(ctx, 'uAdmin', 'ph-a', PIN).error, 'bad_pin');
  const again = login(ctx, 'uAdmin', 'ph-a', PIN2);
  assert.strictEqual(again.ok, true);
  const other = login(ctx, 'u2', 'ph-2').token;
  const protect = ctx.handle_({ token: other, device: 'ph-2', since: 0, changes: [{ kind: 'u', id: 'uAdmin', u: baseU + 40, del: 1, data: { discarded: 1 } }] });
  assert.ok(protect.refusedDiscard && protect.refusedDiscard.indexOf('uAdmin') >= 0);
  assert.strictEqual(JSON.parse(recs(ctx)['u:uAdmin'].data).status, 'active');
  const staffTok = login(ctx, 'uStaff', 'ph-s').token;
  const staffU = recs(ctx)['u:uStaff'].u;
  const rm = ctx.handle_({ token: again.token, device: 'ph-a', since: 0, changes: [{ kind: 'u', id: 'uStaff', u: staffU + 10, del: 1, data: { discarded: 1 } }] });
  assert.strictEqual(rm.accepted, 1);
  assert.strictEqual(ctx.handle_({ token: staffTok, device: 'ph-s', since: 0, changes: [] }).error, 'bad_token');
  const mine = ctx.handle_({ token: again.token, device: 'ph-a', since: 0, changes: [] });
  const mineUser = mine.records.filter(function (r) { return r.kind === 'u' && r.id === 'uAdmin'; })[0];
  assert.ok(mineUser && JSON.parse(mineUser.data).pinH === PIN2);
  const otherView = ctx.handle_({ token: other, device: 'ph-2', since: 0, changes: [] });
  const seen = otherView.records.filter(function (r) { return r.kind === 'u' && r.id === 'uAdmin'; })[0];
  assert.ok(seen && !JSON.parse(seen.data).pinH && JSON.parse(seen.data).hasPin === 1);
});

t('Fuel copy stays read-only and Service records stay in this Sheet', function () {
  boot();
  ctx.props_().deleteProperty('FUEL_PULLED');
  seed(ctx, 'u', 'uAdmin', user('uAdmin', 'JA Kirsten', { admin: 1 }));
  const admin = login(ctx, 'uAdmin', 'ph-a').token;
  const sync = ctx.handle_({ token: admin, device: 'ph-a', since: 0, changes: [{ kind: 'fv', id: 'EVIL', u: Date.now(), del: 0, data: { id: 'EVIL', fleet: 'EVIL' } }] });
  assert.ok(!recs(ctx)['fv:EVIL']);
  assert.ok(recs(ctx)['fv:TRK1']);
  assert.ok(!recs(ctx)['fv:Kouri']);
  assert.ok(!recs(ctx)['fv:STOP1']);
  assert.strictEqual(ctx._fuelWrites.n, 0);
  assert.ok(ctx._opened.length >= 1);
  assert.strictEqual(sync.records.filter(function (r) { return r.kind === 'fv' && r.id === 'TRK1'; }).length, 1);
  assert.strictEqual(ctx._folderNames.length, 0);
});

t('setupOnce adds photoCleanup and does not drop pullFuel', function () {
  boot();
  const out = JSON.parse(ctx.setupOnce());
  assert.strictEqual(out.sheet, 'renamed');
  assert.ok(out.triggers.indexOf('pullFuel:added') >= 0);
  assert.ok(out.triggers.indexOf('photoCleanup:added') >= 0);
  const hours = {};
  ctx._triggers.forEach(function (tr) { hours[tr.fn] = tr.hour; });
  assert.strictEqual(hours.photoCleanup, 3);
  assert.strictEqual(hours.pullFuel, 5);
  assert.ok(ctx._triggers.some(function (tr) { return tr.fn === 'linkFlush' && tr.every === 5; }));
  assert.strictEqual(out.link, 'off (no secret/url)');
  assert.strictEqual(ctx.doGet().getContent(), JSON.stringify({ ok: true, app: 'TMC Service sync', hint: 'running' }));
});

t('artifact setupOnce keeps link armed without printing secrets', function () {
  if (!fs.existsSync(ART) || !fs.existsSync(LIVE)) { results.push('    skipped (artifact or live backup not on disk)'); return; }
  const art = fs.readFileSync(ART, 'utf8');
  const live = fs.readFileSync(LIVE, 'utf8');
  const box = load(art);
  const out = JSON.parse(box.setupOnce());
  assert.strictEqual(out.link, 'on');
  const got = box.props_().getProperty('LINK_SECRET');
  const url = box.props_().getProperty('LINK_URL');
  const want = linkConsts(live);
  if (got !== want.secret || url !== want.url) throw new Error('artifact setupOnce did not store the live link constants');
  assert.ok(box._triggers.some(function (tr) { return tr.fn === 'photoCleanup' && tr.hour === 3; }));
  assert.ok(box._triggers.some(function (tr) { return tr.fn === 'linkFlush'; }));
  assert.strictEqual(box._folderNames.length, 0);
});

const failed = results.filter(function (r) { return r.indexOf('FAIL') === 0; });
results.forEach(function (r) { console.log(r); });
console.log(failed.length ? failed.length + ' failed' : results.filter(function (r) { return r.indexOf('ok') === 0; }).length + ' passed');
process.exit(failed.length ? 1 : 0);
