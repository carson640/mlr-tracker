/**
 * MLR Equipment Tracker — shared backend (Google Apps Script).
 * Container-bound to the "MLR Equipment Tracker (Data)" Google Sheet.
 *
 * Data tabs (auto-created on first run):
 *   Assets   — one row per piece of equipment
 *   Log      — activity log (append-only)
 *   Signoffs — cleaning/QC sign-offs (append-only, stored as JSON)
 *   Meta     — internal: version counter for change detection
 *
 * The client (Index.html) calls getState / getVersion / seedIfEmpty / pushDelta
 * via google.script.run. Writes are serialized with LockService so two people
 * editing from different jobsites can't clobber each other.
 */

var ASSET_COLS = ['id','assetNumber','barcode','category','type','manufacturer','model',
  'serialNumber','capacity','rentalCost','location','custodian','expectedReturn','status',
  'notes','cleanedBy','qcBy','lastCleanedAt','source','updatedAt','updatedBy','floor','room','vehicle','destination','jobNumber','jobName'];
var LOG_COLS = ['id','time','user','txt'];
var SIGNOFF_COLS = ['id','json'];

function doGet(e) {
  // API mode for the standalone live-camera scanner page: ?api=state returns JSON. Does not affect the normal app load.
  var __p = (e && e.parameter) || {};
  if (__p.api) { return apiResponse_(__p); }
  // Phone-camera deep link: ?scan=NUMBER (or ?find=NUMBER) from a QR label opens that piece.
  var scan = (e && e.parameter && (e.parameter.scan || e.parameter.find)) || '';
  var out;
  if (scan) {
    var content = HtmlService.createHtmlOutputFromFile('Index').getContent();
    var inject = '<script>window.__SCAN_PARAM=' + JSON.stringify(String(scan)) + ';<\/script>';
    content = content.indexOf('</head>') >= 0 ? content.replace('</head>', inject + '</head>') : (inject + content);
    out = HtmlService.createHtmlOutput(content);
  } else {
    out = HtmlService.createHtmlOutputFromFile('Index');
  }
  return out
    .setTitle('Equipment Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() { return SpreadsheetApp.getActive(); }

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSheets_() {
  var aSh = sheet_('Assets', ASSET_COLS);
  // Migration: if columns were added later (e.g. floor/room), extend the Assets header.
  if (aSh.getMaxColumns() < ASSET_COLS.length) {
    aSh.insertColumnsAfter(aSh.getMaxColumns(), ASSET_COLS.length - aSh.getMaxColumns());
  }
  if (aSh.getLastColumn() < ASSET_COLS.length) {
    aSh.getRange(1, 1, 1, ASSET_COLS.length).setValues([ASSET_COLS]).setFontWeight('bold');
  }
  sheet_('Log', LOG_COLS);
  sheet_('Signoffs', SIGNOFF_COLS);
  var meta = ss_().getSheetByName('Meta') || sheet_('Meta', ['key', 'value']);
  if (meta.getLastRow() < 2) meta.getRange(2, 1, 1, 2).setValues([['version', 0]]);
  return meta;
}

function getVersion() {
  var meta = ensureSheets_();
  var v = meta.getRange(2, 2).getValue();
  return Number(v) || 0;
}

function bumpVersion_() {
  var meta = ensureSheets_();
  var v = (Number(meta.getRange(2, 2).getValue()) || 0) + 1;
  meta.getRange(2, 2).setValue(v);
  return v;
}

function rowsToObjects_(sh, cols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0]) continue; // skip blank id
    var o = {};
    for (var c = 0; c < cols.length; c++) {
      var val = values[i][c];
      o[cols[c]] = (val === '' || val === null || val === undefined) ? '' : String(val);
    }
    out.push(o);
  }
  return out;
}

function getAssets_() { return rowsToObjects_(sheet_('Assets', ASSET_COLS), ASSET_COLS); }

function getLog_() {
  return rowsToObjects_(sheet_('Log', LOG_COLS), LOG_COLS);
}

function getSignoffs_() {
  var rows = rowsToObjects_(sheet_('Signoffs', SIGNOFF_COLS), SIGNOFF_COLS);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    try { var o = JSON.parse(rows[i].json); o.id = rows[i].id; out.push(o); } catch (e) {}
  }
  return out;
}

/** Full snapshot for the client. */
function getState() {
  ensureSheets_();
  return {
    assets: getAssets_(),
    log: getLog_(),
    signoffs: getSignoffs_(),
    version: getVersion()
  };
}

/** Build {id: rowNumber} for the Assets sheet (rowNumber is 1-based sheet row). */
function assetRowIndex_(sh) {
  var last = sh.getLastRow();
  var idx = {};
  if (last < 2) return idx;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0]) idx[ids[i][0]] = i + 2;
  return idx;
}

function assetToRow_(a) {
  var row = [];
  for (var c = 0; c < ASSET_COLS.length; c++) {
    var k = ASSET_COLS[c];
    row.push(a[k] == null ? '' : a[k]);
  }
  return row;
}

/** One-time population when the sheet is empty. Returns full state. */
function seedIfEmpty(seedAssets) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSheets_();
    var sh = sheet_('Assets', ASSET_COLS);
    if (sh.getLastRow() >= 2) return getState(); // already seeded by someone else
    if (seedAssets && seedAssets.length) {
      var rows = seedAssets.map(assetToRow_);
      sh.getRange(2, 1, rows.length, ASSET_COLS.length).setValues(rows);
    }
    bumpVersion_();
    return getState();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Apply a delta from one client.
 * payload = { changed:[asset...], deletedIds:[id...], newLogs:[{id,time,user,txt}...],
 *             newSignoffs:[{id,...}...], user:string }
 * Returns { ok:true, version }.
 */
function pushDelta(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSheets_();
    payload = payload || {};
    var stamp = new Date().toISOString();
    var user = payload.user || '';

    // ---- Assets: upsert changed + deletes in ONE read + ONE write ----
    // (A bulk check-out can change several hundred rows; doing one block write keeps it ~1-2s.)
    var aSh = sheet_('Assets', ASSET_COLS);
    var changed = payload.changed || [];
    var del = (payload.deletedIds || []).slice();
    if (changed.length || del.length) {
      var lastRow = aSh.getLastRow();
      var data = (lastRow >= 2) ? aSh.getRange(2, 1, lastRow - 1, ASSET_COLS.length).getValues() : [];
      var idx = {};
      for (var i = 0; i < data.length; i++) { if (data[i][0] !== '' && data[i][0] != null) idx[data[i][0]] = i; }
      // deletes
      if (del.length) {
        var delSet = {};
        for (var d = 0; d < del.length; d++) delSet[del[d]] = true;
        data = data.filter(function (r) { return !delSet[r[0]]; });
        idx = {};
        for (var j = 0; j < data.length; j++) { if (data[j][0] !== '' && data[j][0] != null) idx[data[j][0]] = j; }
      }
      // upsert changed
      for (var c = 0; c < changed.length; c++) {
        var a = changed[c];
        if (!a || !a.id) continue;
        a.updatedAt = stamp; a.updatedBy = user;
        var row = assetToRow_(a);
        if (idx[a.id] != null) { data[idx[a.id]] = row; }
        else { idx[a.id] = data.length; data.push(row); }
      }
      var oldCount = (lastRow >= 2) ? lastRow - 1 : 0;
      if (data.length) {
        aSh.getRange(2, 1, data.length, ASSET_COLS.length).setValues(data);
      }
      if (oldCount > data.length) {
        aSh.getRange(data.length + 2, 1, oldCount - data.length, ASSET_COLS.length).clearContent();
      }
    }

    // ---- Log: append new ----
    var newLogs = payload.newLogs || [];
    if (newLogs.length) {
      var lSh = sheet_('Log', LOG_COLS);
      var lrows = newLogs.map(function (l) { return [l.id || '', l.time || '', l.user || '', l.txt || '']; });
      lSh.getRange(lSh.getLastRow() + 1, 1, lrows.length, LOG_COLS.length).setValues(lrows);
    }

    // ---- Signoffs: append new (as JSON) ----
    var newSign = payload.newSignoffs || [];
    if (newSign.length) {
      var sSh = sheet_('Signoffs', SIGNOFF_COLS);
      var srows = newSign.map(function (s) { return [s.id || '', JSON.stringify(s)]; });
      sSh.getRange(sSh.getLastRow() + 1, 1, srows.length, SIGNOFF_COLS.length).setValues(srows);
    }

    var v = bumpVersion_();
    return { ok: true, version: v };
  } finally {
    lock.releaseLock();
  }
}

/* ── External API for the standalone live-camera scanner (additive; the normal app never calls these) ── */
function apiResponse_(p) {
  var data = (p.api === 'state') ? getState()
           : (p.api === 'version') ? getVersion()
           : { ok: false, error: 'unknown api' };
  return jsonOut_(data, p.callback);
}

function jsonOut_(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out = { ok: false, error: 'no data' };
  try {
    var body = (e && e.postData && e.postData.contents) || '{}';
    var req = JSON.parse(body);
    if (req.action === 'loadout') { out = applyLoadout_(req); }
    else if (req.action === 'delta') { out = pushDelta(req.payload || {}); }
    else { out = { ok: false, error: 'unknown action' }; }
  } catch (err) { out = { ok: false, error: String(err) }; }
  return jsonOut_(out, (e && e.parameter && e.parameter.callback) || '');
}

/* Apply a scanned load-out: mark each piece Checked out, onto a vehicle, headed to a job. One read + one write. */
function applyLoadout_(req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSheets_();
    var sh = sheet_('Assets', ASSET_COLS);
    var lastRow = sh.getLastRow();
    var data = (lastRow >= 2) ? sh.getRange(2, 1, lastRow - 1, ASSET_COLS.length).getValues() : [];
    var col = {}; for (var c = 0; c < ASSET_COLS.length; c++) col[ASSET_COLS[c]] = c;
    function norm(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
    var byCode = {};
    for (var i = 0; i < data.length; i++) {
      var r = data[i], keys = [r[col.assetNumber], r[col.barcode], r[col.serialNumber]];
      for (var k = 0; k < keys.length; k++) { var kk = norm(keys[k]); if (kk && !(kk in byCode)) byCode[kk] = i; }
    }
    var codes = req.codes || [];
    var veh = req.vehicle || '', dest = req.destination || '', cust = req.custodian || '';
    var floor = req.floor || '', room = req.room || '', ret = req.expectedReturn || '';
    var jobNo = req.jobNumber || '', jobNm = req.jobName || '';
    var stamp = new Date().toISOString(), user = req.user || 'scanner';
    var applied = 0, notfound = [];
    for (var x = 0; x < codes.length; x++) {
      var idx = byCode[norm(codes[x])];
      if (idx == null) { notfound.push(String(codes[x])); continue; }
      var row = data[idx];
      row[col.status] = 'Checked out';
      row[col.custodian] = cust || row[col.custodian] || 'Inventory';
      row[col.vehicle] = veh;
      if (veh) { row[col.location] = veh; row[col.destination] = dest; }
      else { if (dest) row[col.location] = dest; row[col.destination] = ''; }
      row[col.floor] = floor; row[col.room] = room; row[col.expectedReturn] = ret;
      row[col.jobNumber] = jobNo; row[col.jobName] = jobNm;
      row[col.updatedAt] = stamp; row[col.updatedBy] = user;
      applied++;
    }
    if (applied) {
      sh.getRange(2, 1, data.length, ASSET_COLS.length).setValues(data);
      sheet_('Log', LOG_COLS).appendRow([Utilities.getUuid(), stamp, user,
        'Live-scanned ' + applied + ' piece' + (applied === 1 ? '' : 's') + ' onto ' + (veh || '(no vehicle)') + (jobNo ? (' for job ' + jobNo) : '') + (dest ? (' to ' + dest) : '') + '.']);
      bumpVersion_();
    }
    return { ok: true, applied: applied, notfound: notfound, version: getVersion() };
  } finally {
    lock.releaseLock();
  }
}

/* ───────────────────────────────────────────────────────────────────────────
 * ONE-TIME CLEANUP — tidy the Assets columns so serial / specs / make-model
 * live in the right fields. Makes a full backup tab first, so it's reversible.
 * Run manually from the editor (select normalizeAssets, click Run).
 * ─────────────────────────────────────────────────────────────────────────── */
function normalizeAssets() {
  var lock = LockService.getScriptLock();
  lock.waitLock(120000);
  try {
    ensureSheets_();
    var ss = ss_();
    var sh = sheet_('Assets', ASSET_COLS);
    var last = sh.getLastRow();
    if (last < 2) return 'No data rows to tidy.';

    // 1) Full safety backup (a complete copy of the Assets tab).
    var stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'GMT', 'yyyyMMdd-HHmmss');
    var backup = 'Assets_backup_' + stamp;
    sh.copyTo(ss).setName(backup);

    var rng = sh.getRange(2, 1, last - 1, ASSET_COLS.length);
    var data = rng.getValues();
    var col = {}; for (var c = 0; c < ASSET_COLS.length; c++) col[ASSET_COLS[c]] = c;

    var S = function (v) { return String(v == null ? '' : v).trim(); };
    var isSpec = function (v) { v = S(v); return /\d/.test(v) && /(cfm|gallons?|gal\b|pints?|ppd|amps?|volts?|watts?|hp\b|btu|psi|gpm|inch(?:es)?|"|liters?|litres?|micron|µm|sq\.?\s?ft|cu\.?\s?ft)\b/i.test(v); };
    var isMoney = function (v) { v = S(v); return v === '$' || /^\$/.test(v) || /^\$?\d[\d,]*\.\d{2}$/.test(v); };
    var esc = function (s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); };
    var lastTok = function (s) { var p = S(s).split(/\s+/); return p.length ? p[p.length - 1] : ''; };
    var stripLast = function (s) { return S(s).replace(/\s*\S+\s*$/, '').trim(); };

    var n = 0;
    for (var i = 0; i < data.length; i++) {
      var r = data[i]; if (!S(r[col.id])) continue;
      var serial = S(r[col.serialNumber]), model = S(r[col.model]), mfr = S(r[col.manufacturer]), cap = S(r[col.capacity]);
      var before = [serial, model, mfr, cap].join('');

      // a spec ("500 CFM") sitting in the serial field → move it to capacity
      if (isSpec(serial)) { if (!isSpec(cap)) cap = serial; serial = ''; }
      // money / placeholder junk in the serial field → clear it
      if (serial === '—' || serial === '-' || isMoney(serial)) serial = '';

      if (serial) {
        // a real serial that also trails the name → remove the trailing copy from the name
        var re = new RegExp('[\\s#:_-]*' + esc(serial) + '\\s*$', 'i');
        if (re.test(model)) model = model.replace(re, '').trim();
        else if (re.test(mfr)) mfr = mfr.replace(re, '').trim();
      } else {
        // serial typed onto the end of the name (5+ consecutive digits) → pull it into the serial column
        var lm = lastTok(model);
        if (/\d{5,}/.test(lm) && !isSpec(lm)) { serial = lm.replace(/[^A-Za-z0-9-]/g, ''); model = stripLast(model); }
        else { var lf = lastTok(mfr); if (/\d{5,}/.test(lf) && !isSpec(lf)) { serial = lf.replace(/[^A-Za-z0-9-]/g, ''); mfr = stripLast(mfr); } }
      }

      // a stray serial-like number sitting in capacity → move to an empty serial; clear money in capacity
      if (cap && !isSpec(cap)) { if (isMoney(cap)) cap = ''; else if (/^\d{5,}$/.test(cap)) { if (!serial) serial = cap; cap = ''; } }

      // duplicate make == model → keep one
      if (mfr && model && mfr.toLowerCase() === model.toLowerCase()) mfr = '';
      // collapse a whole-string doubling inside the model ("Moab 85 Moab 85")
      var mt = model.split(/\s+/);
      if (mt.length >= 2 && mt.length % 2 === 0) { var h = mt.length / 2; if (mt.slice(0, h).join(' ').toLowerCase() === mt.slice(h).join(' ').toLowerCase()) model = mt.slice(0, h).join(' '); }

      var after = [serial, model, mfr, cap].join('');
      if (after !== before) {
        r[col.serialNumber] = serial; r[col.model] = model; r[col.manufacturer] = mfr; r[col.capacity] = cap;
        r[col.updatedAt] = new Date().toISOString(); r[col.updatedBy] = 'cleanup';
        n++;
      }
    }

    if (n) {
      rng.setValues(data);
      bumpVersion_();
      sheet_('Log', LOG_COLS).appendRow([Utilities.getUuid(), new Date().toISOString(), 'cleanup',
        'Tidied ' + n + ' row' + (n === 1 ? '' : 's') + ' — routed specs/serials/names to the right columns. Backup tab: ' + backup]);
    }
    return 'Done. Tidied ' + n + ' of ' + data.length + ' rows. Backup tab: ' + backup;
  } finally {
    lock.releaseLock();
  }
}
