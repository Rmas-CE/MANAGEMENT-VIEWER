/**
 * VaultSync — VIEWER MIRROR (separate Google Sheet for the Management Viewer)
 * ---------------------------------------------------------------------------
 * The boss viewer reads ONLY from this sheet. Every few minutes this script takes a
 * fresh copy (snapshot) of the Inventory Tool data and saves it here, so the bosses
 * never connect to the original Inventory Google Sheet. Nothing in the original sheet
 * or its Code.gs is changed, and nothing is ever written back to it.
 *
 * SETUP (one time):
 *  1) Create a NEW Google Sheet, e.g. "VaultSync Viewer Mirror".
 *  2) Extensions › Apps Script › delete the sample code › paste this whole file › Save.
 *  3) Reload the Google Sheet. A menu "VaultSync Viewer" appears › Set up (first time).
 *     Allow the permissions, then type the Inventory Tool Admin password.
 *     It takes the first snapshot and turns on the automatic refresh (every 5 minutes).
 *  4) Deploy › New deployment › type: Web app
 *        Execute as: Me      Who has access: Anyone      › Deploy
 *     Copy the Web app URL (ends in /exec) and paste it into inventory.html (MIRROR_URL).
 *
 * If the Admin password is changed: menu › Set / change Admin password.
 * The "Status" tab shows the last refresh time and any error.
 */
var SOURCE_URL = 'https://script.google.com/macros/s/AKfycbxJ7UqqR7WKe9iBu1y9QUm4UdFsZT-oNLnRdwNsG1M2i3DrjvW0YnXY1QHb4H22SQse/exec'; // original Inventory Tool (read only)
var REFRESH_MINUTES = 5;                 // allowed: 1, 5, 10, 15, 30
var CHUNK = 45000;                       // characters per cell (a cell holds up to 50,000)
var DETAIL_ACTIONS = ['getDeploymentPhotos', 'getReceipt', 'getPendingItem'];   // opened one at a time (proof photos, receipts)

/* ============================== MENU / SETUP ============================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('VaultSync Viewer')
    .addItem('Refresh now', 'refreshNow')
    .addItem('Set / change Admin password', 'setPassword')
    .addSeparator()
    .addItem('Set up (first time)', 'setup')
    .addToUi();
}

function setup() {
  if (!setPassword()) return;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshSnapshot').timeBased().everyMinutes(REFRESH_MINUTES).create();
  var r = refreshSnapshot();
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Done! First snapshot saved (' + r.info + ').\nAuto refresh: every ' + REFRESH_MINUTES + ' minutes.\n\nNext: Deploy › New deployment › Web app (Execute as: Me, Who has access: Anyone) and paste the URL into inventory.html.'
    : 'Could not take the snapshot:\n' + r.error + '\n\nCheck the Admin password (menu › Set / change Admin password).');
}

function setPassword() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Inventory Tool Admin password', 'Used only by this script to READ the inventory data. It is never shown on the viewer page.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText()) return false;
  PropertiesService.getScriptProperties().setProperty('ADMIN_PW', res.getResponseText());
  return true;
}

function refreshNow() {
  var r = refreshSnapshot();
  SpreadsheetApp.getUi().alert(r.ok ? 'Refreshed — ' + r.info : 'Refresh failed: ' + r.error);
}

/* ============================== SNAPSHOT ============================== */
function refreshSnapshot() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: 'Another refresh is running.' };
  try {
    var pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PW');
    if (!pw) throw new Error('Admin password not set yet (menu › Set up).');
    var data = callSource_('getData', { adminPassword: pw });
    var alerts = null;
    try { alerts = callSource_('getAlerts', { adminPassword: pw }); } catch (e) {}
    var at = new Date();
    var json = JSON.stringify({ at: at.toISOString(), data: data, alerts: alerts });
    writeSnapshot_(json);
    var info = (data.requests || []).length + ' requests, ' + (data.posmMaster || []).length + ' POSM, ' + Math.round(json.length / 1024) + ' KB';
    status_([['Last refresh', at], ['Result', 'OK'], ['Details', info], ['Auto refresh', 'every ' + REFRESH_MINUTES + ' minutes'], ['Source', SOURCE_URL]]);
    return { ok: true, info: info };
  } catch (e) {
    // The previous snapshot stays in place, so the viewer keeps showing the last good data.
    status_([['Last attempt', new Date()], ['Result', 'ERROR'], ['Details', String(e.message || e)], ['Auto refresh', 'every ' + REFRESH_MINUTES + ' minutes'], ['Source', SOURCE_URL]]);
    return { ok: false, error: String(e.message || e) };
  } finally {
    lock.releaseLock();
  }
}

function callSource_(action, payload) {
  var body = { action: action };
  Object.keys(payload || {}).forEach(function (k) { body[k] = payload[k]; });
  var resp = UrlFetchApp.fetch(SOURCE_URL, {
    method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify(body),
    muteHttpExceptions: true, followRedirects: true
  });
  var j;
  try { j = JSON.parse(resp.getContentText()); } catch (e) { throw new Error('Inventory Tool did not answer (HTTP ' + resp.getResponseCode() + ').'); }
  if (!j.success) throw new Error(j.error || 'Inventory Tool returned an error.');
  return j;
}

// Two snapshot tabs take turns, so a viewer never reads a half-written copy.
function writeSnapshot_(json) {
  var ss = SpreadsheetApp.getActive(), props = PropertiesService.getScriptProperties();
  var next = props.getProperty('ACTIVE') === 'Snapshot_A' ? 'Snapshot_B' : 'Snapshot_A';
  var sh = ss.getSheetByName(next) || ss.insertSheet(next);
  sh.clear();
  var rows = [];
  for (var i = 0; i < json.length; i += CHUNK) rows.push([json.substr(i, CHUNK)]);
  sh.getRange(1, 1, rows.length, 1).setNumberFormat('@').setValues(rows);
  if (!sh.isSheetHidden()) sh.hideSheet();
  props.setProperty('ACTIVE', next);
  props.setProperty('VERSION', String(Date.now()));
}

function readSnapshot_() {
  var props = PropertiesService.getScriptProperties();
  var ver = props.getProperty('VERSION'), active = props.getProperty('ACTIVE');
  if (!ver || !active) return null;
  var cache = CacheService.getScriptCache();
  var meta = cache.get('snap:' + ver);
  if (meta) {
    var keys = []; for (var i = 0; i < Number(meta); i++) keys.push('snap:' + ver + ':' + i);
    var got = cache.getAll(keys);
    if (keys.every(function (k) { return got[k] != null; })) return JSON.parse(keys.map(function (k) { return got[k]; }).join(''));
  }
  var sh = SpreadsheetApp.getActive().getSheetByName(active);
  if (!sh || sh.getLastRow() < 1) return null;
  var json = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(function (r) { return r[0]; }).join('');
  try {   // keep it in memory for 10 minutes so most viewer loads skip the sheet read
    var parts = {}, n = 0;
    for (var p = 0; p < json.length; p += 90000) parts['snap:' + ver + ':' + (n++)] = json.substr(p, 90000);
    if (n <= 900) { cache.putAll(parts, 600); cache.put('snap:' + ver, String(n), 600); }
  } catch (e) {}
  return JSON.parse(json);
}

function status_(rows) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Status') || ss.insertSheet('Status', 0);
  sh.clear();
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, rows.length, 1).setFontWeight('bold');
  sh.autoResizeColumns(1, 2);
}

/* ============================== WEB APP (viewer reads here) ============================== */
function doGet(e) {
  var a = (e && e.parameter && e.parameter.action) || '';
  if (a === 'checkSeeded') return out_({ seeded: true, orgSeeded: true });
  if (a === 'getImages') return out_(getImages_(String(e.parameter.codes || '')));
  if (a === 'status') { var s = readSnapshot_(); return out_({ success: !!s, at: s ? s.at : null }); }
  return out_({ success: false, error: 'View only.' });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) {}
  var a = String(body.action || '');
  if (a === 'getData' || a === 'getAlerts') {
    var s = readSnapshot_();
    if (!s) return out_({ success: false, error: 'Viewer data is not ready yet — open the VaultSync Viewer Mirror sheet › menu › Set up.' });
    if (a === 'getAlerts') return out_(s.alerts || { success: true });
    var d = s.data; d.success = true; d._snapshotAt = s.at;
    return out_(d);
  }
  if (DETAIL_ACTIONS.indexOf(a) >= 0 && body.id) return out_(detail_(a, String(body.id)));
  return out_({ success: false, error: 'View only — changes cannot be made from the management viewer.' });
}

// Proof photos / receipts are only fetched when a boss opens one, then kept for 6 hours.
function detail_(action, id) {
  var cache = CacheService.getScriptCache(), key = 'd:' + action + ':' + id;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  try {
    var r = callSource_(action, { id: id, adminPassword: PropertiesService.getScriptProperties().getProperty('ADMIN_PW') });
    var txt = JSON.stringify(r);
    if (txt.length < 95000) cache.put(key, txt, 21600);
    return r;
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

function getImages_(codes) {
  var list = codes.split(',').filter(String).slice(0, 60), cache = CacheService.getScriptCache(), results = {}, need = [];
  var got = cache.getAll(list.map(function (c) { return 'img:' + c; }));
  list.forEach(function (c) { if (got['img:' + c] != null) { if (got['img:' + c]) results[c] = got['img:' + c]; } else need.push(c); });
  if (need.length) {
    try {
      var resp = UrlFetchApp.fetch(SOURCE_URL + '?action=getImages&codes=' + encodeURIComponent(need.join(',')), { muteHttpExceptions: true, followRedirects: true });
      var r = (JSON.parse(resp.getContentText()) || {}).results || {};
      var put = {};
      need.forEach(function (c) { var v = r[c] || ''; if (v) results[c] = v; if (v.length < 95000) put['img:' + c] = v; });
      cache.putAll(put, 21600);
    } catch (e) {}
  }
  return { results: results };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
