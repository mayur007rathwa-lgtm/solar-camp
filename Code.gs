/**
 * MGVCL Bodeli Division – Solar Camp Planning (18-11-2026 to 30-11-2026)
 * Google Apps Script backend for the GitHub Pages SDN entry page.
 *
 * STEPS: 1) Paste this code in Extensions > Apps Script of a blank Google Sheet
 *        2) Run setup() once  3) Deploy > New deployment > Web app
 *           (Execute as: Me, Who has access: Anyone) and copy the Web App URL.
 */

var SDNS = ['Bodeli', 'Pavijetpur', 'Tejgadh', 'Chhotaudepur', 'Kawant', 'Naswadi'];
var START = new Date(2026, 10, 18), END = new Date(2026, 10, 30);
var DEFAULT_TARGET = 25;

var FIELDS = ['id', 'date', 'sdn', 'village', 'venue',
  'je_name', 'je_desig', 'je_mobile',
  's1_name', 's1_desig', 's2_name', 's2_desig',
  'a1_name', 'a1_person', 'a1_mobile',
  'a2_name', 'a2_person', 'a2_mobile',
  'a3_name', 'a3_person', 'a3_mobile',
  'target', 'reg_upto', 'reg_above', 'vendor', 'install', 'remarks',
  'updated_at', 'updated_by'];

var HEADERS = ['ID', 'Date', 'SDN', 'Village / Location', 'Camp Venue',
  'JE / DE Name', 'JE / DE Designation', 'JE / DE Mobile',
  'Staff Name - 1', 'Designation - 1', 'Staff Name - 2', 'Designation - 2',
  'Agency - 1 Name', 'Agency - 1 Person Name', 'Agency - 1 Mobile No.',
  'Agency - 2 Name', 'Agency - 2 Person Name', 'Agency - 2 Mobile No.',
  'Agency - 3 Name', 'Agency - 3 Person Name', 'Agency - 3 Mobile No.',
  'Target Registration (Minimum)', 'Registration Done (Nos.) Up to 2.18 KW',
  'Registration Done (Nos.) Above 2.18 KW', 'Vendor Selection Done (Nos.)',
  'No. of Installation (Nos.)', 'Remarks', 'Last Updated', 'Updated By'];

var NUM = ['target', 'reg_upto', 'reg_above', 'vendor', 'install'];

/* ---------- one-time setup ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var e = ss.getSheetByName('Entries') || ss.insertSheet('Entries');
  if (e.getLastRow() < 2) {
    e.clear();
    e.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF').setWrap(true);
    var rows = [];
    for (var d = new Date(START); d <= END; d.setDate(d.getDate() + 1)) {
      var ds = Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
      SDNS.forEach(function (s) {
        var r = FIELDS.map(function () { return ''; });
        r[0] = ds + '_' + s + '_1'; r[1] = ds; r[2] = s;
        r[FIELDS.indexOf('target')] = DEFAULT_TARGET;
        rows.push(r);
      });
    }
    e.getRange(2, 1, rows.length, FIELDS.length).setNumberFormat('@').setValues(rows);
    e.setFrozenRows(1);
  }
  var u = ss.getSheetByName('Users') || ss.insertSheet('Users');
  if (u.getLastRow() < 2) {
    u.clear();
    u.getRange(1, 1, 1, 3).setValues([['Login', 'PIN', 'Role']]).setFontWeight('bold');
    var list = SDNS.map(function (s) { return [s, pin(), 'SDN']; });
    list.push(['Division', pin(), 'ADMIN']);
    u.getRange(2, 1, list.length, 3).setNumberFormat('@').setValues(list);
  }
  var first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 2) ss.deleteSheet(first);
  SpreadsheetApp.getUi && Logger.log('Setup done. See Users sheet for PINs.');
}
function pin() { return String(Math.floor(1000 + Math.random() * 9000)); }

/* ---------- API ---------- */
function doGet() {
  return out({ ok: true, msg: 'Solar Camp API running' });
}

function doPost(ev) {
  try {
    var req = JSON.parse(ev.postData.contents);
    var user = auth(req.login, req.pin);
    if (!user) return out({ ok: false, error: 'Invalid login or PIN' });
    if (req.action === 'login') return out({ ok: true, role: user.role, login: user.login });
    if (req.action === 'list') return out({ ok: true, rows: readRows(user.role === 'ADMIN' ? null : user.login) });
    if (req.action === 'save') return out(saveRows(user, req.rows || [], req.deleted || []));
    if (req.action === 'photos') return out({ ok: true, photos: listPhotos(user.role === 'ADMIN' ? null : user.login) });
    if (req.action === 'upload') return out(uploadPhoto(user, req));
    if (req.action === 'delphoto') return out(deletePhoto(user, req.pid));
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function auth(login, p) {
  var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getDisplayValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === login && v[i][1] === String(p)) return { login: v[i][0], role: v[i][2] };
  }
  return null;
}

function readRows(sdn) {
  var v = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entries').getDataRange().getDisplayValues();
  var res = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    if (sdn && v[i][2] !== sdn) continue;
    var o = {};
    FIELDS.forEach(function (f, j) { o[f] = v[i][j]; });
    res.push(o);
  }
  return res;
}

/* SDN saves its full set of rows: existing IDs are updated, new IDs appended, removed extra rows deleted.
   Division (ADMIN) saves only the rows it changed (+ list of deleted IDs); other rows stay as they are. */
function saveRows(user, rows, deleted) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entries');
    var v = sh.getDataRange().getDisplayValues();
    var admin = user.role === 'ADMIN';
    var sdn = admin ? null : user.login;
    var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm');
    var byId = {}, del = {};
    rows.forEach(function (r) {
      if (sdn && r.sdn !== sdn) throw new Error('You can save only ' + sdn + ' SDN rows');
      if (SDNS.indexOf(r.sdn) < 0) throw new Error('Invalid SDN: ' + r.sdn);
      byId[r.id] = r;
    });
    (deleted || []).forEach(function (id) { if (!/_1$/.test(id)) del[id] = true; });
    var keep = [v[0]];
    for (var i = 1; i < v.length; i++) {
      var id = v[i][0], rowSdn = v[i][2];
      if (byId[id] && (admin || rowSdn === sdn)) { keep.push(toArr(byId[id], now, user.login)); delete byId[id]; }
      else if (admin && del[id]) { /* extra row removed by Division */ }
      else if (!admin && rowSdn === sdn && !/_1$/.test(id) && rows.length) { /* extra row removed by SDN */ }
      else keep.push(v[i]);
    }
    Object.keys(byId).forEach(function (k) { keep.push(toArr(byId[k], now, user.login)); });
    keep = [keep[0]].concat(keep.slice(1).sort(function (a, b) {
      return (a[1] + SDNS.indexOf(a[2]) + a[0]).localeCompare(b[1] + SDNS.indexOf(b[2]) + b[0]);
    }));
    sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), FIELDS.length).clearContent();
    sh.getRange(1, 1, keep.length, FIELDS.length).setNumberFormat('@').setValues(keep);
    return { ok: true, saved: rows.length, at: now };
  } finally {
    lock.releaseLock();
  }
}

function toArr(r, now, by) {
  return FIELDS.map(function (f) {
    if (f === 'updated_at') return now;
    if (f === 'updated_by') return by;
    var x = r[f] == null ? '' : String(r[f]).trim();
    if (NUM.indexOf(f) >= 0 && x !== '' && !/^\d+$/.test(x)) x = '';
    return x;
  });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- PHOTOS (Google Drive) ---------- */
var MAX_PHOTOS = 10;
var PH_HEAD = ['Photo ID', 'Entry ID', 'Date', 'SDN', 'Village', 'File ID', 'File Link', 'Uploaded At', 'Uploaded By'];

/* Run this once from the editor after pasting the new code, to give Drive permission */
function authorizeDrive() {
  rootFolder();
  photoSheet();
  Logger.log('Drive permission OK. Photo folder: ' + rootFolder().getUrl());
}

function photoSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Photos');
  if (!sh) {
    sh = ss.insertSheet('Photos');
    sh.getRange(1, 1, 1, PH_HEAD.length).setValues([PH_HEAD])
      .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rootFolder() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('PHOTO_ROOT');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('Solar Camp Photos - Bodeli Division (18-30 Nov 2026)');
  p.setProperty('PHOTO_ROOT', f.getId());
  return f;
}

function subFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function listPhotos(sdn) {
  var v = photoSheet().getDataRange().getDisplayValues(), res = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    if (sdn && v[i][3] !== sdn) continue;
    res.push({ pid: v[i][0], entry_id: v[i][1], date: v[i][2], sdn: v[i][3], village: v[i][4],
      file_id: v[i][5], view: v[i][6], at: v[i][7], by: v[i][8] });
  }
  return res;
}

function uploadPhoto(user, req) {
  var admin = user.role === 'ADMIN';
  if (!admin && req.sdn !== user.login) return { ok: false, error: 'You can upload only ' + user.login + ' SDN photos' };
  if (SDNS.indexOf(req.sdn) < 0) return { ok: false, error: 'Invalid SDN' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(req.date || '')) return { ok: false, error: 'Invalid date' };
  if (!req.data) return { ok: false, error: 'No photo data' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var have = listPhotos(req.sdn).filter(function (p) { return p.entry_id === req.entryId; }).length;
    if (have >= MAX_PHOTOS) return { ok: false, error: 'Maximum ' + MAX_PHOTOS + ' photos allowed for one camp' };
    var dmy = req.date.split('-').reverse().join('-');
    var folder = subFolder(subFolder(rootFolder(), dmy), req.sdn);
    var now = new Date();
    var stamp = Utilities.formatDate(now, 'Asia/Kolkata', 'dd-MM-yyyy HH:mm');
    var name = dmy + '_' + req.sdn + '_' + (req.village || 'Camp').replace(/[^A-Za-z0-9 ]/g, '').trim().replace(/ +/g, '-') +
      '_' + Utilities.formatDate(now, 'Asia/Kolkata', 'HHmmss') + '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(req.data), 'image/jpeg', name);
    var file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    var rec = { pid: 'P' + now.getTime() + Math.floor(Math.random() * 1000), entry_id: req.entryId, date: req.date,
      sdn: req.sdn, village: req.village || '', file_id: file.getId(), view: file.getUrl(), at: stamp, by: user.login };
    photoSheet().appendRow([rec.pid, rec.entry_id, rec.date, rec.sdn, rec.village, rec.file_id, rec.view, rec.at, rec.by]);
    return { ok: true, photo: rec };
  } finally {
    lock.releaseLock();
  }
}

function deletePhoto(user, pid) {
  var sh = photoSheet(), v = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] !== pid) continue;
    if (user.role !== 'ADMIN' && v[i][3] !== user.login) return { ok: false, error: 'Not allowed' };
    try { DriveApp.getFileById(v[i][5]).setTrashed(true); } catch (e) {}
    sh.deleteRow(i + 1);
    return { ok: true };
  }
  return { ok: false, error: 'Photo not found' };
}
