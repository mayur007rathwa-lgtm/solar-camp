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
    if (req.action === 'save') return out(saveRows(user, req.rows || []));
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

/* SDN saves its full set of rows: existing IDs are updated, new IDs appended, removed extra rows deleted */
function saveRows(user, rows) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entries');
    var v = sh.getDataRange().getDisplayValues();
    var sdn = user.role === 'ADMIN' ? null : user.login;
    var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm');
    var byId = {};
    rows.forEach(function (r) {
      if (sdn && r.sdn !== sdn) throw new Error('You can save only ' + sdn + ' SDN rows');
      byId[r.id] = r;
    });
    var keep = [v[0]];
    for (var i = 1; i < v.length; i++) {
      var id = v[i][0], rowSdn = v[i][2];
      var mine = !sdn || rowSdn === sdn;
      if (mine && byId[id]) { keep.push(toArr(byId[id], now, user.login)); delete byId[id]; }
      else if (mine && !byId[id] && !/_1$/.test(id) && rows.length) { /* extra row removed by SDN */ }
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
