/**
 * Ento field app — Google Sheet receiver
 *
 * Setup (about 10 minutes, once):
 *  1. Create a new Google Sheet. Name it e.g. "Ento surveillance data".
 *  2. Extensions -> Apps Script. Delete anything there and paste this whole file.
 *  3. Click Deploy -> New deployment -> Select type "Web app".
 *       Description:  ento receiver
 *       Execute as:   Me
 *       Who has access: Anyone
 *     Deploy, authorise when Google asks, then copy the Web app URL
 *     (it looks like https://script.google.com/macros/s/AKfy.../exec).
 *  4. In the app: Dashboard -> Settings -> paste the URL into "Google Sheet URL",
 *     then press "Sync now".
 *
 * What it does: each record is appended to a tab named after its form type
 * (household, install, collect, adult, larvae, lab). Columns are created from the
 * record's own keys, so new fields appear automatically at the end of the header row.
 * A record already present (same record_id) is updated in place rather than duplicated.
 *
 * A tab named _history logs every write: when it arrived, whether it was new or an
 * update, which fields changed, and who sent it. Nothing is ever overwritten there,
 * so it is the audit trail of the dataset. Tabs whose name starts with "_" are
 * skipped when the app pulls data back, so the log never becomes records.
 */

var LOG_SHEET = '_history';
// Bumped whenever this file changes, so the app can show which version it reached.
var RECEIVER_VERSION = '2026-08-16 delete+log+admin scope+accounts';
var LOG_COLS = ['logged_at', 'action', 'record_id', 'form_type', 'household_id', 'trap_id',
  'trap_type', 'collected_date', 'period_label', 'collector', 'cluster', 'changed_fields', 'sheet_row'];

/* ─────────────────────────────────────────────────────────────────────────────
 * Accounts
 *
 * A tab named _users holds one row per person allowed to use the app:
 *   username | name | pin_hash | role | cluster | active | created_at | last_login
 *
 * Nobody types a PIN into this sheet directly — use Ento → Add or reset a user,
 * which hashes it. The app posts { action:'login', username, pin }; a successful
 * login returns a signed token that every later sync must carry.
 *
 * Writes and deletes always require a valid token. Reads (action=records) are
 * left open by default so the dashboard keeps working; set ALLOW_ANON_READ to
 * false once the dashboard also passes a token.
 * ────────────────────────────────────────────────────────────────────────── */

var USERS_SHEET = '_users';
var USER_COLS = ['username', 'name', 'pin_hash', 'role', 'cluster', 'active', 'created_at', 'last_login'];
var TOKEN_DAYS = 30;
var ALLOW_ANON_READ = true;

// The signing secret. Set ENTO_SECRET in Project Settings → Script Properties;
// the fallback below only exists so a fresh copy works before that is done.
function serverKey_() {
  return PropertiesService.getScriptProperties().getProperty('ENTO_SECRET')
    || 'ento-change-this-secret-2026';
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

// Same recipe as the app, so the phone can check a cached PIN while offline.
function pinHash_(username, pin) {
  return sha256Hex_('ento:' + String(username).toLowerCase() + ':' + String(pin));
}

function sign_(text) {
  var bytes = Utilities.computeHmacSha256Signature(text, serverKey_());
  return bytes.map(function (b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findUser_(username) {
  var sh = usersSheet_();
  if (sh.getLastRow() < 2) return null;
  var vals = sh.getRange(1, 1, sh.getLastRow(), USER_COLS.length).getValues();
  var header = vals[0].map(function (h) { return String(h).trim(); });
  for (var i = 1; i < vals.length; i++) {
    var row = {};
    for (var c = 0; c < header.length; c++) row[header[c]] = vals[i][c];
    if (String(row.username || '').trim().toLowerCase() === String(username).trim().toLowerCase()) {
      row._row = i + 1;
      return row;
    }
  }
  return null;
}

function login_(body) {
  var username = String(body.username || '').trim();
  var pin = String(body.pin || '').trim();
  if (!username || !pin) return { ok: false, error: 'username and PIN are required' };
  var user = findUser_(username);
  if (!user) return { ok: false, error: 'no account with that username' };
  if (String(user.active).toLowerCase() === 'no' || String(user.active).toLowerCase() === 'false') {
    return { ok: false, error: 'this account has been deactivated' };
  }
  if (String(user.pin_hash || '').trim().toLowerCase() !== pinHash_(user.username, pin)) {
    return { ok: false, error: 'wrong PIN' };
  }
  var sh = usersSheet_();
  var lastCol = USER_COLS.indexOf('last_login') + 1;
  sh.getRange(user._row, lastCol).setValue(new Date());
  var exp = new Date().getTime() + TOKEN_DAYS * 86400000;
  var payload = String(user.username).trim() + '|' + exp;
  return {
    ok: true,
    token: Utilities.base64EncodeWebSafe(payload + '|' + sign_(payload)),
    user: {
      username: String(user.username).trim(),
      name: String(user.name || user.username).trim(),
      role: String(user.role || 'collector').trim(),
      cluster: String(user.cluster || '').trim()
    },
    version: RECEIVER_VERSION
  };
}

function checkToken_(token) {
  if (!token) return { ok: false, error: 'sign in required' };
  var parts;
  try {
    parts = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString().split('|');
  } catch (err) {
    return { ok: false, error: 'bad token' };
  }
  if (parts.length !== 3) return { ok: false, error: 'bad token' };
  if (sign_(parts[0] + '|' + parts[1]) !== parts[2]) return { ok: false, error: 'bad token' };
  if (Number(parts[1]) < new Date().getTime()) return { ok: false, error: 'session expired — sign in again' };
  return { ok: true, username: parts[0] };
}

function logRows(ss, rows) {
  if (!rows.length) return;
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);   // logged_at
    sheet.setColumnWidth(11, 320);  // changed_fields
  }
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, LOG_COLS.length).setValues(rows);
  // Readable timestamps, and the newest entry always visible on open.
  sheet.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  try { sheet.autoResizeColumns(2, 4); } catch (e) {}
}

/**
 * Deletion. The phone posts { deletes: [{ household_id: 'BTB-0012' }, { record_id: '...' }] }
 * when a household or a single record is deleted in the app. Every matching row is
 * removed from every data tab and the removal is logged in _history, so the audit
 * trail keeps what was deleted after the row itself is gone.
 */
function deleteRows_(ss, targets, stamp, log) {
  var hh = {}, ids = {};
  targets.forEach(function (t) {
    if (!t) return;
    if (t.household_id) hh[String(t.household_id)] = 1;
    if (t.record_id) ids[String(t.record_id)] = 1;
  });
  var removed = 0;

  ss.getSheets().forEach(function (sheet) {
    var nameOfSheet = sheet.getName();
    if (nameOfSheet.charAt(0) === '_') return;
    var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;
    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idCol = header.indexOf('record_id');
    var hhCol = header.indexOf('household_id');
    var typeCol = header.indexOf('form_type');
    if (idCol < 0 && hhCol < 0) return;
    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // bottom up, so deleting does not shift the rows still to check
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var rid = idCol >= 0 ? String(row[idCol] || '') : '';
      var hid = hhCol >= 0 ? String(row[hhCol] || '') : '';
      if (!(rid && ids[rid]) && !(hid && hh[hid])) continue;
      sheet.deleteRow(i + 2);
      removed++;
      log.push([stamp, 'deleted (app)', rid, typeCol >= 0 ? row[typeCol] : nameOfSheet,
        hid, '', '', '', '', '', '', 'row removed', '']);
    }
  });
  return removed;
}

function doPost(e) {
  var out = { ok: false };
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'login') return jsonOut_(login_(body));
    var auth = checkToken_(body.token);
    if (!auth.ok) return jsonOut_({ ok: false, error: auth.error, auth: 'required' });
    body.auth_username = auth.username;
    var records = body.records || [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var written = 0;
    var log = [];
    var stamp = new Date();
    var deleted = body.deletes && body.deletes.length
      ? deleteRows_(ss, body.deletes, stamp, log) : 0;

    records.forEach(function (rec) {
      var type = rec.form_type || rec.type || 'other';
      var sheet = ss.getSheetByName(type) || ss.insertSheet(type);
      var header = sheet.getLastRow() > 0
        ? sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
        : [];

      // add any new keys to the header
      Object.keys(rec).forEach(function (k) {
        if (header.indexOf(k) === -1) header.push(k);
      });
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
      sheet.getRange(1, 1, 1, header.length).setFontWeight('bold');
      sheet.setFrozenRows(1);

      var row = header.map(function (k) {
        var v = rec[k];
        return v === undefined || v === null ? '' : v;
      });

      // Date columns as real dates, so the sheet can sort and chart them.
      var dateCols = ['collected_date', 'registered_date', 'synced_at', 'collected_at_utc', 'entered_at_utc'];

      // update in place when this record_id is already there
      var idCol = header.indexOf('record_id') + 1;
      var existing = 0;
      if (idCol > 0 && sheet.getLastRow() > 1) {
        var ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(rec.record_id)) { existing = i + 2; break; }
        }
      }
      // what changed, for the audit trail
      var changed = [];
      if (existing) {
        var before = sheet.getRange(existing, 1, 1, header.length).getValues()[0];
        for (var c = 0; c < header.length; c++) {
          if (String(before[c] === undefined ? '' : before[c]) !== String(row[c])) changed.push(header[c]);
        }
        sheet.getRange(existing, 1, 1, row.length).setValues([row]);
      } else {
        sheet.appendRow(row);
      }
      var atRow = existing || sheet.getLastRow();
      dateCols.forEach(function (name) {
        var col = header.indexOf(name) + 1;
        if (col > 0) {
          sheet.getRange(atRow, col).setNumberFormat(
            name === 'collected_date' || name === 'registered_date' ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm');
        }
      });
      log.push([
        stamp, existing ? (changed.length ? 'updated' : 'unchanged') : 'new',
        rec.record_id || '', type, rec.household_id || '', rec.trap_id || '',
        rec.trap_type || rec.method || '',
        rec.collected_date || '', rec.period_label || '', rec.collector || '', rec.cluster || '',
        existing ? changed.join(', ') : 'all', atRow,
      ]);
      written += 1;
    });

    logRows(ss, log);
    out = { ok: true, count: written, deleted: deleted, version: RECEIVER_VERSION, at: new Date().toISOString() };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Pull: the app calls this to read every record back out of the sheet, so a phone
 * can show households and history collected by other phones.
 *   ...exec                  -> { ok: true, service: 'ento receiver' }
 *   ...exec?action=records   -> { ok: true, records: [...] }
 */
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  if (!ALLOW_ANON_READ && (action === 'records' || action === 'history')) {
    var g = checkToken_(e && e.parameter ? e.parameter.token : '');
    if (!g.ok) return jsonOut_({ ok: false, error: g.error, auth: 'required' });
  }
  if (action === 'history') {
    var log = [];
    try {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
      if (sh && sh.getLastRow() > 1) {
        var vals = sh.getRange(2, 1, sh.getLastRow() - 1, LOG_COLS.length).getValues();
        // newest first, capped so the phone never downloads the whole log
        for (var i = vals.length - 1; i >= 0 && log.length < 300; i--) {
          var row = {};
          for (var c = 0; c < LOG_COLS.length; c++) {
            var v = vals[i][c];
            if (v instanceof Date) v = Utilities.formatDate(v, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
            row[LOG_COLS[c]] = v === null ? '' : v;
          }
          log.push(row);
        }
      }
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, version: RECEIVER_VERSION, history: log }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action !== 'records') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'ento receiver', version: RECEIVER_VERSION }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var out = [];
  // Optional admin scope, any level: province, district, commune, village. Each value
  // matches either the gazetteer code or the name, so both are accepted.
  var LEVELS = ['province', 'district', 'commune', 'village'];
  var want = {};
  LEVELS.forEach(function (lv) {
    var v = e && e.parameter ? String(e.parameter[lv] || '').trim() : '';
    if (v) want[lv] = v;
  });
  try {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    sheets.forEach(function (sheet) {
      if (sheet.getName().charAt(0) === '_') return;
      if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
      var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      var header = values[0];
      for (var i = 1; i < values.length; i++) {
        var rec = {};
        var empty = true;
        for (var c = 0; c < header.length; c++) {
          var key = String(header[c] || '').trim();
          if (!key) continue;
          var v = values[i][c];
          if (v instanceof Date) v = Utilities.formatDate(v, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
          if (v !== '' && v !== null) empty = false;
          rec[key] = v === null ? '' : v;
        }
        if (!empty && rec.record_id) {
          // Admin scope keeps the payload small for a single team.
          var skip = false;
          LEVELS.forEach(function (lv) {
            if (skip || !want[lv]) return;
            var code = String(rec[lv + '_code'] || '').replace(/^0+/, '');
            var name = String(rec[lv] || '').toLowerCase();
            var asked = want[lv];
            if (code !== asked.replace(/^0+/, '') && name !== asked.toLowerCase()) skip = true;
          });
          if (skip) continue;
          out.push(rec);
        }
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, records: out }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ento')
    .addItem('Add or reset a user', 'addOrResetUser')
    .addItem('List users', 'listUsers')
    .addItem('Deactivate a user', 'deactivateUser')
    .addSeparator()
    .addItem('Create history log', 'ensureHistorySheet')
    .addToUi();
}

// Recreates the _history tab if it was deleted. Past entries cannot be recovered;
// logging resumes from the next write or deletion.
function ensureHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(LOG_SHEET)) {
    SpreadsheetApp.getUi().alert('The ' + LOG_SHEET + ' tab is already there.');
    return;
  }
  logRows(ss, [[new Date(), 'log created', '', '', '', '', '', '', '', '', '', 'logging resumes here', '']]);
  SpreadsheetApp.getUi().alert('Created the ' + LOG_SHEET + ' tab. Every write and deletion from now on is logged there.');
}

/* ─── Account admin, from the Ento menu ─────────────────────────────────── */

function addOrResetUser() {
  var ui = SpreadsheetApp.getUi();
  var u = ui.prompt('Username (no spaces, e.g. sokha)', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  var username = u.getResponseText().trim();
  if (!username) return;
  var p = ui.prompt('PIN for ' + username + ' (4–8 digits)', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var pin = p.getResponseText().trim();
  if (!pin) return;

  var existing = findUser_(username);
  var sh = usersSheet_();
  if (existing) {
    sh.getRange(existing._row, USER_COLS.indexOf('pin_hash') + 1).setValue(pinHash_(username, pin));
    sh.getRange(existing._row, USER_COLS.indexOf('active') + 1).setValue('yes');
    ui.alert('PIN reset for ' + username + '. The account is active.');
    return;
  }
  var n = ui.prompt('Full name for ' + username + ' (this is stamped on records)', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  var r = ui.prompt('Role: collector, supervisor or admin', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var c = ui.prompt('Cluster (optional, e.g. C2)', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  sh.appendRow([username, n.getResponseText().trim() || username, pinHash_(username, pin),
    r.getResponseText().trim() || 'collector', c.getResponseText().trim(), 'yes', new Date(), '']);
  ui.alert('Added ' + username + '. They can sign in on the app now.');
}

function deactivateUser() {
  var ui = SpreadsheetApp.getUi();
  var u = ui.prompt('Username to deactivate', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  var user = findUser_(u.getResponseText().trim());
  if (!user) { ui.alert('No account with that username.'); return; }
  usersSheet_().getRange(user._row, USER_COLS.indexOf('active') + 1).setValue('no');
  ui.alert(user.username + ' can no longer sign in. Records already collected are kept.');
}

function listUsers() {
  var sh = usersSheet_();
  if (sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('No users yet. Ento → Add or reset a user.'); return; }
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, USER_COLS.length).getValues();
  var lines = vals.map(function (r) {
    return r[0] + '  —  ' + r[1] + '  (' + (r[3] || 'collector') + ', ' +
      (String(r[5]).toLowerCase() === 'no' ? 'inactive' : 'active') + ')' +
      (r[7] ? '  last signed in ' + Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '');
  });
  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

