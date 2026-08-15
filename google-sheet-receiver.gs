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
var RECEIVER_VERSION = '2026-08-15 delete+log+admin scope';
var LOG_COLS = ['logged_at', 'action', 'record_id', 'form_type', 'household_id', 'trap_id',
  'trap_type', 'collected_date', 'period_label', 'collector', 'cluster', 'changed_fields', 'sheet_row'];

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

/* ─────────────────────────────────────────────────────────────────────────────
 * KoboToolbox import
 *
 * Pulls submissions straight from Kobo's API and writes them into this
 * spreadsheet through the same upsert path the phone app uses, so a record
 * imported twice updates rather than duplicates.
 *
 * Setup, once:
 *   1. In Kobo: Account Settings → Security → API key. Copy the token.
 *   2. Here: Project Settings (gear) → Script Properties → Add property
 *        KOBO_TOKEN = <the token>
 *      Keeping it in properties means the token is not in this file.
 *   3. Fill KOBO_ASSETS below with each form's uid (the string in the Kobo URL
 *      after /forms/, e.g. aXbYcZ...) and which form_type it maps to.
 *   4. Run importFromKobo once from the editor and accept the permissions.
 *   5. Run setupKoboTrigger once to import nightly.
 * ────────────────────────────────────────────────────────────────────────── */

var KOBO_BASE = 'https://kf.kobotoolbox.org';

// Fallback token, used only when the KOBO_TOKEN script property is empty.
// Prefer the property: gear icon → Script Properties → KOBO_TOKEN.
var KOBO_TOKEN_FALLBACK = '99b42bc60d930bf65fd78e85d2bace9a0c19d3ba';

var KOBO_ASSETS = [
  { uid: 'av4u3erwyNjH6EEtTcaBY2', form_type: 'household' },  // HH registration
  { uid: 'ayqLxooYAqWSgGTee9QJpP', form_type: 'install' },    // Trap installation
  { uid: 'apwRgJTHxnXumB4AuUa8Gn', form_type: 'lab' }         // Collection & lab result
];

// Kobo question names → the app's field names. Anything not listed is copied
// through with its group prefix stripped.
var KOBO_FIELDS = {
  hh_id: 'household_id',
  hh_code: 'household_id',
  village_name: 'village',
  commune_name: 'commune',
  district_name: 'district',
  province_name: 'province',
  install_date: 'collected_date',
  lab_date: 'collected_date',
  today: 'collected_date',
  total_containers: 'containers_inspected',
  total_positive: 'containers_positive',
  ovi_eggs_indoor: 'eggs_in',
  ovi_eggs_outdoor: 'eggs_out',
  ovi_eggs_total: 'eggs',
  prok_aeg_female: 'aegypti_f',
  prok_aeg_male: 'aegypti_m',
  prok_alb_female: 'albopictus_f',
  prok_alb_male: 'albopictus_m',
  prok_culex: 'culex',
  prok_adults_total: 'total_catch',
  larv_aeg: 'aegypti',
  larv_alb: 'albopictus',
  larv_culex: 'culex',
  larv_total: 'total_catch'
};

// Kobo's romanised / English place names → gazetteer code + Khmer name.
var KOBO_PLACE = {
  province: { 'Battambang': ['02', 'បាត់ដំបង'], 'Battambang Province': ['02', 'បាត់ដំបង'] },
  district: { 'Battambang City': ['0203', 'បាត់ដំបង'], 'Battambang': ['0203', 'បាត់ដំបង'],
    'Krong Battambang': ['0203', 'បាត់ដំបង'] },
  commune: { 'Toul Ta Ek': ['020301', 'ទួលតាឯក'], 'Preak Peahsdach': ['020302', 'ព្រែកព្រះសេ្ដច'],
    'Rattanak': ['020303', 'រតនៈ'], 'Svay Por': ['020310', 'ស្វាយពោ'] },
  village: { 'Toul Ta Ek': ['0203010400', 'ទួលតាឯក'], 'Otakom': ['0203010200', 'អូរតាគាំ ២'],
    'Kampong Krabei': ['0203100200', 'កំពង់ក្របី'], 'Kamakor': ['0203100400', 'កម្មករ'],
    'Rumchek': ['0203030500', 'រំចេក ៥'], 'Rattanak': ['0203030800', 'រតនៈ'],
    'Numkiem': ['0203020600', 'នំក្រៀប'], 'Ocheay': ['0203020400', 'អូរខ្ជាយ'] }
};

function koboPlaces_(rec) {
  ['province', 'district', 'commune', 'village'].forEach(function (level) {
    var nm = String(rec[level] || '').replace(/_v$/, '').trim();
    var hit = KOBO_PLACE[level][nm];
    if (hit) { if (hit[0]) rec[level + '_code'] = hit[0]; rec[level] = hit[1]; }
  });
  if (rec.commune_code && String(rec.commune_code).length === 6) {
    if (!rec.district_code) rec.district_code = String(rec.commune_code).slice(0, 4);
    if (!rec.province_code) rec.province_code = String(rec.commune_code).slice(0, 2);
    if (!rec.district) rec.district = 'បាត់ដំបង';
    if (!rec.province) rec.province = 'បាត់ដំបង';
  }
  if (rec.village) {
    rec.village = String(rec.village).replace(/_v$/, '');
    if (!rec.village_code) {
      rec.village_code = (rec.commune_code || 'x') + '-'
        + String(rec.village).replace(/[^A-Za-z0-9]+/g, '').toLowerCase();
    }
  }
  return rec;
}

function koboToken_() {
  var t = PropertiesService.getScriptProperties().getProperty('KOBO_TOKEN') || KOBO_TOKEN_FALLBACK;
  if (!t) throw new Error('Set KOBO_TOKEN in Script Properties.');
  return t;
}

function koboFetch_(uid) {
  var url = KOBO_BASE + '/api/v2/assets/' + uid + '/data/?format=json&limit=30000';
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Token ' + koboToken_() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Kobo returned ' + res.getResponseCode() + ' for ' + uid + ': ' + res.getContentText().slice(0, 200));
  }
  return JSON.parse(res.getContentText()).results || [];
}

// One Kobo submission → one app record.
function koboMap_(sub, formType, uid) {
  var rec = { form_type: formType, source: 'kobo' };
  Object.keys(sub).forEach(function (raw) {
    if (raw.charAt(0) === '_' || raw.indexOf('meta/') === 0 || raw.indexOf('formhub/') === 0) return;
    var name = raw.indexOf('/') >= 0 ? raw.split('/').pop() : raw;
    var key = KOBO_FIELDS[name] || name;
    var v = sub[raw];
    rec[key] = v === null || v === undefined ? '' : v;
  });
  // Kobo's own submission id keeps the record stable across imports.
  rec.record_id = 'KB-' + uid.slice(-4) + '-' + sub._id;
  if (!rec.collected_date && sub._submission_time) {
    rec.collected_date = String(sub._submission_time).slice(0, 10);
  }
  rec.kobo_submission_id = sub._id;
  koboPlaces_(rec);
  rec.synced_at = new Date().toISOString();
  if (sub._geolocation && sub._geolocation.length === 2 && sub._geolocation[0] !== null) {
    if (!rec.lat) { rec.lat = sub._geolocation[0]; rec.lon = sub._geolocation[1]; }
  }
  return rec;
}

// Run this first: logs one submission per form so the field names can be checked
// against KOBO_FIELDS before importing everything.
function peekKobo() {
  KOBO_ASSETS.forEach(function (a) {
    var subs = koboFetch_(a.uid);
    Logger.log('%s (%s): %s submissions', a.form_type, a.uid, subs.length);
    if (subs.length) {
      Logger.log('fields: %s', Object.keys(subs[0]).join(', '));
      Logger.log('mapped: %s', JSON.stringify(koboMap_(subs[0], a.form_type, a.uid)));
    }
  });
}

// Apps Script stops a run at six minutes, so the import works through the
// submissions in batches and remembers where it stopped. Run it again (or let
// the nightly trigger run) to continue; it finishes when the cursor resets to 0.
var KOBO_BATCH = 120;

function importFromKobo() {
  if (!KOBO_ASSETS.length) throw new Error('Fill KOBO_ASSETS with your form uids first.');
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty('KOBO_CURSOR') || 0);
  var started = new Date().getTime();

  var all = [];
  KOBO_ASSETS.forEach(function (a) {
    var subs;
    try {
      subs = koboFetch_(a.uid);
    } catch (err) {
      Logger.log('SKIPPED %s (%s): %s', a.form_type, a.uid, err.message);
      return;
    }
    subs.forEach(function (sub) { all.push(koboMap_(sub, a.form_type, a.uid)); });
  });
  if (!all.length) throw new Error('Kobo returned no submissions. Check the token, the server URL and the form uids.');

  var slice = all.slice(cursor, cursor + KOBO_BATCH);
  var out = doPost({ postData: { contents: JSON.stringify({ records: slice }) } });
  var parsed = JSON.parse(out.getContent());
  if (!parsed.ok) throw new Error('Writing to the sheet failed: ' + (parsed.error || 'unknown'));

  var next = cursor + slice.length;
  var done = next >= all.length;
  props.setProperty('KOBO_CURSOR', done ? '0' : String(next));
  Logger.log('Kobo import: %s of %s submissions written this run (%s s). %s',
    slice.length, all.length, Math.round((new Date().getTime() - started) / 1000),
    done ? 'Complete.' : 'Run again to continue from ' + next + '.');
  return { ok: true, written: parsed.written, of: all.length, done: done };
}

// Start the import over from the first submission.
function resetKoboCursor() {
  PropertiesService.getScriptProperties().setProperty('KOBO_CURSOR', '0');
  Logger.log('Kobo import cursor reset.');
}

function setupKoboTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importFromKobo') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importFromKobo').timeBased().everyDays(1).atHour(2).create();
  Logger.log('Nightly Kobo import scheduled for 02:00.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ento')
    .addItem('Check Kobo fields', 'peekKobo')
    .addItem('Import from Kobo now', 'importFromKobo')
    .addItem('Schedule nightly Kobo import', 'setupKoboTrigger')
    .addItem('Reset Kobo import position', 'resetKoboCursor')
    .addSeparator()
    .addItem('Create history log', 'ensureHistorySheet')
    .addSeparator()
    .addItem('Remove all Kobo rows', 'deleteKoboRows')
    .addItem('Stop Kobo import', 'stopKoboImport')
    .addToUi();
}

/* ─────────────────────────────────────────────────────────────────────────────
 * Clearing the Kobo data out again
 *
 * Ento → Remove all Kobo rows       deletes every imported row from every tab
 * Ento → Stop Kobo import           removes the nightly trigger and the cursor
 *
 * App records are untouched: only rows whose record_id starts "KB-" or whose
 * source column says "kobo" are removed, and each deletion is logged in _history.
 * ────────────────────────────────────────────────────────────────────────── */

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

function deleteKoboRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stamp = new Date();
  var log = [];
  var removed = 0;

  ss.getSheets().forEach(function (sheet) {
    var nameOfSheet = sheet.getName();
    if (nameOfSheet === LOG_SHEET) return;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;

    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idCol = header.indexOf('record_id');
    var srcCol = header.indexOf('source');
    var hhCol = header.indexOf('household_id');
    var typeCol = header.indexOf('form_type');
    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // bottom up, so deleting does not shift the rows still to check
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var id = idCol >= 0 ? String(row[idCol] || '') : '';
      var src = srcCol >= 0 ? String(row[srcCol] || '') : '';
      if (id.indexOf('KB-') !== 0 && src.toLowerCase() !== 'kobo') continue;
      sheet.deleteRow(i + 2);
      removed++;
      log.push([stamp, 'deleted (kobo)', id, typeCol >= 0 ? row[typeCol] : nameOfSheet,
        hhCol >= 0 ? row[hhCol] : '', '', '', '', '', '', '', 'row removed', '']);
    }
  });

  logRows(ss, log);
  Logger.log('Removed %s Kobo rows.', removed);
  SpreadsheetApp.getUi().alert('Removed ' + removed + ' Kobo rows. App records were not touched.');
  return removed;
}

function stopKoboImport() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importFromKobo') { ScriptApp.deleteTrigger(t); n++; }
  });
  PropertiesService.getScriptProperties().deleteProperty('KOBO_CURSOR');
  Logger.log('Removed %s Kobo trigger(s). Nothing will import automatically now.', n);
}
