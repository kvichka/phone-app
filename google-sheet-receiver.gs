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
var LOG_COLS = ['logged_at', 'action', 'record_id', 'form_type', 'household_id', 'trap_id',
  'collected_date', 'period_label', 'collector', 'cluster', 'changed_fields', 'sheet_row'];

function logRows(ss, rows) {
  if (!rows.length) return;
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LOG_COLS.length).setValues(rows);
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
      log.push([
        stamp, existing ? (changed.length ? 'updated' : 'unchanged') : 'new',
        rec.record_id || '', type, rec.household_id || '', rec.trap_id || '',
        rec.collected_date || '', rec.period_label || '', rec.collector || '', rec.cluster || '',
        existing ? changed.join(', ') : 'all', atRow,
      ]);
      written += 1;
    });

    logRows(ss, log);
    out = { ok: true, count: written, at: new Date().toISOString() };
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
    return ContentService.createTextOutput(JSON.stringify({ ok: true, history: log }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action !== 'records') {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'ento receiver' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var out = [];
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
        if (!empty && rec.record_id) out.push(rec);
      }
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true, records: out }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
