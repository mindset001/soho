// ============================================================
// Soho Boulevard — Google Apps Script Backend
//
// SETUP STEPS:
//   1. Open your Google Spreadsheet (create one if needed).
//   2. Click Extensions > Apps Script.
//   3. Replace all code with this file.
//   4. Set TOKEN below to a strong secret (any long random string).
//   5. Click Deploy > New deployment > Web App.
//      - Execute as: Me
//      - Who has access: Anyone
//   6. Copy the Web App URL (/exec) into SHEETS_URL in index.html.
//   7. Copy your TOKEN into SHEETS_TOKEN in index.html.
//   8. Re-deploy (New deployment or Manage deployments > edit) after
//      any code changes — editing the script alone does not update
//      the live URL.
// ============================================================

const TOKEN = '7d4f8c1a9b3e6f2d5a8c7e1f4b9d2a6c8f3e7b1d5a9c2f6e4b8d1a7c3f9e5b2;

// Sheet tab names
const SHEET_APPS  = 'Applications';
const SHEET_EVTS  = 'Events';

// Column order — must match the fields written by index.html
const APPS_COLS = ['id','memberNo','ts','name','email','phone','birthday','interest','clubs','tier','interestedIn','status'];
const EVTS_COLS = ['id','ts','event','name','email','phone','guests','note','status'];

// ── Entry points ──────────────────────────────────────────────

function doGet(e) {
  const key   = (e.parameter.key   || '').trim();
  const token = (e.parameter.token || '').trim();
  if (token !== TOKEN) return respond({ error: 'Unauthorized' });
  return respond({ data: readData(key) });
}

function doPost(e) {
  var payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch (_) { return respond({ error: 'Bad JSON' }); }

  if (payload.token !== TOKEN) return respond({ error: 'Unauthorized' });

  if (payload.action === 'save') {
    writeData(payload.key, payload.data || []);
    return respond({ ok: true });
  }
  if (payload.action === 'sendEmail') {
    sendEmails(payload.recipients || [], payload.subject || '', payload.body || '');
    return respond({ ok: true });
  }
  return respond({ error: 'Unknown action' });
}

// ── Read ──────────────────────────────────────────────────────

function readData(key) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var cols = colsForKey(key);
  if (!cols) return [];

  var sh   = getOrCreateSheet(ss, sheetNameForKey(key), cols);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var rows = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return rows
    .filter(function(r) { return r[0]; })
    .map(function(r) {
      var obj = {};
      cols.forEach(function(h, i) {
        var v = r[i];
        // Pipe-separated arrays → real arrays
        if (h === 'interest' || h === 'clubs' || h === 'interestedIn') {
          v = (v && String(v).length) ? String(v).split('|').filter(Boolean) : [];
        }
        obj[h] = (v === undefined || v === null) ? '' : v;
      });
      return obj;
    });
}

// ── Write ─────────────────────────────────────────────────────

function writeData(key, list) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var cols = colsForKey(key);
  if (!cols) return;

  var sh      = getOrCreateSheet(ss, sheetNameForKey(key), cols);
  var lastRow = sh.getLastRow();

  // Clear existing data rows (keep header row 1)
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, cols.length).clearContent();
  }
  if (!list.length) return;

  var rows = list.map(function(rec) {
    return cols.map(function(h) {
      var v = rec[h];
      return Array.isArray(v) ? v.join('|') : (v === undefined || v === null ? '' : v);
    });
  });
  sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
}

// ── Email ─────────────────────────────────────────────────────

function sendEmails(recipients, subject, body) {
  recipients.forEach(function(r) {
    var first   = ((r.name || '').trim().split(' ')[0]) || 'there';
    var hasToken = /\{name\}|\{first\}/i.test(body);
    var outBody = hasToken
      ? body.replace(/\{name\}/gi, r.name || 'there').replace(/\{first\}/gi, first)
      : 'Hi ' + first + ',\n\n' + body;
    var outSubj = subject
      .replace(/\{name\}/gi, r.name || 'there')
      .replace(/\{first\}/gi, first);
    try {
      MailApp.sendEmail({ to: r.email, subject: outSubj, body: outBody });
    } catch (err) {
      Logger.log('Failed to send to ' + r.email + ': ' + err.message);
    }
  });
}

// ── Helpers ───────────────────────────────────────────────────

function colsForKey(key) {
  if (key === 'sb_applications') return APPS_COLS;
  if (key === 'sb_event_regs')   return EVTS_COLS;
  return null;
}

function sheetNameForKey(key) {
  if (key === 'sb_applications') return SHEET_APPS;
  if (key === 'sb_event_regs')   return SHEET_EVTS;
  return key;
}

function getOrCreateSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#F6F2E9')
      .setFontColor('#241C13');
    sh.setFrozenRows(1);
  }
  return sh;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
