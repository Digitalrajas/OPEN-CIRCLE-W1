/**
 * Open Circle W1 — Registration backend (Google Apps Script)
 *
 * Paste this into the Google Sheet's Extensions → Apps Script editor,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Each form submission from the landing page is appended as one row.
 */

const SHEET_NAME = 'Registrations';
// Registrations close 17 Oct, 24:00 IST — same deadline as the landing page countdown
const DEADLINE = new Date('2026-10-18T00:00:00+05:30');
const HEADERS = ['Timestamp', 'Full Name', 'WhatsApp', 'Email', 'College', 'City', 'Has Laptop'];

function doPost(e) {
  if (new Date() >= DEADLINE) {
    return jsonResponse({ ok: false, error: 'closed' });
  }

  const p = (e && e.parameter) || {};
  if (!p.fullName || !p.whatsapp || !p.email) {
    return jsonResponse({ ok: false, error: 'missing_fields' });
  }

  // Serialise concurrent submissions so rows never overwrite each other
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet().appendRow([
      new Date(),
      clean(p.fullName),
      clean(p.whatsapp),
      clean(p.email),
      clean(p.college),
      clean(p.city),
      p.laptop === 'no' ? 'No' : 'Yes',
    ]);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({ ok: true });
}

// Lets you open the web app URL in a browser to confirm it is deployed
function doGet() {
  return jsonResponse({ ok: true, status: 'Open Circle W1 registration endpoint is live' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Trim, cap length, and stop values like "+91…" or "=IMPORTXML(…)" being read as formulas
function clean(value) {
  const text = String(value || '').trim().slice(0, 200);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
