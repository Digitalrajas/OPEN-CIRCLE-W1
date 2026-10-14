/**
 * Open Circle W1 — Registration backend (Google Apps Script)
 *
 * Paste this into the Google Sheet's Extensions → Apps Script editor,
 * then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * To update later: Deploy → Manage deployments → Edit → Version: New version (keeps the same URL).
 * Each valid form submission from the landing page is appended as one row.
 */

const SHEET_NAME = 'Registrations';
// Registrations close 17 Oct, 24:00 IST — same deadline as the landing page countdown
const DEADLINE = new Date('2026-10-18T00:00:00+05:30');
const HEADERS = ['Timestamp', 'Full Name', 'WhatsApp', 'Email', 'College', 'City', 'Has Laptop'];

// Only sent to the browser after a successful registration, so it never appears in the page source
const WHATSAPP_GROUP_LINK = 'https://chat.whatsapp.com/IPeYLX67rzzA3shVftx7ri';

// Flood protection: max submissions accepted across everyone per rolling minute
const MAX_SUBMISSIONS_PER_MINUTE = 30;

function doPost(e) {
  if (new Date() >= DEADLINE) {
    return jsonResponse({ ok: false, error: 'closed' });
  }

  const p = (e && e.parameter) || {};

  // Honeypot: real visitors never see the "website" field, bots fill every input.
  // Pretend success so the bot has no signal to adapt to.
  if (p.website) {
    return jsonResponse({ ok: true });
  }

  const fullName = String(p.fullName || '').trim();
  const email = String(p.email || '').trim().toLowerCase();
  const phone = normalizeIndianMobile(p.whatsapp);

  if (fullName.length < 2 || fullName.length > 80) {
    return jsonResponse({ ok: false, error: 'invalid_name' });
  }
  if (!phone) {
    return jsonResponse({ ok: false, error: 'invalid_phone' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 120) {
    return jsonResponse({ ok: false, error: 'invalid_email' });
  }

  if (isRateLimited()) {
    return jsonResponse({ ok: false, error: 'busy' });
  }

  // Serialise concurrent submissions so the duplicate check and append are atomic
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();

    if (isAlreadyRegistered(sheet, phone, email)) {
      return jsonResponse({ ok: true, duplicate: true, groupLink: WHATSAPP_GROUP_LINK });
    }

    sheet.appendRow([
      new Date(),
      clean(fullName),
      // Leading apostrophe keeps "+91…" as text instead of a formula/number
      "'+91 " + phone,
      clean(email),
      clean(p.college),
      clean(p.city),
      p.laptop === 'no' ? 'No' : 'Yes',
    ]);
  } finally {
    lock.releaseLock();
  }

  return jsonResponse({ ok: true, groupLink: WHATSAPP_GROUP_LINK });
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

// Accepts "98765 43210", "+91 98765-43210", "09876543210", etc. Returns the 10-digit number or ''.
function normalizeIndianMobile(value) {
  let digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 12 && digits.indexOf('91') === 0) digits = digits.slice(2);
  if (digits.length === 11 && digits.charAt(0) === '0') digits = digits.slice(1);
  return /^[6-9]\d{9}$/.test(digits) ? digits : '';
}

function isAlreadyRegistered(sheet, phone, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  // Columns C (WhatsApp) and D (Email)
  const rows = sheet.getRange(2, 3, lastRow - 1, 2).getDisplayValues();
  return rows.some(([rowPhone, rowEmail]) =>
    String(rowPhone).replace(/\D/g, '').slice(-10) === phone ||
    String(rowEmail).trim().toLowerCase() === email
  );
}

function isRateLimited() {
  const cache = CacheService.getScriptCache();
  const key = 'submissions_' + Math.floor(Date.now() / 60000);
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), 120);
  return count > MAX_SUBMISSIONS_PER_MINUTE;
}

// Trim, cap length, and stop values like "=IMPORTXML(…)" being read as formulas
function clean(value) {
  const text = String(value || '').trim().slice(0, 200);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
