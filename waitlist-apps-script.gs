/**
 * 3-LOW waitlist → Google Sheet
 *
 * Setup (one time, ~5 minutes):
 *   1. Create a Google Sheet (e.g. "3-LOW Waitlist").
 *   2. In the Sheet: Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *      Click Deploy and approve the permissions prompt.
 *   4. Copy the Web app URL (ends in /exec) and paste it into index.html:
 *        const WAITLIST_ENDPOINT = 'https://script.google.com/macros/s/.../exec';
 *
 * Each signup adds a row to a "Waitlist" tab: Timestamp | Name | Email | Source.
 * If you edit this script later, use Deploy → Manage deployments → Edit → New version
 * so the same URL keeps working.
 */

const SHEET_NAME = 'Waitlist';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};

    // Honeypot: real people never fill this hidden field; bots often do.
    if (p.website) return reply_({ ok: true });

    const email = String(p.email || '').trim().slice(0, 320);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return reply_({ ok: false, error: 'invalid email' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Source']);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      safe_(p.name, 200),
      safe_(email, 320),
      safe_(p.source, 100),
    ]);
    return reply_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Stop anything typed into the form from being run as a spreadsheet formula.
function safe_(value, max) {
  const s = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
