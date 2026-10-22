// Google Apps Script — paste into script.google.com (Extensions > Apps Script in your Sheet)
// Deploy > New deployment > Web app > Execute as: Me, Access: Anyone. Copy the /exec URL into index.html (SHEETS_URL).
const SHEET = 'RSVP';
function sh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (s.getLastRow() === 0) s.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'Jumlah Tamu', 'Ucapan']);
  return s;
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    sh_().appendRow([new Date(), String(d.name).slice(0, 80), d.attend, Number(d.guests) || 0, String(d.message || '').slice(0, 500)]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}
function doGet() {
  const rows = sh_().getDataRange().getValues().slice(1)
    .filter(r => r[1] && r[4]).reverse().slice(0, 100)
    .map(r => ({ name: r[1], message: r[4] }));
  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}
