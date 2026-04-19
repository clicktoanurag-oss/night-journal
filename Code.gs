// Night Journal — Google Apps Script backend
// Paste this entire file into your Apps Script editor

const SHEET_NAME = 'entries';
const SECRET_KEY = 'anurag2024night';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','date','saved_at','gratitude','achievements','challenges','open_points','identity','next_steps']);
  }
  return sheet;
}

function doGet(e) {
  if (!e.parameter || e.parameter.key !== SECRET_KEY) return deny();

  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return respond([]);

  const headers = rows[0];
  const jsonFields = ['gratitude','achievements','challenges','identity','next_steps'];

  const entries = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    jsonFields.forEach(f => {
      try { obj[f] = JSON.parse(obj[f]); } catch { obj[f] = []; }
    });
    return obj;
  }).reverse(); // newest first

  return respond(entries);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents);
  if (payload.key !== SECRET_KEY) return deny();

  const sheet = getSheet();

  if (payload.action === 'delete') {
    const rows = sheet.getDataRange().getValues();
    const idCol = rows[0].indexOf('id');
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idCol]) === String(payload.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    return respond({ ok: true });
  }

  // Save new entry
  const newId = Date.now();
  sheet.appendRow([
    newId,
    payload.date        || '',
    new Date().toISOString(),
    JSON.stringify(payload.gratitude    || []),
    JSON.stringify(payload.achievements || []),
    JSON.stringify(payload.challenges   || []),
    payload.open_points || '',
    JSON.stringify(payload.identity     || []),
    JSON.stringify(payload.next_steps   || []),
  ]);

  return respond({ id: newId });
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function deny() {
  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
    .setMimeType(ContentService.MimeType.JSON);
}
