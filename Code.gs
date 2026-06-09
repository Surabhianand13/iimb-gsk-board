// ============================================================
// IIMB GSK Women Support Group — Google Apps Script Backend
// Paste this entire file into your Google Apps Script editor
// ============================================================

const SHEET_NAME = 'Posts';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action || (e.postData ? JSON.parse(e.postData.contents).action : null);

  let result;
  try {
    if (action === 'getPosts') {
      result = getPosts();
    } else if (action === 'addPost') {
      const data = JSON.parse(e.postData.contents);
      result = addPost(data.name, data.url);
    } else if (action === 'deletePost') {
      const data = JSON.parse(e.postData.contents);
      result = deletePost(data.id);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'name', 'url', 'timestamp']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getPosts() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { posts: [] };

  const cutoff = Date.now() - 3 * 24 * 60 * 60 * 1000;
  const posts = [];

  for (let i = 1; i < rows.length; i++) {
    const [id, name, url, timestamp] = rows[i];
    if (!id) continue;
    const ts = Number(timestamp);
    if (ts < cutoff) continue; // skip older than 3 days
    posts.push({ id: String(id), name, url, ts });
  }

  // Clean up old rows (older than 3 days) to keep sheet tidy
  cleanOldRows(sheet, cutoff);

  return { posts };
}

function addPost(name, url) {
  const sheet = getSheet();
  const id = Date.now().toString();
  const ts = Date.now();
  sheet.appendRow([id, name, url, ts]);
  return { success: true, id, ts };
}

function deletePost(id) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Post not found' };
}

function cleanOldRows(sheet, cutoff) {
  const rows = sheet.getDataRange().getValues();
  // Go bottom-up to safely delete rows
  for (let i = rows.length - 1; i >= 1; i--) {
    const ts = Number(rows[i][3]);
    if (ts && ts < cutoff) {
      sheet.deleteRow(i + 1);
    }
  }
}
