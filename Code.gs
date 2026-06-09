// ============================================================
// IIMB GSK Women Support Group — Google Apps Script Backend
// ============================================================

const SHEET_NAME = 'Posts';

const MIN_NAME_LENGTH = 2;
const MAX_NAME_LENGTH = 50;
const MAX_POSTS_PER_DAY = 30;
const MAX_POSTS_PER_NAME = 2;

const LINKEDIN_PATTERNS = [
  /^https:\/\/(www\.)?linkedin\.com\/posts\//,
  /^https:\/\/(www\.)?linkedin\.com\/feed\/update\//,
  /^https:\/\/(www\.)?linkedin\.com\/pulse\//,
  /^https:\/\/lnkd\.in\//
];

const BLOCKED_NAME_WORDS = ['test', 'spam', 'bot', 'admin', 'null', 'undefined', 'script'];

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || 'getPosts';
    const callback = params.callback || '';

    let result;
    if (action === 'getPosts') {
      result = getPosts();
    } else if (action === 'addPost') {
      result = addPost(params.name || '', params.url || '');
    } else if (action === 'deletePost') {
      result = deletePost(params.id || '');
    } else {
      result = { error: 'Unknown action' };
    }

    const json = JSON.stringify(result);
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    const msg = JSON.stringify({ error: err.toString() });
    return ContentService
      .createTextOutput(msg)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.action === 'addPost') result = addPost(data.name, data.url);
    else if (data.action === 'deletePost') result = deletePost(data.id);
    else result = { error: 'Unknown action' };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function validatePost(name, url) {
  if (!name || typeof name !== 'string') return 'Name is required.';
  const n = name.trim();
  if (n.length < MIN_NAME_LENGTH) return 'Name is too short. Please use your real name.';
  if (n.length > MAX_NAME_LENGTH) return 'Name is too long.';
  if (/^[^a-zA-Z]+$/.test(n)) return 'Name must contain letters.';
  if (/<|>|javascript:|data:/i.test(n)) return 'Invalid characters in name.';
  const lower = n.toLowerCase();
  for (const w of BLOCKED_NAME_WORDS) { if (lower === w) return 'Please use your real name.'; }

  if (!url || typeof url !== 'string') return 'URL is required.';
  const u = url.trim();
  if (!u.startsWith('https://')) return 'Only secure (https) URLs allowed.';
  if (!LINKEDIN_PATTERNS.some(p => p.test(u))) return 'Only LinkedIn post URLs are allowed.';
  if (u.length > 500) return 'URL is too long.';
  return null;
}

function checkRateLimits(name, sheet) {
  const rows = sheet.getDataRange().getValues();
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);
  const todayTs = todayStart.getTime();
  let totalToday = 0, nameToday = 0;
  const lower = name.trim().toLowerCase();
  for (let i = 1; i < rows.length; i++) {
    const ts = Number(rows[i][3]);
    if (ts >= todayTs) {
      totalToday++;
      if (String(rows[i][1]).trim().toLowerCase() === lower) nameToday++;
    }
  }
  if (totalToday >= MAX_POSTS_PER_DAY) return 'Daily post limit reached. Try again tomorrow.';
  if (nameToday >= MAX_POSTS_PER_NAME) return 'You have already added ' + MAX_POSTS_PER_NAME + ' posts today.';
  return null;
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
    if (ts < cutoff) continue;
    posts.push({ id: String(id), name, url, ts });
  }
  cleanOldRows(sheet, cutoff);
  return { posts };
}

function addPost(name, url) {
  const err = validatePost(name, url);
  if (err) return { error: err };
  const sheet = getSheet();
  const rateErr = checkRateLimits(name, sheet);
  if (rateErr) return { error: rateErr };
  const id = Date.now().toString();
  const ts = Date.now();
  sheet.appendRow([id, name.trim(), url.trim(), ts]);
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
  for (let i = rows.length - 1; i >= 1; i--) {
    const ts = Number(rows[i][3]);
    if (ts && ts < cutoff) sheet.deleteRow(i + 1);
  }
}
