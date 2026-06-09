// ============================================================
// IIMB GSK Women Support Group — Google Apps Script Backend
// ============================================================

const SHEET_NAME = 'Posts';

// ── Spam Filter Config ──────────────────────────────────────
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
// ────────────────────────────────────────────────────────────

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback; // JSONP callback name

  let result;
  try {
    if (action === 'getPosts') {
      result = getPosts();
    } else if (action === 'addPost') {
      const name = e.parameter.name || '';
      const url = e.parameter.url || '';
      result = addPost(name, url);
    } else if (action === 'deletePost') {
      const id = e.parameter.id || '';
      result = deletePost(id);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  // Return as JSONP if callback provided, plain JSON otherwise
  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep doPost for direct API use
function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'addPost') result = addPost(data.name, data.url);
    else if (data.action === 'deletePost') result = deletePost(data.id);
    else result = { error: 'Unknown action' };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Spam Validation ─────────────────────────────────────────
function validatePost(name, url) {
  if (!name || typeof name !== 'string') return 'Name is required.';
  const trimmedName = name.trim();
  if (trimmedName.length < MIN_NAME_LENGTH) return 'Name is too short. Please use your real name.';
  if (trimmedName.length > MAX_NAME_LENGTH) return 'Name is too long.';
  if (/^[^a-zA-Z]+$/.test(trimmedName)) return 'Name must contain letters.';
  if (/<|>|javascript:|data:/i.test(trimmedName)) return 'Invalid characters in name.';
  const lowerName = trimmedName.toLowerCase();
  for (const word of BLOCKED_NAME_WORDS) {
    if (lowerName === word) return 'Please use your real name.';
  }

  if (!url || typeof url !== 'string') return 'URL is required.';
  const trimmedUrl = url.trim();
  if (!trimmedUrl.startsWith('https://')) return 'Only secure (https) URLs are allowed.';
  const isLinkedIn = LINKEDIN_PATTERNS.some(p => p.test(trimmedUrl));
  if (!isLinkedIn) return 'Only LinkedIn post URLs are allowed (linkedin.com/posts/ or linkedin.com/feed/update/)';
  if (trimmedUrl.length > 500) return 'URL is too long.';

  return null;
}

// ── Rate Limiting ────────────────────────────────────────────
function checkRateLimits(name, sheet) {
  const rows = sheet.getDataRange().getValues();
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);
  const todayTs = todayStart.getTime();
  let totalToday = 0, nameToday = 0;
  const lowerName = name.trim().toLowerCase();
  for (let i = 1; i < rows.length; i++) {
    const ts = Number(rows[i][3]);
    if (ts >= todayTs) {
      totalToday++;
      if (String(rows[i][1]).trim().toLowerCase() === lowerName) nameToday++;
    }
  }
  if (totalToday >= MAX_POSTS_PER_DAY) return 'Daily post limit reached. Try again tomorrow.';
  if (nameToday >= MAX_POSTS_PER_NAME) return `You've already added ${MAX_POSTS_PER_NAME} posts today.`;
  return null;
}
// ────────────────────────────────────────────────────────────

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
  const validationError = validatePost(name, url);
  if (validationError) return { error: validationError };
  const sheet = getSheet();
  const rateLimitError = checkRateLimits(name, sheet);
  if (rateLimitError) return { error: rateLimitError };
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
