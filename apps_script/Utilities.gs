function generateUuid() {
  return Utilities.getUuid();
}

function nowIso() {
  return new Date().toISOString();
}

function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('Missing SPREADSHEET_ID in Script Properties.');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function ensureSheet(sheetName, headers) {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headers);
  }
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (currentHeaders.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function initSchema() {
  Object.keys(HEADERS).forEach((key) => {
    ensureSheet(key, HEADERS[key]);
  });
}

function logEvent(level, source, message, payload) {
  const sheet = ensureSheet(CONFIG.TABLES.LOGS, HEADERS.logs);
  sheet.appendRow([
    generateUuid(),
    nowIso(),
    level,
    source,
    message,
    payload ? JSON.stringify(payload) : ''
  ]);
}

function normalizeTags(tags) {
  if (!tags) {
    return '';
  }
  if (Array.isArray(tags)) {
    return tags.map((tag) => tag.trim()).filter((tag) => tag).join(',');
  }
  return String(tags)
    .split(',')
    .map((tag) => tag.trim())
    .filter((tag) => tag)
    .join(',');
}

function parsePriority(value) {
  const normalized = String(value || '').toLowerCase();
  if (['alta', 'high', 'urgent'].includes(normalized)) {
    return 'high';
  }
  if (['media', 'medium'].includes(normalized)) {
    return 'medium';
  }
  if (['baja', 'low'].includes(normalized)) {
    return 'low';
  }
  return normalized || 'normal';
}

function parseDate(value) {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toISOString();
}
