function getSpreadsheet_() {
  if (CONFIG.spreadsheetId) {
    return SpreadsheetApp.openById(CONFIG.spreadsheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Missing sheet: ${name}`);
  }
  return sheet;
}

function ensureSchema_() {
  const spreadsheet = getSpreadsheet_();
  Object.keys(CONFIG.tables).forEach((tableName) => {
    let sheet = spreadsheet.getSheetByName(tableName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(tableName);
    }
    const headers = CONFIG.tables[tableName];
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const normalized = current.map((value) => String(value || '').trim());
    const matches = headers.every((header, idx) => normalized[idx] === header);
    if (!matches) {
      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  });
}

function uuid_() {
  const uuid = Utilities.getUuid();
  return uuid.toLowerCase();
}

function nowIso_() {
  return Utilities.formatDate(new Date(), CONFIG.timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function appendRow_(table, rowObject) {
  const sheet = getSheet_(table);
  const headers = CONFIG.tables[table];
  const row = headers.map((header) => rowObject[header] || '');
  sheet.appendRow(row);
}

function getDevice_(payload) {
  return payload.device || payload.userAgent || 'unknown';
}

function normalizeText_(value) {
  return String(value || '').trim();
}

function parseTags_(value) {
  if (!value) return '';
  if (Array.isArray(value)) return value.join(',');
  return String(value)
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean)
    .join(',');
}

function logEvent_(level, message, payload, meta) {
  const entry = {
    log_id: uuid_(),
    timestamp: nowIso_(),
    source: meta?.source || 'system',
    device: meta?.device || 'system',
    actor: meta?.actor || 'system',
    level,
    message,
    payload: payload ? JSON.stringify(payload) : ''
  };
  appendRow_('logs', entry);
}
