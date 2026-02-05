/**
 * Universal Capture & Data Management System
 * Backend: Google Apps Script
 * Data Lake: Google Sheets
 * Channels: Telegram Webhook + HTML Web App
 */

const CONFIG = {
  spreadsheetIdKey: 'SPREADSHEET_ID',
  telegramTokenKey: 'TELEGRAM_BOT_TOKEN',
  telegramSecretKey: 'TELEGRAM_WEBHOOK_SECRET',
  speechApiKey: 'GOOGLE_SPEECH_API_KEY',
  allowedEntitiesKey: 'ALLOWED_ENTITY_IDS',
  defaultEntityIdKey: 'DEFAULT_ENTITY_ID',
  defaultTimezone: 'America/Mexico_City',
  webAppTitle: 'Universal Capture Console',
};

const TABLES = {
  raw_events: [
    'id', 'timestamp', 'source', 'device', 'actor', 'entity_id', 'intent', 'payload', 'status', 'priority'
  ],
  notes: [
    'id', 'timestamp', 'entity_id', 'title', 'body', 'tags', 'source', 'status', 'priority', 'raw_event_id'
  ],
  tasks: [
    'id', 'timestamp', 'entity_id', 'title', 'details', 'due_date', 'status', 'priority', 'source', 'raw_event_id', 'google_task_id'
  ],
  calendar_events: [
    'id', 'timestamp', 'entity_id', 'title', 'details', 'start_time', 'end_time', 'location', 'status', 'source', 'raw_event_id', 'google_event_id'
  ],
  financial_movements: [
    'id', 'timestamp', 'entity_id', 'movement_type', 'amount', 'currency', 'category', 'notes', 'status', 'source', 'raw_event_id'
  ],
  emails_log: [
    'id', 'timestamp', 'entity_id', 'to', 'cc', 'subject', 'body', 'status', 'source', 'raw_event_id', 'gmail_thread_id'
  ],
  entities: [
    'id', 'created_at', 'name', 'type', 'status', 'notes'
  ],
  knowledge_treasure: [
    'id', 'timestamp', 'entity_id', 'title', 'content', 'tags', 'source', 'raw_event_id'
  ],
  tags: [
    'id', 'created_at', 'label', 'notes'
  ],
  relations: [
    'id', 'created_at', 'from_table', 'from_id', 'to_table', 'to_id', 'relation_type', 'notes'
  ],
  logs: [
    'id', 'timestamp', 'level', 'message', 'context'
  ]
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Universal Capture ⚡')
    .addItem('Initialize Schema', 'initializeSchema')
    .addItem('Health Check', 'healthCheck')
    .addToUi();
}

function initializeSchema() {
  const ss = getSpreadsheet_();
  Object.keys(TABLES).forEach((name) => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]);
    sheet.setFrozenRows(1);
  });
  logEvent_('info', 'Schema initialized', { tables: Object.keys(TABLES) });
}

function healthCheck() {
  const ss = getSpreadsheet_();
  const status = {
    spreadsheetId: ss.getId(),
    sheets: ss.getSheets().map((sheet) => sheet.getName()),
    timestamp: new Date().toISOString(),
  };
  SpreadsheetApp.getUi().alert('Health Check OK\n' + JSON.stringify(status, null, 2));
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle(CONFIG.webAppTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const contentType = (e && e.postData && e.postData.type) || 'application/json';
  const rawBody = e && e.postData ? e.postData.contents : '';
  const params = e && e.parameter ? e.parameter : {};

  if (contentType.indexOf('application/json') === -1 && !rawBody) {
    return jsonResponse_({ status: 'error', message: 'Unsupported content type' }, 415);
  }

  const payload = rawBody ? JSON.parse(rawBody) : {};

  if (isTelegramWebhook_(e, params, payload)) {
    const result = handleTelegramWebhook_(payload);
    return jsonResponse_(result);
  }

  if (payload && payload.action === 'query') {
    const result = queryData_(payload);
    return jsonResponse_(result);
  }

  const result = handleCapture_(payload);
  return jsonResponse_(result);
}

function handleCapture_(payload) {
  ensureSchema_();
  const normalized = normalizePayload_(payload);
  const rawEventId = logRawEvent_(normalized);
  const routed = routePayload_(normalized, rawEventId);
  return {
    status: 'ok',
    raw_event_id: rawEventId,
    routed,
  };
}

function handleTelegramWebhook_(payload) {
  const secret = getScriptProperty_(CONFIG.telegramSecretKey);
  if (secret && payload && payload.secret_token && payload.secret_token !== secret) {
    return { status: 'error', message: 'Invalid secret token' };
  }

  const telegramEvent = parseTelegramUpdate_(payload);
  if (!telegramEvent) {
    return { status: 'ignored', message: 'No message payload found' };
  }

  const capturePayload = telegramEventToPayload_(telegramEvent);
  return handleCapture_(capturePayload);
}

function queryData_(payload) {
  ensureSchema_();
  const table = payload.table || 'raw_events';
  const filters = payload.filters || {};
  const sort = payload.sort || null;
  const limit = payload.limit || 50;
  const data = selectFromTable_(table, filters, sort, limit);
  return { status: 'ok', table, data };
}

function createTimeTriggers() {
  ScriptApp.newTrigger('sendDailySummary').timeBased().everyDays(1).atHour(20).create();
  ScriptApp.newTrigger('sendWeeklySummary').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();
}

function sendDailySummary() {
  const summary = summarizeActivity_('daily');
  logEvent_('info', 'Daily summary generated', summary);
}

function sendWeeklySummary() {
  const summary = summarizeActivity_('weekly');
  logEvent_('info', 'Weekly summary generated', summary);
}
