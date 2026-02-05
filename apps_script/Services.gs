function getSpreadsheet_() {
  const id = getScriptProperty_(CONFIG.spreadsheetIdKey);
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function getScriptProperty_(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function ensureSchema_() {
  const ss = getSpreadsheet_();
  Object.keys(TABLES).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]);
      sheet.setFrozenRows(1);
    }
  });
}

function logRawEvent_(payload) {
  const id = Utilities.getUuid();
  const timestamp = new Date().toISOString();
  const row = [
    id,
    timestamp,
    payload.source,
    payload.device,
    payload.actor,
    payload.entity_id,
    payload.intent,
    JSON.stringify(payload),
    payload.status,
    payload.priority
  ];
  appendRow_('raw_events', row);
  return id;
}

function logEvent_(level, message, context) {
  const row = [
    Utilities.getUuid(),
    new Date().toISOString(),
    level,
    message,
    JSON.stringify(context || {})
  ];
  appendRow_('logs', row);
}

function appendRow_(table, row) {
  const sheet = getSpreadsheet_().getSheetByName(table);
  sheet.appendRow(row);
}

function normalizePayload_(payload) {
  const now = new Date().toISOString();
  const defaultEntity = getScriptProperty_(CONFIG.defaultEntityIdKey) || 'entity_default';
  const normalized = {
    id: payload.id || Utilities.getUuid(),
    timestamp: payload.timestamp || now,
    source: payload.source || 'html',
    device: payload.device || 'unknown',
    actor: payload.actor || 'user',
    entity_id: payload.entity_id || defaultEntity,
    intent: payload.intent || detectIntent_(payload),
    status: payload.status || 'new',
    priority: payload.priority || 'normal',
    data: payload.data || payload,
  };
  authorizeEntity_(normalized.entity_id);
  return normalized;
}

function detectIntent_(payload) {
  if (!payload || !payload.text) return 'note';
  const text = payload.text.toLowerCase();
  if (text.indexOf('/task') === 0) return 'task';
  if (text.indexOf('/note') === 0) return 'note';
  if (text.indexOf('/money') === 0) return 'financial';
  if (text.indexOf('/treasure') === 0) return 'treasure';
  if (text.indexOf('/email') === 0) return 'email';
  if (text.indexOf('/event') === 0) return 'calendar';
  return 'note';
}

function routePayload_(payload, rawEventId) {
  switch (payload.intent) {
    case 'task':
      return createTask_(payload, rawEventId);
    case 'calendar':
      return createCalendarEvent_(payload, rawEventId);
    case 'financial':
      return createFinancialMovement_(payload, rawEventId);
    case 'email':
      return createEmailLog_(payload, rawEventId);
    case 'treasure':
      return createKnowledgeTreasure_(payload, rawEventId);
    case 'note':
    default:
      return createNote_(payload, rawEventId);
  }
}

function createNote_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    parsed.title || 'Note',
    parsed.body || payload.text || '',
    parsed.tags.join(','),
    payload.source,
    payload.status,
    payload.priority,
    rawEventId
  ];
  appendRow_('notes', row);
  return { table: 'notes', id: row[0] };
}

function createTask_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const dueDate = parsed.due_date || payload.data.due_date || '';
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    parsed.title || 'Task',
    parsed.body || payload.text || '',
    dueDate,
    payload.status,
    parsed.priority || payload.priority,
    payload.source,
    rawEventId,
    ''
  ];
  appendRow_('tasks', row);
  if (payload.data.sync_google_tasks) {
    row[10] = createGoogleTask_(row[3], row[4], row[5]);
  }
  return { table: 'tasks', id: row[0], google_task_id: row[10] };
}

function createCalendarEvent_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const start = payload.data.start_time || parsed.start_time || '';
  const end = payload.data.end_time || parsed.end_time || '';
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    parsed.title || 'Event',
    parsed.body || payload.text || '',
    start,
    end,
    payload.data.location || '',
    payload.status,
    payload.source,
    rawEventId,
    ''
  ];
  appendRow_('calendar_events', row);
  if (payload.data.sync_calendar) {
    row[11] = createGoogleCalendarEvent_(row[3], row[4], row[5], row[6], row[7]);
  }
  return { table: 'calendar_events', id: row[0], google_event_id: row[11] };
}

function createFinancialMovement_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    parsed.movement_type || payload.data.movement_type || 'expense',
    parsed.amount || payload.data.amount || 0,
    parsed.currency || payload.data.currency || 'MXN',
    parsed.category || payload.data.category || '',
    parsed.body || payload.text || '',
    payload.status,
    payload.source,
    rawEventId
  ];
  appendRow_('financial_movements', row);
  return { table: 'financial_movements', id: row[0] };
}

function createEmailLog_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    payload.data.to || '',
    payload.data.cc || '',
    parsed.title || payload.data.subject || 'Email',
    parsed.body || payload.text || '',
    payload.status,
    payload.source,
    rawEventId,
    ''
  ];
  appendRow_('emails_log', row);
  return { table: 'emails_log', id: row[0] };
}

function createKnowledgeTreasure_(payload, rawEventId) {
  const parsed = parseCommand_(payload.text || '');
  const row = [
    Utilities.getUuid(),
    payload.timestamp,
    payload.entity_id,
    parsed.title || 'Treasure',
    parsed.body || payload.text || '',
    parsed.tags.join(','),
    payload.source,
    rawEventId
  ];
  appendRow_('knowledge_treasure', row);
  return { table: 'knowledge_treasure', id: row[0] };
}

function parseCommand_(text) {
  const cleaned = text.replace(/^\/[a-zA-Z]+\s*/g, '').trim();
  const parts = cleaned.split('|').map((part) => part.trim());
  const title = parts[0] || '';
  const result = {
    title,
    body: parts.slice(1).join(' | '),
    tags: [],
    priority: '',
    due_date: '',
    amount: null,
    currency: '',
    movement_type: '',
    category: '',
    start_time: '',
    end_time: '',
  };

  parts.forEach((part) => {
    if (part.toLowerCase().indexOf('tag:') === 0) {
      result.tags.push(part.split(':').slice(1).join(':').trim());
    }
    if (part.toLowerCase().indexOf('priority') === 0) {
      result.priority = part.split(':').slice(1).join(':').trim();
    }
    if (part.match(/\d{4}-\d{2}-\d{2}/)) {
      result.due_date = part.match(/\d{4}-\d{2}-\d{2}/)[0];
    }
    if (part.match(/\d+(\.\d+)?/)) {
      result.amount = parseFloat(part.match(/\d+(\.\d+)?/)[0]);
    }
  });

  return result;
}

function authorizeEntity_(entityId) {
  const allowed = getScriptProperty_(CONFIG.allowedEntitiesKey);
  if (!allowed) return true;
  const list = allowed.split(',').map((id) => id.trim());
  if (list.indexOf(entityId) === -1) {
    throw new Error('Entity not allowed: ' + entityId);
  }
  return true;
}

function selectFromTable_(table, filters, sort, limit) {
  const sheet = getSpreadsheet_().getSheetByName(table);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  let rows = values.map((row) => {
    const record = {};
    headers.forEach((header, idx) => {
      record[header] = row[idx];
    });
    return record;
  });

  Object.keys(filters || {}).forEach((key) => {
    rows = rows.filter((row) => String(row[key]).toLowerCase().indexOf(String(filters[key]).toLowerCase()) !== -1);
  });

  if (sort && sort.key) {
    rows.sort((a, b) => {
      if (a[sort.key] === b[sort.key]) return 0;
      return sort.direction === 'desc' ? (a[sort.key] < b[sort.key] ? 1 : -1) : (a[sort.key] > b[sort.key] ? 1 : -1);
    });
  }

  return rows.slice(0, limit);
}

function summarizeActivity_(period) {
  const now = new Date();
  const cutoff = new Date();
  if (period === 'weekly') {
    cutoff.setDate(now.getDate() - 7);
  } else {
    cutoff.setDate(now.getDate() - 1);
  }

  const rows = selectFromTable_('raw_events', {}, { key: 'timestamp', direction: 'desc' }, 500);
  const recent = rows.filter((row) => new Date(row.timestamp) >= cutoff);
  return {
    period,
    count: recent.length,
    latest: recent[0] || null
  };
}

function createGoogleTask_(title, notes, dueDate) {
  if (!Tasks) return '';
  const task = {
    title: title,
    notes: notes,
    due: dueDate ? new Date(dueDate).toISOString() : undefined,
  };
  const taskList = Tasks.Tasklists.list().items?.[0];
  if (!taskList) return '';
  const created = Tasks.Tasks.insert(task, taskList.id);
  return created.id;
}

function createGoogleCalendarEvent_(title, description, start, end, location) {
  const calendar = CalendarApp.getDefaultCalendar();
  if (!start || !end) return '';
  const event = calendar.createEvent(title, new Date(start), new Date(end), {
    description,
    location,
  });
  return event.getId();
}

function jsonResponse_(payload, status) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  if (status) {
    output.setContent(JSON.stringify(Object.assign({ status }, payload)));
  }
  return output;
}
