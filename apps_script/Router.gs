function handleCapture(payload) {
  const normalized = normalizeCapture(payload);
  const rawEventId = logRawEvent(normalized);

  const route = normalized.intent || inferIntent(normalized.text || '');
  const context = {
    raw_event_id: rawEventId,
    source: normalized.source,
    device: normalized.device,
    entity_id: normalized.entity_id,
    status: normalized.status,
    priority: normalized.priority
  };

  switch (route) {
    case 'note':
      return createNote(normalized, context);
    case 'task':
      return createTask(normalized, context);
    case 'calendar':
      return createCalendarEvent(normalized, context);
    case 'money':
      return createFinancialMovement(normalized, context);
    case 'email':
      return createEmailLog(normalized, context);
    case 'treasure':
      return createKnowledgeTreasure(normalized, context);
    default:
      return createNote(normalized, context);
  }
}

function normalizeCapture(payload) {
  const normalized = payload || {};
  return {
    text: normalized.text || '',
    intent: normalized.intent || '',
    source: normalized.source || 'unknown',
    device: normalized.device || 'unknown',
    actor: normalized.actor || '',
    entity_id: normalized.entity_id || '',
    status: normalized.status || 'open',
    priority: parsePriority(normalized.priority),
    tags: normalizeTags(normalized.tags),
    metadata: normalized.metadata || {}
  };
}

function inferIntent(text) {
  const lowered = text.toLowerCase();
  if (lowered.startsWith('task') || lowered.includes('tarea')) {
    return 'task';
  }
  if (lowered.startsWith('money') || lowered.includes('gasto') || lowered.includes('ingreso')) {
    return 'money';
  }
  if (lowered.startsWith('email') || lowered.includes('correo')) {
    return 'email';
  }
  if (lowered.startsWith('calendar') || lowered.includes('evento')) {
    return 'calendar';
  }
  if (lowered.startsWith('treasure') || lowered.includes('idea') || lowered.includes('concepto')) {
    return 'treasure';
  }
  return 'note';
}

function logRawEvent(normalized) {
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    intent: normalized.intent || '',
    payload: JSON.stringify(normalized),
    entity_id: normalized.entity_id,
    status: normalized.status,
    priority: normalized.priority,
    tags: normalized.tags,
    correlation_id: generateUuid()
  };
  appendRecord(CONFIG.TABLES.RAW_EVENTS, record, HEADERS.raw_events);
  return record.id;
}

function createNote(normalized, context) {
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: context.entity_id,
    status: context.status,
    priority: context.priority,
    content: normalized.text,
    tags: normalized.tags,
    raw_event_id: context.raw_event_id
  };
  appendRecord(CONFIG.TABLES.NOTES, record, HEADERS.notes);
  return record;
}

function createTask(normalized, context) {
  const parsed = parseTaskText(normalized.text);
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: context.entity_id,
    status: context.status,
    priority: parsePriority(parsed.priority || context.priority),
    title: parsed.title,
    details: parsed.details,
    due_date: parseDate(parsed.due_date),
    raw_event_id: context.raw_event_id,
    external_id: ''
  };
  appendRecord(CONFIG.TABLES.TASKS, record, HEADERS.tasks);
  return record;
}

function createCalendarEvent(normalized, context) {
  const parsed = parseCalendarText(normalized.text);
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: context.entity_id,
    status: context.status,
    priority: context.priority,
    title: parsed.title,
    details: parsed.details,
    start_at: parseDate(parsed.start_at),
    end_at: parseDate(parsed.end_at),
    location: parsed.location,
    raw_event_id: context.raw_event_id,
    external_id: ''
  };
  appendRecord(CONFIG.TABLES.CALENDAR_EVENTS, record, HEADERS.calendar_events);
  return record;
}

function createFinancialMovement(normalized, context) {
  const parsed = parseMoneyText(normalized.text);
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: parsed.entity_id || context.entity_id,
    status: context.status,
    priority: context.priority,
    movement_type: parsed.movement_type,
    amount: parsed.amount,
    currency: parsed.currency,
    category: parsed.category,
    description: parsed.description,
    raw_event_id: context.raw_event_id
  };
  appendRecord(CONFIG.TABLES.FINANCIAL_MOVEMENTS, record, HEADERS.financial_movements);
  return record;
}

function createEmailLog(normalized, context) {
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: context.entity_id,
    status: context.status,
    priority: context.priority,
    subject: normalized.text,
    to: normalized.metadata.to || '',
    cc: normalized.metadata.cc || '',
    body_excerpt: normalized.metadata.body_excerpt || '',
    raw_event_id: context.raw_event_id,
    gmail_id: ''
  };
  appendRecord(CONFIG.TABLES.EMAILS_LOG, record, HEADERS.emails_log);
  return record;
}

function createKnowledgeTreasure(normalized, context) {
  const record = {
    id: generateUuid(),
    timestamp: nowIso(),
    source: context.source,
    device: context.device,
    entity_id: context.entity_id,
    status: context.status,
    priority: context.priority,
    content: normalized.text,
    tags: normalized.tags,
    raw_event_id: context.raw_event_id
  };
  appendRecord(CONFIG.TABLES.KNOWLEDGE_TREASURE, record, HEADERS.knowledge_treasure);
  return record;
}

function parseTaskText(text) {
  const parts = text.replace('/task', '').trim().split('|').map((part) => part.trim());
  return {
    title: parts[0] || text,
    details: parts[3] || '',
    due_date: parts[1] || '',
    priority: parts[2] || ''
  };
}

function parseCalendarText(text) {
  const parts = text.replace('/event', '').replace('/calendar', '').trim().split('|').map((part) => part.trim());
  return {
    title: parts[0] || text,
    start_at: parts[1] || '',
    end_at: parts[2] || '',
    location: parts[3] || '',
    details: parts[4] || ''
  };
}

function parseMoneyText(text) {
  const cleaned = text.replace('/money', '').trim();
  const entityMatch = cleaned.match(/entidad:([\w-]+)/i);
  const amountMatch = cleaned.match(/(\d+[\.,]?\d*)/);
  const currencyMatch = cleaned.match(/(mxn|usd|eur)/i);
  return {
    movement_type: cleaned.toLowerCase().includes('gasto') ? 'expense' : 'income',
    amount: amountMatch ? amountMatch[1].replace(',', '.') : '',
    currency: currencyMatch ? currencyMatch[1].toUpperCase() : 'MXN',
    category: '',
    description: cleaned,
    entity_id: entityMatch ? entityMatch[1] : ''
  };
}
