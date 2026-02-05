function createNote_(normalized, rawEventId) {
  const note = parseNote_(normalized);
  note.raw_event_id = rawEventId;
  appendRow_('notes', note);
  return note;
}

function createTask_(normalized, rawEventId) {
  const task = parseTask_(normalized);
  task.raw_event_id = rawEventId;
  appendRow_('tasks', task);
  return task;
}

function createFinancialMovement_(normalized, rawEventId) {
  const movement = parseFinancialMovement_(normalized);
  movement.raw_event_id = rawEventId;
  appendRow_('financial_movements', movement);
  return movement;
}

function createCalendarEvent_(normalized, rawEventId) {
  const event = parseCalendarEvent_(normalized);
  event.raw_event_id = rawEventId;
  appendRow_('calendar_events', event);
  return event;
}

function createEmailLog_(normalized, rawEventId) {
  const email = parseEmailLog_(normalized);
  email.raw_event_id = rawEventId;
  appendRow_('emails_log', email);
  return email;
}

function createEntity_(normalized) {
  const entity = {
    entity_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    name: normalizeText_(normalized.payload.name || normalized.text),
    type: normalizeText_(normalized.payload.type || 'general'),
    metadata: JSON.stringify(normalized.payload.metadata || {}),
    status: normalized.status,
    priority: normalized.priority
  };
  appendRow_('entities', entity);
  return entity;
}

function createKnowledgeTreasure_(normalized, rawEventId) {
  const treasure = {
    treasure_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id,
    content: normalizeText_(normalized.payload.content || normalized.text),
    tags: parseTags_(normalized.payload.tags || normalized.payload.tag || ''),
    status: normalized.status,
    priority: normalized.priority,
    raw_event_id: rawEventId
  };
  appendRow_('knowledge_treasure', treasure);
  return treasure;
}

function parseNote_(normalized) {
  const parsed = parseCommand_(normalized.text);
  return {
    note_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id || parsed.entity_id,
    title: parsed.title || 'Nota',
    body: parsed.body || normalized.text,
    tags: parseTags_(parsed.tags),
    status: normalized.status,
    priority: normalized.priority
  };
}

function parseTask_(normalized) {
  const parsed = parseCommand_(normalized.text);
  return {
    task_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id || parsed.entity_id,
    title: parsed.title || parsed.body || normalized.text,
    details: parsed.details || '',
    due_date: parsed.due_date || '',
    status: normalized.status,
    priority: parsed.priority || normalized.priority,
    linked_note_id: parsed.linked_note_id || '',
    linked_email_id: parsed.linked_email_id || ''
  };
}

function parseFinancialMovement_(normalized) {
  const parsed = parseCommand_(normalized.text);
  return {
    movement_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id || parsed.entity_id,
    movement_type: parsed.movement_type || 'gasto',
    amount: parsed.amount || '',
    currency: parsed.currency || 'MXN',
    description: parsed.body || normalized.text,
    status: normalized.status,
    priority: normalized.priority
  };
}

function parseCalendarEvent_(normalized) {
  const parsed = parseCommand_(normalized.text);
  return {
    calendar_event_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id || parsed.entity_id,
    title: parsed.title || parsed.body || normalized.text,
    start_time: parsed.start_time || '',
    end_time: parsed.end_time || '',
    location: parsed.location || '',
    status: normalized.status,
    priority: normalized.priority,
    linked_task_id: parsed.linked_task_id || ''
  };
}

function parseEmailLog_(normalized) {
  const parsed = parseCommand_(normalized.text);
  return {
    email_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    entity_id: normalized.entity_id || parsed.entity_id,
    subject: parsed.title || parsed.subject || normalized.text,
    body: parsed.body || '',
    recipient: parsed.recipient || '',
    status: normalized.status,
    priority: normalized.priority,
    linked_task_id: parsed.linked_task_id || ''
  };
}

function parseCommand_(text) {
  if (!text) return {};
  const clean = text.replace(/^\/[a-zA-Z]+\s*/, '').trim();
  const parts = clean.split('|').map((part) => part.trim());
  const result = {
    body: parts[0] || '',
    title: parts[0] || ''
  };
  parts.slice(1).forEach((part) => {
    if (part.match(/\d{4}-\d{2}-\d{2}/)) {
      result.due_date = part;
    } else if (part.match(/alta|media|baja/i)) {
      result.priority = part.toLowerCase();
    } else if (part.match(/mxn|usd|eur/i)) {
      const tokens = part.split(' ');
      result.amount = tokens[0];
      result.currency = tokens[1] || 'MXN';
    } else if (part.includes(':')) {
      const [key, value] = part.split(':');
      result[key.trim()] = value.trim();
    }
  });
  return result;
}
