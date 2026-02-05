function doGet(e) {
  ensureSchema_();
  const template = HtmlService.createTemplateFromFile('HtmlCapture');
  template.webAppUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('Universal Capture').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  ensureSchema_();
  const payload = parsePostPayload_(e);
  const action = payload.action || payload.command || 'capture';
  let response = {};
  if (action === 'query') {
    response = handleQuery_(payload);
  } else if (payload.telegram) {
    response = handleTelegramUpdate_(payload.telegram);
  } else {
    response = handleCapture_(payload, payload.source || 'web');
  }
  return buildResponse_(response);
}

function parsePostPayload_(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try {
      const json = JSON.parse(e.postData.contents);
      return json;
    } catch (error) {
      return {
        raw: e.postData.contents
      };
    }
  }
  return e.parameter || {};
}

function buildResponse_(payload) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function handleCapture_(payload, source) {
  const normalized = normalizeCapturePayload_(payload, source);
  const rawEventId = recordRawEvent_(normalized);
  const intent = normalized.intent || detectIntent_(normalized);
  const result = routeIntent_(intent, normalized, rawEventId);
  return {
    ok: true,
    raw_event_id: rawEventId,
    intent,
    result
  };
}

function normalizeCapturePayload_(payload, source) {
  const actor = payload.actor || payload.user || payload.from || 'unknown';
  return {
    payload,
    intent: payload.intent || payload.type || payload.command || '',
    source,
    actor,
    device: getDevice_(payload),
    text: normalizeText_(payload.text || payload.message || payload.body || ''),
    entity_id: payload.entity_id || payload.entity || '',
    status: payload.status || 'open',
    priority: payload.priority || payload.prioridad || 'media'
  };
}

function recordRawEvent_(normalized) {
  const rawEvent = {
    raw_event_id: uuid_(),
    timestamp: nowIso_(),
    source: normalized.source,
    device: normalized.device,
    actor: normalized.actor,
    intent: normalized.intent,
    payload: JSON.stringify(normalized.payload || {}),
    entity_id: normalized.entity_id,
    status: normalized.status,
    priority: normalized.priority
  };
  appendRow_('raw_events', rawEvent);
  return rawEvent.raw_event_id;
}

function detectIntent_(normalized) {
  const text = normalized.text.toLowerCase();
  if (text.startsWith('/note')) return 'note';
  if (text.startsWith('/task')) return 'task';
  if (text.startsWith('/money')) return 'money';
  if (text.startsWith('/treasure')) return 'treasure';
  if (text.startsWith('/email')) return 'email';
  if (text.startsWith('/event')) return 'calendar';
  return 'note';
}

function routeIntent_(intent, normalized, rawEventId) {
  switch (intent) {
    case 'task':
      return createTask_(normalized, rawEventId);
    case 'money':
      return createFinancialMovement_(normalized, rawEventId);
    case 'treasure':
      return createKnowledgeTreasure_(normalized, rawEventId);
    case 'email':
      return createEmailLog_(normalized, rawEventId);
    case 'calendar':
      return createCalendarEvent_(normalized, rawEventId);
    case 'entity':
      return createEntity_(normalized, rawEventId);
    default:
      return createNote_(normalized, rawEventId);
  }
}

function handleQuery_(payload) {
  const table = payload.table || 'notes';
  const filters = payload.filters || {};
  const limit = Number(payload.limit || 50);
  return {
    ok: true,
    data: queryTable_(table, filters, { limit })
  };
}
