const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  TELEGRAM_BOT_TOKEN: PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN'),
  TELEGRAM_ALLOWED_CHAT_IDS: (PropertiesService.getScriptProperties().getProperty('TELEGRAM_ALLOWED_CHAT_IDS') || '')
    .split(',')
    .map((id) => id.trim())
    .filter((id) => id),
  API_SHARED_SECRET: PropertiesService.getScriptProperties().getProperty('API_SHARED_SECRET'),
  SPEECH_API_KEY: PropertiesService.getScriptProperties().getProperty('SPEECH_API_KEY'),
  DEFAULT_TIMEZONE: Session.getScriptTimeZone(),
  TABLES: {
    RAW_EVENTS: 'raw_events',
    NOTES: 'notes',
    TASKS: 'tasks',
    CALENDAR_EVENTS: 'calendar_events',
    FINANCIAL_MOVEMENTS: 'financial_movements',
    EMAILS_LOG: 'emails_log',
    ENTITIES: 'entities',
    KNOWLEDGE_TREASURE: 'knowledge_treasure',
    TAGS: 'tags',
    RELATIONS: 'relations',
    LOGS: 'logs'
  }
};

const HEADERS = {
  raw_events: [
    'id',
    'timestamp',
    'source',
    'device',
    'actor',
    'intent',
    'payload',
    'entity_id',
    'status',
    'priority',
    'tags',
    'correlation_id'
  ],
  notes: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'content',
    'tags',
    'raw_event_id'
  ],
  tasks: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'title',
    'details',
    'due_date',
    'raw_event_id',
    'external_id'
  ],
  calendar_events: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'title',
    'details',
    'start_at',
    'end_at',
    'location',
    'raw_event_id',
    'external_id'
  ],
  financial_movements: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'movement_type',
    'amount',
    'currency',
    'category',
    'description',
    'raw_event_id'
  ],
  emails_log: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'subject',
    'to',
    'cc',
    'body_excerpt',
    'raw_event_id',
    'gmail_id'
  ],
  entities: [
    'id',
    'timestamp',
    'name',
    'type',
    'status',
    'priority',
    'notes',
    'metadata'
  ],
  knowledge_treasure: [
    'id',
    'timestamp',
    'source',
    'device',
    'entity_id',
    'status',
    'priority',
    'content',
    'tags',
    'raw_event_id'
  ],
  tags: [
    'id',
    'timestamp',
    'label',
    'scope',
    'notes'
  ],
  relations: [
    'id',
    'timestamp',
    'from_table',
    'from_id',
    'to_table',
    'to_id',
    'relation_type',
    'notes'
  ],
  logs: [
    'id',
    'timestamp',
    'level',
    'source',
    'message',
    'payload'
  ]
};
