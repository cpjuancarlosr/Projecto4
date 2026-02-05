const CONFIG = {
  spreadsheetId: '',
  defaultLocale: 'es-MX',
  timezone: 'America/Mexico_City',
  webApp: {
    allowCORS: true
  },
  telegram: {
    botToken: '',
    webhookSecret: ''
  },
  speech: {
    languageCode: 'es-MX'
  },
  tables: {
    raw_events: [
      'raw_event_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'intent',
      'payload',
      'entity_id',
      'status',
      'priority'
    ],
    notes: [
      'note_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'title',
      'body',
      'tags',
      'status',
      'priority',
      'raw_event_id'
    ],
    tasks: [
      'task_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'title',
      'details',
      'due_date',
      'status',
      'priority',
      'raw_event_id',
      'linked_note_id',
      'linked_email_id'
    ],
    calendar_events: [
      'calendar_event_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'title',
      'start_time',
      'end_time',
      'location',
      'status',
      'priority',
      'raw_event_id',
      'linked_task_id'
    ],
    financial_movements: [
      'movement_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'movement_type',
      'amount',
      'currency',
      'description',
      'status',
      'priority',
      'raw_event_id'
    ],
    emails_log: [
      'email_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'subject',
      'body',
      'recipient',
      'status',
      'priority',
      'raw_event_id',
      'linked_task_id'
    ],
    entities: [
      'entity_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'name',
      'type',
      'metadata',
      'status',
      'priority'
    ],
    knowledge_treasure: [
      'treasure_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'entity_id',
      'content',
      'tags',
      'status',
      'priority',
      'raw_event_id'
    ],
    tags: [
      'tag_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'label',
      'color',
      'status',
      'priority'
    ],
    relations: [
      'relation_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'from_table',
      'from_id',
      'to_table',
      'to_id',
      'relation_type',
      'status',
      'priority'
    ],
    logs: [
      'log_id',
      'timestamp',
      'source',
      'device',
      'actor',
      'level',
      'message',
      'payload'
    ]
  }
};
