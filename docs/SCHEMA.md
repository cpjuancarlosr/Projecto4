# Google Sheets Schema

Each sheet is a table. Initialize with `initSchema()`.

## raw_events
- id (UUID)
- timestamp (ISO 8601)
- source
- device
- actor
- intent
- payload (JSON)
- entity_id
- status
- priority
- tags
- correlation_id

## notes
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- content
- tags
- raw_event_id

## tasks
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- title
- details
- due_date
- raw_event_id
- external_id

## calendar_events
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- title
- details
- start_at
- end_at
- location
- raw_event_id
- external_id

## financial_movements
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- movement_type
- amount
- currency
- category
- description
- raw_event_id

## emails_log
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- subject
- to
- cc
- body_excerpt
- raw_event_id
- gmail_id

## entities
- id
- timestamp
- name
- type
- status
- priority
- notes
- metadata

## knowledge_treasure
- id
- timestamp
- source
- device
- entity_id
- status
- priority
- content
- tags
- raw_event_id

## tags
- id
- timestamp
- label
- scope
- notes

## relations
- id
- timestamp
- from_table
- from_id
- to_table
- to_id
- relation_type
- notes

## logs
- id
- timestamp
- level
- source
- message
- payload
