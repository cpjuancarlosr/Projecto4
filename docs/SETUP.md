# Universal Capture System (Apps Script + Sheets + Telegram)

## Overview
This system treats Google Sheets as a relational data lake and Google Apps Script as the orchestration layer. Every capture is logged to `raw_events` first, then routed to a normalized table for long-term usage.

## Spreadsheet Schema (One sheet per table)
Create a single spreadsheet (or use a container-bound script) and ensure these sheets exist with the following headers (order matters):

- raw_events
  - raw_event_id, timestamp, source, device, actor, intent, payload, entity_id, status, priority
- notes
  - note_id, timestamp, source, device, actor, entity_id, title, body, tags, status, priority, raw_event_id
- tasks
  - task_id, timestamp, source, device, actor, entity_id, title, details, due_date, status, priority, raw_event_id, linked_note_id, linked_email_id
- calendar_events
  - calendar_event_id, timestamp, source, device, actor, entity_id, title, start_time, end_time, location, status, priority, raw_event_id, linked_task_id
- financial_movements
  - movement_id, timestamp, source, device, actor, entity_id, movement_type, amount, currency, description, status, priority, raw_event_id
- emails_log
  - email_id, timestamp, source, device, actor, entity_id, subject, body, recipient, status, priority, raw_event_id, linked_task_id
- entities
  - entity_id, timestamp, source, device, actor, name, type, metadata, status, priority
- knowledge_treasure
  - treasure_id, timestamp, source, device, actor, entity_id, content, tags, status, priority, raw_event_id
- tags
  - tag_id, timestamp, source, device, actor, label, color, status, priority
- relations
  - relation_id, timestamp, source, device, actor, from_table, from_id, to_table, to_id, relation_type, status, priority
- logs
  - log_id, timestamp, source, device, actor, level, message, payload

Run `ensureSchema_()` or deploy the web app once to auto-create tables.

## Apps Script Deployment
1. Create a new Apps Script project (standalone or bound to the spreadsheet).
2. Copy the `apps-script` files into the project.
3. Add the spreadsheet ID in `CONFIG.spreadsheetId` if using standalone.
4. Deploy as a Web App (execute as you, accessible to anyone with the link or your domain).
5. Save the Web App URL for Telegram and HTML capture.

## Telegram Bot Setup
1. Create a bot with BotFather and copy the token into `CONFIG.telegram.botToken`.
2. Set the webhook:
   ```
   https://api.telegram.org/bot<YOUR_TOKEN>/setWebhook?url=<WEB_APP_URL>
   ```
3. Send a Telegram message or voice note to the bot. The script will log to `raw_events` and route based on commands.

### Example Commands
- `/note idea sobre cliente X`
- `/task llamar a proveedor | 2026-02-10 | alta`
- `/money gasto 850 MXN gasolina entidad:personal`
- `/treasure palabra favorita: criterio`
- `/email enviar propuesta cliente Y`

Voice notes will attempt transcription using Google Speech-to-Text (requires Cloud Speech API enabled and billing).

## HTML Capture Interface
Visit the Web App URL in a browser to use the mobile-first capture UI. It supports quick capture and SQL-light querying (filters via `campo:valor`).

## Data Flow (End-to-End)
1. Capture enters via Telegram or HTML.
2. `handleCapture_` logs to `raw_events`.
3. Intent detection routes to the normalized table.
4. Integrations optionally sync to Google Tasks/Calendar/Gmail.
5. Queries run via `handleQuery_` for dashboards or exports.

## Extending the System
- Add a new table by adding headers in `CONFIG.tables` and a new parser/creator.
- Extend parsing with domain-specific commands in `parseCommand_`.
- Add new integrations by hooking into `routeIntent_` or post-insert hooks.
