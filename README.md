# Universal Capture & Data Management System (Apps Script + Sheets)

This repository contains a production-ready Google Apps Script backend, Google Sheets schema, and a lightweight HTML capture interface that act as a unified operational memory.

## Core Philosophy
- Google Sheets as the single source of truth (raw data first).
- Apps Script handles validation, routing, automation, and permissions.
- Capture is faster than thinking.
- Immutable raw event log (`raw_events`) for full auditability.

## Repository Layout
- `apps_script/` — Apps Script source code and HTML UI.
- `docs/SCHEMA.md` — Database-style schema for Sheets.
- `docs/SETUP_TELEGRAM.md` — Telegram bot setup.
- `docs/CONFIGURATION.md` — Where to obtain Script Property values.
- `docs/FLOW.md` — End-to-end data flow.

## Quick Start
1. Create a Google Sheet to host the tables.
2. Open Apps Script, paste the contents of `apps_script/` as script files.
3. Add Script Properties:
   - `SPREADSHEET_ID`
   - `TELEGRAM_BOT_TOKEN`
   - `TELEGRAM_ALLOWED_CHAT_IDS`
   - `API_SHARED_SECRET`
   - `SPEECH_API_KEY`
   (See `docs/CONFIGURATION.md` for where to obtain each value.)
4. Run `initSchema()` once to create all tables.
5. Deploy as a Web App for Telegram and HTML capture.

## Core Features
- Universal capture router with intent detection.
- Telegram bot capture (text, voice, image metadata).
- Mobile-first web capture interface with SQL-light queries.
- CRUD services for all tables.
- Hooks for Google Tasks, Calendar, and Gmail.
- Extensible tables and router for new data types.

## Key Apps Script Entrypoints
- `doPost(e)` — Telegram webhook + API endpoint.
- `doGet(e)` — HTML capture UI.
- `handleCapture(payload)` — Central router.
- `runQuery(query)` — SQL-light query engine.

Refer to `docs/` for schema, setup, and flow details.
