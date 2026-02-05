# Universal Capture System - Setup Guide

## 1. Google Sheets Schema (Data Lake)
Create a master spreadsheet and add a sheet per table. Use the `Initialize Schema` menu option in Apps Script to auto-create the tables with headers.

**Tables**
- raw_events
- notes
- tasks
- calendar_events
- financial_movements
- emails_log
- entities
- knowledge_treasure
- tags
- relations
- logs

## 2. Script Properties
Set the following Script Properties in Apps Script (**Project Settings → Script properties**):

| Key | Example | Purpose |
| --- | --- | --- |
| SPREADSHEET_ID | `1AbC...` | Master spreadsheet ID (optional if bound to sheet) |
| TELEGRAM_BOT_TOKEN | `12345:ABC...` | Telegram bot token |
| TELEGRAM_WEBHOOK_SECRET | `secret` | Shared secret for webhook security |
| GOOGLE_SPEECH_API_KEY | `AIza...` | Speech-to-Text API key |
| ALLOWED_ENTITY_IDS | `entity_default,entity_professional` | Whitelist for entity access |
| DEFAULT_ENTITY_ID | `entity_default` | Default entity for captures |

## 3. Telegram Bot Setup
1. Create a bot with `@BotFather` and get the token.
2. Deploy Apps Script as a Web App (`Execute as: Me`, `Who has access: Anyone`).
3. Set webhook with secret:

```
https://api.telegram.org/bot<TELEGRAM_BOT_TOKEN>/setWebhook?url=<WEB_APP_URL>&secret_token=<TELEGRAM_WEBHOOK_SECRET>
```

## 4. HTML Capture Interface
Open the Web App URL in any browser (iPhone, iPad, desktop). The interface supports:
- Notes
- Tasks
- Calendar events
- Financial movements
- Knowledge treasure
- Email log

## 5. Optional Integrations
- **Google Tasks**: Enable the Tasks Advanced Service in Apps Script.
- **Google Calendar**: Uses the default calendar by default.
- **Email logging**: Connect to Gmail with `GmailApp` if you want to send + log.

## 6. Data Flow
1. Capture event arrives via Telegram or HTML Web App.
2. Event is normalized and logged to `raw_events`.
3. Router parses intent and writes to target table.
4. Optional integrations sync to Google Tasks/Calendar.
5. Activity is auditable via `logs`.
