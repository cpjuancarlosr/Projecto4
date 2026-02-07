# Configuration Sources

These values are set in **Apps Script → Project Settings → Script Properties**. Do **not** hardcode real IDs or tokens in source code.

## SPREADSHEET_ID
- **Where to get it:** Open the Google Sheet that will act as your database. The ID is the long string between `/d/` and `/edit` in the URL.
- Example URL: `https://docs.google.com/spreadsheets/d/1AbC...xyz/edit#gid=0`
- The ID is `1AbC...xyz`.

## TELEGRAM_BOT_TOKEN
- **Where to get it:** Create a bot using **@BotFather** in Telegram. BotFather returns a token that looks like `123456789:AA...`.
- Keep this private. If leaked, revoke and regenerate it in BotFather.

## TELEGRAM_ALLOWED_CHAT_IDS
- **Where to get it:** Send a message to your bot, then call:
  - `https://api.telegram.org/bot<YOUR_TOKEN>/getUpdates`
- Inspect the JSON response and grab `message.chat.id` values for the chats you want to allow.
- Use a comma-separated list for multiple IDs (e.g., `12345678,98765432`).

## API_SHARED_SECRET
- **Where to get it:** Generate any strong random string you control.
- This protects the HTML capture UI API calls. If you set it, pass `?secret=YOUR_SECRET` in the web app URL.

## SPEECH_API_KEY
- **Where to get it:** Create a project in Google Cloud Console and enable **Speech-to-Text API**.
- Generate an API key and restrict it to the Speech-to-Text API if possible.
- Required only for voice transcription.
