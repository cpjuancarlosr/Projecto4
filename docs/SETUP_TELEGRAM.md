# Telegram Bot Setup

1. Create a bot with BotFather and copy the token.
2. Deploy the Apps Script project as a Web App.
3. Set Script Properties:
   - `SPREADSHEET_ID` (Google Sheet ID)
   - `TELEGRAM_BOT_TOKEN`
   - `TELEGRAM_ALLOWED_CHAT_IDS` (comma-separated list)
   - `API_SHARED_SECRET` (optional shared secret for the web UI)
   - `SPEECH_API_KEY` (Google Cloud Speech-to-Text API key)
   - See `docs/CONFIGURATION.md` for where to obtain each value.
4. Register webhook:
   - `https://api.telegram.org/bot<token>/setWebhook?url=<web_app_url>`
5. Send a message to the bot:
   - `/note idea sobre cliente X`
   - `/task llamar a proveedor | 2026-02-10 | alta`
   - `/money gasto 850 MXN gasolina entidad:personal`
   - `/treasure palabra favorita: criterio`
   - `/email enviar propuesta cliente Y`

Voice notes are transcribed with Google Speech-to-Text if `SPEECH_API_KEY` is set.
