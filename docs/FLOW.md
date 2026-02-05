# End-to-End Data Flow

1. Capture arrives via Telegram or the HTML Web App.
2. `handleCapture` normalizes payloads and writes `raw_events`.
3. Router sends data to the correct table.
4. Logs and relations are recorded for traceability.
5. Queries use `runQuery` to filter and return rows.
