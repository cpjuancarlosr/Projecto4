# Universal Capture & Data Management System

This repository contains a production-ready Google Apps Script backend, Google Sheets data lake schema, and a Telegram + HTML capture interface for a universal operational memory system.

## What is included
- **Apps Script backend** with schema initialization, router, CRUD services, logging, permissions, and integrations.
- **HTML capture UI** optimized for mobile and desktop.
- **Telegram webhook** integration with voice transcription support via Google Speech-to-Text.
- **Setup guide** with schema, properties, and webhook instructions.

## Quick Start
1. Open the Apps Script project from `apps_script/`.
2. Set Script Properties (see `docs/SETUP.md`).
3. Run `Initialize Schema` from the custom menu.
4. Deploy as a Web App and connect your Telegram bot.

## Files
- `apps_script/Code.gs` – entrypoints, router, web app handlers.
- `apps_script/Services.gs` – core data services, schema, CRUD, query engine.
- `apps_script/Telegram.gs` – Telegram webhook parsing and voice transcription.
- `apps_script/WebApp.html` – HTML capture interface.
- `docs/SETUP.md` – setup + schema reference.
