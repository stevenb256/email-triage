# Copilot instructions for email-triage

This file is a concise reference for automated Copilot sessions working on this repository.

---
1) Build / run / test / lint

- Primary app: Python Flask app implemented in app.py.
  - Activate the existing venv: `source venv/bin/activate` (venv/ is included in the repo snapshot).
  - Ensure credentials: copy `.env.example` → `.env` and set ANTHROPIC_API_KEY.
  - Start the server (dev): `python3 app.py` — service listens on http://localhost:5001 by default.
  - DB: SQLite file: `email_triage.db` created/updated automatically by the app.

- Node: package.json contains a `start` script (`node server.js`) but this repo's primary server is the Flask app in app.py. Use `npm start` only if a `server.js` is present or you add a Node-based server.

- Tests / Lint: No test suite or linter configuration found in repository. (If tests are added, use the project's test runner; no single-test command is present today.)

---
2) High-level architecture (big picture)

- Single-process Flask backend (app.py) that:
  - Runs a background sync worker (thread + asyncio event loop) to talk to a local MCP server/binary and sync Outlook messages.
  - Stores email and thread state in a local SQLite DB (`email_triage.db`) with three main tables: `emails`, `threads`, `meta`.
  - Uses the anthropic Python SDK to call LLM models for analysis and reply generation (constants: ANALYSIS_MODEL, REPLY_MODEL).
  - Exposes REST endpoints under `/api/*` for the frontend to fetch threads, messages, streaming AI-parsed paragraphs, trigger sync/resync, and perform mailbox actions (reply/move/delete/markread).
  - Serves a self-contained single-file frontend: HTML is embedded in the `HTML` string and returned from `/`.

- MCP integration:
  - app.py uses an MCP client (mcp.ClientSession + stdio_client) to call Outlook-specific tools such as `outlook_mail_list_folders`, `outlook_mail_list_messages`, `outlook_mail_get_message`, `outlook_mail_draft_message`, `outlook_mail_send_message`, etc.
  - MCP binary is expected at the path set by MCP_COMMAND (default `/opt/homebrew/bin/McpOutlookLocal`). The MCP session is started at process boot and remains available to call_tool().

- Data flow summary:
  - Background sync lists folders, fetches recent Inbox messages, inserts new rows into `emails`, groups messages to `threads`, and calls the LLM to analyze each thread once per thread.
  - `formatted_body` (JSON) caches AI-parsed paragraphs for each message.

---
3) Key conventions and repo-specific patterns

- Conversation key normalization:
  - conversation_key is derived via `_norm_subject()` which strips common prefixes (RE/FW/etc.) and lowercases the subject; this drives thread grouping.

- Topic normalization:
  - LLM topics are mapped into `CANONICAL_TOPICS` using `_TOPIC_RULES` and `_normalize_topic()`; prefer making changes to these constants (not ad-hoc string checks) when altering topic mapping.

- DB and migrations:
  - `init_db()` creates tables and attempts idempotent ALTER TABLE migrations on startup.
  - `meta` table is used for small cached items (keys: `efforts_subfolders`, `other_folders`, `folders_raw`).

- AI behavior hooks:
  - `analyze_thread()` builds the prompt and enforces a strict JSON-only response. Changes to analysis behavior or prompt text should be done here.
  - `_format_message_with_ai()` and `_format_prompt()` control per-message paragraph parsing.
  - Models live in constants `ANALYSIS_MODEL` and `REPLY_MODEL` at the top of app.py.

- MCP usage pattern:
  - All MCP calls go through `call_tool(name, args)` which waits for the MCP session and returns structured content or parsed JSON.
  - MCP tool names are stable; look for calls like `outlook_mail_list_folders`, `outlook_mail_get_message`, and `outlook_mail_move_message` if modifying mailbox behavior.

- Caching and re-sync:
  - `formatted_body` stores the AI-parsed paragraphs (JSON) used by `/api/format_message` and `/api/format_message_stream`.
  - Use `/api/resync_thread` to rebuild a single thread from mailbox data; `/api/reanalyze_all` re-runs analysis over all threads.

- Background sync timing:
  - SYNC_INTERVAL is 300 seconds (5 minutes). Change at top-level constants if needed.

---
4) Files/areas to inspect when making changes

- app.py — single source of truth for backend, sync logic, AI prompts, DB shape, and the embedded frontend.
- `.env.example` — required environment variables (ANTHROPIC_API_KEY).
- `email_triage.db` — local SQLite DB used for debugging and fast UI loads.
- MCP binary location (MCP_COMMAND in app.py) and the local MCP client code paths if you change how Outlook is integrated.

---
5) AI assistant / CI helper config

- No additional AI assistant configs detected (CLAUDE.md, AGENTS.md, .cursorrules, AIDER_CONVENTIONS.md, .windsurfrules, .clinerules, etc.).

---
If anything here should be expanded (examples, troubleshooting steps, or extra commands for dev containers / CI), say which area to extend and Copilot will add it.
