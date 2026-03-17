# Email Triage App — Requirements

Personal email triage app for a senior tech leader at Microsoft. Single-file Flask app (`app.py`) connecting Outlook via MCP, Claude AI via the Anthropic API, and a SQLite database. Runs on `localhost:5002`.

---

## Architecture

- **Backend**: Flask (Python), single file `app.py`
- **Database**: SQLite (`email_triage.db`) — emails, threads, calendar_events, meta
- **Email integration**: Outlook via MCP (`McpOutlookLocal`) using `stdio_client`
- **AI**: Anthropic Claude — Haiku for analysis/formatting, Sonnet for reply drafts
- **Frontend**: Single embedded HTML template with inline CSS + JS, no build step
- **Sync**: Background thread runs every 5 minutes

---

## Data Model

### `emails` table
- `id`, `subject`, `from_name`, `from_address`, `received_date_time`, `is_read`
- `body_preview` (500 chars), `conversation_key`, `raw_json`, `synced_at`, `formatted_body`, `folder`

### `threads` table
- `conversation_key` (PK), `subject`, `topic`, `action`, `urgency`, `summary`
- `suggested_reply`, `suggested_folder`, `participants` (JSON), `email_ids` (JSON)
- `latest_id`, `message_count`, `has_unread`, `latest_received`, `updated_at`, `is_flagged`

### `calendar_events` table
- `id`, `subject`, `start_time`, `end_time`, `location`, `attendees`, `raw_json`, `synced_at`

### `meta` table
- Key-value store for: `my_email`, `efforts_subfolders`, `other_folders`, `folders_raw`, `next_meeting`

---

## Sync System

- Inbox syncs every **5 minutes** (SYNC_INTERVAL), fetching up to 100 messages at a time
- Non-inbox folders sync on same cycle, 100 messages per folder, paginated
- **Skipped folders**: Drafts, Outbox, Junk Email, Conversation History, RSS Feeds, Sync Issues, Scheduled, Inbox (re-included in first pass)
- On startup: refreshes folders, refreshes calendar, then kicks off first sync
- Sync status tracked (running, phase, progress, emailsAdded, threadsUpdated, lastError)
- **Inbox threads**: AI-analyzed on every new thread; result stored in `threads` table
- **Non-inbox threads**: stored in DB but not AI-analyzed during sync
- Sync can be triggered manually via "⟳ Sync Now" button
- A single thread can be fully re-fetched and re-analyzed via "↺ Resync Thread"
- All threads can be re-analyzed via "⚙ Re-analyze" (background job)

---

## AI Integration

### 1. Thread Analysis (Claude Haiku)
Called for each new inbox thread during sync, and on-demand via resync/reanalyze.

Inputs: all emails in thread, folder lists, optional user-provided reply context

Outputs (stored in `threads`):
- **summary**: 2–4 dense sentences naming people, dates, metrics, decisions, blockers, next steps
- **topic**: Engineering | Product Planning | Finance | Incidents & Outages | Team & HR | Partnerships | FYI & Updates | Strategy & Leadership
- **action**: `reply` | `delete` | `file` | `read` | `done`
- **urgency**: `high` | `medium` | `low`
- **suggestedReply**: complete send-ready reply draft (always generated unless action=delete)
- **suggestedFolder**: exact folder name for filing (when action=file)

Reply tone: direct, warm, confident — no hollow openers. Tailored to thread type (status update, question, incident, cross-team, etc.). Specific names/numbers/details.

### 2. Message Formatting (Claude Haiku)
Per-message, on-demand (lazy-loaded when a message is expanded). Cached after first run.

Output: array of paragraphs, each with:
- Exact paragraph text
- Intent label: Status Update | Request | Decision | Question | Action Item | Context | FYI | Warning | Introduction | Closing
- Emoji for the intent
- Optional fact-check warning (shown with ⚠️)

### 3. Reply Generation (Claude Sonnet 4.6)
On-demand via `/api/generate_reply`. Takes user's intent/notes + thread context. Returns polished professional reply.

### 4. Suggested Reply Refresh
`/api/suggested_reply` — re-runs thread analysis with optional user-provided context for regeneration. Updates cached `suggested_reply` in DB.

---

## Views

### App always boots to Triage Sheet view.

---

### Triage Sheet
- Default view on startup
- Groups threads by AI topic, each group collapsible
- Threads sorted newest-first within each group
- Per thread row:
  - Urgency pill, subject, AI-suggested action pill with icon
  - Expand to see AI summary + inline message previews (newest-first)
  - Buttons: **Reply**, **File**, **Delete**, **Clear** (clear pending action)
- Queued actions shown with colored border (blue=reply, yellow=file, red=delete)
- **⚡ Execute All** button processes all queued actions sequentially with progress indicator
- After execute-all: triage sheet re-renders, completed items removed
- Keyboard shortcut reference shown at bottom

**Triage keyboard shortcuts:**
| Key | Action |
|-----|--------|
| `↑` / `↓` or `j` / `k` | Navigate threads |
| `Enter` / `Space` | Expand/collapse thread or topic |
| `→` / `←` | Expand/collapse |
| `r` | Reply |
| `d` | Mark for delete |
| `f` | Mark for file |
| `x` | Clear action |
| `Esc` | Exit to mailbox |

---

### Mailbox (Folder View)
- Accessible by clicking any folder in the left nav
- Folder list sidebar: Inbox → Efforts (with subfolders) → Partners → Deleted Items → Sent Items → other folders
- Folder view shows threads, sorted newest-first, with:
  - Unread dot (blue)
  - Subject, sender, date, preview
  - Message count badge (if >1 message)
  - On-hover quick action buttons: **↩ Reply**, **🗑 Delete**
- Keyboard hint bar at bottom of folder list

**Mailbox list keyboard shortcuts:**
| Key | Action |
|-----|--------|
| `j` / `↓` | Next thread |
| `k` / `↑` | Previous thread |
| `Enter` / `→` | Open thread |
| `r` | Quick reply |
| `d` | Delete |
| `f` | File |
| `u` | Toggle read/unread |
| `Esc` / `←` | Back to folder |

**Thread detail keyboard shortcuts (while thread is open):**
| Key | Action |
|-----|--------|
| `j` / `↓` | Next thread in folder |
| `k` / `↑` | Previous thread in folder |
| `r` | Reply |
| `d` | Delete |
| `f` | File |
| `Esc` / `←` | Back to folder list |

---

### Thread Detail View
Shared between email (triage) view and mailbox view.

**Thread header** contains:
- ✕ Close button (when in mailbox context)
- Urgency pill + action pill + unread dot (all on same line as subject)
- Subject (truncated)
- Date
- Participant avatars (colored by name hash, up to 5) + names
- Message count
- AI Summary box (🤖 AI Summary label)
- Action buttons: **↩ Reply**, **📁 File** (or "📁 {suggestedFolder}"), **🚩 Flag**, **🗑 Delete**

**Message list** (below header):
- Edge-to-edge cards, no horizontal padding
- Each card shows: sender avatar, sender name + subject preview on same line, date, format toggle button
- Expanded card shows: To/CC recipients, message body
- Body view toggle: **AI view** (default, formatted paragraphs with intent pills) | **HTML** (iframe) | **Original** (plain text)
- AI view shows fact-check warnings inline
- Messages sorted **newest-first**

For threads opened from folders not yet in the triage data: synthetic thread object is created so action buttons always appear; background analysis runs to populate suggested reply.

---

### Calendar View
- Day view by default, with Week view toggle
- Navigation: prev/next arrows, Today button, Day/Week toggle
- **Day view**: vertical time grid (7am–9pm), 30-min slots, events positioned by start/end time
- **Week view**: 7-column grid (Mon–Sun) with same time range; all-day events row at top
- Events colored by subject hash, show title + time range + location
- Scroll position reset to 8am on render
- No AI meeting prep (removed)

---

### Search
- Triggered by typing in the search bar (Enter or 2+ chars)
- Full-text search across subject, from name/address, body preview
- Results show subject, sender, date, preview, folder badge
- Click result: opens thread in correct context (triage or mailbox depending on folder)

---

## Reply System

**Reply All behavior:**
- Sender of latest message → To field
- All To recipients from latest message → To field
- All CC recipients from latest message → CC field
- **Current user is filtered out** from all recipient fields

**Reply modal:**
- Pre-populated with AI-generated suggested reply on open (uses cached value or fetches fresh)
- Shows "Generating..." spinner while fetching
- **↺ Regenerate** button: takes whatever is currently typed as context, calls `/api/suggested_reply` with that context, replaces textarea content
- Editable To/CC recipient tags (click × to remove)
- Sends via Outlook MCP (`ReplyAll` operation)
- After send: thread removed from triage/mailbox

---

## Compose (New Message)

- **✉ New Message** button always visible in title bar
- Large compose modal with:
  - **To** field with people picker autocomplete
  - **CC** field with people picker autocomplete
  - **Subject** text input
  - **Body** textarea (large, resizable)
- People picker:
  - Queries `/api/people` — deduplicated list of all senders from received email history, excluding self
  - Pre-cached at startup for instant suggestions
  - Filtered as you type (name or address, case-insensitive)
  - Arrow keys to navigate, Enter or comma to add, Backspace to remove last
  - Also accepts raw email addresses directly
- Sends via Outlook MCP
- Toast notification on success or failure

---

## Sidebar

- **📋 Triage Sheet** button
- Folder tree (Inbox, Efforts + subfolders, Partners, Deleted Items, Sent Items, other folders)
- **📅 Calendar** button
- **Today's schedule**: compact list of today's meetings (time in light blue, meeting name on same line), past meetings faded
- **Week hours**: "Xh in meetings this week" computed from Mon–Fri calendar events
- Bottom bar: message count, thread count, sync status, sync progress bar

---

## Header / Title Bar

Always visible across all views. Contains:
- Brand: "Email"
- Search input
- **✉ New Message** button
- **⟳ Sync Now** button
- **↺ Resync Thread** button (enabled when a thread is selected)
- **⚙ Re-analyze** button

---

## Delete Behavior
- Moves messages to Deleted Items folder via Outlook MCP
- Removes thread from local SQLite DB
- After deleting while viewing a thread: advances to next thread in list, or returns to folder/list if none

---

## File / Move Behavior
- Moves messages to chosen folder via Outlook MCP
- Folder picker groups Efforts subfolders separately from other folders
- Pre-selects AI-suggested folder when available
- After filing: thread removed from triage view

---

## Mark Read / Unread
- Updates Outlook and local DB
- Unread threads show blue dot in mailbox list
- Toggleable via `u` keyboard shortcut in folder view

---

## Flag
- Flags/unflags thread in Outlook and local DB
- Flag button updates immediately without reload

---

## Visual Design

- Dark navy theme (`#0a1628` base, `#0d2040` surface, `#1a3252` border)
- Low-contrast grey text replaced with readable cyan-blue (`#5ba4cf`)
- Message list extends edge-to-edge (no horizontal padding)
- Thread header: pills + subject + date all on one line
- Avatars: initials, color-coded deterministically by name
- Monospace font throughout (`Monaco`, `Menlo`, `Courier New`)
- Sidebar resizable via drag handle
- Animations: modal open (scale + fade), row deletion (opacity fade)

---

## Configuration

- `MCP_COMMAND`: path to `McpOutlookLocal` binary
- `DB_PATH`: path to SQLite DB (default: same directory as `app.py`)
- `PORT`: 5002
- `SYNC_INTERVAL`: 300 seconds
- `INBOX_FETCH` / `FOLDER_FETCH`: 100 messages per request
- `ANALYSIS_MODEL`: `claude-haiku-4-5-20251001`
- `REPLY_MODEL`: `claude-sonnet-4-6`
- `ANTHROPIC_API_KEY`: from `.env` or environment

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/threads` | All threads grouped by topic |
| GET | `/api/updates?since=` | Incremental poll for new/changed threads |
| GET | `/api/thread_messages?conversationKey=` | All messages for a thread (newest-first) |
| GET | `/api/format_message?id=` | AI-formatted paragraphs for a message (cached) |
| GET | `/api/format_message_stream?id=` | SSE stream of AI formatting |
| GET | `/api/calendar?start=&end=` | Calendar events in range |
| GET | `/api/mailbox/folders` | Folder tree with counts |
| GET | `/api/mailbox/folder?folder=&offset=` | Threads in a folder |
| GET | `/api/search?q=` | Full-text search |
| GET | `/api/my_email` | Current user's email address |
| GET | `/api/people?q=` | Deduplicated sender list for people picker |
| POST | `/api/sync_now` | Trigger immediate sync |
| POST | `/api/resync_thread` | Re-fetch + re-analyze single thread |
| POST | `/api/reanalyze_all` | Re-analyze all threads |
| POST | `/api/suggested_reply` | Get/refresh AI suggested reply (accepts context) |
| POST | `/api/generate_reply` | Generate reply from user intent prompt |
| POST | `/api/reply/{latest_id}` | Send reply-all via Outlook |
| POST | `/api/send_new` | Send new message via Outlook |
| POST | `/api/delete` | Delete messages (move to Deleted Items) |
| POST | `/api/move` | Move messages to folder |
| POST | `/api/markread` / `/api/mark_read` | Mark messages read or unread |
| POST | `/api/flag` | Flag/unflag thread |
| POST | `/api/meeting_prep` | Generate meeting preparation notes (exists, unused in UI) |
