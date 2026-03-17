#!/usr/bin/env python3
"""
Email Triage — Outlook MCP + Anthropic API + SQLite
• Syncs inbox every 5 min, only processing new messages
• Analyzes each thread with Claude Haiku — one call per thread
• Web UI loads instantly from DB, updates incrementally every 10s
"""

import asyncio
import json
import os
import re
import sqlite3
import threading
import time
import webbrowser
from datetime import datetime, timezone, timedelta

import anthropic
from flask import Flask, jsonify, render_template_string, request
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# ─── Config ────────────────────────────────────────────────────────────────────

MCP_COMMAND     = "/opt/homebrew/bin/McpOutlookLocal"
DB_PATH         = os.path.join(os.path.dirname(__file__), "email_triage.db")
PORT            = 5002
SYNC_INTERVAL   = 300   # 5 minutes
INBOX_FETCH     = 100
FOLDER_FETCH    = 100   # messages per non-inbox folder per sync cycle
ANALYSIS_MODEL  = "claude-haiku-4-5-20251001"
REPLY_MODEL     = "claude-sonnet-4-6"

SKIP_SYNC_FOLDERS = {
    "Drafts", "Outbox", "Junk Email",
    "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled", "Inbox", "",
}

_env_path = os.path.join(os.path.dirname(__file__), ".env")
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

# ─── Database ──────────────────────────────────────────────────────────────────

_thread_local = threading.local()


def get_db() -> sqlite3.Connection:
    if not hasattr(_thread_local, "conn"):
        conn = sqlite3.connect(DB_PATH, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        _thread_local.conn = conn
    return _thread_local.conn


def init_db():
    db = get_db()
    db.executescript("""
    CREATE TABLE IF NOT EXISTS emails (
        id                  TEXT PRIMARY KEY,
        subject             TEXT,
        from_name           TEXT,
        from_address        TEXT,
        received_date_time  TEXT,
        is_read             INTEGER DEFAULT 0,
        body_preview        TEXT,
        conversation_key    TEXT,
        raw_json            TEXT,
        synced_at           TEXT,
        formatted_body      TEXT
    );
    CREATE TABLE IF NOT EXISTS threads (
        conversation_key    TEXT PRIMARY KEY,
        subject             TEXT,
        topic               TEXT,
        action              TEXT,
        urgency             TEXT,
        summary             TEXT,
        suggested_reply     TEXT,
        suggested_folder    TEXT,
        participants        TEXT,
        email_ids           TEXT,
        latest_id           TEXT,
        message_count       INTEGER DEFAULT 0,
        has_unread          INTEGER DEFAULT 0,
        latest_received     TEXT,
        updated_at          TEXT
    );
    CREATE TABLE IF NOT EXISTS meta (
        key     TEXT PRIMARY KEY,
        value   TEXT
    );
    CREATE TABLE IF NOT EXISTS calendar_events (
        id              TEXT PRIMARY KEY,
        subject         TEXT,
        start_time      TEXT,
        end_time        TEXT,
        location        TEXT,
        attendees       TEXT,
        raw_json        TEXT,
        synced_at       TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_emails_conv_key ON emails(conversation_key);
    CREATE INDEX IF NOT EXISTS idx_threads_updated  ON threads(updated_at);
    CREATE INDEX IF NOT EXISTS idx_threads_urgency  ON threads(urgency);
    CREATE INDEX IF NOT EXISTS idx_calendar_start ON calendar_events(start_time);
    """)
    db.commit()
    # Migrations: add columns if not present (idempotent)
    for migration in [
        "ALTER TABLE emails ADD COLUMN formatted_body TEXT",
        "ALTER TABLE threads ADD COLUMN is_flagged INTEGER DEFAULT 0",
        "ALTER TABLE emails ADD COLUMN folder TEXT",
    ]:
        try:
            db.execute(migration)
            db.commit()
        except Exception:
            pass


def meta_get(key: str, default=None):
    row = get_db().execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def meta_set(key: str, value: str):
    db = get_db()
    db.execute("INSERT OR REPLACE INTO meta(key,value) VALUES(?,?)", (key, value))
    db.commit()


def get_my_email() -> str:
    """Detect the current user's email address from Sent Items or meta cache."""
    cached = meta_get("my_email", "")
    if cached:
        return cached
    db = get_db()
    row = db.execute(
        "SELECT from_address FROM emails WHERE folder='Sent Items' AND from_address != '' LIMIT 1"
    ).fetchone()
    if row:
        email = row["from_address"]
        meta_set("my_email", email)
        return email
    return ""


def _thread_to_dict(row) -> dict:
    d = dict(row)
    try:
        d["participants"] = json.loads(d.get("participants") or "[]")
    except Exception:
        d["participants"] = []
    try:
        d["emailIds"] = json.loads(d.get("email_ids") or "[]")
    except Exception:
        d["emailIds"] = []
    return {
        "conversationKey": d["conversation_key"],
        "subject":         d["subject"] or "",
        "topic":           d["topic"] or "General",
        "action":          d["action"] or "read",
        "urgency":         d["urgency"] or "low",
        "summary":         d["summary"] or "",
        "suggestedReply":  d["suggested_reply"] or "",
        "suggestedFolder": d["suggested_folder"] or "",
        "participants":    d["participants"],
        "emailIds":        d["emailIds"],
        "latestId":        d["latest_id"] or "",
        "messageCount":    d["message_count"] or 0,
        "hasUnread":       bool(d["has_unread"]),
        "isFlagged":       bool(d.get("is_flagged", 0)),
        "latestReceived":  d["latest_received"] or "",
        "updatedAt":       d["updated_at"] or "",
    }


def remove_thread(conv_key: str):
    db = get_db()
    db.execute("DELETE FROM emails WHERE conversation_key=?", (conv_key,))
    db.execute("DELETE FROM threads WHERE conversation_key=?", (conv_key,))
    db.commit()


# ─── MCP Session ───────────────────────────────────────────────────────────────

_loop = asyncio.new_event_loop()
_session: ClientSession | None = None
_session_ready = threading.Event()


async def _run_mcp():
    global _session
    params = StdioServerParameters(command=MCP_COMMAND, args=[])
    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            _session = session
            _session_ready.set()
            await asyncio.Event().wait()


def _bg_loop():
    asyncio.set_event_loop(_loop)
    _loop.run_forever()


threading.Thread(target=_bg_loop, daemon=True).start()
asyncio.run_coroutine_threadsafe(_run_mcp(), _loop)


def call_tool(name: str, args: dict):
    if not _session_ready.wait(timeout=20):
        raise RuntimeError("MCP session not ready")
    future = asyncio.run_coroutine_threadsafe(_session.call_tool(name, args), _loop)
    result = future.result(timeout=30)
    if result.isError:
        raise RuntimeError(f"MCP error: {result.content[0].text if result.content else 'unknown'}")
    if result.structuredContent:
        return result.structuredContent
    if result.content:
        try:
            return json.loads(result.content[0].text)
        except Exception:
            return result.content[0].text
    return None


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _norm_subject(subj: str) -> str:
    s = re.sub(r'^(RE|FW|FWD|AW|R|RES|SV)[\s:]+', '', str(subj or ''), flags=re.IGNORECASE)
    return s.strip().lower() or "no-subject"


def _clean(s, n=None) -> str:
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s or ''))
    return s[:n] if n else s


def _utcnow() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _folder_lists(folders_raw: list):
    skip = {"Drafts", "Sent Items", "Outbox", "Deleted Items", "Junk Email", ""}
    names = [
        f.get("display_name") or f.get("displayName", "")
        for f in folders_raw
        if (f.get("display_name") or f.get("displayName", "")) not in skip
    ]
    efforts = [n for n in names if n.startswith("Efforts")]
    other   = [n for n in names if not n.startswith("Efforts")]
    return efforts, other


# ─── Topic normalization ───────────────────────────────────────────────────────

CANONICAL_TOPICS = [
    "Engineering", "Incidents & Outages", "Product Planning", "Partnerships",
    "Finance", "Team & HR", "Customer Issues", "Legal & Compliance",
    "Events & Travel", "FYI & Updates", "Strategy & Leadership",
    "Architecture & Design", "External Communications",
]

# Ordered rules: first match wins (most specific first)
_TOPIC_RULES = [
    (["incident", "outage", "sev "],                                     "Incidents & Outages"),
    (["financ", "budget", "expense", "billing", "payment"],              "Finance"),
    (["legal", "compliance", "gdpr", "regulation"],                      "Legal & Compliance"),
    (["travel", "conference", "offsite", "summit"],                      "Events & Travel"),
    (["partnership", "customer stor", "customer engag"],                 "Partnerships"),
    (["team", " hr ", "hiring", "recruit", "headcount", "people ops"],   "Team & HR"),
    (["customer issue", "client issue", "support ticket"],               "Customer Issues"),
    (["architect", "system design"],                                     "Architecture & Design"),
    (["strateg", "leadership", "executive", "vision", "okr"],            "Strategy & Leadership"),
    (["external comm", "press", "announcement"],                         "External Communications"),
    (["product plan", "product updat", "product launch", "product feat",
      "product metric", "product rev", "product eval", "product qual",
      "product dev", "roadmap", "feature", "launch", "sprint"],          "Product Planning"),
    (["engineer", "infrastructure", "replatform", "deploy", "migration",
      "latency", "reliab", "scale", "resource alloc", "capacity",
      "cost", "performance", "metric", "tools", "develop"],              "Engineering"),
    (["project update", "status", "progress", "fyi", "update"],          "FYI & Updates"),
]


def _normalize_topic(raw: str) -> str:
    """Map a free-form LLM topic to a canonical category."""
    r = raw.lower()
    for c in CANONICAL_TOPICS:
        if c.lower() == r:
            return c                    # exact match
    for keywords, canonical in _TOPIC_RULES:
        if any(kw in r for kw in keywords):
            return canonical
    return "FYI & Updates"             # fallback


# ─── Anthropic Analysis ────────────────────────────────────────────────────────

_ai = None


def _get_ai():
    global _ai
    if _ai is None:
        _ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    return _ai


def analyze_thread(emails: list, efforts_folders: list, other_folders: list, reply_context: str = "") -> dict:
    emails = sorted(emails, key=lambda e: e.get("received_date_time", ""))
    participants = list(dict.fromkeys(
        (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
        for e in emails
        if (e.get("from_name") or e.get("from_address"))
    ))[:8]

    context = emails[-8:]
    msgs_text = "\n\n".join(
        f"From: {_clean(e.get('from_name') or e.get('from_address','Unknown'), 50)} "
        f"| {(e.get('received_date_time',''))[:10]}\n"
        f"{_clean(e.get('body_preview','(no preview)'), 2000)}"
        for e in context
    )

    if efforts_folders:
        efforts_list = ", ".join(efforts_folders[:15])
        folder_guidance = (
            f"\nEFFORTS SUBFOLDERS (pick one of these for filing — use the exact name): {efforts_list}"
        )
        if other_folders:
            folder_guidance += f"\nOther folders: {', '.join(other_folders[:6])}"
    else:
        folder_guidance = f"\nAvailable folders: {', '.join(other_folders[:10])}" if other_folders else ""

    subject = _clean(emails[-1].get("subject", "(no subject)"), 100)

    prompt = f"""You are a world-class executive communication assistant for a senior engineering/product leader at a large tech company.
Analyze this email thread and return ONLY valid JSON.

SUBJECT: {subject}
PARTICIPANTS: {', '.join(participants)}
TOTAL MESSAGES: {len(emails)}{folder_guidance}

MESSAGES (chronological, most recent last):
{msgs_text}

INSTRUCTIONS:

1. Determine the best action: reply | delete | file | read | done
   - reply: thread is waiting on the leader, a question was asked, action or decision required
   - delete: spam, automated notification with zero value, or social/marketing
   - file: substantive content worth keeping for reference
   - read: informational FYI, team-wide broadcast, no response needed
   - done: already fully resolved

2. Write "suggestedReply" — a complete, send-ready reply in FIRST PERSON.
   ALWAYS write a reply UNLESS action is "delete". Even for status updates and read-only threads,
   draft a reply that a leader would find useful to send.

   TONE: Direct, warm, confident. Like a senior leader who respects people's time and genuinely cares
   about the team. Never sycophantic, never hollow. No "I hope this finds you well", no "Thanks for sharing".

   TAILOR THE REPLY TO THE THREAD TYPE:

   A) STATUS UPDATES / PROGRESS REPORTS (action=read or file):
      — Acknowledge the specific work done. Name actual people. Reference specific metrics, milestones,
        or decisions mentioned in the thread.
      — Show you actually read it: reference a detail that proves it ("The latency drop from X to Y is great to see.")
      — Express genuine appreciation for the effort, not just the result.
      — If you spot anything worth a follow-up question or push, add it naturally.
      — Length: 2-4 sentences.

   B) QUESTIONS / REQUESTS WAITING ON YOU (action=reply):
      — Answer the question or fulfill the request directly and completely.
      — If you need more information first, ask exactly what you need — be specific.
      — State any decision you're making and the reason in one sentence.
      — Length: 3-5 sentences.

   C) VAGUE, UNCLEAR, OR MISSING-CONTEXT THREADS:
      — Do NOT pretend to understand. Push for the specific clarity needed.
      — Ask 1-3 sharp, targeted questions: What's the current state? What's the ask? What's the timeline?
        What decision needs to be made and by whom?
      — Be direct: "I want to engage on this but need a bit more context first…"
      — Length: 2-4 sentences.

   D) INCIDENT / OUTAGE / HIGH-URGENCY THREADS (urgency=high):
      — Acknowledge you're aware. State your immediate priority or what you're unblocking.
      — Offer your help or decision clearly: "I can free up [name] from X to help on this."
      — Ask the one most critical follow-up question if resolution is unclear.
      — Length: 2-4 sentences.

   E) CROSS-TEAM / PARTNERSHIP / EXTERNAL THREADS:
      — Professional but warm. Align on next steps. Name the right owner if it's not you.
      — Length: 3-5 sentences.

   BE SPECIFIC: Use actual names, numbers, and details from the thread. Generic replies are useless.
   Only use empty string "" for "suggestedReply" if action is "delete".

3. For "suggestedFolder": REQUIRED when action=file. Pick the single best name from the EFFORTS SUBFOLDERS
   list above using the exact name. Leave "" if action is not file.

Return ONLY this JSON (no markdown fences, no explanation):
{{
  "summary": "2-4 sentences that are DENSE WITH SPECIFICS. Name every person involved and their role/relationship. Quote or closely paraphrase the key ask, decision, or status. Include concrete details: numbers, dates, system names, project names, decisions made, blockers, next steps. If there is an open action item or question directed at the reader, state it explicitly. NO vague generalities — if you say 'progress was shared' instead of the actual progress, that is wrong.",
  "topic": "broad category label (e.g. Engineering, Product Planning, Finance, Incidents & Outages, Team & HR, Partnerships, FYI & Updates, Strategy & Leadership)",
  "action": "reply OR delete OR file OR read OR done",
  "urgency": "high OR medium OR low",
  "suggestedReply": "complete draft reply or empty string only if deleting",
  "suggestedFolder": "exact folder name or empty string"
}}"""

    if reply_context:
        prompt += f"\n\nNOTE: The user has provided the following context/notes for the reply. Incorporate this into your suggestedReply:\n{reply_context}"

    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r'^```[a-z]*\n?', '', raw)
        raw = re.sub(r'\n?```$', '', raw.strip())
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError(f"No JSON found: {raw[:200]}")
        result = json.loads(m.group())
        result["topic"] = _normalize_topic(result.get("topic", ""))
        return result
    except Exception as ex:
        print(f"  Analysis error: {ex}")
        return {
            "summary": f"Could not analyze thread: {ex}",
            "topic": "FYI & Updates",
            "action": "read",
            "urgency": "low",
            "suggestedReply": "",
            "suggestedFolder": "",
        }


def _format_prompt(body: str, from_name: str, date: str) -> str:
    return f"""You are an expert email analyst helping a senior tech leader understand an email.

FROM: {_clean(from_name, 80)}  |  DATE: {date}
EMAIL BODY:
{_clean(body, 8000)}

Break this email into its natural paragraphs. For each paragraph:
1. Provide the exact paragraph text (verbatim)
2. Classify the intent from EXACTLY one of: Status Update | Request | Decision | Question | Action Item | Context | FYI | Warning | Introduction | Closing
3. Choose an appropriate emoji for that intent
4. Fact-check: if the paragraph makes a specific claim that seems incorrect or worth verifying, provide a short concern string (1-2 sentences). Otherwise use null.

Return ONLY valid JSON (no markdown fences):
{{"paragraphs":[{{"text":"...","intent":"...","emoji":"...","fact_concern":null}}]}}"""


def _parse_format_response(raw: str, body: str) -> list:
    raw = re.sub(r'^```[a-z]*\n?', '', raw.strip())
    raw = re.sub(r'\n?```$', '', raw.strip())
    m = re.search(r'\{[\s\S]*\}', raw)
    if not m:
        raise ValueError(f"No JSON: {raw[:100]}")
    result = json.loads(m.group())
    return result.get("paragraphs", [])


def _format_message_with_ai(msg: dict) -> list:
    """Format a single message into AI-annotated paragraphs with intent + fact-check."""
    body = msg.get("body") or msg.get("body_preview") or ""
    if not body.strip():
        return [{"text": "(no content)", "intent": "FYI", "emoji": "📭", "fact_concern": None}]
    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    date = (msg.get("received_date_time") or "")[:10]
    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=3000,
            messages=[{"role": "user", "content": _format_prompt(body, from_name, date)}],
        )
        return _parse_format_response(resp.content[0].text, body)
    except Exception as ex:
        print(f"  Format error: {ex}")
        paras = [p.strip() for p in body.split('\n\n') if p.strip()][:20]
        return [{"text": p, "intent": "FYI", "emoji": "📄", "fact_concern": None} for p in paras]


# ─── Sync Worker ───────────────────────────────────────────────────────────────

_sync_status = {
    "running": False,
    "lastSync": None,
    "lastError": None,
    "emailsAdded": 0,
    "threadsUpdated": 0,
    "phase": "",
    "progress": "",
    "done": 0,
    "total": 0,
}
_sync_lock = threading.Lock()


def _refresh_folders() -> tuple:
    try:
        fdrs = call_tool("outlook_mail_list_folders", {"top": 50})
        top_level = fdrs.get("folders", fdrs.get("value", [])) if isinstance(fdrs, dict) else []

        efforts_id = None
        for f in top_level:
            name = f.get("display_name") or f.get("displayName", "")
            if name.lower() == "efforts":
                efforts_id = f.get("id") or f.get("folderId")
                break

        efforts_subfolders = []
        if efforts_id:
            sub = call_tool("outlook_mail_list_folders", {"parent_folder_id": efforts_id, "top": 50})
            sub_list = sub.get("folders", sub.get("value", [])) if isinstance(sub, dict) else []
            efforts_subfolders = [
                f.get("display_name") or f.get("displayName", "")
                for f in sub_list
                if (f.get("display_name") or f.get("displayName", ""))
            ]

        skip = {"Drafts", "Sent Items", "Outbox", "Deleted Items", "Junk Email",
                "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled", "Inbox", ""}
        other = [
            f.get("display_name") or f.get("displayName", "")
            for f in top_level
            if (f.get("display_name") or f.get("displayName", "")) not in skip
            and (f.get("display_name") or f.get("displayName", "")).lower() != "efforts"
        ]

        meta_set("efforts_subfolders", json.dumps(efforts_subfolders))
        meta_set("other_folders", json.dumps(other))
        meta_set("folders_raw", json.dumps(top_level))
        print(f"  Folders: {len(efforts_subfolders)} Efforts subfolders, {len(other)} other")
        return efforts_subfolders, other
    except Exception as e:
        print(f"Warning: could not refresh folders: {e}")
        return json.loads(meta_get("efforts_subfolders", "[]")), json.loads(meta_get("other_folders", "[]"))


def _refresh_calendar():
    """Fetch upcoming calendar events and store them in the local DB. Stores a meta key 'next_meeting' with JSON for the soonest future event."""
    try:
        import time as _time
        tz_name = _time.tzname[0] if _time.tzname else "UTC"
        # Try to get a proper IANA timezone name
        try:
            import subprocess
            result = subprocess.run(["readlink", "/etc/localtime"], capture_output=True, text=True)
            if result.returncode == 0 and "zoneinfo/" in result.stdout:
                tz_name = result.stdout.strip().split("zoneinfo/")[-1]
        except Exception:
            pass

        now = datetime.now(timezone.utc)
        # Format without Z suffix as required by the tool
        start = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S")
        end = (now + timedelta(days=60)).strftime("%Y-%m-%dT%H:%M:%S")
        resp = None
        try:
            print(f"  Calling outlook_calendar_list_events: start={start}, end={end}, tz={tz_name}")
            resp = call_tool("outlook_calendar_list_events", {
                "start_datetime": start,
                "end_datetime": end,
                "time_zone": tz_name,
            })
            print(f"  Calendar resp type={type(resp).__name__}: {str(resp)[:300]}")
        except Exception as ex:
            print(f"  calendar tool failed: {ex}")
            resp = None
        events = []
        if isinstance(resp, dict):
            if "events" in resp:
                events = resp.get("events")
            elif "value" in resp and isinstance(resp.get("value"), list):
                events = resp.get("value")
            else:
                # maybe the resp itself is a single event or list
                if isinstance(resp.get("items"), list):
                    events = resp.get("items")
        elif isinstance(resp, list):
            events = resp

        db = get_db()
        now_iso = _utcnow()
        def _ev_get(ev, *keys):
            """Case-insensitive multi-key lookup for event dicts."""
            for k in keys:
                v = ev.get(k) or ev.get(k.lower()) or ev.get(k[0].upper() + k[1:]) or ev.get(k.upper())
                if v:
                    return v
            return None

        def _get_dt(o):
            if isinstance(o, dict):
                return (_ev_get(o, "dateTime", "DateTime") or
                        _ev_get(o, "date", "Date") or
                        _ev_get(o, "date_time") or '')
            return str(o or '')

        for ev in events:
            ev_id = _ev_get(ev, "id", "Id", "eventId", "uid") or ''
            subj = _ev_get(ev, "subject", "Subject", "title", "Title") or ''
            start_obj = _ev_get(ev, "start", "Start", "startDateTime", "start_time") or {}
            end_obj = _ev_get(ev, "end", "End", "endDateTime", "end_time") or {}
            start_time = _get_dt(start_obj)
            end_time = _get_dt(end_obj)
            loc_raw = _ev_get(ev, "location", "Location", "place")
            location = ''
            if isinstance(loc_raw, dict):
                location = (_ev_get(loc_raw, "displayName", "DisplayName") or
                            _ev_get(loc_raw, "address", "Address") or '')
            else:
                location = str(loc_raw or '')
            attendees = []
            for a in (_ev_get(ev, "attendees", "Attendees", "participants") or []):
                if isinstance(a, dict):
                    name = (_ev_get(a, "displayName", "DisplayName") or
                            _ev_get(a, "name", "Name", "email", "EmailAddress", "address") or '')
                    attendees.append(name)
                else:
                    attendees.append(str(a))
            db.execute(
                "INSERT OR REPLACE INTO calendar_events (id,subject,start_time,end_time,location,attendees,raw_json,synced_at) VALUES(?,?,?,?,?,?,?,?)",
                (ev_id, subj, start_time, end_time, location, json.dumps(attendees), json.dumps(ev), now_iso)
            )
        db.commit()

        # Compute next meeting (compare in local time since MCP returns local-tz datetimes)
        now_local = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        row = db.execute("SELECT id,subject,start_time,end_time,location,attendees FROM calendar_events WHERE start_time > ? ORDER BY start_time ASC LIMIT 1", (now_local,)).fetchone()
        if row:
            nm = {"id": row[0], "subject": row[1], "start_time": row[2], "end_time": row[3], "location": row[4], "attendees": json.loads(row[5] or "[]")}
            meta_set("next_meeting", json.dumps(nm))
            return nm
        else:
            meta_set("next_meeting", json.dumps({}))
            return None
    except Exception as e:
        print(f"Warning: could not refresh calendar: {e}")
        try:
            return json.loads(meta_get("next_meeting", "{}") or "{}")
        except Exception:
            return None


def _insert_messages(db, emails: list, folder: str) -> int:
    """INSERT OR IGNORE a list of raw message dicts into the emails table. Returns count added."""
    now = _utcnow()
    added = 0
    for e in emails:
        if not e.get("id"):
            continue
        cur = db.execute(
            "INSERT OR IGNORE INTO emails "
            "(id,subject,from_name,from_address,received_date_time,"
            " is_read,body_preview,conversation_key,raw_json,synced_at,folder) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (
                e["id"],
                e.get("subject", ""),
                e.get("from_name", ""),
                e.get("from_address", ""),
                e.get("received_date_time", ""),
                1 if e.get("is_read") else 0,
                _clean(e.get("body_preview", ""), 500),
                _norm_subject(e.get("subject", "")),
                json.dumps(e),
                now,
                folder,
            ),
        )
        if cur.rowcount:
            added += 1
    db.commit()
    return added


def _do_sync():
    _sync_status.update({"phase": "fetching", "progress": "Fetching folder list…", "done": 0, "total": 0})
    efforts, other = _refresh_folders()
    try:
        _sync_status["progress"] = "Refreshing calendar…"
        nm = _refresh_calendar()
        if nm:
            _sync_status["nextMeeting"] = nm
    except Exception as e:
        print(f"  Calendar refresh failed: {e}")

    db = get_db()

    # ── Phase 1: Inbox — store + AI analyze ────────────────────────────────────
    _sync_status["progress"] = "Fetching inbox…"
    inbox_emails = []
    inbox_before = None
    while True:
        _args = {"folder": "Inbox", "top": INBOX_FETCH}
        if inbox_before:
            _args["received_before"] = inbox_before
        _r = call_tool("outlook_mail_list_messages", _args)
        _page = _r.get("messages", []) if isinstance(_r, dict) else []
        if not _page:
            break
        inbox_emails.extend(_page)
        if len(_page) < INBOX_FETCH:
            break
        _oldest = min((e.get("received_date_time","") for e in _page if e.get("received_date_time")), default="")
        if not _oldest or _oldest == inbox_before:
            break
        inbox_before = _oldest
        _sync_status["progress"] = f"Fetching inbox… ({len(inbox_emails)} so far)"

    # Determine which inbox messages are new before inserting
    inbox_ids = [e["id"] for e in inbox_emails if e.get("id")]
    existing_inbox = set()
    if inbox_ids:
        placeholders = ",".join("?" * len(inbox_ids))
        existing_inbox = {
            r[0] for r in db.execute(
                f"SELECT id FROM emails WHERE id IN ({placeholders})", inbox_ids
            ).fetchall()
        }
    new_inbox = [e for e in inbox_emails if e.get("id") and e["id"] not in existing_inbox]

    if inbox_emails:
        _insert_messages(db, inbox_emails, "Inbox")

    # AI analyze new inbox threads
    threads_updated = 0
    if new_inbox:
        print(f"Sync: {len(new_inbox)} new inbox email(s)")
        affected_keys = list({_norm_subject(e.get("subject", "")) for e in new_inbox})
        total = len(affected_keys)
        _sync_status.update({"phase": "analyzing", "done": 0, "total": total})

        for idx, ck in enumerate(affected_keys):
            # Only consider inbox emails for thread analysis
            rows = db.execute(
                "SELECT * FROM emails WHERE conversation_key=? AND folder='Inbox' "
                "ORDER BY received_date_time ASC",
                (ck,)
            ).fetchall()
            if not rows:
                continue
            thread_emails = [dict(r) for r in rows]
            display_subj = _clean(thread_emails[-1].get("subject", ck), 55)
            _sync_status["progress"] = f"Analyzing {idx+1}/{total}: \"{display_subj}\""

            try:
                result = analyze_thread(thread_emails, efforts, other)
            except Exception as ex:
                print(f"  Failed to analyze {ck!r}: {ex}")
                _sync_status["done"] = idx + 1
                continue

            latest = thread_emails[-1]
            participants = list(dict.fromkeys(
                (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
                for e in thread_emails
                if (e.get("from_name") or e.get("from_address"))
            ))[:8]
            has_unread = any(not e.get("is_read") for e in thread_emails)

            db.execute(
                "INSERT OR REPLACE INTO threads "
                "(conversation_key,subject,topic,action,urgency,summary,"
                " suggested_reply,suggested_folder,participants,email_ids,"
                " latest_id,message_count,has_unread,latest_received,updated_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    ck,
                    latest["subject"],
                    result.get("topic", "General"),
                    result.get("action", "read"),
                    result.get("urgency", "low"),
                    result.get("summary", ""),
                    result.get("suggestedReply", ""),
                    result.get("suggestedFolder", ""),
                    json.dumps(participants),
                    json.dumps([e["id"] for e in thread_emails]),
                    latest["id"],
                    len(thread_emails),
                    1 if has_unread else 0,
                    latest["received_date_time"],
                    _utcnow(),
                ),
            )
            db.commit()
            threads_updated += 1
            _sync_status["done"] = idx + 1
            print(f"  ✓ {display_subj!r} → {result.get('action')} [{result.get('urgency')}]")

    # ── Phase 2: All other folders — store only, no AI ─────────────────────────
    top_level = json.loads(meta_get("folders_raw", "[]"))
    sync_folders = []
    for f in top_level:
        name = f.get("display_name") or f.get("displayName", "")
        if name not in SKIP_SYNC_FOLDERS:
            sync_folders.append(name)
    for sub in efforts:
        sync_folders.append(f"Efforts/{sub}")

    total_folders = len(sync_folders)
    _sync_status.update({"phase": "fetching", "done": 0, "total": total_folders})
    for fi, folder_name in enumerate(sync_folders):
        _sync_status["progress"] = f"Syncing {fi+1}/{total_folders}: {folder_name}…"
        try:
            # Paginate backwards using received_before until we get no new messages
            folder_total_added = 0
            before = None
            while True:
                args = {"folder": folder_name, "top": FOLDER_FETCH}
                if before:
                    args["received_before"] = before
                result = call_tool("outlook_mail_list_messages", args)
                folder_emails = result.get("messages", []) if isinstance(result, dict) else []
                if not folder_emails:
                    break
                added = _insert_messages(db, folder_emails, folder_name)
                folder_total_added += added
                # If we added nothing new this page, the rest will also be known — stop
                if added == 0:
                    break
                # If fewer results than requested, we've hit the end
                if len(folder_emails) < FOLDER_FETCH:
                    break
                # Advance cursor to oldest message in this batch
                oldest = min(e.get("received_date_time", "") for e in folder_emails if e.get("received_date_time"))
                if not oldest or oldest == before:
                    break
                before = oldest
            if folder_total_added:
                print(f"  {folder_name}: {folder_total_added} new message(s)")
        except Exception as ex:
            print(f"  Warning: could not sync folder '{folder_name}': {ex}")
        _sync_status["done"] = fi + 1

    _sync_status.update({"phase": "done", "progress": f"Done — {threads_updated} thread(s) updated."})
    return len(new_inbox), threads_updated


def run_sync():
    if not _sync_lock.acquire(blocking=False):
        return
    _sync_status.update({"running": True, "lastError": None})
    try:
        added, updated = _do_sync()
        _sync_status.update({
            "running": False,
            "lastSync": _utcnow(),
            "emailsAdded": added,
            "threadsUpdated": updated,
        })
    except Exception as e:
        _sync_status.update({"running": False, "lastError": str(e)})
        print(f"Sync error: {e}")
    finally:
        _sync_lock.release()


def _sync_loop():
    _session_ready.wait(timeout=30)
    while True:
        run_sync()
        time.sleep(SYNC_INTERVAL)


# ─── Flask ─────────────────────────────────────────────────────────────────────

app = Flask(__name__)


@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/threads")
def api_threads():
    db = get_db()
    # Sort by recency — groups and threads within groups will both be newest-first
    rows = db.execute("""
        SELECT * FROM threads ORDER BY latest_received DESC
    """).fetchall()
    threads = [_thread_to_dict(r) for r in rows]

    # Build groups preserving insertion order (first occurrence of topic = most recent thread)
    groups: dict[str, list] = {}
    order: list[str] = []
    for t in threads:
        topic = t["topic"]
        if topic not in groups:
            groups[topic] = []
            order.append(topic)
        groups[topic].append(t)

    # order already reflects most-recent-thread-first per topic
    result = [{"topic": k, "threads": groups[k]} for k in order]
    latest_ts = max((t["updatedAt"] for t in threads), default="")

    try:
        nm = json.loads(meta_get("next_meeting", "{}") or "{}")
    except Exception:
        nm = {}
    return jsonify({
        "groups": result,
        "latestTs": latest_ts,
        "threadCount": len(threads),
        "emailCount": db.execute("SELECT COUNT(*) FROM emails").fetchone()[0],
        "syncStatus": {**_sync_status},
        "nextMeeting": nm,
    })


@app.route("/api/updates")
def api_updates():
    since = request.args.get("since", "")
    if not since:
        since = _utcnow()
    db = get_db()
    rows = db.execute(
        "SELECT * FROM threads WHERE updated_at > ? ORDER BY updated_at ASC",
        (since,)
    ).fetchall()
    threads = [_thread_to_dict(r) for r in rows]
    latest_ts = threads[-1]["updatedAt"] if threads else since
    try:
        nm = json.loads(meta_get("next_meeting", "{}") or "{}")
    except Exception:
        nm = {}
    return jsonify({
        "threads": threads,
        "latestTs": latest_ts,
        "syncStatus": {**_sync_status},
        "nextMeeting": nm,
    })


@app.route("/api/calendar")
def api_calendar():
    db = get_db()
    start_str = request.args.get("start")
    end_str = request.args.get("end")
    if not start_str:
        start_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%S")
    if not end_str:
        end_str = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S")
    rows = db.execute(
        "SELECT id,subject,start_time,end_time,location,attendees FROM calendar_events "
        "WHERE start_time >= ? AND start_time <= ? ORDER BY start_time ASC",
        (start_str, end_str)
    ).fetchall()
    events = [{"id": r[0], "subject": r[1], "start_time": r[2], "end_time": r[3],
               "location": r[4], "attendees": json.loads(r[5] or "[]")} for r in rows]
    return jsonify({"events": events})


@app.route("/api/meeting_prep", methods=["POST"])
def api_meeting_prep():
    data = request.json or {}
    subject    = data.get("subject", "(No title)")
    attendees  = data.get("attendees", [])
    start_time = data.get("start_time", "")
    end_time   = data.get("end_time", "")
    location   = data.get("location", "")

    names = ", ".join(
        a.get("name") or a.get("email", "") for a in (attendees or [])[:12]
    ) or "not listed"

    try:
        st = datetime.fromisoformat(start_time) if start_time else None
        time_str = st.strftime("%A %b %d at %-I:%M %p") if st else ""
    except Exception:
        time_str = start_time

    prompt = (
        f"You are preparing a senior tech leader at Microsoft for an upcoming meeting.\n\n"
        f"Meeting: {subject}\n"
        f"Time: {time_str}\n"
        f"Location: {location or 'Not specified'}\n"
        f"Attendees: {names}\n\n"
        f"Provide:\n"
        f"1. A 1-2 sentence heads-up: what this meeting is likely about and what the leader should be ready for.\n"
        f"2. Exactly 3 concise, specific topics or questions worth raising or keeping in mind.\n\n"
        f'Respond ONLY with valid JSON: {{"headsup": "...", "topics": ["...", "...", "..."]}}'
    )
    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=400,
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.content[0].text.strip()
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            result = json.loads(m.group())
            return jsonify({"ok": True,
                            "headsup": result.get("headsup", ""),
                            "topics":  result.get("topics", [])})
    except Exception as e:
        print(f"meeting_prep error: {e}")
    return jsonify({"ok": False, "headsup": "", "topics": []})


_FOLDER_ICONS = {
    "Inbox": "📥", "Sent Items": "📤", "Archive": "🗄️",
    "Drafts": "📝", "Deleted Items": "🗑️", "Junk Email": "🚫",
}
_FOLDERS_SKIP_DISPLAY = {
    "Drafts", "Outbox", "Junk Email",
    "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled",
}


@app.route("/api/mailbox/folders")
def api_mailbox_folders():
    top_level = json.loads(meta_get("folders_raw", "[]"))
    efforts_subs = json.loads(meta_get("efforts_subfolders", "[]"))
    db = get_db()
    # Get message counts per folder
    counts = {r["folder"]: r["cnt"] for r in db.execute(
        "SELECT folder, COUNT(*) as cnt FROM emails GROUP BY folder"
    ).fetchall()}
    def subfolder_count(path):
        return counts.get(path, 0)
    folder_map = {}
    for f in top_level:
        name = f.get("display_name") or f.get("displayName", "")
        if not name or name in _FOLDERS_SKIP_DISPLAY:
            continue
        if name.lower() == "efforts":
            children = []
            for s in sorted(efforts_subs):
                path = f"Efforts/{s}"
                children.append({"name": s, "path": path, "icon": "📂", "count": subfolder_count(path)})
            folder_map["Efforts"] = {"name": "Efforts", "icon": "📁", "count": 0, "children": children}
        else:
            folder_map[name] = {"name": name, "icon": _FOLDER_ICONS.get(name, "📁"), "count": counts.get(name, 0)}

    # Fixed display order
    ORDER = ["Inbox", "Efforts", "Partners", "Deleted Items", "Sent Items"]
    folders = []
    for name in ORDER:
        if name in folder_map:
            folders.append(folder_map.pop(name))
    # Append any remaining folders not in ORDER
    folders.extend(folder_map.values())
    return jsonify({"folders": folders})


@app.route("/api/mailbox/folder")
def api_mailbox_folder():
    folder = request.args.get("folder", "").strip()
    if not folder:
        return jsonify({"threads": [], "total": 0, "folder": ""})
    offset = int(request.args.get("offset", 0))
    limit = 100
    db = get_db()
    # Latest email per conversation_key in this folder
    rows = db.execute("""
        SELECT e.id, e.subject, e.from_name, e.from_address,
               e.received_date_time, e.body_preview, e.is_read, e.conversation_key,
               g.cnt, g.unread
        FROM emails e
        JOIN (
            SELECT conversation_key,
                   MAX(received_date_time) AS latest,
                   COUNT(*) AS cnt,
                   SUM(CASE WHEN is_read=0 THEN 1 ELSE 0 END) AS unread
            FROM emails WHERE folder=?
            GROUP BY conversation_key
        ) g ON e.conversation_key = g.conversation_key
              AND e.received_date_time = g.latest
              AND e.folder = ?
        GROUP BY e.conversation_key
        ORDER BY e.received_date_time DESC
        LIMIT ? OFFSET ?
    """, (folder, folder, limit, offset)).fetchall()
    total = db.execute(
        "SELECT COUNT(DISTINCT conversation_key) FROM emails WHERE folder=?", (folder,)
    ).fetchone()[0]
    threads = [{
        "id": r["id"], "subject": r["subject"] or "(No subject)",
        "fromName": r["from_name"] or "", "fromAddress": r["from_address"] or "",
        "date": r["received_date_time"] or "", "preview": r["body_preview"] or "",
        "isRead": bool(r["is_read"]), "conversationKey": r["conversation_key"],
        "messageCount": r["cnt"], "unreadCount": r["unread"],
    } for r in rows]
    return jsonify({"threads": threads, "total": total, "folder": folder})


@app.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify({"results": [], "query": q, "count": 0})
    like = f"%{q}%"
    db = get_db()
    rows = db.execute("""
        SELECT id, subject, from_name, from_address, received_date_time,
               body_preview, folder, is_read, conversation_key
        FROM emails
        WHERE subject LIKE ? OR from_name LIKE ? OR from_address LIKE ? OR body_preview LIKE ?
        ORDER BY received_date_time DESC
        LIMIT 100
    """, (like, like, like, like)).fetchall()
    return jsonify({"results": [dict(r) for r in rows], "query": q, "count": len(rows)})


@app.route("/api/sync_now", methods=["POST"])
def api_sync_now():
    if not _sync_status["running"]:
        threading.Thread(target=run_sync, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})


@app.route("/api/resync_thread", methods=["POST"])
def api_resync_thread():
    conv_key = (request.json or {}).get("conversationKey", "")
    if not conv_key:
        return jsonify({"ok": False, "error": "conversationKey required"}), 400
    if _sync_status["running"]:
        return jsonify({"ok": False, "error": "Sync already running"}), 409

    def _do_resync():
        db = get_db()
        try:
            # ── Step 1: Collect known IDs from the thread record (authoritative source)
            thread_row = db.execute(
                "SELECT * FROM threads WHERE conversation_key=?", (conv_key,)
            ).fetchone()
            stored_ids = []
            if thread_row:
                try:
                    stored_ids = json.loads(dict(thread_row).get("email_ids") or "[]")
                except Exception:
                    pass
            # Also grab any IDs already in the emails table (may differ if data drifted)
            db_ids = [r[0] for r in db.execute(
                "SELECT id FROM emails WHERE conversation_key=?", (conv_key,)
            ).fetchall()]
            all_ids = list(dict.fromkeys(stored_ids + db_ids))  # deduplicated, stored_ids first

            if not all_ids:
                _sync_status.update({"running": False, "lastError": "No messages found for thread."})
                return

            total_steps = len(all_ids) * 2 + 1  # fetch + format + analyze
            _sync_status.update({"phase": "fetching", "done": 0, "total": total_steps,
                                  "progress": f"Blowing away cached data for {len(all_ids)} message(s)…"})

            # ── Step 2: Nuke all existing data for this thread
            db.execute("DELETE FROM emails WHERE conversation_key=?", (conv_key,))
            db.execute("DELETE FROM threads WHERE conversation_key=?", (conv_key,))
            db.commit()

            # ── Step 3: Re-fetch every message fresh from Outlook
            _sync_status["progress"] = f"Re-fetching {len(all_ids)} messages from Outlook…"
            fresh_msgs = []   # list of normalized msg dicts with full body
            now = _utcnow()
            for i, msg_id in enumerate(all_ids):
                _sync_status["done"] = i
                _sync_status["progress"] = f"Fetching message {i+1}/{len(all_ids)}…"
                try:
                    resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
                    if isinstance(resp, dict) and resp:
                        raw = resp["messages"][0] if isinstance(resp.get("messages"), list) and resp["messages"] else resp
                    else:
                        print(f"  Resync: empty response for {msg_id}, skipping")
                        continue
                    nm = _normalize_msg(raw)
                    fresh_msgs.append((msg_id, raw, nm))

                    # Re-insert with FULL quote-stripped body stored (no truncation limit)
                    ck = _norm_subject(raw.get("subject", "") or nm.get("subject", ""))
                    full_body = nm.get("body", "")
                    db.execute(
                        "INSERT OR REPLACE INTO emails "
                        "(id,subject,from_name,from_address,received_date_time,"
                        " is_read,body_preview,conversation_key,raw_json,synced_at,formatted_body) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,NULL)",
                        (
                            msg_id,
                            raw.get("subject", nm.get("subject", "")),
                            nm.get("from_name", ""),
                            nm.get("from_address", ""),
                            nm.get("received_date_time", ""),
                            1 if raw.get("is_read") else 0,
                            full_body,   # store full quote-stripped body, no char limit
                            conv_key,    # keep original conv_key so thread stays together
                            json.dumps(raw),
                            now,
                        )
                    )
                    db.commit()
                except Exception as ex:
                    print(f"  Resync: failed to re-fetch {msg_id}: {ex}")

            if not fresh_msgs:
                _sync_status.update({"running": False, "lastError": "Could not re-fetch any messages."})
                return

            # ── Step 4: Re-analyze the thread (now with full bodies in body_preview)
            thread_emails = [dict(r) for r in db.execute(
                "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)
            ).fetchall()]
            display_subj = _clean(thread_emails[-1].get("subject", conv_key), 55)
            _sync_status.update({"done": len(all_ids), "progress": f"Re-analyzing \"{display_subj}\"…"})

            efforts = json.loads(meta_get("efforts_subfolders", "[]"))
            other   = json.loads(meta_get("other_folders", "[]"))
            result  = analyze_thread(thread_emails, efforts, other)

            latest       = thread_emails[-1]
            participants = list(dict.fromkeys(
                (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
                for e in thread_emails if (e.get("from_name") or e.get("from_address"))
            ))[:8]
            fetched_ids  = [e["id"] for e in thread_emails]
            has_unread   = any(not e.get("is_read") for e in thread_emails)

            db.execute(
                "INSERT OR REPLACE INTO threads "
                "(conversation_key,subject,topic,action,urgency,summary,"
                " suggested_reply,suggested_folder,participants,email_ids,"
                " latest_id,message_count,has_unread,latest_received,updated_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    conv_key,
                    latest["subject"],
                    result.get("topic", "General"),
                    result.get("action", "read"),
                    result.get("urgency", "low"),
                    result.get("summary", ""),
                    result.get("suggestedReply", ""),
                    result.get("suggestedFolder", ""),
                    json.dumps(participants),
                    json.dumps(fetched_ids),
                    latest["id"],
                    len(thread_emails),
                    1 if has_unread else 0,
                    latest["received_date_time"],
                    _utcnow(),
                )
            )
            db.commit()

            # ── Step 5: Re-run AI formatting on every message (warm cache with clean content)
            _sync_status["progress"] = f"Re-formatting {len(fresh_msgs)} messages with AI…"
            for i, (msg_id, _raw, nm) in enumerate(fresh_msgs):
                _sync_status["done"] = len(all_ids) + 1 + i
                try:
                    paras = _format_message_with_ai(nm)
                    db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                               (json.dumps(paras), msg_id))
                    db.commit()
                except Exception as ex:
                    print(f"  Resync format error for {msg_id}: {ex}")

            _sync_status.update({
                "running": False, "threadsUpdated": 1, "lastSync": _utcnow(),
                "phase": "done",
                "progress": f"Thread fully resynced — {len(thread_emails)} message(s) rebuilt.",
            })
        except Exception as ex:
            _sync_status.update({"running": False, "lastError": str(ex)})
            print(f"Resync thread error: {ex}")

    def _run_resync():
        if not _sync_lock.acquire(blocking=False):
            return
        _sync_status.update({"running": True, "lastError": None})
        try:
            _do_resync()
        finally:
            _sync_lock.release()

    if not _sync_status["running"]:
        threading.Thread(target=_run_resync, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})


@app.route("/api/reanalyze_all", methods=["POST"])
def api_reanalyze_all():
    if _sync_status["running"]:
        return jsonify({"ok": False, "error": "Sync already running"})

    def _do_reanalyze():
        efforts = json.loads(meta_get("efforts_subfolders", "[]"))
        other   = json.loads(meta_get("other_folders", "[]"))
        db = get_db()
        keys = [r[0] for r in db.execute("SELECT conversation_key FROM threads").fetchall()]
        total = len(keys)
        _sync_status.update({"running": True, "phase": "analyzing", "done": 0, "total": total,
                              "progress": f"Re-analyzing {total} threads…"})
        updated = 0
        for idx, ck in enumerate(keys):
            rows = db.execute(
                "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (ck,)
            ).fetchall()
            if not rows:
                _sync_status["done"] = idx + 1
                continue
            thread_emails = [dict(r) for r in rows]
            display_subj = _clean(thread_emails[-1].get("subject", ck), 55)
            _sync_status["progress"] = f"Re-analyzing {idx+1}/{total}: \"{display_subj}\""
            try:
                result = analyze_thread(thread_emails, efforts, other)
                db.execute(
                    "UPDATE threads SET topic=?,action=?,urgency=?,summary=?,"
                    "suggested_reply=?,suggested_folder=?,updated_at=? WHERE conversation_key=?",
                    (_normalize_topic(result.get("topic","")), result.get("action","read"),
                     result.get("urgency","low"), result.get("summary",""),
                     result.get("suggestedReply",""), result.get("suggestedFolder",""),
                     _utcnow(), ck)
                )
                db.commit()
                updated += 1
            except Exception as ex:
                print(f"  Re-analyze error for {ck}: {ex}")
            _sync_status["done"] = idx + 1
        _sync_status.update({"running": False, "lastSync": _utcnow(), "threadsUpdated": updated,
                              "phase": "done", "progress": f"Re-analyzed {updated}/{total} threads."})

    threading.Thread(target=_do_reanalyze, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})


@app.route("/api/folders")
def api_folders():
    efforts = json.loads(meta_get("efforts_subfolders", "[]"))
    other   = json.loads(meta_get("other_folders", "[]"))
    return jsonify({"folders": efforts + other, "effortsFolders": efforts})


def _parse_recipients(raw) -> list:
    """Normalize a recipients list from various Outlook API shapes."""
    result = []
    for r in (raw or []):
        if not isinstance(r, dict):
            continue
        name = (r.get("name") or r.get("display_name") or r.get("displayName") or
                r.get("emailAddress", {}).get("name") or "")
        addr = (r.get("address") or r.get("email") or
                r.get("emailAddress", {}).get("address") or "")
        if addr:
            result.append({"name": name.strip(), "address": addr.strip()})
    return result


def _strip_quoted_html(html: str) -> str:
    """Remove quoted/forwarded message blocks from HTML before text extraction."""
    # Each of these markers signals the start of quoted/forwarded content in Outlook/Gmail.
    # Content above the marker is always the new text; everything at or after is history.
    markers = [
        'id="mail-editor-reference-message-container"',  # Outlook mobile
        'id="divRplyFwdMsg"',                            # Outlook desktop reply/forward
        'id="appendonsend"',                             # Outlook append-on-send
        'class="gmail_quote"',                           # Gmail
        'id="divTaggedContent"',                         # Another Outlook variant
    ]
    lower = html.lower()
    cut = len(html)
    for marker in markers:
        idx = lower.find(marker.lower())
        if 0 < idx < cut:
            # Walk back to the opening < of the tag containing this attribute
            tag_start = html.rfind('<', 0, idx)
            if tag_start != -1:
                cut = tag_start
    return html[:cut]


def _normalize_msg(m: dict) -> dict:
    from_name    = m.get("from_name") or ""
    from_address = m.get("from_address") or ""
    received     = m.get("received_date_time") or ""

    raw_html = m.get("body_content") or ""
    if raw_html:
        # Strip quoted/forwarded history before any other processing
        raw_html = _strip_quoted_html(raw_html)
        # Replace block-level elements with newlines to preserve paragraph structure
        raw_html = re.sub(r'<br\s*/?>', '\n', raw_html, flags=re.IGNORECASE)
        raw_html = re.sub(r'</?(?:div|p|tr|li|blockquote|hr)[^>]*>', '\n', raw_html, flags=re.IGNORECASE)
        # Strip all remaining tags
        body_text = re.sub(r'<[^>]+>', '', raw_html)
        body_text = re.sub(r'&nbsp;', ' ', body_text)
        body_text = re.sub(r'&#\d+;|&[a-z]+;', ' ', body_text)
        body_text = re.sub(r'[ \t]{2,}', ' ', body_text)
        body_text = re.sub(r'\n{3,}', '\n\n', body_text).strip()
        # Fallback text-level stripping for non-HTML quote patterns
        # (handles "From: name <email>\nDate: ..." style inline quoting)
        text_cut = re.search(
            r'\n[Ff]rom:\s.{3,120}\n\s*(?:[Ss]ent|[Dd]ate|[Tt]o|[Cc]c)\s*:',
            body_text
        )
        if text_cut:
            body_text = body_text[:text_cut.start()].strip()
    else:
        body_text = m.get("body_preview") or ""

    to_recips  = _parse_recipients(m.get("to_recipients") or m.get("toRecipients"))
    cc_recips  = _parse_recipients(m.get("cc_recipients") or m.get("ccRecipients"))

    # Sanitize raw HTML for safe iframe rendering
    body_html = ""
    raw_content = m.get("body_content") or ""
    if raw_content and (m.get("body_content_type", "").upper() == "HTML" or raw_content.lstrip().startswith("<")):
        h = raw_content
        h = re.sub(r'<script\b[^>]*>.*?</script>', '', h, flags=re.IGNORECASE | re.DOTALL)
        h = re.sub(r'<style\b[^>]*>.*?</style>', lambda mo: mo.group(), h, flags=re.IGNORECASE | re.DOTALL)  # keep styles
        h = re.sub(r'\s+on\w+="[^"]*"', '', h, flags=re.IGNORECASE)
        h = re.sub(r"\s+on\w+='[^']*'", '', h, flags=re.IGNORECASE)
        # Neutralize CID inline images — show placeholder
        h = re.sub(r'src=["\']cid:[^"\']*["\']',
                   'src="" alt="[inline image — not available]" style="display:inline-block;padding:4px 8px;background:#eee;color:#666;font-size:11px;border-radius:3px"',
                   h, flags=re.IGNORECASE)
        body_html = h

    return {
        "id":                 m.get("id", ""),
        "subject":            m.get("subject", ""),
        "from_name":          from_name,
        "from_address":       from_address,
        "received_date_time": received,
        "is_read":            m.get("is_read"),
        "body":               body_text,
        "body_html":          body_html,
        "to_recipients":      to_recips,
        "cc_recipients":      cc_recips,
    }


@app.route("/api/thread_messages")
def api_thread_messages():
    ids = request.args.getlist("id")
    conv_key = request.args.get("conversationKey")
    if not ids and conv_key:
        db = get_db()
        rows = db.execute(
            "SELECT id FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC",
            (conv_key,)
        ).fetchall()
        ids = [r["id"] for r in rows]
    if not ids:
        return jsonify({"messages": []})
    db = get_db()
    rows = db.execute(
        "SELECT * FROM emails WHERE id IN ({})".format(",".join("?" * len(ids))),
        ids
    ).fetchall()
    db_msgs = {r["id"]: dict(r) for r in rows}

    result = [_normalize_msg(db_msgs.get(msg_id, {"id": msg_id})) for msg_id in ids]
    result.sort(key=lambda m: m.get("received_date_time", ""), reverse=True)
    return jsonify({"messages": result})


@app.route("/api/format_message")
def api_format_message():
    msg_id = request.args.get("id", "")
    db = get_db()
    row = db.execute("SELECT * FROM emails WHERE id=?", (msg_id,)).fetchone()

    # Return cached formatted body if available
    if row and row["formatted_body"]:
        try:
            return jsonify({"paragraphs": json.loads(row["formatted_body"]), "cached": True})
        except Exception:
            pass

    fallback = dict(row) if row else {"id": msg_id}
    try:
        resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
        if isinstance(resp, dict) and "messages" in resp:
            raw = resp["messages"][0] if resp["messages"] else fallback
        else:
            raw = fallback
        msg = _normalize_msg(raw)
    except Exception:
        msg = _normalize_msg(fallback)
    paragraphs = _format_message_with_ai(msg)

    # Persist to DB so future opens are instant
    if row:
        try:
            db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                       (json.dumps(paragraphs), msg_id))
            db.commit()
        except Exception:
            pass

    return jsonify({"paragraphs": paragraphs})


@app.route("/api/format_message_stream")
def api_format_message_stream():
    from flask import Response, stream_with_context
    msg_id = request.args.get("id", "")
    db = get_db()
    row = db.execute("SELECT * FROM emails WHERE id=?", (msg_id,)).fetchone()

    # Serve from cache immediately as a single done event
    if row and row["formatted_body"]:
        try:
            paras = json.loads(row["formatted_body"])
            def _cached():
                yield f"data: {json.dumps({'type':'done','paragraphs':paras})}\n\n"
            return Response(stream_with_context(_cached()), mimetype="text/event-stream",
                            headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})
        except Exception:
            pass

    # Fetch full message body
    fallback = dict(row) if row else {"id": msg_id}
    try:
        resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
        if isinstance(resp, dict) and "messages" in resp:
            raw = resp["messages"][0] if resp["messages"] else fallback
        elif isinstance(resp, dict) and resp:
            raw = resp
        else:
            raw = fallback
        msg = _normalize_msg(raw)
    except Exception:
        msg = _normalize_msg(fallback)

    body = (msg.get("body") or msg.get("body_preview") or "").strip()
    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    date = (msg.get("received_date_time") or "")[:10]

    if not body:
        def _empty():
            paras = [{"text": "(no content)", "intent": "FYI", "emoji": "📭", "fact_concern": None}]
            yield f"data: {json.dumps({'type':'done','paragraphs':paras})}\n\n"
        return Response(stream_with_context(_empty()), mimetype="text/event-stream",
                        headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})

    prompt = _format_prompt(body, from_name, date)

    @stream_with_context
    def _stream():
        full_text = ""
        try:
            with _get_ai().messages.stream(
                model=ANALYSIS_MODEL,
                max_tokens=3000,
                messages=[{"role": "user", "content": prompt}],
            ) as stream:
                for chunk in stream.text_stream:
                    full_text += chunk
                    yield f"data: {json.dumps({'type':'token','text':chunk})}\n\n"
        except Exception as ex:
            print(f"  Stream format error: {ex}")
            paras = [{"text": p.strip(), "intent": "FYI", "emoji": "📄", "fact_concern": None}
                     for p in body.split('\n\n') if p.strip()][:20]
            yield f"data: {json.dumps({'type':'done','paragraphs':paras})}\n\n"
            return

        try:
            paras = _parse_format_response(full_text, body)
        except Exception:
            paras = [{"text": p.strip(), "intent": "FYI", "emoji": "📄", "fact_concern": None}
                     for p in body.split('\n\n') if p.strip()][:20]

        if row:
            try:
                db2 = get_db()
                db2.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                            (json.dumps(paras), msg_id))
                db2.commit()
            except Exception:
                pass

        yield f"data: {json.dumps({'type':'done','paragraphs':paras})}\n\n"

    return Response(_stream(), mimetype="text/event-stream",
                    headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})


@app.route("/api/generate_reply", methods=["POST"])
def api_generate_reply():
    conv_key    = request.json.get("conversationKey", "")
    user_prompt = request.json.get("userPrompt", "").strip()
    if not user_prompt:
        return jsonify({"error": "userPrompt required"}), 400

    db = get_db()
    thread_row = db.execute("SELECT * FROM threads WHERE conversation_key=?", (conv_key,)).fetchone()
    email_rows = db.execute(
        "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)
    ).fetchall()
    thread = dict(thread_row) if thread_row else {}
    emails = [dict(r) for r in email_rows]

    subject = thread.get("subject") or (emails[-1].get("subject", "") if emails else "")
    context_emails = emails[-6:]
    msgs_text = "\n\n".join(
        f"From: {_clean(e.get('from_name') or e.get('from_address','Unknown'),50)} | {(e.get('received_date_time',''))[:10]}\n"
        f"{_clean(e.get('body_preview','(no preview)'), 600)}"
        for e in context_emails
    )

    prompt = f"""You are helping a senior tech leader craft a professional email reply.

THREAD SUBJECT: {subject}
THREAD CONTEXT (oldest first, most recent last):
{msgs_text}

THE LEADER'S CORE MESSAGE (what they want to say — stay true to this):
"{user_prompt}"

Write a polished, professional reply that:
1. Leads with and stays grounded in the leader's core intent — this is non-negotiable
2. Uses specific names, decisions, and details from the thread to make it feel personal and grounded
3. Is warm but direct — no filler phrases, no corporate speak, no "I hope this finds you well"
4. Uses 1-3 emojis placed naturally (not forced) to add energy and approachability
5. Has clear paragraph breaks for readability
6. Ends with clear next steps, a question, or a crisp closing — whichever fits
7. Length: match the complexity. Simple acknowledgement = 2-3 sentences. Complex topic = 4-7 sentences.

Return ONLY the reply body text. No subject line, no "From:", no markdown fences."""

    try:
        resp = _get_ai().messages.create(
            model=REPLY_MODEL,
            max_tokens=1200,
            messages=[{"role": "user", "content": prompt}],
        )
        return jsonify({"reply": resp.content[0].text.strip()})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/suggested_reply", methods=["POST"])
def api_suggested_reply():
    payload = request.json or {}
    conv_key = payload.get("conversationKey", "")
    reply_context = payload.get("context", "")
    if not conv_key:
        return jsonify({"error": "conversationKey required"}), 400
    db = get_db()
    rows = db.execute("SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)).fetchall()
    if not rows:
        return jsonify({"error": "No messages found for thread."}), 404
    emails = [dict(r) for r in rows]
    efforts = json.loads(meta_get("efforts_subfolders", "[]"))
    other = json.loads(meta_get("other_folders", "[]"))
    try:
        result = analyze_thread(emails, efforts, other, reply_context=reply_context)
        latest = emails[-1]
        participants = list(dict.fromkeys(
            (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
            for e in emails if (e.get("from_name") or e.get("from_address"))
        ))[:8]
        email_ids = [e["id"] for e in emails]
        has_unread = any(not e.get("is_read") for e in emails)
        db.execute(
            "INSERT OR REPLACE INTO threads "
            "(conversation_key,subject,topic,action,urgency,summary,"
            " suggested_reply,suggested_folder,participants,email_ids,"
            " latest_id,message_count,has_unread,latest_received,updated_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                conv_key,
                latest.get("subject", ""),
                result.get("topic", "General"),
                result.get("action", "read"),
                result.get("urgency", "low"),
                result.get("summary", ""),
                result.get("suggestedReply", ""),
                result.get("suggestedFolder", ""),
                json.dumps(participants),
                json.dumps(email_ids),
                latest.get("id", ""),
                len(emails),
                1 if has_unread else 0,
                latest.get("received_date_time", ""),
                _utcnow(),
            )
        )
        db.commit()
        return jsonify({"reply": result.get("suggestedReply", "")})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/reply/<latest_id>", methods=["POST"])
def api_reply(latest_id):
    body     = request.json.get("body", "")
    conv_key = request.json.get("conversationKey", "")
    to_list  = request.json.get("to", [])   # list of email address strings
    cc_list  = request.json.get("cc", [])
    try:
        draft_args = {
            "source_message_id": latest_id,
            "operation": "ReplyAll",
            "bodyText": body,
        }
        if to_list:
            draft_args["to"] = to_list
        if cc_list:
            draft_args["cc"] = cc_list
        draft = call_tool("outlook_mail_draft_message", draft_args)
        draft_id = None
        if isinstance(draft, dict):
            draft_id = (draft.get("draft_id") or draft.get("id") or
                        (draft.get("widgetState") or {}).get("draftId"))
        if not draft_id:
            return jsonify({"error": f"Could not get draft ID: {draft}"}), 500
        call_tool("outlook_mail_send_message", {"draft_id": draft_id})
        if conv_key:
            remove_thread(conv_key)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/my_email")
def api_my_email():
    return jsonify({"email": get_my_email()})


@app.route("/api/people")
def api_people():
    q = request.args.get("q", "").lower().strip()
    my = get_my_email().lower()
    db = get_db()
    rows = db.execute(
        "SELECT from_name, from_address FROM emails WHERE from_address != '' ORDER BY from_name"
    ).fetchall()
    seen = {}
    for r in rows:
        addr = (r["from_address"] or "").strip().lower()
        name = (r["from_name"] or "").strip()
        if not addr or addr == my:
            continue
        if addr not in seen:
            seen[addr] = name or r["from_address"]
    result = [{"name": name, "address": addr} for addr, name in seen.items()]
    if q:
        result = [p for p in result if q in p["address"].lower() or q in (p["name"] or "").lower()]
    result.sort(key=lambda p: (p["name"] or p["address"]).lower())
    return jsonify({"people": result[:60]})


@app.route("/api/send_new", methods=["POST"])
def api_send_new():
    payload = request.json or {}
    to_list = payload.get("to", [])
    cc_list = payload.get("cc", [])
    subject = payload.get("subject", "")
    body = payload.get("body", "")
    try:
        draft_args = {
            "operation": "New",
            "subject": subject,
            "bodyText": body,
        }
        if to_list:
            draft_args["to"] = to_list
        if cc_list:
            draft_args["cc"] = cc_list
        draft = call_tool("outlook_mail_draft_message", draft_args)
        draft_id = None
        if isinstance(draft, dict):
            draft_id = (draft.get("draft_id") or draft.get("id") or
                        (draft.get("widgetState") or {}).get("draftId"))
        if not draft_id:
            return jsonify({"error": f"Could not get draft ID: {draft}"}), 500
        call_tool("outlook_mail_send_message", {"draft_id": draft_id})
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/delete", methods=["POST"])
def api_delete():
    ids = request.json.get("ids", [])
    conv_key = request.json.get("conversationKey", "")
    for msg_id in ids:
        try:
            print(f"  Moving to Deleted Items: {msg_id[:40]}…")
            call_tool("outlook_mail_move_message", {
                "message_id": msg_id,
                "destination_folder": "Deleted Items",
            })
            print(f"  ✓ Moved: {msg_id[:40]}")
        except Exception as e:
            # Any MCP error (404, unknown, stale ID) means the message is inaccessible —
            # treat as already gone and continue cleaning up locally.
            print(f"  ~ Skipping inaccessible message ({e}): {msg_id[:40]}")
    if conv_key:
        remove_thread(conv_key)
    return jsonify({"ok": True})


@app.route("/api/move", methods=["POST"])
def api_move():
    ids = request.json.get("ids", [])
    folder = request.json.get("folder", "")
    conv_key = request.json.get("conversationKey", "")
    errors = []
    for msg_id in ids:
        try:
            call_tool("outlook_mail_move_message", {"message_id": msg_id, "destination_folder": folder})
        except Exception as e:
            errors.append(str(e))
    if conv_key:
        remove_thread(conv_key)
    return jsonify({"ok": not errors})


@app.route("/api/markread", methods=["POST"])
@app.route("/api/mark_read", methods=["POST"])
def api_markread():
    payload = request.json or {}
    ids = payload.get("ids", [])
    conv_key = payload.get("conversationKey", "")
    read = payload.get("read", True)
    # If no ids provided, look up by conv_key
    if not ids and conv_key:
        db = get_db()
        rows = db.execute("SELECT id FROM emails WHERE conversation_key=?", (conv_key,)).fetchall()
        ids = [r["id"] for r in rows]
    if not ids:
        return jsonify({"ok": True})
    try:
        call_tool("outlook_mail_mark_read", {"message_ids": ids, "is_read": read})
        db = get_db()
        id_ph = ",".join("?" * len(ids))
        db.execute(f"UPDATE emails SET is_read=? WHERE id IN ({id_ph})", [1 if read else 0] + ids)
        db.execute("UPDATE threads SET has_unread=? WHERE conversation_key=?", (0 if read else 1, conv_key))
        db.commit()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/flag", methods=["POST"])
def api_flag():
    conv_key = request.json.get("conversationKey", "")
    flagged   = request.json.get("flagged", True)   # True = flag, False = unflag
    if not conv_key:
        return jsonify({"error": "missing conversationKey"}), 400
    db = get_db()
    db.execute("UPDATE threads SET is_flagged=? WHERE conversation_key=?",
               (1 if flagged else 0, conv_key))
    db.commit()
    return jsonify({"ok": True, "isFlagged": flagged})


# ─── HTML ──────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Email Triage</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;font-family:'Monaco','Menlo','Courier New',monospace;}
html,body{height:100%;background:#0a1628;color:#c9d1d9;overflow:hidden;}

/* Layout */
.app{display:flex;flex-direction:column;height:100vh;}
.header{display:flex;align-items:center;justify-content:space-between;padding:0 18px;height:50px;background:#0d2040;border-bottom:1px solid #1e3d6b;flex-shrink:0;gap:12px;}
.header-brand{display:flex;align-items:center;gap:8px;flex-shrink:0;}
.header-brand h1{font-size:14px;font-weight:700;color:#e6edf3;letter-spacing:-0.3px;}
.header-center{flex:1;display:flex;align-items:center;justify-content:center;}
.header-right{display:flex;align-items:center;gap:8px;flex-shrink:0;}
.body{display:flex;flex:1;overflow:hidden;}

/* Sidebar */
.sidebar{width:260px;flex-shrink:0;background:#0a1628;border-right:none;display:flex;flex-direction:column;overflow:hidden;}
.sidebar-scroll{flex:1;overflow-y:auto;display:flex;flex-direction:column;}
.resize-handle{width:5px;cursor:ew-resize;background:transparent;flex-shrink:0;transition:background .15s;}
.resize-handle:hover,.resize-handle.dragging{background:#58a6ff;}
.sidebar-hdr{padding:12px 14px 6px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.14em;color:#5ba4cf;}
.topic-group{}
.tg-header{display:flex;align-items:center;gap:6px;padding:6px 14px;cursor:pointer;color:#7d8fa3;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;border-top:1px solid #0d2040;user-select:none;}
.tg-header:hover{color:#a8bccc;background:rgba(255,255,255,.03);}
.tg-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.tg-count{font-size:10px;background:#0d2040;color:#5a7a9e;border-radius:8px;padding:1px 6px;flex-shrink:0;}
.tg-chevron{font-size:9px;color:#5ba4cf;transition:transform .2s;}
.topic-group.collapsed .tg-chevron{transform:rotate(-90deg);}
.topic-group.collapsed .tg-threads{display:none;}
.thread-item{display:flex;align-items:flex-start;padding:7px 12px 7px 14px;cursor:pointer;border-left:3px solid transparent;transition:all .12s;gap:7px;}
.thread-item:hover{background:#0d2040;}
.thread-item.active{background:#122545;border-left-color:#58a6ff;}
.thread-item.urg-high{border-left-color:#f85149!important;}
.thread-item.urg-high.active{background:#200f18!important;}
.ti-dot{width:5px;height:5px;border-radius:50%;background:#58a6ff;flex-shrink:0;margin-top:5px;}
.ti-dot-empty{width:5px;flex-shrink:0;}
.ti-body{flex:1;overflow:hidden;min-width:0;}
.ti-subj{font-size:11.5px;color:#8b949e;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.4;}
.thread-item.unread .ti-subj{color:#f0c000;font-weight:700;}
.thread-item.active .ti-subj{color:#e6edf3;font-weight:600;}
.thread-item.active.unread .ti-subj{color:#ffe066;font-weight:700;}
.tg-header.has-unread{color:#f0c000;}
.tg-header.has-unread:hover{color:#ffe066;}
.tg-header.has-flagged .tg-name::after{content:' 🚩';font-size:11px;}
.thread-item.flagged .ti-subj{color:#f85149;}
.thread-item.flagged .ti-subj::after{content:' 🚩';font-size:10px;}
.thread-item.active.flagged .ti-subj{color:#ff7b72;}
.btn-flag{background:rgba(248,81,73,.1);color:#f85149;border:1px solid rgba(248,81,73,.25);}
.btn-flag.flagged{background:rgba(248,81,73,.25);color:#ff7b72;border-color:rgba(248,81,73,.5);}
.ti-meta{font-size:10px;color:#5ba4cf;margin-top:2px;display:flex;gap:5px;}

/* Right pane */
.right-pane{flex:1;overflow:hidden;display:flex;flex-direction:column;background:#0a1628;border-left:1px solid #1e3d6b;}
#first-load{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;}
.empty-pane{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;color:#5ba4cf;}
.empty-pane .ep-icon{font-size:36px;opacity:.35;}
.empty-pane .ep-txt{font-size:12px;}

/* Thread detail */
.thread-detail{flex:1;display:flex;flex-direction:column;min-height:0;overflow:hidden;}
.thread-hdr{padding:11px 18px 9px;background:linear-gradient(160deg,#0d2040 0%,#122545 100%);border-bottom:1px solid #1e3d6b;flex-shrink:0;}
.th-top{display:flex;align-items:center;gap:7px;margin-bottom:5px;flex-wrap:wrap;}
.th-badges{display:none;}
.urg-pill{display:inline-flex;align-items:center;padding:2px 9px;border-radius:10px;font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;}
.urg-high{background:rgba(248,81,73,.15);color:#f85149;border:1px solid rgba(248,81,73,.3);}
.urg-medium{background:rgba(210,153,34,.15);color:#d29922;border:1px solid rgba(210,153,34,.3);}
.urg-low{background:rgba(63,185,80,.12);color:#3fb950;border:1px solid rgba(63,185,80,.25);}
.act-pill{display:inline-flex;align-items:center;padding:2px 9px;border-radius:10px;font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;background:#1a3252;color:#8b949e;border:1px solid #243f65;}
.act-reply{background:rgba(31,111,235,.15);color:#58a6ff;border-color:rgba(88,166,255,.25);}
.act-delete{background:rgba(248,81,73,.12);color:#f85149;border-color:rgba(248,81,73,.25);}
.act-file{background:rgba(63,185,80,.1);color:#3fb950;border-color:rgba(63,185,80,.2);}
.th-subject{font-size:14px;font-weight:700;color:#e6edf3;line-height:1.3;flex:1;letter-spacing:-.2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}
.th-date{font-size:11px;color:#5ba4cf;flex-shrink:0;white-space:nowrap;}
.th-participants{display:flex;align-items:center;gap:8px;margin-bottom:6px;}
.avatars{display:flex;}
.avatar{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;font-size:8px;font-weight:700;color:#fff;border:2px solid #0a1628;margin-right:-4px;flex-shrink:0;}
.th-names{font-size:11px;color:#8b949e;margin-left:8px;}
.th-msgcount{font-size:10px;color:#8b949e;background:#1a3252;border-radius:7px;padding:2px 7px;}
.th-summary{background:rgba(88,166,255,.05);border:1px solid rgba(88,166,255,.12);border-radius:7px;padding:7px 11px;font-size:11px;color:#8b949e;line-height:1.6;margin-bottom:7px;}
.th-summary-lbl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#58a6ff;margin-bottom:3px;}
.th-actions{display:flex;gap:6px;flex-wrap:wrap;}

/* Buttons */
.btn{display:inline-flex;align-items:center;gap:5px;padding:6px 13px;border-radius:7px;font-size:11.5px;font-weight:600;border:none;cursor:pointer;transition:all .15s;white-space:nowrap;font-family:inherit;}
.btn:hover{filter:brightness(1.12);transform:translateY(-1px);}
.btn:active{transform:scale(.96) translateY(0);}
.btn-reply{background:#1f6feb;color:#fff;}
.btn-file{background:rgba(63,185,80,.15);color:#3fb950;border:1px solid rgba(63,185,80,.3);}
.btn-delete{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.25);}
.btn-ghost{background:#1a3252;color:#8b949e;border:1px solid #243f65;}
.btn-sm{padding:5px 10px;font-size:11px;}

/* Messages section */
.msgs-section{flex:1;min-height:0;overflow-y:auto;padding:0 0 16px;}
.msgs-label{display:flex;align-items:center;gap:8px;padding:7px 18px 5px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:#5ba4cf;border-bottom:1px solid #21262d;margin-bottom:6px;}
.msgs-label span{color:#58a6ff;}

/* Message cards */
.msg-card{background:#0d2040;border-left:none;border-right:none;border-top:none;border-bottom:1px solid #1a3252;margin-bottom:0;overflow:hidden;transition:border-color .15s;border-radius:0;}
.msg-card:hover{border-color:#2a4d7a;}
.msg-card.open{border-color:#2a4d7a;}
.msg-hdr{display:flex;align-items:center;gap:9px;padding:9px 13px;cursor:pointer;user-select:none;}
.msg-hdr:hover{background:rgba(255,255,255,.03);}
.msg-from-wrap{flex:1;display:flex;align-items:baseline;gap:7px;min-width:0;overflow:hidden;}
.msg-from{font-size:12px;font-weight:600;color:#c9d1d9;flex-shrink:0;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.msg-preview{font-size:11px;color:#06b6d4;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}
.msg-date{font-size:10.5px;color:#5ba4cf;flex-shrink:0;}
.msg-recips{padding:1px 13px 5px 46px;font-size:10px;color:#8b949e;display:flex;gap:10px;flex-wrap:wrap;line-height:1.6;}
.msg-recip-lbl{color:#556070;margin-right:2px;}
.msg-chevron{color:#5ba4cf;font-size:9px;flex-shrink:0;transition:transform .2s;}
.msg-card.open .msg-chevron{transform:rotate(180deg);}
.msg-body{display:none;border-top:1px solid #1a3252;padding:15px;}
.msg-card.open .msg-body{display:block;}

/* Paragraph intent blocks */
.para-blk{margin-bottom:13px;}
.intent-pill{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:9px;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;margin-bottom:5px;}
.i-status-update{background:rgba(88,166,255,.12);color:#58a6ff;border:1px solid rgba(88,166,255,.2);}
.i-request{background:rgba(219,109,40,.12);color:#d18616;border:1px solid rgba(209,134,22,.22);}
.i-decision{background:rgba(188,140,255,.12);color:#bc8cff;border:1px solid rgba(188,140,255,.22);}
.i-question{background:rgba(210,153,34,.12);color:#d29922;border:1px solid rgba(210,153,34,.22);}
.i-action-item{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.22);}
.i-context{background:rgba(139,148,158,.1);color:#8b949e;border:1px solid rgba(139,148,158,.18);}
.i-fyi{background:rgba(56,189,248,.1);color:#38bdf8;border:1px solid rgba(56,189,248,.18);}
.i-warning{background:rgba(245,158,11,.12);color:#f0883e;border:1px solid rgba(240,136,62,.22);}
.i-introduction{background:rgba(99,102,241,.12);color:#818cf8;border:1px solid rgba(129,140,248,.22);}
.i-closing{background:rgba(63,185,80,.1);color:#3fb950;border:1px solid rgba(63,185,80,.18);}
.para-txt{font-size:12px;color:#c9d1d9;line-height:1.8;}
.para-txt .eml{color:#58a6ff;background:rgba(88,166,255,.08);border-radius:3px;padding:0 3px;}
.para-txt a.link{color:#79b8ff;text-decoration:underline;text-underline-offset:2px;word-break:break-all;}
.para-txt a.link:hover{color:#a5d0ff;}
.fact-warn{display:flex;align-items:flex-start;gap:6px;background:rgba(240,136,62,.07);border:1px solid rgba(240,136,62,.18);border-radius:6px;padding:6px 10px;margin-top:5px;font-size:11px;color:#f0883e;line-height:1.5;}
.msg-ai-loading{display:flex;align-items:center;gap:8px;color:#5ba4cf;font-size:11.5px;padding:6px 0;}
.stream-wrap{padding:10px 16px;}
.stream-para{font-size:12px;color:#c9d1d9;line-height:1.8;margin-bottom:10px;white-space:pre-wrap;animation:sFadeIn .25s ease;}
.stream-cursor{display:inline-block;width:2px;height:13px;background:#06b6d4;vertical-align:middle;margin-left:1px;animation:sBlink .7s step-end infinite;}
@keyframes sFadeIn{from{opacity:0;transform:translateY(3px)}to{opacity:1;transform:translateY(0)}}
@keyframes sBlink{50%{opacity:0}}

/* Sync dot (used in sidebar footer) */
.sync-dot{width:7px;height:7px;border-radius:50%;background:#3fb950;flex-shrink:0;}
.sync-dot.syncing{background:#58a6ff;animation:pdot 1s infinite;}
.sync-dot.error{background:#f85149;}
@keyframes pdot{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.5;transform:scale(.8);}}
.new-badge{display:inline-flex;align-items:center;background:#1f6feb;color:#fff;border-radius:9px;font-size:10px;font-weight:600;padding:1px 7px;animation:bpop .3s cubic-bezier(.34,1.56,.64,1);}
@keyframes bpop{from{transform:scale(0);opacity:0;}to{transform:scale(1);opacity:1;}}
.badge-sm{display:inline-flex;align-items:center;padding:2px 8px;border-radius:8px;font-size:10.5px;background:#1a3252;color:#8b949e;}

/* Loading */
.spinner{width:28px;height:28px;border:2px solid #21262d;border-top-color:#58a6ff;border-radius:50%;animation:spin .75s linear infinite;}
.spinner-sm{width:12px;height:12px;border-width:2px;display:inline-block;vertical-align:middle;margin-right:4px;flex-shrink:0;}
@keyframes spin{to{transform:rotate(360deg);}}
.load-txt{color:#8b949e;font-size:12px;}

/* Modals */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);backdrop-filter:blur(4px);z-index:100;align-items:center;justify-content:center;}
.modal-overlay.open{display:flex;animation:mfade .15s ease;}
@keyframes mfade{from{opacity:0;}to{opacity:1;}}
.modal{background:#0d2040;border:1px solid #243f65;border-radius:12px;padding:20px;width:100%;box-shadow:0 24px 64px rgba(0,0,0,.6);animation:min .2s cubic-bezier(.34,1.56,.64,1);}
@keyframes min{from{opacity:0;transform:scale(.94) translateY(8px);}to{opacity:1;transform:scale(1) translateY(0);}}
.modal-lg{max-width:600px;}
.modal-xl{max-width:720px;}
.modal-sm{max-width:400px;}
.modal h3{font-size:14px;font-weight:700;margin-bottom:4px;color:#e6edf3;}
.modal-sub{font-size:11.5px;color:#8b949e;margin-bottom:12px;}
textarea{width:100%;background:#0a1628;border:1px solid #243f65;border-radius:8px;padding:10px 12px;font-size:12px;line-height:1.65;resize:vertical;min-height:150px;outline:none;font-family:'Monaco','Menlo','Courier New',monospace;color:#c9d1d9;transition:border-color .15s;}
textarea:focus{border-color:#58a6ff;}
select{width:100%;background:#0a1628;border:1px solid #243f65;border-radius:8px;padding:9px 12px;font-size:12px;outline:none;color:#c9d1d9;cursor:pointer;font-family:'Monaco','Menlo','Courier New',monospace;}
select:focus{border-color:#58a6ff;}
.modal-footer{display:flex;justify-content:flex-end;gap:8px;margin-top:14px;}
.del-subj{font-size:13px;font-weight:600;color:#e6edf3;margin-bottom:6px;background:rgba(248,81,73,.09);padding:9px 12px;border-radius:7px;border:1px solid rgba(248,81,73,.22);}
.del-warn{font-size:11.5px;color:#8b949e;}
/* Recipient fields */
.recip-row{display:flex;align-items:flex-start;gap:8px;margin-bottom:6px;}
.recip-label{font-size:11px;font-weight:700;color:#5a7a9e;width:24px;flex-shrink:0;padding-top:8px;letter-spacing:.04em;}
.recip-field{display:flex;flex-wrap:wrap;gap:5px;align-items:center;padding:5px 8px;border:1px solid #243f65;border-radius:8px;background:#0a1628;min-height:36px;flex:1;cursor:default;}
.recip-tag{display:inline-flex;align-items:center;gap:3px;background:#1a3252;color:#c9d1d9;border:1px solid #2a4d7a;border-radius:5px;padding:2px 8px;font-size:11px;}
.recip-tag .rm{cursor:pointer;color:#5a7a9e;margin-left:2px;font-size:14px;line-height:1;font-weight:300;}
.recip-tag .rm:hover{color:#f85149;}
.recip-empty{font-size:11px;color:#5ba4cf;font-style:italic;padding:4px 2px;}
.reply-intent-hint{font-size:11px;color:#5a7a9e;margin-bottom:6px;}
.generating-overlay{display:flex;align-items:center;gap:8px;color:#58a6ff;font-size:12px;padding:8px 0;}
/* Compose */
.recip-input{background:none;border:none;outline:none;color:#c9d1d9;font-size:12px;font-family:inherit;min-width:160px;padding:4px 6px;width:100%;}
.compose-subject-input{flex:1;background:#0a1628;border:1px solid #243f65;border-radius:8px;padding:8px 12px;font-size:13px;color:#c9d1d9;outline:none;font-family:inherit;width:100%;box-sizing:border-box;}
.compose-subject-input:focus{border-color:#58a6ff;}
.people-dropdown{position:absolute;top:100%;left:0;right:0;background:#0d2040;border:1px solid #2a4d7a;border-radius:8px;z-index:9999;max-height:220px;overflow-y:auto;display:none;box-shadow:0 8px 24px rgba(0,0,0,.5);}
.people-dropdown.open{display:block;}
.pd-item{padding:8px 12px;cursor:pointer;font-size:12px;display:flex;flex-direction:column;gap:1px;}
.pd-item:hover,.pd-item.active{background:#1a3252;}
.pd-item-name{color:#c9d1d9;font-weight:600;}
.pd-item-addr{color:#5ba4cf;font-size:11px;}

/* Inline reply */
.th-inline-reply{margin-top:12px;background:#080f1e;border:1px solid #1a3252;border-radius:8px;padding:11px 13px;}
.th-ir-lbl{font-size:10px;font-weight:600;color:#5ba4cf;text-transform:uppercase;letter-spacing:.6px;margin-bottom:6px;}
.th-ir-ta{width:100%;background:#0d2040;border:1px solid #1a3252;border-radius:6px;color:#c9d1d9;font-size:12px;line-height:1.6;padding:8px 10px;resize:vertical;min-height:72px;box-sizing:border-box;font-family:inherit;}
.th-ir-ta:focus{outline:none;border-color:#2a4d7a;}
.th-ir-actions{display:flex;gap:7px;margin-top:7px;}

/* Triage sheet */
.triage-pane{padding:20px 24px;overflow-y:auto;display:flex;flex-direction:column;height:100%;box-sizing:border-box;}
.triage-hdr{display:flex;align-items:center;gap:10px;margin-bottom:16px;flex-shrink:0;}
.triage-title{font-size:16px;font-weight:700;color:#e6edf3;flex:1;}
.triage-queue-count{font-size:11.5px;color:#8b949e;}
.triage-kb-hint{font-size:10px;color:#5ba4cf;margin-left:auto;}
.triage-rows{flex:1;overflow-y:auto;}
/* Topic groups */
.triage-topic-group{margin-bottom:12px;}
.triage-topic-hdr{display:flex;align-items:center;gap:8px;padding:5px 10px;cursor:pointer;border-radius:7px;user-select:none;margin-bottom:4px;outline:none;}
.triage-topic-hdr:hover{background:#0d2040;}
.triage-topic-chevron{font-size:10px;color:#5ba4cf;transition:transform .15s;display:inline-block;width:10px;}
.triage-topic-hdr.open .triage-topic-chevron{transform:rotate(90deg);}
.triage-topic-label{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#e3b429;}
.triage-topic-badge{background:#1a3252;color:#8b949e;font-size:10px;font-weight:700;padding:1px 7px;border-radius:8px;}
.triage-topic-rows{display:flex;flex-direction:column;gap:6px;}
/* Rows */
.triage-row{background:#0d2040;border:1px solid #1a3252;border-radius:10px;transition:border-color .15s;outline:none;}
.triage-row.ts-delete{border-color:#9a1c1c;}
.triage-row.ts-file{border-color:#b45309;}
.triage-row.ts-done{border-color:#1a7f37;opacity:.5;}
.triage-row-summary{padding:11px 14px 0;cursor:pointer;}
.triage-row-top{display:flex;align-items:flex-start;gap:8px;margin-bottom:5px;}
.triage-row-expand-chevron{font-size:9px;color:#5ba4cf;margin-top:3px;flex-shrink:0;width:9px;transition:transform .12s;}
.triage-row.expanded .triage-row-expand-chevron{transform:rotate(90deg);}
.triage-subj{font-size:12px;font-weight:600;color:#e6edf3;flex:1;line-height:1.35;}
.triage-sum{font-size:11px;color:#8b949e;line-height:1.5;padding-bottom:8px;}
/* Action rec pill */
.action-rec{display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap;flex-shrink:0;}
.action-rec-reply{background:rgba(31,111,235,.18);color:#58a6ff;border:1px solid rgba(88,166,255,.3);}
.action-rec-delete{background:rgba(248,81,73,.12);color:#f85149;border:1px solid rgba(248,81,73,.25);}
.action-rec-file{background:rgba(180,83,9,.15);color:#e08020;border:1px solid rgba(180,83,9,.3);}
.action-rec-read{background:rgba(72,79,88,.2);color:#8b949e;border:1px solid #2a5a8a;}
.action-rec-done{background:rgba(26,127,55,.12);color:#3fb950;border:1px solid rgba(63,185,80,.25);}
/* Inline messages */
.triage-msgs{border-top:1px solid #162030;margin:0;background:#080f1e;}
.triage-msg-row{display:flex;align-items:center;gap:8px;padding:5px 14px;cursor:pointer;border-bottom:1px solid #0d2040;user-select:none;}
.triage-msg-row:last-child{border-bottom:none;}
.triage-msg-row:hover{background:#0d2040;}
.triage-msg-chev{font-size:8px;color:#5ba4cf;width:8px;flex-shrink:0;transition:transform .1s;}
.triage-msg-row.open .triage-msg-chev{transform:rotate(90deg);}
.triage-msg-from{font-size:10.5px;font-weight:600;color:#c9d1d9;width:110px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.triage-msg-prev{font-size:10.5px;color:#5ba4cf;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.triage-msg-date{font-size:10px;color:#5ba4cf;flex-shrink:0;}
.triage-msg-body{padding:6px 14px 8px 30px;font-size:11px;color:#8b949e;line-height:1.6;white-space:pre-wrap;border-bottom:1px solid #0d2040;}
.triage-msg-body:last-child{border-bottom:none;}
/* Actions bar */
.triage-btns{display:flex;gap:5px;padding:8px 14px;align-items:center;flex-wrap:wrap;border-top:1px solid #162030;}
.triage-qlbl{margin-left:auto;font-size:10px;font-weight:600;}
.ts-delete .triage-qlbl{color:#f85149;}
.ts-file .triage-qlbl{color:#b45309;}
.btn-ts-del{color:#f85149;border-color:#f8514966;}
.btn-ts-del:hover,.btn-ts-del.active{background:#9a1c1c;color:#fff;border-color:#9a1c1c;}
.btn-ts-file{color:#b45309;border-color:#b4530966;}
.btn-ts-file:hover,.btn-ts-file.active{background:#b45309;color:#fff;border-color:#b45309;}
/* Keyboard focus */
.triage-kb-focus{outline:2px solid #58a6ff !important;outline-offset:-2px;}
.triage-sidebar-btn{display:block;width:calc(100% - 16px);margin:0 8px 8px;padding:7px 10px;background:#0d2040;border:1px solid #1a3252;border-radius:7px;color:#8b949e;font-size:12px;cursor:pointer;text-align:left;transition:all .15s;}
.triage-sidebar-btn:hover{background:#1a3252;color:#c9d1d9;}
.triage-sidebar-btn.active{background:#1a3252;color:#06b6d4;border-color:#06b6d4;}
/* Nav buttons */
.nav-btns{display:flex;gap:3px;}
.nav-btn{padding:5px 11px;border-radius:6px;font-size:11.5px;font-weight:600;border:1px solid transparent;background:none;color:#8b949e;cursor:pointer;transition:all .15s;white-space:nowrap;}
.nav-btn:hover{color:#c9d1d9;background:#1a3252;}
.nav-btn.active{background:#1a3252;color:#58a6ff;border-color:#58a6ff55;}
/* Search */
.search-wrap{flex:1;max-width:380px;position:relative;margin:0 12px;}
.search-input{width:100%;box-sizing:border-box;background:#0d2040;border:1px solid #1a3252;color:#e6edf3;border-radius:7px;padding:6px 12px 6px 30px;font-size:12px;outline:none;font-family:inherit;transition:border-color .15s;}
.search-input:focus{border-color:#58a6ff66;background:#111d33;}
.search-input::placeholder{color:#5ba4cf;}
.search-icon{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:#5ba4cf;font-size:13px;pointer-events:none;}
/* Search pane */
.search-pane{flex:1;overflow-y:auto;padding:16px 20px;}
.search-hdr{font-size:11.5px;color:#8b949e;margin-bottom:10px;padding-bottom:8px;border-bottom:1px solid #1a3252;}
.search-row{display:flex;align-items:flex-start;gap:10px;padding:9px 12px;border-radius:8px;border:1px solid #1a3252;background:#0d2040;margin-bottom:5px;cursor:pointer;transition:border-color .15s;}
.search-row:hover{border-color:#2a4d7a;}
.search-row-body{flex:1;min-width:0;}
.search-row-subj{font-size:12.5px;font-weight:600;color:#e6edf3;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.search-row-meta{font-size:10.5px;color:#8b949e;margin-top:2px;}
.search-row-preview{font-size:11px;color:#5ba4cf;margin-top:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.search-row-folder{font-size:10px;color:#5ba4cf;background:#1a3252;border-radius:4px;padding:1px 6px;flex-shrink:0;margin-top:2px;white-space:nowrap;}
/* Mailbox folder tree */
.folder-tree{padding:4px 0;}
.folder-item{display:flex;align-items:center;gap:7px;padding:5px 10px;border-radius:6px;margin:1px 6px;cursor:pointer;font-size:12px;color:#8b949e;transition:all .15s;}
.folder-item:hover{background:#0d2040;color:#c9d1d9;}
.folder-item.active{background:#1a3252;color:#58a6ff;}
.folder-item-name{flex:1;}
.folder-item-count{font-size:10px;color:#5ba4cf;font-weight:600;}
.folder-group-hdr{display:flex;align-items:center;gap:7px;padding:6px 10px;border-radius:6px;margin:1px 6px;cursor:pointer;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#7d8fa3;transition:all .15s;}
.folder-group-hdr:hover{background:#0d2040;color:#c9d1d9;}
.folder-group-chevron{font-size:9px;margin-left:auto;transition:transform .2s;}
.folder-group-hdr.open .folder-group-chevron{transform:rotate(180deg);}
.folder-group-children{padding-left:10px;display:none;}
.folder-group-children.open{display:block;}
/* Mailbox right pane */
.mailbox-pane{display:flex;flex-direction:column;flex:1;overflow:hidden;}
.mailbox-folder-hdr{padding:11px 18px;font-size:13px;font-weight:700;color:#e6edf3;border-bottom:1px solid #1a3252;flex-shrink:0;display:flex;align-items:baseline;gap:8px;}
.mailbox-folder-count{font-size:11px;color:#5ba4cf;font-weight:400;}
.mailbox-list{flex:1;overflow-y:auto;}
.mbox-row{display:flex;align-items:flex-start;gap:9px;padding:9px 16px;border-bottom:1px solid #12253f;cursor:pointer;transition:background .1s;}
.mbox-row:hover{background:#0d2040;}
.mbox-row.active{background:#112240;}
.mbox-row.focused{background:#0f2848;box-shadow:inset 3px 0 0 #58a6ff;}
.mbox-row.focused.active{background:#122a4a;box-shadow:inset 3px 0 0 #58a6ff;}
.mbox-row.focused .mbox-actions{display:flex;}
.mbox-kb-hint{padding:5px 16px 6px;font-size:10px;color:#3d5a7a;border-top:1px solid #0e1d30;flex-shrink:0;letter-spacing:.02em;display:flex;gap:12px;flex-wrap:wrap;}
.mbox-kb-hint kbd{background:#0d1f35;border:1px solid #1a3252;border-radius:3px;padding:0px 4px;font-size:9px;color:#5ba4cf;font-family:inherit;margin-right:2px;}
.mbox-dot{width:7px;height:7px;border-radius:50%;background:#58a6ff;flex-shrink:0;margin-top:5px;}
.mbox-dot-empty{width:7px;height:7px;flex-shrink:0;}
.mbox-body{flex:1;min-width:0;}
.mbox-subj{font-size:12.5px;color:#c9d1d9;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.mbox-row.unread .mbox-subj{color:#e6edf3;font-weight:600;}
.mbox-meta{display:flex;justify-content:space-between;margin-top:2px;}
.mbox-from{font-size:11px;color:#8b949e;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;}
.mbox-date{font-size:10.5px;color:#5ba4cf;flex-shrink:0;margin-left:8px;}
.mbox-preview{font-size:11px;color:#5ba4cf;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-top:1px;}
.mbox-cnt{font-size:10px;color:#5ba4cf;background:#1a3252;border-radius:10px;padding:1px 6px;flex-shrink:0;}
.mailbox-empty{display:flex;align-items:center;justify-content:center;flex:1;color:#5ba4cf;font-size:13px;padding:40px;}
.mbox-back{font-size:11px;color:#8b949e;background:#1a3252;border:1px solid #243f65;border-radius:6px;cursor:pointer;padding:3px 10px;transition:all .15s;}
.mbox-back:hover{color:#c9d1d9;background:#243f65;}
.mbox-actions{display:none;gap:4px;flex-shrink:0;align-items:center;margin-left:4px;}
.mbox-row:hover .mbox-actions{display:flex;}
.mbox-act-btn{padding:3px 7px;border-radius:5px;font-size:10px;font-weight:600;border:1px solid;cursor:pointer;line-height:1.4;background:transparent;white-space:nowrap;transition:all .12s;}
.mbox-act-reply{color:#58a6ff;border-color:#58a6ff55;}
.mbox-act-reply:hover{background:#1f6feb33;}
.mbox-act-del{color:#f85149;border-color:#f8514955;}
.mbox-act-del:hover{background:#9a1c1c44;}
/* Bottom status bar in sidebar */
.sidebar-footer{flex-shrink:0;border-top:1px solid #1a3252;padding:8px 12px;margin-top:auto;}
.sidebar-counts{font-size:10.5px;color:#5ba4cf;margin-bottom:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.sidebar-sync-row{display:flex;align-items:center;gap:6px;}
.sidebar-sync-txt{font-size:10.5px;color:#5ba4cf;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.sidebar-sync-bar-wrap{height:2px;background:#1a3252;border-radius:2px;overflow:hidden;display:none;margin-top:4px;}
.sidebar-sync-bar{height:100%;background:linear-gradient(90deg,#58a6ff,#bc8cff);border-radius:2px;transition:width .4s ease;}


/* Calendar pane */
.cal-pane{display:flex;flex-direction:column;flex:1;overflow:hidden;background:#0a1628;}
.cal-nav{display:flex;align-items:center;gap:10px;padding:12px 18px;border-bottom:1px solid #1a3252;flex-shrink:0;}
.cal-nav-title{flex:1;font-size:14px;font-weight:700;color:#e6edf3;}
.cal-nav-btn{background:#1a3252;border:1px solid #243f65;color:#8b949e;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:12px;}
.cal-nav-btn:hover{color:#c9d1d9;background:#243f65;}
.cal-today-btn{background:#1f6feb22;border-color:#1f6feb55;color:#58a6ff;}
.cal-today-btn:hover{background:#1f6feb44;}
.cal-scroll-wrap{flex:1;overflow:auto;}
.cal-grid{display:grid;grid-template-columns:52px repeat(7,minmax(0,1fr));min-width:560px;background:#0a1628;}
.cal-hdr-cell{padding:8px 4px;text-align:center;font-size:11px;font-weight:700;color:#8b949e;border-bottom:1px solid #1a3252;border-right:1px solid #1a3252;background:#0a1628;position:sticky;top:0;z-index:2;min-width:0;}
.cal-hdr-cell.today{color:#58a6ff;}
.cal-hdr-corner{border-bottom:1px solid #1a3252;background:#0a1628;position:sticky;top:0;z-index:2;}
.cal-hdr-day{font-size:13px;font-weight:700;}
.cal-hdr-dow{font-size:10px;text-transform:uppercase;letter-spacing:.08em;opacity:.7;}
.cal-time-label{font-size:10px;color:#5ba4cf;text-align:right;padding:0 6px 0 0;line-height:1;padding-top:2px;border-right:1px solid #1a3252;flex-shrink:0;}
.cal-cell{border-right:1px solid #1a3252;border-bottom:1px solid #12253f;position:relative;min-width:0;overflow:visible;}
.cal-cell.today-col{background:rgba(88,166,255,.03);}
.cal-cell.hour-start{border-top:1px solid #1a3252;}
.cal-event{position:absolute;left:2px;right:2px;border-radius:4px;padding:2px 5px;font-size:10.5px;font-weight:600;overflow:hidden;cursor:pointer;z-index:1;line-height:1.35;transition:filter .15s;}
.cal-event:hover{filter:brightness(1.2);}
.cal-event-title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cal-event-time{font-size:9.5px;opacity:.8;white-space:nowrap;}
.cal-all-day-cell{border-bottom:2px solid #1a3252;border-right:1px solid #1a3252;padding:2px 2px;min-height:20px;min-width:0;overflow:hidden;}
.cal-all-day-event{background:rgba(88,166,255,.18);border:1px solid rgba(88,166,255,.35);color:#8bb8f8;border-radius:3px;padding:1px 5px;font-size:10px;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin-bottom:1px;}
.cal-loading{display:flex;align-items:center;justify-content:center;flex:1;color:#5ba4cf;font-size:13px;}
.cal-view-btn{min-width:38px;text-align:center;}
.cal-view-btn.active{background:#1f6feb33;border-color:#1f6feb77;color:#58a6ff;}
/* Day view event cards */
.cal-day-event{left:3px!important;right:3px!important;border-radius:6px!important;overflow:hidden!important;cursor:default!important;}
.cal-day-ev-hdr{margin-bottom:3px;}
.cal-day-ev-title{font-size:11.5px;font-weight:700;line-height:1.3;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.cal-day-ev-time{font-size:10px;opacity:.8;margin-bottom:4px;}
.cal-day-ev-prep{font-size:10.5px;line-height:1.5;overflow:hidden;border-top:1px solid rgba(255,255,255,.1);padding-top:5px;margin-top:2px;}
.cal-prep-loading{opacity:.5;font-size:10px;font-style:italic;}
.cal-prep-headsup{margin-bottom:5px;line-height:1.5;opacity:.9;}
.cal-prep-topics{display:flex;flex-direction:column;gap:2px;}
.cal-prep-topic{opacity:.8;font-size:10px;}
/* Today's schedule in sidebar */
.today-cal{border-top:1px solid #1a3252;margin-top:8px;padding:8px 0 4px;}
.today-cal-hdr{padding:4px 14px 4px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.14em;color:#5ba4cf;}
.week-hours-line{padding:0 14px 6px;font-size:10px;color:#60a5fa;font-weight:600;}
.today-cal-empty{padding:4px 14px;font-size:11px;color:#5ba4cf;font-style:italic;}
.today-ev{display:flex;align-items:center;gap:6px;padding:3px 10px;border-radius:6px;margin:0 6px 1px;cursor:default;}
.today-ev:hover{background:#0d2040;}
.today-ev-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0;}
.today-ev-time{font-size:10px;color:#60a5fa;flex-shrink:0;font-variant-numeric:tabular-nums;}
.today-ev-title{font-size:11px;color:#c9d1d9;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;}
</style>
</head>
<body>
<div class="app">
  <header class="header">
    <div class="header-brand">
      <h1>Email</h1>
    </div>
    <div class="header-center">
      <div class="search-wrap">
        <span class="search-icon">⌕</span>
        <input type="text" id="search-input" class="search-input" placeholder="Search mail…"
               onkeydown="if(event.key==='Enter')doSearch(this.value);else if(event.key==='Escape')clearSearch()">
      </div>
    </div>
    <div class="header-right">
      <div id="triage-actions" style="display:flex;gap:6px;align-items:center;">
        <button class="btn btn-reply btn-sm" onclick="openCompose()" style="font-weight:700">✉ New Message</button>
        <button class="btn btn-ghost btn-sm" onclick="triggerSync()">⟳ Sync Now</button>
        <button class="btn btn-ghost btn-sm" id="resync-thread-btn" onclick="resyncThread()" title="Re-fetch &amp; re-analyze the selected thread" disabled>↺ Resync Thread</button>
        <button class="btn btn-ghost btn-sm" onclick="reanalyzeAll()" id="reanalyze-btn">⚙ Re-analyze</button>
      </div>
    </div>
  </header>
  <div class="body">
    <nav class="sidebar" id="sidebar">
      <div class="sidebar-scroll">
        <div id="sidebar-nav">
          <button class="triage-sidebar-btn" id="triage-sidebar-btn" onclick="openTriageSheet()">📋 Triage Sheet</button>
          <div class="sidebar-hdr">Folders</div>
          <div id="folder-tree"></div>
          <button class="triage-sidebar-btn" id="nav-calendar" onclick="switchTab('calendar')" style="margin-top:8px">📅 Calendar</button>
          <div class="today-cal" id="today-cal">
            <div class="today-cal-hdr">Today</div>
            <div class="week-hours-line" id="week-hours-line"></div>
            <div id="today-cal-list"><div class="today-cal-empty">Loading…</div></div>
          </div>
        </div>
      </div>
      <div class="sidebar-footer">
        <div class="sidebar-counts" id="sidebar-counts"></div>
        <div class="sidebar-sync-row">
          <div class="sync-dot" id="sync-dot"></div>
          <span class="sidebar-sync-txt" id="sync-txt">Connecting…</span>
          <span id="new-badge-wrap"></span>
        </div>
        <div class="sidebar-sync-bar-wrap" id="sync-bar-wrap">
          <div class="sidebar-sync-bar" id="sync-bar" style="width:0%"></div>
        </div>
      </div>
    </nav>
    <div class="resize-handle" id="resize-handle"></div>
    <div class="right-pane" id="right-pane">
      <div id="first-load">
        <div class="spinner"></div>
        <div class="load-txt" id="load-msg">Loading inbox...</div>
      </div>
      <div class="empty-pane" id="empty-pane" style="display:none">
        <div class="ep-icon">✉</div>
        <div class="ep-txt">Select a thread to read</div>
      </div>
      <div class="thread-detail" id="thread-detail" style="display:none">
        <div class="thread-hdr" id="thread-hdr"></div>
        <div class="msgs-section" id="msgs-section"></div>
      </div>
      <div class="triage-pane" id="triage-pane" style="display:none"></div>
      <div class="mailbox-pane" id="mailbox-pane" style="display:none">
        <div class="mailbox-folder-hdr">
          <span id="mailbox-folder-name">Select a folder</span>
          <span class="mailbox-folder-count" id="mailbox-folder-count"></span>
        </div>
        <div class="mailbox-list" id="mailbox-list">
          <div class="mailbox-empty">Select a folder</div>
        </div>
        <div class="mbox-kb-hint">
          <span><kbd>j</kbd><kbd>k</kbd> navigate</span>
          <span><kbd>Enter</kbd> open</span>
          <span><kbd>r</kbd> reply</span>
          <span><kbd>d</kbd> delete</span>
          <span><kbd>f</kbd> file</span>
          <span><kbd>u</kbd> mark read</span>
          <span><kbd>Esc</kbd> back</span>
        </div>
      </div>
      <div class="search-pane" id="search-pane" style="display:none">
        <div class="search-hdr" id="search-hdr"></div>
        <div id="search-results"></div>
      </div>
      <div class="cal-pane" id="calendar-pane" style="display:none">
        <div class="cal-nav">
          <button class="cal-nav-btn" onclick="calMove(-1)">&#8249;</button>
          <div class="cal-nav-title" id="cal-title"></div>
          <button class="cal-nav-btn cal-today-btn" onclick="calGoToday()">Today</button>
          <button class="cal-nav-btn" onclick="calMove(1)">&#8250;</button>
          <div style="width:1px;background:#1a3252;height:18px;margin:0 2px;flex-shrink:0"></div>
          <button class="cal-nav-btn cal-view-btn active" id="cal-view-day" onclick="calSetView('day')">Day</button>
          <button class="cal-nav-btn cal-view-btn" id="cal-view-week" onclick="calSetView('week')">Week</button>
        </div>
        <div class="cal-loading" id="cal-loading">Loading calendar…</div>
        <div class="cal-scroll-wrap" id="cal-scroll-wrap" style="display:none">
          <div class="cal-grid" id="cal-grid"></div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- Reply modal -->
<div class="modal-overlay" id="reply-modal">
  <div class="modal modal-lg">
    <h3>↩ Reply All</h3>
    <div class="modal-sub" id="reply-sub"></div>
    <div class="recip-row">
      <span class="recip-label">TO</span>
      <div class="recip-field" id="reply-to-field"></div>
    </div>
    <div class="recip-row">
      <span class="recip-label">CC</span>
      <div class="recip-field" id="reply-cc-field"></div>
    </div>
    <div style="position:relative;margin-top:10px;">
      <textarea id="reply-body" style="min-height:260px;resize:vertical;width:100%;box-sizing:border-box;padding-right:100px;" placeholder="Generating reply…"></textarea>
      <div id="reply-generating" style="position:absolute;top:10px;right:10px;font-size:11px;color:#60a5fa;display:none"><div class="spinner spinner-sm" style="display:inline-block;margin-right:4px;vertical-align:middle"></div>Generating…</div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
      <button class="btn btn-ghost btn-sm" onclick="regenerateReply()" id="reply-regen-btn">↺ Regenerate</button>
      <button class="btn btn-reply" onclick="sendReply()">✉ Send Reply</button>
    </div>
  </div>
</div>

<!-- Compose modal -->
<div class="modal-overlay" id="compose-modal">
  <div class="modal modal-xl">
    <h3>✉ New Message</h3>
    <div class="recip-row" style="position:relative;">
      <span class="recip-label">TO</span>
      <div style="flex:1;position:relative;">
        <div class="recip-field" id="compose-to-field" style="cursor:text;" onclick="focusComposeInput('to')"></div>
        <input id="compose-to-input" class="recip-input" type="text" placeholder="Add recipient…" autocomplete="off"
               oninput="peopleSuggest(this,'to')" onkeydown="peopleSuggestKey(event,'to')">
        <div class="people-dropdown" id="pd-to"></div>
      </div>
    </div>
    <div class="recip-row" style="position:relative;">
      <span class="recip-label">CC</span>
      <div style="flex:1;position:relative;">
        <div class="recip-field" id="compose-cc-field" style="cursor:text;" onclick="focusComposeInput('cc')"></div>
        <input id="compose-cc-input" class="recip-input" type="text" placeholder="Add recipient…" autocomplete="off"
               oninput="peopleSuggest(this,'cc')" onkeydown="peopleSuggestKey(event,'cc')">
        <div class="people-dropdown" id="pd-cc"></div>
      </div>
    </div>
    <div class="recip-row">
      <span class="recip-label">Subject</span>
      <input id="compose-subject" type="text" class="compose-subject-input" placeholder="Subject…">
    </div>
    <textarea id="compose-body" style="margin-top:10px;min-height:340px;resize:vertical;width:100%;box-sizing:border-box;" placeholder="Write your message…"></textarea>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
      <button class="btn btn-reply" onclick="sendNewMessage()">✉ Send</button>
    </div>
  </div>
</div>

<!-- File modal -->
<div class="modal-overlay" id="file-modal">
  <div class="modal modal-sm">
    <h3>📁 File Thread</h3>
    <div class="modal-sub" id="file-sub"></div>
    <select id="folder-select"></select>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
      <button class="btn btn-file" onclick="fileThread()">File Thread</button>
    </div>
  </div>
</div>

<!-- Delete modal -->
<div class="modal-overlay" id="delete-modal">
  <div class="modal modal-sm">
    <h3>🗑 Delete Thread</h3>
    <div class="del-subj" id="del-subj"></div>
    <div class="del-warn">All <span id="del-count"></span> messages will be permanently deleted.</div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
      <button class="btn btn-delete" onclick="confirmDelete()">Delete</button>
    </div>
  </div>
</div>

<script>
// ── State ──────────────────────────────────────────────────────────────────────
let state = {
  groups: [],
  threadMap: {},
  selectedKey: null,
  collapsedTopics: new Set(),
  formatCache: {},        // msgId → [{text,intent,emoji,fact_concern}]
  expandedMsgs: new Set(),// indices of expanded messages
  currentMsgs: [],        // messages for selected thread (newest-first)
  latestTs: '',
  folders: [],
  effortsFolders: [],
  pollTimer: null,
  triageActions: {},      // map of conversationKey → {type: 'send'|'delete'|'file', reply: string}
  triageView: false,
  mailboxContext: false,  // true when thread opened from mailbox view
  collapsedTriageTopics: new Set(),
  triageFocusIdx: -1,
  expandedTriageRows: new Set(),
  triageMsgCache: {},
  showOriginal: {},       // msgId -> true to show original text instead of AI-formatted
};
let _activeThread = null;

// ── Avatar helpers ─────────────────────────────────────────────────────────────
const AV_COLORS = ['#1f6feb','#1a7f37','#9a1c1c','#7d4e00','#6e3cc1','#b45309','#0284c7','#be185d'];
function avColor(name) {
  let h = 0; for (const c of String(name)) h = (h*31+c.charCodeAt(0))&0xffffffff;
  return AV_COLORS[Math.abs(h)%AV_COLORS.length];
}
function initials(name) {
  const p = String(name||'').trim().split(/\s+/);
  return p.length>=2?(p[0][0]+p[1][0]).toUpperCase():(p[0]||'?').slice(0,2).toUpperCase();
}

// ── Init ───────────────────────────────────────────────────────────────────────
async function init() {
  const d = await fetch('/api/threads').then(r=>r.json()).catch(()=>null);
  if (!d) { setTimeout(init,3000); return; }
  if (d.groups.length===0 && d.syncStatus.running) {
    document.getElementById('load-msg').textContent='Syncing inbox for the first time…';
    updateSyncStatus(d.syncStatus);
    setTimeout(init,3000); return;
  }
  state.groups = d.groups;
  state.threadMap = {};
  state.groups.forEach(g=>g.threads.forEach(t=>{state.threadMap[t.conversationKey]=t;}));
  state.collapsedTopics = new Set(state.groups.map(g=>g.topic));
  state.latestTs = d.latestTs;
  document.getElementById('first-load').style.display='none';
  updateCounts(d.emailCount, Object.keys(state.threadMap).length);
  updateSyncStatus(d.syncStatus);

  renderSidebar();
  renderTodayCal();
  openTriageSheet();
  schedulePoll();

  // Load my email for reply-all filtering + preload people cache
  fetch('/api/my_email').then(r=>r.json()).then(d=>{ if(d.email) MY_EMAIL=d.email; }).catch(()=>{});
  fetch('/api/people').then(r=>r.json()).then(d=>{ if(d.people) _pdCache=d.people; }).catch(()=>{});
}

// ── Poll ───────────────────────────────────────────────────────────────────────
function schedulePoll() {
  clearTimeout(state.pollTimer);
  state.pollTimer = setTimeout(pollUpdates,10000);
}
async function pollUpdates() {
  const d = await fetch(`/api/updates?since=${encodeURIComponent(state.latestTs)}`).then(r=>r.json()).catch(()=>null);
  if (!d) { schedulePoll(); return; }
  updateSyncStatus(d.syncStatus);
  // Update next meeting display if present
  try {
    const nm = d.nextMeeting || {};
    const nmEl = document.getElementById('next-meeting');
    if (nm && nm.subject && nm.start_time) {
      nmEl.textContent = `Next: ${esc(nm.subject)} · ${fmtUntil(nm.start_time)}`;
    } else if (nmEl) { nmEl.textContent = ''; }
  } catch(e) {}
  if (d.threads && d.threads.length>0) {
    let newCount=0;
    for (const t of d.threads) {
      if (t.updatedAt>state.latestTs) state.latestTs=t.updatedAt;
      const isNew=!state.threadMap[t.conversationKey];
      state.threadMap[t.conversationKey]=t;
      if (isNew) { newCount++; _insertGroup(t); } else { _updateGroup(t); }
    }
    renderSidebar();
    updateCounts(null, Object.keys(state.threadMap).length);
    if (newCount>0) _showNewBadge(newCount);
    if (state.selectedKey && state.threadMap[state.selectedKey]) {
      _renderThreadHdr(state.threadMap[state.selectedKey]);
    }
  }
  schedulePoll();
}
function _insertGroup(t) {
  let g=state.groups.find(g=>g.topic===t.topic);
  if (!g){g={topic:t.topic,threads:[]};state.groups.unshift(g);}
  g.threads.unshift(t);
}
function _updateGroup(t) {
  for (const g of state.groups) {
    const i=g.threads.findIndex(x=>x.conversationKey===t.conversationKey);
    if (i>=0){if(g.topic!==t.topic){g.threads.splice(i,1);_insertGroup(t);}else{g.threads[i]=t;}return;}
  }
  _insertGroup(t);
}

// ── Sidebar ────────────────────────────────────────────────────────────────────
function renderSidebar() {
  // Topic list removed; sidebar now shows folder tree (via initMailbox).
  // Highlight active thread in folder tree if applicable.
}
function _threadItemHTML(t) {
  const active = t.conversationKey===state.selectedKey;
  const urgCls = t.urgency==='high'?' urg-high':'';
  const unreadCls = t.hasUnread?' unread':'';
  const flaggedCls = t.isFlagged?' flagged':'';
  return `<div class="thread-item${active?' active':''}${urgCls}${unreadCls}${flaggedCls}" data-key="${esc(t.conversationKey)}" onclick="selectThread(this.getAttribute('data-key'))">
    ${t.hasUnread?'<div class="ti-dot"></div>':'<div class="ti-dot-empty"></div>'}
    <div class="ti-body">
      <div class="ti-subj">${esc(t.subject||'(No subject)')}</div>
      <div class="ti-meta"><span>${esc((t.participants||[])[0]||'')}</span><span>${esc(fmtDate(t.latestReceived||''))}</span></div>
    </div>
  </div>`;
}
function toggleGroup(topic) {
  state.collapsedTopics.has(topic)?state.collapsedTopics.delete(topic):state.collapsedTopics.add(topic);
  renderSidebar();
}

// ── Select thread ──────────────────────────────────────────────────────────────
async function selectThread(convKey) {
  state.selectedKey = convKey;
  state.expandedMsgs = new Set();
  state.currentMsgs = [];
  state.triageView = false;
  state.mailboxContext = false;
  const rb=document.getElementById('resync-thread-btn');
  if(rb){rb.disabled=false;rb.textContent='↺ Resync Thread';}
  renderSidebar();
  const t = state.threadMap[convKey];
  if (!t) return;
  document.getElementById('empty-pane').style.display='none';
  document.getElementById('triage-pane').style.display='none';
  document.getElementById('triage-sidebar-btn').classList.remove('active');
  document.getElementById('thread-detail').style.display='flex';
  _renderThreadHdr(t);
  const sec = document.getElementById('msgs-section');
  sec.innerHTML=`<div class="msg-ai-loading"><div class="spinner spinner-sm"></div> Loading messages…</div>`;
  if (!t.emailIds||!t.emailIds.length){sec.innerHTML+='<div style="color:#5ba4cf;font-size:12px;padding:10px 0">No messages found.</div>';return;}
  const params=(t.emailIds||[]).map(id=>`id=${encodeURIComponent(id)}`).join('&');
  const d=await fetch('/api/thread_messages?'+params).then(r=>r.json()).catch(()=>({messages:[]}));
  let msgs=d.messages||[];
  msgs=msgs.slice().sort((a,b)=>(b.received_date_time||'')>(a.received_date_time||'')?1:-1);
  state.currentMsgs=msgs;
  _renderMsgs(msgs,t);
  if (msgs.length>0) toggleMsg(0);
  // Auto mark-read
  if (t.hasUnread && t.emailIds && t.emailIds.length) {
    fetch('/api/markread',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({ids:t.emailIds,conversationKey:convKey})}).then(()=>{
      t.hasUnread=false;
      renderSidebar();
    }).catch(()=>{});
  }
}

function _renderThreadHdr(t) {
  const enc=encodeThread(t);
  const urgCls={high:'urg-high',medium:'urg-medium',low:'urg-low'}[t.urgency]||'urg-low';
  const actCls={reply:'act-reply',delete:'act-delete',file:'act-file'}[t.action]||'';
  const parts=t.participants||[];
  const avHTML=parts.slice(0,5).map(p=>`<span class="avatar" title="${esc(p)}" style="background:${avColor(p)}">${initials(p)}</span>`).join('');
  const dateStr=fmtDate(t.latestReceived||'');
  const fileLabel=t.suggestedFolder?`📁 ${esc(t.suggestedFolder)}`:'📁 File';
  let fileBtnHtml=`<button class="btn btn-file btn-sm" onclick="openFile('${enc}')">${fileLabel}</button>`;
  document.getElementById('thread-hdr').innerHTML=`
    <div class="th-top">
      ${state.mailboxContext?`<button class="mbox-back" onclick="backToMailboxList()">✕ Close</button>`:''}
      <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
      <span class="act-pill ${actCls}">${t.action||'read'}</span>
      ${t.hasUnread?'<span style="width:6px;height:6px;border-radius:50%;background:#58a6ff;display:inline-block;flex-shrink:0"></span>':''}
      <div class="th-subject">${esc(t.subject||'(No subject)')}</div>
      <div class="th-date">${esc(dateStr)}</div>
    </div>
    <div class="th-participants">
      <div class="avatars">${avHTML}</div>
      <span class="th-names">${esc(parts.slice(0,4).join(', '))}${parts.length>4?' +'+(parts.length-4):''}</span>
      <span class="th-msgcount">${t.messageCount||0} msg${(t.messageCount||0)!==1?'s':''}</span>
    </div>
    ${t.summary?`<div class="th-summary"><div class="th-summary-lbl">🤖 AI Summary</div>${esc(t.summary)}</div>`:''}
    <div class="th-actions">
      <button class="btn btn-reply btn-sm" onclick="openReply('${enc}')">↩ Reply</button>
      ${fileBtnHtml}
      <button class="btn btn-flag btn-sm${t.isFlagged?' flagged':''}" id="flag-btn-${esc(t.conversationKey)}" onclick="toggleFlag('${enc}')">${t.isFlagged?'🚩 Flagged':'🚩 Flag'}</button>
      <button class="btn btn-delete btn-sm" onclick="openDelete('${enc}')">🗑 Delete</button>
    </div>`;
}

function _renderMsgs(msgs, t) {
  const sec=document.getElementById('msgs-section');
  if (!msgs.length){sec.innerHTML=`<div style="color:#5ba4cf;font-size:12px;padding:16px 0">No messages found.</div>`;return;}
  let html=`<div class="msgs-label"></div>`;
  html+=msgs.map((m,i)=>_msgCardHTML(m,i)).join('');
  sec.innerHTML=html;
}

function _msgCardHTML(m, idx) {
  const from=m.from_name||m.from_address||'Unknown';
  const date=fmtDate((m.received_date_time||'').slice(0,19));
  const bodyText=String(m.body||m.body_preview||'').trim();
  const preview=bodyText.slice(0,100).replace(/\n+/g,' ');
  const isOpen=state.expandedMsgs.has(idx);
  const toList=(m.to_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const ccList=(m.cc_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const recipRow=(toList||ccList)?`<div class="msg-recips">`
    +(toList?`<span><span class="msg-recip-lbl">To:</span>${toList}</span>`:'')
    +(ccList?`<span><span class="msg-recip-lbl">CC:</span>${ccList}</span>`:'')
    +`</div>`:'';
  const fmtOriginal = !!state.showOriginal[m.id];
  const hasHtml = !!(m.body_html);
  const fmtBtnLabel = fmtOriginal ? 'AI view' : (hasHtml ? 'HTML' : 'Original');
  return `<div class="msg-card${isOpen?' open':''}" id="mc-${idx}">
    <div class="msg-hdr" onclick="toggleMsg(${idx})">
      <span class="avatar" style="background:${avColor(from)};width:24px;height:24px;font-size:8.5px;border:2px solid #0a1628;flex-shrink:0">${initials(from)}</span>
      <span class="msg-from-wrap"><span class="msg-from">${esc(from)}</span><span class="msg-preview">${esc(preview)}</span></span>
      <span class="msg-date">${esc(date)}</span>
      <button class="btn btn-ghost btn-sm" id="fmt-btn-${idx}" onclick="(event||window.event).stopPropagation(); toggleFormatView(${idx})">${fmtBtnLabel}</button>
      <span class="msg-chevron">▾</span>
    </div>
    ${recipRow}
    <div class="msg-body" id="mb-${idx}">${isOpen?_bodyContent(idx):''}
    </div>
  </div>`;
}

function _bodyContent(idx) {
  const m=state.currentMsgs[idx];
  if (!m) return '';
  const showOriginal = !!state.showOriginal[m.id];
  const originalText = String(m.body||m.body_preview||'').trim();
  if (showOriginal) {
    if (m.body_html) {
      const safe = m.body_html.replace(/"/g, '&quot;');
      return `<iframe sandbox="allow-same-origin" srcdoc="${safe}"
        style="width:100%;border:none;min-height:200px;display:block;background:#fff;border-radius:4px;"
        onload="this.style.height=Math.min(700,this.contentDocument.body.scrollHeight+20)+'px'"></iframe>`;
    }
    return `<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(originalText)}</div>`;
  }
  if (state.formatCache[m.id]) return _renderParas(state.formatCache[m.id]);
  setTimeout(()=>loadFormatted(idx),0);
  return `<div class="stream-wrap" id="sw-${idx}"><span class="stream-cursor"></span></div>`;
}

function toggleMsg(idx) {
  const card=document.getElementById('mc-'+idx);
  const body=document.getElementById('mb-'+idx);
  if (!card||!body) return;
  if (state.expandedMsgs.has(idx)) {
    state.expandedMsgs.delete(idx);
    card.classList.remove('open');
    body.innerHTML='';
  } else {
    state.expandedMsgs.add(idx);
    card.classList.add('open');
    body.innerHTML=_bodyContent(idx);
  }
}

function loadFormatted(idx) {
  const m=state.currentMsgs[idx];
  if (!m||!state.expandedMsgs.has(idx)) return;
  const bodyEl=document.getElementById('mb-'+idx);
  if (!bodyEl) return;

  // Cached — render immediately, no stream needed
  if (state.formatCache[m.id]) {
    bodyEl.innerHTML=_renderParas(state.formatCache[m.id]);
    return;
  }

  // Set up streaming container
  bodyEl.innerHTML='<div class="stream-wrap" id="sw-'+idx+'"></div>';
  const wrap=document.getElementById('sw-'+idx);
  let shownCount=0;
  let accumulated='';

  const es=new EventSource(`/api/format_message_stream?id=${encodeURIComponent(m.id)}`);

  es.onmessage=(evt)=>{
    const data=JSON.parse(evt.data);
    if (data.type==='token') {
      accumulated+=data.text;
      if (!wrap||!state.expandedMsgs.has(idx)) {es.close();return;}
      // Extract completed "text":"..." values as paragraphs become available
      const matches=[...accumulated.matchAll(/"text":\s*"((?:[^"\\]|\\.)*)"/g)];
      for (let i=shownCount;i<matches.length;i++) {
        const txt=matches[i][1].replace(/\\n/g,'\n').replace(/\\"/g,'"').replace(/\\\\/g,'\\');
        const div=document.createElement('div');
        div.className='stream-para';
        div.textContent=txt;
        wrap.appendChild(div);
      }
      shownCount=matches.length;
      // Keep cursor at end
      let cur=wrap.querySelector('.stream-cursor');
      if (!cur){cur=document.createElement('span');cur.className='stream-cursor';wrap.appendChild(cur);}
      else wrap.appendChild(cur); // move to end
    } else if (data.type==='done') {
      es.close();
      state.formatCache[m.id]=data.paragraphs||[];
      // Only overwrite the body with AI view if the user hasn't switched to Original
      if (bodyEl&&state.expandedMsgs.has(idx) && !state.showOriginal[m.id]) bodyEl.innerHTML=_renderParas(state.formatCache[m.id]);
    }
  };

  es.onerror=()=>{
    es.close();
    if (bodyEl&&state.expandedMsgs.has(idx)) {
      const fallback=String(m.body||m.body_preview||'').trim();
      bodyEl.innerHTML=`<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(fallback)}</div>`;
    }
  };
}

function toggleFormatView(idx) {
  const m = state.currentMsgs[idx];
  if (!m) return;
  state.showOriginal[m.id] = !state.showOriginal[m.id];
  const bodyEl = document.getElementById('mb-'+idx);
  const btn = document.getElementById('fmt-btn-'+idx);
  if (btn) btn.textContent = state.showOriginal[m.id] ? 'AI view' : (m.body_html ? 'HTML' : 'Original');
  if (state.expandedMsgs.has(idx) && bodyEl) {
    bodyEl.innerHTML = _bodyContent(idx);
    // If switching to AI view and content isn't cached yet, trigger load
    if (!state.showOriginal[m.id] && !state.formatCache[m.id]) loadFormatted(idx);
  }
}

// intent → CSS class suffix
const INTENT_CLS = {
  'Status Update':'status-update','Request':'request','Decision':'decision',
  'Question':'question','Action Item':'action-item','Context':'context',
  'FYI':'fyi','Warning':'warning','Introduction':'introduction','Closing':'closing'
};

function _renderParas(paras) {
  if (!paras||!paras.length) return '<div style="color:#5ba4cf;font-size:12px">(no content)</div>';
  return paras.map(p=>{
    const cls='i-'+(INTENT_CLS[p.intent]||'context');
    const factHtml=p.fact_concern
      ?`<div class="fact-warn"><span>⚠️</span><span>${esc(p.fact_concern)}</span></div>`:'';
    return `<div class="para-blk">
      <div class="intent-pill ${cls}">${esc(p.emoji||'')} ${esc(p.intent||'FYI')}</div>
      <div class="para-txt">${linkify(p.text||'')}</div>
      ${factHtml}
    </div>`;
  }).join('');
}

function highlightEmails(html) {
  return html.replace(/([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})/g,
    '<span class="eml">$1</span>');
}
function linkify(text) {
  // Split on URLs, linkify those parts, highlight emails in non-URL parts
  const parts=text.split(/(https?:\/\/[^\s<>"')\]]+)/g);
  return parts.map((part,i)=>{
    if(i%2===1){const href=esc(part);return `<a href="${href}" target="_blank" rel="noopener noreferrer" class="link">${href}</a>`;}
    return highlightEmails(esc(part));
  }).join('');
}

// ── Encode/decode thread ───────────────────────────────────────────────────────
function encodeThread(t) {
  try {
    const j=JSON.stringify({conversationKey:t.conversationKey,latestId:t.latestId,emailIds:t.emailIds,subject:t.subject,messageCount:t.messageCount,suggestedReply:t.suggestedReply,suggestedFolder:t.suggestedFolder});
    return btoa(unescape(encodeURIComponent(j))).replace(/=/g,'');
  } catch(e){return '';}
}
function decodeThread(s) {
  try { return JSON.parse(decodeURIComponent(escape(atob(s.replace(/-/g,'+').replace(/_/g,'/'))))); }
  catch{return {};}
}

// ── Sync status ────────────────────────────────────────────────────────────────
function updateSyncStatus(ss) {
  if (!ss) return;
  const dot=document.getElementById('sync-dot');
  const txt=document.getElementById('sync-txt');
  const wrap=document.getElementById('sync-bar-wrap');
  const bar=document.getElementById('sync-bar');
  if (ss.running) {
    dot.className='sync-dot syncing';
    txt.textContent=ss.progress||'Syncing…';
    if (ss.total>0){const pct=Math.round((ss.done/ss.total)*100);wrap.style.display='block';bar.style.width=Math.max(4,pct)+'%';}
    else wrap.style.display='none';
  } else {
    if(wrap) wrap.style.display='none';
    if (ss.lastError){dot.className='sync-dot error';txt.textContent='Sync error: '+ss.lastError;}
    else if (ss.lastSync){
      dot.className='sync-dot';
      const mins=Math.round((Date.now()-new Date(ss.lastSync))/60000);
      const ago=mins<1?'just now':mins===1?'1 min ago':`${mins} min ago`;
      txt.textContent=`Synced ${ago}`+(ss.threadsUpdated>0?` · ${ss.threadsUpdated} updated`:'');
    } else {dot.className='sync-dot syncing';txt.textContent='Waiting for first sync…';}
  }
}
function _showNewBadge(n) {
  const w=document.getElementById('new-badge-wrap');
  w.innerHTML=`<span class="new-badge">${n} new</span>`;
  setTimeout(()=>{w.innerHTML='';},5000);
}
function updateCounts(emailCount,threadCount) {
  const el=document.getElementById('sidebar-counts');
  if (!el) return;
  el.textContent=emailCount!==null?`${emailCount} emails · ${threadCount} threads`:`${threadCount} threads`;
}
async function triggerSync() {
  const d=await fetch('/api/sync_now',{method:'POST'}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
}
async function resyncThread() {
  const convKey=state.selectedKey;
  if (!convKey) return;
  const btn=document.getElementById('resync-thread-btn');
  btn.disabled=true; btn.textContent='↺ Resyncing…';
  const d=await fetch('/api/resync_thread',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({conversationKey:convKey})}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
  // Reload the thread view after resync completes
  setTimeout(async()=>{
    btn.disabled=false; btn.textContent='↺ Resync Thread';
    if (state.selectedKey===convKey) await selectThread(convKey);
  }, 1500);
}
async function reanalyzeAll() {
  const btn=document.getElementById('reanalyze-btn');
  btn.disabled=true; btn.textContent='⚙ Re-analyzing…';
  const d=await fetch('/api/reanalyze_all',{method:'POST'}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
  setTimeout(()=>{btn.disabled=false;btn.textContent='⚙ Re-analyze';},3000);
}

// ── Modals ─────────────────────────────────────────────────────────────────────
// ── Reply (2-step) ─────────────────────────────────────────────────────────────
let _replyState = {thread:null, to:[], cc:[]};
let MY_EMAIL = '';

async function openReply(enc) {
  _replyState.thread = decodeThread(enc);
  _activeThread = _replyState.thread;
  const t = _replyState.thread;

  // Build reply-all recipients: sender + all To recipients, CC from latest message; filter self
  // Use newest message (index 0 after newest-first sort)
  const latest = state.currentMsgs[0];
  _replyState.to = [];
  _replyState.cc = [];
  const myAddr = MY_EMAIL.toLowerCase();
  const addUniq = (list, r) => {
    if (!r || !r.address) return;
    if (r.address.toLowerCase() === myAddr) return; // exclude self
    if (!list.find(x=>x.address.toLowerCase()===r.address.toLowerCase())) list.push(r);
  };
  if (latest) {
    // Sender goes to To
    if (latest.from_address) addUniq(_replyState.to, {name:latest.from_name||latest.from_address, address:latest.from_address});
    // All To recipients
    for (const r of (latest.to_recipients||[])) addUniq(_replyState.to, r);
    // All CC recipients
    for (const r of (latest.cc_recipients||[])) addUniq(_replyState.cc, r);
  }

  document.getElementById('reply-sub').textContent = `Re: ${t.subject||''}`;
  const bodyEl = document.getElementById('reply-body');
  bodyEl.value = '';
  bodyEl.placeholder = 'Generating reply…';
  _renderRecipFields();
  document.getElementById('reply-modal').classList.add('open');
  document.getElementById('reply-generating').style.display = 'flex';

  // Auto-populate with suggested reply
  try {
    // Use cached suggestedReply if available
    const thread = state.threadMap[t.conversationKey];
    if (thread && thread.suggestedReply) {
      bodyEl.value = thread.suggestedReply;
      bodyEl.placeholder = '';
    } else {
      const d = await fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({conversationKey: t.conversationKey})}).then(r=>r.json()).catch(()=>null);
      if (d && d.reply) {
        bodyEl.value = d.reply;
        if (thread) thread.suggestedReply = d.reply;
      }
      bodyEl.placeholder = '';
    }
  } catch(e) { bodyEl.placeholder = 'Write your reply…'; }
  document.getElementById('reply-generating').style.display = 'none';
  setTimeout(()=>{ bodyEl.focus(); bodyEl.setSelectionRange(0,0); }, 50);
}

async function regenerateReply() {
  const t = _replyState.thread;
  if (!t) return;
  const bodyEl = document.getElementById('reply-body');
  const context = bodyEl.value.trim();
  const btn = document.getElementById('reply-regen-btn');
  btn.disabled = true; btn.textContent = '↺ Generating…';
  document.getElementById('reply-generating').style.display = 'flex';
  try {
    const d = await fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: t.conversationKey, context})}).then(r=>r.json()).catch(()=>null);
    if (d && d.reply) {
      bodyEl.value = d.reply;
      const thread = state.threadMap[t.conversationKey];
      if (thread) thread.suggestedReply = d.reply;
    }
  } finally {
    document.getElementById('reply-generating').style.display = 'none';
    btn.disabled = false; btn.textContent = '↺ Regenerate';
  }
}


function _renderRecipFields() {
  _renderTags('reply-to-field', _replyState.to, 'to');
  _renderTags('reply-cc-field', _replyState.cc, 'cc');
}
function _renderTags(fieldId, list, field) {
  const el = document.getElementById(fieldId);
  if (!list.length) { el.innerHTML='<span class="recip-empty">none</span>'; return; }
  el.innerHTML = list.map(r=>
    `<span class="recip-tag" data-field="${field}" data-addr="${esc(r.address)}">${esc(r.name||r.address)}<span class="rm" data-rm-field="${field}" data-rm-addr="${esc(r.address)}">×</span></span>`
  ).join('');
  el.querySelectorAll('.rm').forEach(btn=>{
    btn.addEventListener('click', ()=>removeRecip(btn.dataset.rmField, btn.dataset.rmAddr));
  });
}
function removeRecip(field, address) {
  _replyState[field] = _replyState[field].filter(r=>r.address!==address);
  _renderRecipFields();
}
function openFile(enc) {
  _activeThread=decodeThread(enc);
  const suggested=_activeThread.suggestedFolder||'';
  document.getElementById('file-sub').textContent=`"${_activeThread.subject||''}"`;
  const populate=(folders,effortsFolders)=>{
    const es=new Set(effortsFolders||[]);
    const eff=folders.filter(n=>es.has(n)),oth=folders.filter(n=>!es.has(n));
    let h='';
    if (eff.length) h+=`<optgroup label="Efforts">`+eff.map(n=>`<option value="${esc(n)}"${n===suggested?' selected':''}>${esc(n)}</option>`).join('')+`</optgroup>`;
    if (oth.length) h+=`<optgroup label="Other">`+oth.map(n=>`<option value="${esc(n)}"${n===suggested?' selected':''}>${esc(n)}</option>`).join('')+`</optgroup>`;
    document.getElementById('folder-select').innerHTML=h;
  };
  if (state.folders.length){populate(state.folders,state.effortsFolders);document.getElementById('file-modal').classList.add('open');}
  else fetch('/api/folders').then(r=>r.json()).then(d=>{state.folders=d.folders||[];state.effortsFolders=d.effortsFolders||[];populate(state.folders,state.effortsFolders);document.getElementById('file-modal').classList.add('open');});
}
async function openDelete(enc) {
  const t=decodeThread(enc);
  await _act('/api/delete',{ids:t.emailIds,conversationKey:t.conversationKey},t.conversationKey);
}
function closeModals() {
  document.querySelectorAll('.modal-overlay').forEach(m=>m.classList.remove('open'));
  _activeThread=null;
  // Close any open people dropdowns
  document.querySelectorAll('.people-dropdown').forEach(d=>d.classList.remove('open'));
}

// ── Compose ────────────────────────────────────────────────────────────────────
let _composeState = {to:[], cc:[]};
let _pdCache = null;

function openCompose() {
  _composeState = {to:[], cc:[]};
  document.getElementById('compose-to-field').innerHTML='';
  document.getElementById('compose-cc-field').innerHTML='';
  document.getElementById('compose-to-input').value='';
  document.getElementById('compose-cc-input').value='';
  document.getElementById('compose-subject').value='';
  document.getElementById('compose-body').value='';
  document.getElementById('compose-modal').classList.add('open');
  setTimeout(()=>document.getElementById('compose-to-input').focus(), 50);
}

function focusComposeInput(field) {
  document.getElementById(`compose-${field}-input`).focus();
}

function _renderComposeTags(field) {
  const el = document.getElementById(`compose-${field}-field`);
  el.innerHTML = _composeState[field].map(r=>
    `<span class="recip-tag" data-addr="${esc(r.address)}">${esc(r.name||r.address)}<span class="rm" data-rm-field="${field}" data-rm-addr="${esc(r.address)}">×</span></span>`
  ).join('');
  el.querySelectorAll('.rm').forEach(btn=>{
    btn.addEventListener('click', ()=>{ _composeState[btn.dataset.rmField]=_composeState[btn.dataset.rmField].filter(r=>r.address!==btn.dataset.rmAddr); _renderComposeTags(btn.dataset.rmField); });
  });
}

function _addComposeRecip(field, person) {
  if (!person.address) return;
  if (!_composeState[field].find(r=>r.address.toLowerCase()===person.address.toLowerCase()))
    _composeState[field].push(person);
  _renderComposeTags(field);
  const inp = document.getElementById(`compose-${field}-input`);
  inp.value = '';
  document.getElementById(`pd-${field}`).classList.remove('open');
}

let _pdTimers = {};
async function peopleSuggest(inp, field) {
  const q = inp.value.trim();
  if (!q) { document.getElementById(`pd-${field}`).classList.remove('open'); return; }
  clearTimeout(_pdTimers[field]);
  _pdTimers[field] = setTimeout(async ()=>{
    let people;
    if (_pdCache) {
      people = _pdCache.filter(p=>(p.name||'').toLowerCase().includes(q.toLowerCase())||(p.address||'').toLowerCase().includes(q.toLowerCase())).slice(0,12);
    } else {
      const d = await fetch(`/api/people?q=${encodeURIComponent(q)}`).then(r=>r.json()).catch(()=>null);
      people = d ? d.people : [];
    }
    const dd = document.getElementById(`pd-${field}`);
    if (!people.length) { dd.classList.remove('open'); return; }
    dd.innerHTML = people.map((p,i)=>
      `<div class="pd-item" data-i="${i}"><span class="pd-item-name">${esc(p.name||p.address)}</span><span class="pd-item-addr">${esc(p.address)}</span></div>`
    ).join('');
    dd._people = people;
    dd.querySelectorAll('.pd-item').forEach((item,i)=>{
      item.addEventListener('mousedown', e=>{ e.preventDefault(); _addComposeRecip(field, dd._people[i]); });
    });
    dd.classList.add('open');
  }, 150);
}

function peopleSuggestKey(e, field) {
  const dd = document.getElementById(`pd-${field}`);
  const inp = document.getElementById(`compose-${field}-input`);
  if (e.key==='Enter' || e.key===',') {
    e.preventDefault();
    // If dropdown has active item, use it; otherwise treat as raw email
    const active = dd.querySelector('.pd-item.active');
    if (active && dd._people) {
      const i = parseInt(active.dataset.i);
      _addComposeRecip(field, dd._people[i]);
    } else if (inp.value.includes('@')) {
      _addComposeRecip(field, {name:'', address:inp.value.trim()});
    }
    return;
  }
  if (e.key==='ArrowDown'||e.key==='ArrowUp') {
    e.preventDefault();
    const items = [...dd.querySelectorAll('.pd-item')];
    if (!items.length) return;
    const cur = dd.querySelector('.pd-item.active');
    const idx = cur ? items.indexOf(cur) : -1;
    if (cur) cur.classList.remove('active');
    const next = items[(idx + (e.key==='ArrowDown'?1:-1) + items.length) % items.length];
    next.classList.add('active');
    return;
  }
  if (e.key==='Escape') { dd.classList.remove('open'); }
  if (e.key==='Backspace' && !inp.value && _composeState[field].length) {
    _composeState[field].pop();
    _renderComposeTags(field);
  }
}

async function sendNewMessage() {
  const to = _composeState.to.map(r=>r.address).filter(Boolean);
  const cc = _composeState.cc.map(r=>r.address).filter(Boolean);
  const subject = document.getElementById('compose-subject').value.trim();
  const body = document.getElementById('compose-body').value.trim();
  if (!to.length) { alert('Please add at least one recipient.'); return; }
  if (!subject) { alert('Please add a subject.'); return; }
  closeModals();
  const btn = null; // not available after close
  const d = await fetch('/api/send_new', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({to, cc, subject, body})}).then(r=>r.json()).catch(()=>null);
  if (d && d.ok) {
    showToast('Message sent');
  } else {
    showToast('Send failed: '+(d&&d.error||'unknown error'), true);
  }
}

function showToast(msg, isError=false) {
  let el = document.getElementById('app-toast');
  if (!el) { el = document.createElement('div'); el.id='app-toast'; el.style.cssText='position:fixed;bottom:24px;right:24px;background:#1a3a5c;color:#c9d1d9;border:1px solid #2a5a8a;border-radius:8px;padding:10px 18px;font-size:13px;z-index:99999;transition:opacity .3s;'; document.body.appendChild(el); }
  if (isError) el.style.borderColor='#f85149';
  else el.style.borderColor='#2a5a8a';
  el.textContent = msg;
  el.style.opacity='1';
  clearTimeout(el._t);
  el._t = setTimeout(()=>el.style.opacity='0', 3000);
}

// ── Actions ────────────────────────────────────────────────────────────────────
async function sendReply() {
  const body=document.getElementById('reply-body').value.trim();
  if (!body) return;
  const t=_replyState.thread||_activeThread;
  const to=_replyState.to.map(r=>r.address).filter(Boolean);
  const cc=_replyState.cc.map(r=>r.address).filter(Boolean);
  closeModals();
  await _act('/api/reply/'+t.latestId,{body,conversationKey:t.conversationKey,to,cc},t.conversationKey);
}
async function fileThread() {
  const folder=document.getElementById('folder-select').value;
  const t=_activeThread; closeModals();
  await _act('/api/move',{ids:t.emailIds,folder,conversationKey:t.conversationKey},t.conversationKey);
}
async function quickFile(enc,folder) {
  const t=decodeThread(enc);
  await _act('/api/move',{ids:t.emailIds,folder,conversationKey:t.conversationKey},t.conversationKey);
}
async function confirmDelete() {
  const t=_activeThread; closeModals();
  await _act('/api/delete',{ids:t.emailIds,conversationKey:t.conversationKey},t.conversationKey);
}
async function toggleFlag(enc) {
  const t=decodeThread(enc);
  const thread=state.threadMap[t.conversationKey];
  if (!thread) return;
  const nowFlagged=!thread.isFlagged;
  const d=await fetch('/api/flag',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({conversationKey:t.conversationKey,flagged:nowFlagged})}).then(r=>r.json()).catch(()=>null);
  if (!d||!d.ok) return;
  thread.isFlagged=nowFlagged;
  // Update flag button immediately
  const btn=document.getElementById('flag-btn-'+t.conversationKey);
  if (btn){
    btn.textContent=nowFlagged?'🚩 Flagged':'🚩 Flag';
    btn.className='btn btn-flag btn-sm'+(nowFlagged?' flagged':'');
  }
  renderSidebar();
}
async function _act(url,body,convKey) {
  const res=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const d=await res.json();
  if (!d.ok) return alert('Error: '+(d.error||'Unknown error'));
  // Find next mailbox row before removing from DOM
  const mboxRow = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  const nextRow = mboxRow?.nextElementSibling;
  if (mboxRow) mboxRow.remove();
  delete state.threadMap[convKey];
  for (const g of state.groups) g.threads=g.threads.filter(t=>t.conversationKey!==convKey);
  state.groups=state.groups.filter(g=>g.threads.length>0);
  if (state.selectedKey===convKey) {
    state.selectedKey=null;
    const rb=document.getElementById('resync-thread-btn');
    if(rb) rb.disabled=true;
    if (state.mailboxContext && nextRow && nextRow.dataset.key) {
      // Advance to next thread in folder
      openMailboxThread(nextRow.dataset.key, nextRow.dataset.folder || mailboxCurrentFolder);
    } else if (state.mailboxContext) {
      backToMailboxList();
    } else {
      document.getElementById('thread-detail').style.display='none';
      document.getElementById('empty-pane').style.display='flex';
    }
  }
  renderSidebar();
  updateCounts(null,Object.keys(state.threadMap).length);
}

// ── Inline reply ───────────────────────────────────────────────────────────────
async function sendInlineReply(enc) {
  const t = decodeThread(enc);
  const ta = document.getElementById('inline-reply-'+enc);
  const body = ta ? ta.value.trim() : '';
  if (!body) return;
  const res = await fetch('/api/reply/'+t.latestId, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({body, conversationKey: t.conversationKey, to: [], cc: []})
  }).then(r=>r.json()).catch(()=>null);
  if (!res || !res.ok) { alert('Error: '+(res&&res.error||'Unknown error')); return; }
  delete state.threadMap[t.conversationKey];
  for (const g of state.groups) g.threads = g.threads.filter(th=>th.conversationKey!==t.conversationKey);
  state.groups = state.groups.filter(g=>g.threads.length>0);
  if (state.selectedKey===t.conversationKey) {
    state.selectedKey = null;
    document.getElementById('thread-detail').style.display='none';
    document.getElementById('empty-pane').style.display='flex';
    const rb=document.getElementById('resync-thread-btn');
    if(rb) rb.disabled=true;
  }
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
}

async function regenerateInlineReply(enc) {
  const t = decodeThread(enc);
  const ta = document.getElementById('inline-reply-'+enc);
  if (ta) { ta.value = '⏳ Generating...'; ta.disabled = true; }
  try {
    const res = await fetch('/api/suggested_reply', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: t.conversationKey})
    }).then(r=>r.json());
    if (res.reply) {
      const thread = state.threadMap[t.conversationKey];
      if (thread) thread.suggestedReply = res.reply;
      if (ta) { ta.value = res.reply; ta.disabled = false; }
      else {
        // textarea may not exist yet (was "Generate reply..." button) — re-render header
        const fullThread = state.threadMap[t.conversationKey];
        if (fullThread) _renderThreadHdr(fullThread);
      }
    } else {
      if (ta) { ta.value = ''; ta.disabled = false; }
    }
  } catch(e) {
    if (ta) { ta.value = ''; ta.disabled = false; }
    alert('Error regenerating reply: '+e);
  }
}

// ── Triage sheet ───────────────────────────────────────────────────────────────
function openTriageSheet() {
  state.triageView = true;
  state.triageFocusIdx = -1;
  state.selectedKey = null;
  document.getElementById('empty-pane').style.display='none';
  document.getElementById('thread-detail').style.display='none';
  const pane = document.getElementById('triage-pane');
  pane.style.display='flex';
  // Attach delegated click handler once
  if (!pane._triageDelegated) {
    pane._triageDelegated = true;
    pane.addEventListener('click', _triagePaneClick);
  }
  renderSidebar();
  renderTriageSheet();
  initMailbox(); // ensure folder tree is populated
  document.addEventListener('keydown', _triageKeydown);
}

function _triagePaneClick(e) {
  // Triage action buttons
  const btn = e.target.closest('[data-triage-action]');
  if (btn) {
    e.stopPropagation();
    const row = btn.closest('[data-convkey]');
    if (!row) return;
    const convKey = row.dataset.convkey;
    const action  = btn.dataset.triageAction;
    if      (action==='reply')  triageOpenReply(convKey);
    else if (action==='file')   triageMark(convKey,'file');
    else if (action==='delete') triageMark(convKey,'delete');
    else if (action==='clear')  triageMark(convKey,null);
    return;
  }
  // Row expand/collapse (summary area)
  const summary = e.target.closest('[data-triage-expand]');
  if (summary) {
    const row = summary.closest('[data-convkey]');
    if (row) triageToggleExpand(row.dataset.convkey);
    return;
  }
  // Inline message row toggle
  const msgRow = e.target.closest('[data-triage-msg]');
  if (msgRow) {
    const row = msgRow.closest('[data-convkey]');
    if (row) triageToggleMsgBody(row.dataset.convkey, parseInt(msgRow.dataset.triageMsg));
    return;
  }
  // Topic header collapse/expand
  const topicHdr = e.target.closest('[data-triage-topic]');
  if (topicHdr) {
    toggleTriageTopic(topicHdr.dataset.topic);
  }
}

const ACTION_REC = {
  reply:  {cls:'action-rec-reply',  icon:'↩', label:'Reply Needed'},
  delete: {cls:'action-rec-delete', icon:'🗑', label:'Safe to Delete'},
  file:   {cls:'action-rec-file',   icon:'📁', label:'File for Reference'},
  read:   {cls:'action-rec-read',   icon:'👁', label:'FYI Only'},
  done:   {cls:'action-rec-done',   icon:'✓',  label:'Resolved'},
};

function _triageRowHTML(t) {
  const convKey = t.conversationKey;
  const action = state.triageActions[convKey];
  const actionCls = action ? ' ts-'+action.type : '';
  const expanded = state.expandedTriageRows.has(convKey);
  const urgCls = {high:'urg-high',medium:'urg-medium',low:'urg-low'}[t.urgency]||'urg-low';
  const rec = ACTION_REC[t.action] || ACTION_REC.read;
  const statusLbl = action ? (action.type==='delete'?'🗑 Queued':action.type==='file'?'📁 Queued':'') : '';
  const msgsHtml = expanded ? _triageMsgsHTML(convKey) : '';
  return `<div class="triage-row${actionCls}${expanded?' expanded':''}" id="triage-row-${esc(convKey)}" data-convkey="${esc(convKey)}">
    <div class="triage-row-summary" data-triage-expand="1">
      <div class="triage-row-top">
        <span class="triage-row-expand-chevron">▶</span>
        <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
        <span class="triage-subj">${esc(t.subject||'(No subject)')}</span>
        <span class="action-rec ${rec.cls}">${rec.icon} ${rec.label}</span>
      </div>
      ${t.summary?`<div class="triage-sum" style="padding-left:17px">${esc(t.summary)}</div>`:''}
    </div>
    <div class="triage-msgs" id="triage-msgs-${esc(convKey)}" style="${expanded?'':'display:none'}">${msgsHtml}</div>
    <div class="triage-btns">
      <button class="btn btn-reply btn-sm" data-triage-action="reply">↩ Reply</button>
      <button class="btn btn-ghost btn-sm btn-ts-file${action&&action.type==='file'?' active':''}" data-triage-action="file">📁 File</button>
      <button class="btn btn-ghost btn-sm btn-ts-del${action&&action.type==='delete'?' active':''}" data-triage-action="delete">🗑 Delete</button>
      ${action?`<button class="btn btn-ghost btn-sm" data-triage-action="clear">✕</button>`:''}
      <span class="triage-qlbl">${esc(statusLbl)}</span>
    </div>
  </div>`;
}

function _triageMsgsHTML(convKey) {
  const msgs = state.triageMsgCache[convKey];
  if (!msgs) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block"></div> Loading…</div>';
  if (!msgs.length) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf">No messages</div>';
  return [...msgs].reverse().map((m,i) => {
    const from = m.from_name||m.from_address||'?';
    const preview = String(m.body||m.body_preview||'').replace(/\s+/g,' ').trim().slice(0,80);
    const date = fmtDate((m.received_date_time||'').slice(0,19));
    return `<div class="triage-msg-row" id="tmr-${esc(convKey)}-${i}" data-triage-msg="${i}">
      <span class="triage-msg-chev">▶</span>
      <span class="triage-msg-from">${esc(from)}</span>
      <span class="triage-msg-prev">${esc(preview)}</span>
      <span class="triage-msg-date">${esc(date)}</span>
    </div>
    <div class="triage-msg-body" id="tmb-${esc(convKey)}-${i}" style="display:none">${esc(String(m.body||m.body_preview||'').trim())}</div>`;
  }).join('');
}

async function triageToggleExpand(convKey) {
  const row = document.getElementById('triage-row-'+convKey);
  const msgsEl = document.getElementById('triage-msgs-'+convKey);
  if (!row || !msgsEl) return;
  const expanding = !state.expandedTriageRows.has(convKey);
  if (expanding) {
    state.expandedTriageRows.add(convKey);
    row.classList.add('expanded');
    msgsEl.style.display = '';
    if (!state.triageMsgCache[convKey]) {
      msgsEl.innerHTML = '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block;margin-right:6px"></div>Loading…</div>';
      const r = await fetch(`/api/thread_messages?conversationKey=${encodeURIComponent(convKey)}`).then(r=>r.json()).catch(()=>null);
      const msgs = (r&&r.messages||[]).slice().sort((a,b)=>(b.received_date_time||'')>(a.received_date_time||'')?1:-1);
      state.triageMsgCache[convKey] = msgs;
      msgsEl.innerHTML = _triageMsgsHTML(convKey);
    }
  } else {
    state.expandedTriageRows.delete(convKey);
    row.classList.remove('expanded');
    msgsEl.style.display = 'none';
  }
}

function triageToggleMsgBody(convKey, idx) {
  const row = document.getElementById(`tmr-${convKey}-${idx}`);
  const body = document.getElementById(`tmb-${convKey}-${idx}`);
  if (!body) return;
  const open = body.style.display !== 'none';
  body.style.display = open ? 'none' : '';
  if (row) row.classList.toggle('open', !open);
}

function triageOpenReply(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  closeTriageSheet();
  // Navigate to thread then open reply modal
  selectThread(convKey);
  setTimeout(() => openReply(encodeThread(thread)), 500);
}

function toggleTriageTopic(topic) {
  if (state.collapsedTriageTopics.has(topic)) state.collapsedTriageTopics.delete(topic);
  else state.collapsedTriageTopics.add(topic);
  const grp = document.getElementById('ttg-'+btoa(unescape(encodeURIComponent(topic))).replace(/[^a-zA-Z0-9]/g,''));
  if (!grp) { renderTriageSheet(); return; }
  const rows = grp.querySelector('.triage-topic-rows');
  const hdr  = grp.querySelector('.triage-topic-hdr');
  const collapsed = state.collapsedTriageTopics.has(topic);
  if (rows) rows.style.display = collapsed ? 'none' : '';
  if (hdr)  hdr.classList.toggle('open', !collapsed);
  _triageUpdateFocus();
}

function renderTriageSheet() {
  const pane = document.getElementById('triage-pane');
  const queuedCount = Object.keys(state.triageActions).length;
  // Sort threads within each group newest-first, then sort groups by their latest thread
  const sortedGroups = state.groups.map(g => ({
    ...g,
    threads: [...g.threads].sort((a,b) => (b.latestReceived||'').localeCompare(a.latestReceived||''))
  })).sort((a,b) => {
    const aLat = a.threads[0]?.latestReceived || '';
    const bLat = b.threads[0]?.latestReceived || '';
    return bLat.localeCompare(aLat);
  });
  const groupsHtml = sortedGroups.map(g => {
    const topic = g.topic || 'Uncategorized';
    const safeId = 'ttg-'+btoa(unescape(encodeURIComponent(topic))).replace(/[^a-zA-Z0-9]/g,'');
    const collapsed = state.collapsedTriageTopics.has(topic);
    return `<div class="triage-topic-group" id="${safeId}">
      <div class="triage-topic-hdr${collapsed?'':' open'}" data-topic="${esc(topic)}" data-triage-topic="1">
        <span class="triage-topic-chevron">▶</span>
        <span class="triage-topic-label">${esc(topic)}</span>
        <span class="triage-topic-badge">${g.threads.length}</span>
      </div>
      <div class="triage-topic-rows" style="${collapsed?'display:none':''}">
        ${g.threads.map(_triageRowHTML).join('')}
      </div>
    </div>`;
  }).join('');
  pane.innerHTML = `<div class="triage-hdr">
    <span class="triage-title">📋 Triage Sheet</span>
    <span class="triage-queue-count" id="triage-queue-count">${queuedCount} queued</span>
    <button class="btn btn-reply btn-sm" id="triage-execute-btn" onclick="executeAllActions()"${queuedCount===0?' disabled':''}>⚡ Execute All</button>
    <span class="triage-kb-hint">↑↓ navigate · Enter expand · R reply · D delete · F file · Esc back</span>
  </div>
  <div class="triage-rows">${groupsHtml}</div>`;
}

// ── Triage keyboard navigation ──────────────────────────────────────────────
function _triageNavList() {
  const list = [];
  for (const g of state.groups) {
    const topic = g.topic || 'Uncategorized';
    list.push({type:'topic', topic});
    if (!state.collapsedTriageTopics.has(topic)) {
      for (const t of g.threads) list.push({type:'thread', convKey:t.conversationKey});
    }
  }
  return list;
}

function _triageUpdateFocus() {
  document.querySelectorAll('.triage-kb-focus').forEach(el=>el.classList.remove('triage-kb-focus'));
  const nav = _triageNavList();
  const item = nav[state.triageFocusIdx];
  if (!item) return;
  let el;
  if (item.type==='topic') {
    el = document.querySelector(`.triage-topic-hdr[data-topic="${CSS.escape(item.topic)}"]`);
  } else {
    el = document.getElementById('triage-row-'+item.convKey);
  }
  if (el) { el.classList.add('triage-kb-focus'); el.scrollIntoView({block:'nearest',behavior:'smooth'}); }
}

function _triageKeydown(e) {
  if (!state.triageView) return;
  // Don't intercept if user is typing in an input
  if (e.target.tagName==='INPUT'||e.target.tagName==='TEXTAREA') return;
  const nav = _triageNavList();
  let idx = state.triageFocusIdx;

  if (e.key==='ArrowDown') {
    e.preventDefault();
    idx = idx < nav.length-1 ? idx+1 : idx;
  } else if (e.key==='ArrowUp') {
    e.preventDefault();
    idx = idx > 0 ? idx-1 : 0;
  } else if (e.key==='ArrowRight'||e.key==='ArrowLeft') {
    e.preventDefault();
    const item = nav[idx];
    if (item&&item.type==='topic') {
      if (e.key==='ArrowRight') state.collapsedTriageTopics.delete(item.topic);
      else state.collapsedTriageTopics.add(item.topic);
      toggleTriageTopic(item.topic);
    } else if (item&&item.type==='thread'&&e.key==='ArrowRight') {
      triageToggleExpand(item.convKey);
    } else if (item&&item.type==='thread'&&e.key==='ArrowLeft') {
      state.expandedTriageRows.delete(item.convKey);
      const row=document.getElementById('triage-row-'+item.convKey);
      const msgsEl=document.getElementById('triage-msgs-'+item.convKey);
      if(row){row.classList.remove('expanded');}
      if(msgsEl){msgsEl.style.display='none';}
    }
    state.triageFocusIdx = idx;
    _triageUpdateFocus(); return;
  } else if (e.key==='Enter'||e.key===' ') {
    e.preventDefault();
    const item = nav[idx];
    if (!item) { idx=0; }
    else if (item.type==='topic') toggleTriageTopic(item.topic);
    else triageToggleExpand(item.convKey);
  } else if (e.key==='r'||e.key==='R') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageOpenReply(item.convKey); }
    return;
  } else if (e.key==='d'||e.key==='D') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, state.triageActions[item.convKey]?.type==='delete'?null:'delete'); }
    return;
  } else if (e.key==='f'||e.key==='F') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, state.triageActions[item.convKey]?.type==='file'?null:'file'); }
    return;
  } else if (e.key==='x'||e.key==='X') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, null); }
    return;
  } else if (e.key==='Escape') {
    e.preventDefault(); closeTriageSheet(); return;
  } else return;

  state.triageFocusIdx = idx;
  _triageUpdateFocus();
}

function closeTriageSheet() {
  state.triageView = false;
  state.triageFocusIdx = -1;
  document.removeEventListener('keydown', _triageKeydown);
  switchTab('mailbox');
}

function triageMark(convKey, type) {
  if (type === null) delete state.triageActions[convKey];
  else state.triageActions[convKey] = {type};
  // Update row visual
  const row = document.getElementById('triage-row-'+convKey);
  if (row) {
    const expanded = state.expandedTriageRows.has(convKey);
    row.className = 'triage-row' + (type?' ts-'+type:'') + (expanded?' expanded':'');
    const qlbl = row.querySelector('.triage-qlbl');
    if (qlbl) qlbl.textContent = type==='delete'?'🗑 Queued':type==='file'?'📁 Queued':'';
    row.querySelectorAll('.btn-ts-del,.btn-ts-file').forEach(b=>b.classList.remove('active'));
    if (type==='delete'){const b=row.querySelector('.btn-ts-del');if(b)b.classList.add('active');}
    else if (type==='file'){const b=row.querySelector('.btn-ts-file');if(b)b.classList.add('active');}
    // Rebuild clear button
    const btns = row.querySelector('.triage-btns');
    if (btns) {
      let clr = btns.querySelector('.btn-ts-clr');
      if (type && !clr) {
        const b=document.createElement('button');b.className='btn btn-ghost btn-sm btn-ts-clr';
        b.textContent='✕';b.onclick=()=>triageMark(convKey,null);
        btns.insertBefore(b, btns.querySelector('.triage-qlbl'));
      } else if (!type && clr) clr.remove();
    }
  }
  const queuedCount = Object.keys(state.triageActions).length;
  const countEl = document.getElementById('triage-queue-count');
  if (countEl) countEl.textContent = queuedCount+' queued';
  const execBtn = document.getElementById('triage-execute-btn');
  if (execBtn) execBtn.disabled = queuedCount === 0;
}

async function executeAllActions() {
  const entries = Object.entries(state.triageActions);
  if (!entries.length) return;
  const execBtn = document.getElementById('triage-execute-btn');
  let done = 0;
  const total = entries.length;
  for (const [convKey, action] of entries) {
    if (execBtn) execBtn.textContent = `Executing ${done+1}/${total}...`;
    const thread = state.threadMap[convKey];
    if (!thread) { done++; continue; }
    try {
      if (action.type === 'send') {
        // Open reply modal for this thread so user can compose
        closeTriageSheet();
        selectThread(convKey);
        const enc = encodeThread(thread);
        setTimeout(()=>openReply(enc), 400);
        break; // handle one reply at a time
      } else if (action.type === 'delete') {
        await fetch('/api/delete', {
          method: 'POST',
          headers: {'Content-Type':'application/json'},
          body: JSON.stringify({ids: thread.emailIds, conversationKey: convKey})
        });
      } else if (action.type === 'file') {
        await fetch('/api/move', {
          method: 'POST',
          headers: {'Content-Type':'application/json'},
          body: JSON.stringify({ids: thread.emailIds, folder: thread.suggestedFolder||'', conversationKey: convKey})
        });
      }
    } catch(e) {
      console.error('Execute action error for '+convKey, e);
    }
    // Mark row as done
    const row = document.getElementById('triage-row-'+convKey);
    if (row) {
      row.className = 'triage-row ts-done';
      const qlbl = row.querySelector('.triage-qlbl');
      if (qlbl) qlbl.textContent = '✓ Done';
    }
    // Remove from state
    delete state.triageActions[convKey];
    delete state.threadMap[convKey];
    for (const g of state.groups) g.threads = g.threads.filter(t=>t.conversationKey!==convKey);
    state.groups = state.groups.filter(g=>g.threads.length>0);
    done++;
  }
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
  // Re-render triage sheet so completed items are removed
  renderTriageSheet();
  const execBtn2 = document.getElementById('triage-execute-btn');
  if (execBtn2) { execBtn2.textContent = `✓ ${done} action${done!==1?'s':''} done`; execBtn2.disabled = true; }
}

// ── Helpers ────────────────────────────────────────────────────────────────────
function fmtDate(s) {
  if (!s) return '';
  const d=new Date(s),now=new Date(),diff=now-d;
  if (isNaN(d)) return '';
  if (diff<3600000){const m=Math.round(diff/60000);return m<1?'just now':`${m}m`;}
  if (diff<86400000) return d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
  if (diff<604800000) return d.toLocaleDateString([],{weekday:'short'});
  return d.toLocaleDateString([],{month:'short',day:'numeric'});
}
function fmtUntil(s) {
  if (!s) return '';
  const d=new Date(s),now=new Date();
  if (isNaN(d)) return '';
  const diff=d-now; // ms until event
  if (diff<=0) return 'now';
  const mins=Math.round(diff/60000);
  if (mins<60) return `in ${mins}m`;
  const hrs=Math.floor(diff/3600000);
  const rem=Math.round((diff%3600000)/60000);
  if (hrs<24) return rem>0?`in ${hrs}h ${rem}m`:`in ${hrs}h`;
  const days=Math.floor(diff/86400000);
  const time=d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
  if (days===0) return `today ${time}`;
  if (days===1) return `tomorrow ${time}`;
  const dow=d.toLocaleDateString([],{weekday:'short'});
  return `${dow} ${time}`;
}
function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

document.querySelectorAll('.modal-overlay').forEach(m=>
  m.addEventListener('click',e=>{if(e.target===m)closeModals();}));

// ── Tab switching ──────────────────────────────────────────────────────────────
let activeTab = 'mailbox';
function switchTab(tab) {
  activeTab = tab;
  clearSearch();
  if (state.triageView) { state.triageView = false; document.removeEventListener('keydown', _triageKeydown); }
  if (tab !== 'mailbox') _mboxUnregisterKeys();
  // Calendar sidebar button highlight
  const calBtn = document.getElementById('nav-calendar');
  if (calBtn) calBtn.classList.toggle('active', tab==='calendar');
  // Title bar actions always visible
  document.getElementById('triage-actions').style.display = 'flex';
  // Sidebar always visible
  document.getElementById('sidebar').style.display = '';
  document.getElementById('resize-handle').style.display = '';
  // Hide all right-pane views
  ['empty-pane','thread-detail','triage-pane','mailbox-pane','calendar-pane','search-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  // Show tab-specific view
  if (tab === 'email') {
    const hasThread = !!document.getElementById('thread-detail').dataset.loaded;
    document.getElementById(hasThread ? 'thread-detail' : 'empty-pane').style.display = '';
    initMailbox();
  } else if (tab === 'mailbox') {
    document.getElementById('mailbox-pane').style.display = 'flex';
    initMailbox();
  } else if (tab === 'calendar') {
    document.getElementById('calendar-pane').style.display = 'flex';
    renderCalendar();
  }
}

// ── Mailbox ────────────────────────────────────────────────────────────────────
let mailboxFolderLoaded = false;
async function initMailbox() {
  if (mailboxFolderLoaded) return;
  mailboxFolderLoaded = true;
  const r = await fetch('/api/mailbox/folders').then(r=>r.json()).catch(()=>null);
  if (!r) return;
  const tree = document.getElementById('folder-tree');
  const fItem = (f, path) => {
    const cnt = f.count ? `<span class="folder-item-count">${f.count.toLocaleString()}</span>` : '';
    return `<div class="folder-item" data-folder="${esc(path)}" onclick="selectMailboxFolder('${esc(path)}',this)"><span>${f.icon}</span><span class="folder-item-name">${esc(f.name)}</span>${cnt}</div>`;
  };
  tree.innerHTML = r.folders.map(f => {
    if (f.children && f.children.length) {
      return `<div>
        <div class="folder-group-hdr" onclick="this.classList.toggle('open');this.nextElementSibling.classList.toggle('open')">
          <span>${f.icon}</span><span class="folder-item-name">${esc(f.name)}</span>
          <span class="folder-group-chevron">▾</span>
        </div>
        <div class="folder-group-children">
          ${f.children.map(c=>fItem(c, c.path||c.name)).join('')}
        </div>
      </div>`;
    }
    return fItem(f, f.name);
  }).join('');
}

let mailboxCurrentFolder = null;
async function selectMailboxFolder(folder, el) {
  mailboxCurrentFolder = folder;
  // Always switch to mailbox tab — ensures triage/calendar/other panes are hidden
  if (activeTab !== 'mailbox' || state.triageView) switchTab('mailbox');
  document.querySelectorAll('.folder-item.active').forEach(e=>e.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('mailbox-folder-name').textContent = folder.split('/').pop();
  document.getElementById('mailbox-folder-count').textContent = '';
  document.getElementById('mailbox-list').innerHTML = '<div class="mailbox-empty">Loading…</div>';
  // Ensure mailbox pane is shown (user might have drilled into a thread)
  document.getElementById('mailbox-pane').style.display = 'flex';
  document.getElementById('thread-detail').style.display = 'none';
  const r = await fetch(`/api/mailbox/folder?folder=${encodeURIComponent(folder)}`).then(r=>r.json()).catch(()=>null);
  if (!r) { document.getElementById('mailbox-list').innerHTML = '<div class="mailbox-empty">Error loading folder</div>'; return; }
  document.getElementById('mailbox-folder-count').textContent = `${r.total} thread${r.total!==1?'s':''}`;
  const list = document.getElementById('mailbox-list');
  if (!r.threads.length) { list.innerHTML = '<div class="mailbox-empty">No messages</div>'; return; }
  list.innerHTML = r.threads.map(t => {
    const ck = esc(t.conversationKey);
    const fl = esc(folder);
    return `<div class="mbox-row${!t.isRead?' unread':''}" data-key="${ck}" data-folder="${fl}"
        onclick="openMailboxThread(this.dataset.key, this.dataset.folder)">
      ${!t.isRead ? '<div class="mbox-dot"></div>' : '<div class="mbox-dot-empty"></div>'}
      <div class="mbox-body">
        <div class="mbox-subj">${esc(t.subject)}</div>
        <div class="mbox-meta">
          <span class="mbox-from">${esc(t.fromName||t.fromAddress)}</span>
          <span class="mbox-date">${esc(fmtDate(t.date))}</span>
        </div>
        <div class="mbox-preview">${esc(t.preview)}</div>
      </div>
      ${t.messageCount>1?`<div class="mbox-cnt">${t.messageCount}</div>`:''}
      <div class="mbox-actions" onclick="event.stopPropagation()">
        <button class="mbox-act-btn mbox-act-reply" onclick="mboxQuickReply(this.closest('.mbox-row').dataset.key, this.closest('.mbox-row').dataset.folder)">↩ Reply</button>
        <button class="mbox-act-btn mbox-act-del" onclick="mboxQuickDelete(this.closest('.mbox-row').dataset.key)">🗑</button>
      </div>
    </div>`;
  }).join('');
  _mboxFocusIdx = -1;
  _mboxRegisterKeys();
}

let _mboxLoadSeq = 0;
async function openMailboxThread(convKey, folder) {
  const seq = ++_mboxLoadSeq;
  document.querySelectorAll('.mbox-row.active').forEach(e=>e.classList.remove('active'));
  try { document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`)?.classList.add('active'); } catch(e){}
  state.selectedKey = convKey;
  state.mailboxContext = true;
  state.expandedMsgs = new Set();
  state.currentMsgs = [];
  const hdrEl = document.getElementById('thread-hdr');
  const msgsEl = document.getElementById('msgs-section');
  // Hide everything else, show thread-detail
  ['empty-pane','triage-pane','mailbox-pane','search-pane','calendar-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  document.getElementById('thread-detail').style.display = 'flex';
  document.getElementById('thread-detail').dataset.loaded = '1';
  hdrEl.innerHTML = '<div style="padding:20px"><div class="spinner"></div></div>';
  msgsEl.innerHTML = '';
  const r = await fetch(`/api/thread_messages?conversationKey=${encodeURIComponent(convKey)}`).then(r=>r.json()).catch(()=>null);
  // Stale response — another thread was clicked while this was loading
  if (seq !== _mboxLoadSeq) return;
  if (!r) {
    hdrEl.innerHTML = `<div style="padding:14px 22px 10px;display:flex;justify-content:space-between;align-items:center">
      <button class="mbox-back" onclick="backToMailboxList()">✕ Close</button>
      <span style="color:#8b949e">Error loading messages</span></div>`;
    return;
  }
  state.currentMsgs = r.messages || [];
  // newest-first; latest for latestId is index 0
  const latestMsg = state.currentMsgs[0] || {};
  const thread = state.threadMap[convKey];
  if (thread) {
    _renderThreadHdr(thread);
  } else {
    // Build a minimal thread object so we always show action buttons
    const syntheticThread = {
      conversationKey: convKey,
      subject: latestMsg.subject || '(No subject)',
      latestId: latestMsg.id || '',
      emailIds: state.currentMsgs.map(m=>m.id),
      messageCount: state.currentMsgs.length,
      participants: [...new Set(state.currentMsgs.map(m=>m.from_name||m.from_address).filter(Boolean))],
      urgency: 'low', action: 'read',
      summary: '', suggestedReply: '', suggestedFolder: folder,
      hasUnread: state.currentMsgs.some(m=>!m.is_read),
    };
    state.threadMap[convKey] = syntheticThread;
    _renderThreadHdr(syntheticThread);
    // Trigger analysis in background so summary populates
    fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: convKey})}).then(r=>r.json()).then(d=>{
      if (d && d.reply) {
        const t = state.threadMap[convKey];
        if (t) { t.suggestedReply = d.reply; }
      }
    }).catch(()=>{});
  }
  const msgs = state.currentMsgs;
  msgsEl.innerHTML = msgs.map((m,i)=>_msgCardHTML(m,i)).join('');
}

function backToMailboxList() {
  state.mailboxContext = false;
  document.getElementById('thread-detail').style.display = 'none';
  document.getElementById('thread-detail').dataset.loaded = '';
  document.getElementById('mailbox-pane').style.display = 'flex';
  // Restore keyboard focus to the row that was open
  if (state.selectedKey) {
    const rows = _mboxGetRows();
    const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
    _mboxSetFocus(idx >= 0 ? idx : 0);
  }
  _mboxRegisterKeys();
}

// ── Mailbox keyboard navigation ─────────────────────────────────────────────────
let _mboxFocusIdx = -1;

function _mboxGetRows() {
  return [...document.querySelectorAll('#mailbox-list .mbox-row')];
}

function _mboxSetFocus(idx, scroll=true) {
  document.querySelectorAll('.mbox-row.focused').forEach(r=>r.classList.remove('focused'));
  const rows = _mboxGetRows();
  if (idx < 0 || idx >= rows.length) { _mboxFocusIdx = -1; return; }
  _mboxFocusIdx = idx;
  rows[idx].classList.add('focused');
  if (scroll) rows[idx].scrollIntoView({block:'nearest', behavior:'smooth'});
}

function _mboxRegisterKeys() {
  document.removeEventListener('keydown', _mboxKeydown);
  document.addEventListener('keydown', _mboxKeydown);
}
function _mboxUnregisterKeys() {
  document.removeEventListener('keydown', _mboxKeydown);
}

function _mboxKeydown(e) {
  // Ignore when typing in inputs or a modal is open
  if (['INPUT','TEXTAREA','SELECT'].includes(e.target.tagName)) return;
  if (document.querySelector('.modal-overlay.open')) return;

  const listEl = document.getElementById('mailbox-pane');
  const threadEl = document.getElementById('thread-detail');
  const inList = listEl && listEl.style.display !== 'none';
  const inThread = threadEl && threadEl.style.display !== 'none' && threadEl.dataset.loaded;

  if (!inList && !inThread) return;

  const rows = _mboxGetRows();

  if (inList) {
    switch(e.key) {
      case 'ArrowDown': case 'j':
        e.preventDefault();
        _mboxSetFocus(_mboxFocusIdx < 0 ? 0 : Math.min(_mboxFocusIdx + 1, rows.length - 1));
        return;
      case 'ArrowUp': case 'k':
        e.preventDefault();
        _mboxSetFocus(_mboxFocusIdx <= 0 ? 0 : _mboxFocusIdx - 1);
        return;
      case 'Enter': case 'ArrowRight': {
        e.preventDefault();
        const row = rows[_mboxFocusIdx];
        if (row) openMailboxThread(row.dataset.key, row.dataset.folder || mailboxCurrentFolder);
        return;
      }
      case 'r': {
        const row = rows[_mboxFocusIdx];
        if (row) mboxQuickReply(row.dataset.key, row.dataset.folder || mailboxCurrentFolder);
        return;
      }
      case 'd': {
        e.preventDefault();
        const row = rows[_mboxFocusIdx];
        if (row) {
          const nextIdx = Math.min(_mboxFocusIdx, rows.length - 2);
          mboxQuickDelete(row.dataset.key).then(()=>setTimeout(()=>_mboxSetFocus(nextIdx), 50));
        }
        return;
      }
      case 'f': {
        const row = rows[_mboxFocusIdx];
        if (row) {
          const thread = state.threadMap[row.dataset.key];
          if (thread) openFile(encodeThread(thread));
        }
        return;
      }
      case 'u': {
        const row = rows[_mboxFocusIdx];
        if (row) _mboxToggleRead(row.dataset.key);
        return;
      }
    }
  }

  if (inThread) {
    switch(e.key) {
      case 'Escape': case 'ArrowLeft':
        e.preventDefault();
        backToMailboxList();
        return;
      case 'ArrowDown': case 'j': {
        e.preventDefault();
        const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
        if (idx < rows.length - 1) {
          _mboxFocusIdx = idx + 1;
          openMailboxThread(rows[idx+1].dataset.key, rows[idx+1].dataset.folder || mailboxCurrentFolder);
        }
        return;
      }
      case 'ArrowUp': case 'k': {
        e.preventDefault();
        const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
        if (idx > 0) {
          _mboxFocusIdx = idx - 1;
          openMailboxThread(rows[idx-1].dataset.key, rows[idx-1].dataset.folder || mailboxCurrentFolder);
        }
        return;
      }
      case 'r': {
        const thread = state.threadMap[state.selectedKey];
        if (thread) openReply(encodeThread(thread));
        return;
      }
      case 'd': {
        e.preventDefault();
        if (state.selectedKey) mboxQuickDelete(state.selectedKey);
        return;
      }
      case 'f': {
        const thread = state.threadMap[state.selectedKey];
        if (thread) openFile(encodeThread(thread));
        return;
      }
    }
  }
}

async function _mboxToggleRead(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  const markRead = thread.hasUnread;
  await fetch('/api/mark_read', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conversationKey: convKey, read: markRead})}).catch(()=>{});
  thread.hasUnread = !markRead;
  if (markRead) thread.isRead = true;
  // Update unread dot on the row
  const row = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  if (row) {
    row.classList.toggle('unread', !markRead);
    const dot = row.querySelector('.mbox-dot, .mbox-dot-empty');
    if (dot) { dot.className = !markRead ? 'mbox-dot' : 'mbox-dot-empty'; }
  }
}

async function mboxQuickReply(convKey, folder) {
  await openMailboxThread(convKey, folder);
  const thread = state.threadMap[convKey];
  if (thread) setTimeout(() => openReply(encodeThread(thread)), 300);
}

async function mboxQuickDelete(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  // Find next row before removing
  const row = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  const nextRow = row?.nextElementSibling;
  // Animate out
  if (row) { row.style.opacity='0'; row.style.transition='opacity .15s'; }
  await fetch('/api/delete', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({ids: thread.emailIds, conversationKey: convKey})});
  if (row) row.remove();
  delete state.threadMap[convKey];
  for (const g of state.groups) g.threads = g.threads.filter(t=>t.conversationKey!==convKey);
  state.groups = state.groups.filter(g=>g.threads.length>0);
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
  // Update folder count
  const countEl = document.getElementById('mailbox-folder-count');
  if (countEl) {
    const cur = parseInt(countEl.textContent) || 0;
    countEl.textContent = `${Math.max(0,cur-1)} threads`;
  }
  // If we were viewing this thread, open the next one or go back
  if (state.selectedKey === convKey) {
    if (nextRow && nextRow.dataset.key) {
      openMailboxThread(nextRow.dataset.key, nextRow.dataset.folder || mailboxCurrentFolder);
    } else {
      backToMailboxList();
    }
  }
}

// ── Search ─────────────────────────────────────────────────────────────────────
let searchTimeout = null;
async function doSearch(q) {
  q = (q||'').trim();
  if (q.length < 2) return;
  // Hide all right-pane content, show search pane
  ['empty-pane','thread-detail','triage-pane','mailbox-pane','calendar-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  const pane = document.getElementById('search-pane');
  pane.style.display = 'flex';
  pane.style.flexDirection = 'column';
  document.getElementById('search-hdr').textContent = `Searching for "${q}"…`;
  document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">Searching…</div>';
  const r = await fetch(`/api/search?q=${encodeURIComponent(q)}`).then(r=>r.json()).catch(()=>null);
  if (!r) { document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">Error</div>'; return; }
  document.getElementById('search-hdr').textContent = `${r.count} result${r.count!==1?'s':''} for "${q}"`;
  if (!r.results.length) { document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">No results</div>'; return; }
  document.getElementById('search-results').innerHTML = r.results.map(e => `
    <div class="search-row" onclick="openSearchResult('${esc(e.conversation_key)}','${esc(e.folder||'')}','${esc(e.id)}')">
      <div class="search-row-body">
        <div class="search-row-subj">${esc(e.subject||'(No subject)')}</div>
        <div class="search-row-meta">${esc(e.from_name||e.from_address||'')} · ${esc(fmtDate((e.received_date_time||'').slice(0,19)))}</div>
        <div class="search-row-preview">${esc(e.body_preview||'')}</div>
      </div>
      <div class="search-row-folder">${esc(e.folder||'Unknown')}</div>
    </div>`).join('');
}

function clearSearch() {
  document.getElementById('search-pane').style.display = 'none';
}

async function openSearchResult(convKey, folder, emailId) {
  document.getElementById('search-pane').style.display = 'none';
  const inInbox = (folder||'').toLowerCase() === 'inbox';
  if (inInbox && state.threadMap[convKey]) {
    switchTab('email');
    selectThread(convKey);
  } else {
    switchTab('mailbox');
    // Small delay to let mailbox init
    await initMailbox();
    // Pre-highlight the folder and load the thread
    const folderEl = document.querySelector(`.folder-item[data-folder="${CSS.escape(folder)}"]`);
    await openMailboxThread(convKey, folder);
  }
}

// ── Calendar ───────────────────────────────────────────────────────────────────
const CAL_COLORS = [
  ['#1f6feb','#58a6ff'],['#388bfd22','#79c0ff'],['#1a7f37','#3fb950'],
  ['#6e40c9','#bc8cff'],['#b45309','#d18616'],['#0e7490','#06b6d4'],
  ['#9a1c1c','#f85149'],
];
let calWeekOffset = 0;
let calDayOffset  = 0;   // days from today (day view)
let calViewMode   = 'day';
let calEvents     = [];
let calPrepCache  = {};  // eventId -> {headsup, topics}

function calGetWeekStart(offset) {
  const d = new Date();
  const day = d.getDay();
  const mon = new Date(d);
  mon.setDate(d.getDate() - ((day+6)%7) + offset*7);
  mon.setHours(0,0,0,0);
  return mon;
}
function calSetView(mode) {
  calViewMode = mode;
  document.getElementById('cal-view-day')?.classList.toggle('active', mode==='day');
  document.getElementById('cal-view-week')?.classList.toggle('active', mode==='week');
  renderCalendar();
}
function calMove(dir) {
  if (calViewMode==='day') calDayOffset += dir; else calWeekOffset += dir;
  renderCalendar();
}
function calGoToday() { calWeekOffset = 0; calDayOffset = 0; renderCalendar(); }

async function renderCalendar() {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const DOW_LONG = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  let startISO, endISO, title;

  if (calViewMode === 'day') {
    const base = new Date(); base.setHours(0,0,0,0);
    const day = new Date(base); day.setDate(base.getDate() + calDayOffset);
    const next = new Date(day); next.setDate(day.getDate() + 1);
    startISO = day.toISOString().slice(0,19);
    endISO   = next.toISOString().slice(0,19);
    title = `${DOW_LONG[day.getDay()]}, ${months[day.getMonth()]} ${day.getDate()}, ${day.getFullYear()}`;
  } else {
    const weekStart = calGetWeekStart(calWeekOffset);
    const weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate()+7);
    startISO = weekStart.toISOString().slice(0,19);
    endISO   = weekEnd.toISOString().slice(0,19);
    const s = weekStart, e = new Date(weekEnd); e.setDate(e.getDate()-1);
    title = s.getMonth()===e.getMonth()
      ? `${months[s.getMonth()]} ${s.getDate()} – ${e.getDate()}, ${s.getFullYear()}`
      : `${months[s.getMonth()]} ${s.getDate()} – ${months[e.getMonth()]} ${e.getDate()}, ${s.getFullYear()}`;
  }
  document.getElementById('cal-title').textContent = title;

  document.getElementById('cal-loading').style.display = 'flex';
  document.getElementById('cal-scroll-wrap').style.display = 'none';
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(startISO)}&end=${encodeURIComponent(endISO)}`);
    const d = await r.json();
    calEvents = d.events || [];
  } catch(e) { calEvents = []; }
  document.getElementById('cal-loading').style.display = 'none';
  document.getElementById('cal-scroll-wrap').style.display = '';

  const grid = document.getElementById('cal-grid');
  if (calViewMode === 'day') {
    grid.style.gridTemplateColumns = '52px 1fr';
    buildDayView();
  } else {
    grid.style.gridTemplateColumns = '52px repeat(7,minmax(0,1fr))';
    buildCalGrid(calGetWeekStart(calWeekOffset));
  }
}

function buildCalGrid(weekStart) {
  const grid = document.getElementById('cal-grid');
  const today = new Date(); today.setHours(0,0,0,0);
  const days = Array.from({length:7}, (_,i) => { const d=new Date(weekStart); d.setDate(weekStart.getDate()+i); return d; });
  const SLOT_H = 24; // px per 30-min slot
  const HOUR_START = 7, HOUR_END = 21; // visible hours
  const SLOTS = (HOUR_END - HOUR_START) * 2;

  // Separate all-day events
  const allDayEvs = calEvents.filter(ev => {
    const st = ev.start_time || '';
    return st.length <= 10 || /T00:00:00/.test(st) && /T00:00:00/.test(ev.end_time||'');
  });
  const timedEvs = calEvents.filter(ev => !allDayEvs.includes(ev));

  // Assign colors per event title hash
  function evColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length];
  }

  let html = '';

  // Corner + day headers
  html += `<div class="cal-hdr-corner"></div>`;
  days.forEach((d,i) => {
    const isToday = d.getTime()===today.getTime();
    const dow = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][i];
    html += `<div class="cal-hdr-cell${isToday?' today':''}">
      <div class="cal-hdr-dow">${dow}</div>
      <div class="cal-hdr-day">${d.getDate()}</div>
    </div>`;
  });

  // All-day row
  html += `<div style="font-size:9px;color:#5ba4cf;text-align:right;padding:2px 4px 2px 0;border-right:1px solid #1a3252;border-bottom:2px solid #1a3252;">all day</div>`;
  days.forEach((d,i) => {
    const isToday = d.getTime()===today.getTime();
    const dayStr = d.toISOString().slice(0,10);
    const evs = allDayEvs.filter(ev => (ev.start_time||'').startsWith(dayStr));
    html += `<div class="cal-all-day-cell${isToday?' today-col':''}">`;
    evs.forEach(ev => { html += `<div class="cal-all-day-event" title="${esc(ev.subject)}">${esc(ev.subject)}</div>`; });
    html += `</div>`;
  });

  // Time slots
  for (let slot=0; slot<SLOTS; slot++) {
    const totalMins = (HOUR_START * 60) + slot * 30;
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const isHour = m===0;
    if (isHour) {
      const label = h===12?'12pm':h>12?`${h-12}pm`:`${h}am`;
      html += `<div class="cal-time-label" style="height:${SLOT_H}px;${isHour?'':'border-top:none'}">${label}</div>`;
    } else {
      html += `<div class="cal-time-label" style="height:${SLOT_H}px;border-top:none"></div>`;
    }
    days.forEach((d,di) => {
      const isToday = d.getTime()===today.getTime();
      html += `<div class="cal-cell${isToday?' today-col':''}${isHour?' hour-start':''}" style="height:${SLOT_H}px"></div>`;
    });
  }

  grid.innerHTML = html;

  // Position timed events as absolute overlays
  // We need to position them inside the correct cell after render
  // Use a post-render approach: collect cells by [day][slot]
  const cells = grid.querySelectorAll('.cal-cell');
  const cellMap = {}; // "dayIdx-slot" -> cell el
  let ci = 0;
  for (let slot=0; slot<SLOTS; slot++) {
    for (let di=0; di<7; di++) {
      cellMap[`${di}-${slot}`] = cells[ci++];
    }
  }

  // For each timed event, find its day col and slot range
  timedEvs.forEach((ev, idx) => {
    const st = new Date(ev.start_time);
    const et = new Date(ev.end_time || ev.start_time);
    const evDay = new Date(st); evDay.setHours(0,0,0,0);
    const di = days.findIndex(d => d.getTime()===evDay.getTime());
    if (di < 0) return;

    const startMins = st.getHours()*60 + st.getMinutes();
    const endMins = et.getHours()*60 + et.getMinutes() || startMins + 30;
    const startSlot = Math.max(0, Math.floor((startMins - HOUR_START*60)/30));
    const durationSlots = Math.max(1, Math.ceil((endMins - startMins)/30));

    const anchorCell = cellMap[`${di}-${startSlot}`];
    if (!anchorCell) return;

    const topOffset = ((startMins - HOUR_START*60) % 30) / 30 * SLOT_H;
    const height = Math.max(SLOT_H-2, durationSlots * SLOT_H - 2);
    const [bg, fg] = evColor(ev.subject||'');
    const timeStr = st.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});

    const el = document.createElement('div');
    el.className = 'cal-event';
    el.style.cssText = `top:${topOffset}px;height:${height}px;background:${bg}33;border:1px solid ${bg}88;color:${fg};`;
    el.title = `${ev.subject}\\n${timeStr}${ev.location?' · '+ev.location:''}`;
    const timeEl = durationSlots > 1 ? `<div class="cal-event-time">${timeStr}</div>` : '';
    el.innerHTML = `<div class="cal-event-title">${esc(ev.subject||'(No title)')}</div>${timeEl}`;
    anchorCell.style.position = 'relative';
    anchorCell.appendChild(el);
  });

  // Scroll to 8am (1hr from HOUR_START=7 = 2 slots)
  document.getElementById('cal-scroll-wrap').scrollTop = 2 * SLOT_H;
}

function buildDayView() {
  const grid = document.getElementById('cal-grid');
  const base = new Date(); base.setHours(0,0,0,0);
  const day  = new Date(base); day.setDate(base.getDate() + calDayOffset);
  const today = new Date(); today.setHours(0,0,0,0);
  const isToday = day.getTime() === today.getTime();
  const SLOT_H = 32;
  const HOUR_START = 7, HOUR_END = 21;
  const SLOTS = (HOUR_END - HOUR_START) * 2;
  const dayStr = day.toISOString().slice(0,10);

  function evColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length];
  }

  const allDayEvs = calEvents.filter(ev => {
    const st = ev.start_time||'';
    return st.length<=10 || (/T00:00:00/.test(st) && /T00:00:00/.test(ev.end_time||''));
  });
  const timedEvs = calEvents.filter(ev => !allDayEvs.includes(ev));

  const dow = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][day.getDay()];
  let html = `<div class="cal-hdr-corner"></div>
    <div class="cal-hdr-cell${isToday?' today':''}">
      <div class="cal-hdr-dow">${dow}</div>
      <div class="cal-hdr-day">${day.getDate()}</div>
    </div>
    <div style="font-size:9px;color:#5ba4cf;text-align:right;padding:2px 4px 2px 0;border-right:1px solid #1a3252;border-bottom:2px solid #1a3252;">all day</div>
    <div class="cal-all-day-cell${isToday?' today-col':''}">`;
  allDayEvs.filter(ev=>(ev.start_time||'').startsWith(dayStr)).forEach(ev=>{
    html+=`<div class="cal-all-day-event">${esc(ev.subject)}</div>`;
  });
  html += `</div>`;

  for (let slot=0; slot<SLOTS; slot++) {
    const totalMins = (HOUR_START*60) + slot*30;
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const isHour = m===0;
    const label = isHour ? (h===12?'12pm':h>12?`${h-12}pm`:`${h}am`) : '';
    html += `<div class="cal-time-label" style="height:${SLOT_H}px;${isHour?'':'border-top:none'}">${label}</div>`;
    html += `<div class="cal-cell cal-day-cell${isToday?' today-col':''}${isHour?' hour-start':''}" style="height:${SLOT_H}px"></div>`;
  }
  grid.innerHTML = html;

  const cells = Array.from(grid.querySelectorAll('.cal-day-cell'));

  timedEvs.forEach(ev => {
    const st = new Date(ev.start_time);
    const et = new Date(ev.end_time || ev.start_time);
    const startMins = st.getHours()*60 + st.getMinutes();
    const endMins   = et.getHours()*60 + et.getMinutes() || startMins+30;
    const startSlot = Math.max(0, Math.floor((startMins - HOUR_START*60)/30));
    const durSlots  = Math.max(2, Math.ceil((endMins - startMins)/30));
    const anchor    = cells[startSlot];
    if (!anchor) return;

    const topOffset = ((startMins - HOUR_START*60) % 30) / 30 * SLOT_H;
    const height    = Math.max(SLOT_H*2-2, durSlots * SLOT_H - 2);
    const [bg, fg]  = evColor(ev.subject||'');
    const timeStr   = st.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const endStr    = et.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const prepId    = 'prep-'+ev.id.replace(/[^a-zA-Z0-9]/g,'_');

    const el = document.createElement('div');
    el.className = 'cal-event cal-day-event';
    el.style.cssText = `top:${topOffset}px;height:${height}px;`+
      `background:${bg}44;border-left:3px solid ${bg};`+
      `border-top:1px solid ${bg}88;border-right:1px solid ${bg}44;`+
      `border-bottom:1px solid ${bg}44;color:${fg};padding:5px 8px;`;
    el.innerHTML = `<div class="cal-day-ev-hdr">
        <div class="cal-day-ev-title">${esc(ev.subject||'(No title)')}</div>
        <div class="cal-day-ev-time">${timeStr}–${endStr}${ev.location?' · '+esc(ev.location):''}</div>
      </div>`;
    anchor.style.position = 'relative';
    anchor.appendChild(el);
  });

  document.getElementById('cal-scroll-wrap').scrollTop = 2 * SLOT_H;
}

async function loadMeetingPrep(ev, prepId) {
  const cacheKey = ev.id;
  if (calPrepCache[cacheKey]) { _renderPrep(prepId, calPrepCache[cacheKey]); return; }
  try {
    const r = await fetch('/api/meeting_prep', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({subject:ev.subject, attendees:ev.attendees,
        start_time:ev.start_time, end_time:ev.end_time, location:ev.location})
    }).then(r=>r.json());
    if (r.ok) { calPrepCache[cacheKey]=r; _renderPrep(prepId, r); }
    else { const el=document.getElementById(prepId); if(el) el.innerHTML=''; }
  } catch(e) { const el=document.getElementById(prepId); if(el) el.innerHTML=''; }
}
function _renderPrep(prepId, prep) {
  const el = document.getElementById(prepId);
  if (!el) return;
  const topics = (prep.topics||[]).map(t=>`<div class="cal-prep-topic">• ${esc(t)}</div>`).join('');
  el.innerHTML = `<div class="cal-prep-headsup">${esc(prep.headsup||'')}</div>`+
    (topics ? `<div class="cal-prep-topics">${topics}</div>` : '');
}

// ── Today sidebar widget ───────────────────────────────────────────────────────
async function renderTodayCal() {
  const today = new Date();
  const start = new Date(today); start.setHours(0,0,0,0);
  const end   = new Date(today); end.setHours(23,59,59,0);
  const startISO = start.toISOString().slice(0,19);
  const endISO   = end.toISOString().slice(0,19);
  let events = [];
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(startISO)}&end=${encodeURIComponent(endISO)}`);
    const d = await r.json();
    events = (d.events||[]).filter(ev => {
      // exclude all-day (no time component or midnight-to-midnight)
      const st = ev.start_time||'';
      return st.length > 10 && !/T00:00:00/.test(st);
    });
  } catch(e) {}
  const list = document.getElementById('today-cal-list');
  if (!list) return;
  if (!events.length) { list.innerHTML = '<div class="today-cal-empty">No meetings today</div>'; return; }
  function evDotColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length][1];
  }
  list.innerHTML = events.map(ev => {
    const st = new Date(ev.start_time);
    const timeStr = st.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const color = evDotColor(ev.subject||'');
    const past = st < today;
    return `<div class="today-ev" style="${past?'opacity:.45':''}">
      <div class="today-ev-dot" style="background:${color}"></div>
      <span class="today-ev-time">${timeStr}</span>
      <span class="today-ev-title" title="${esc(ev.subject)}">${esc(ev.subject||'(No title)')}</span>
    </div>`;
  }).join('');
  // Also update week hours
  updateWeekHours();
}

async function updateWeekHours() {
  const today = new Date();
  const dow = today.getDay();
  const mon = new Date(today); mon.setDate(today.getDate() - ((dow+6)%7)); mon.setHours(0,0,0,0);
  const sat = new Date(mon); sat.setDate(mon.getDate()+5);
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(mon.toISOString().slice(0,19))}&end=${encodeURIComponent(sat.toISOString().slice(0,19))}`);
    const d = await r.json();
    const evs = (d.events||[]).filter(ev => {
      const st=ev.start_time||''; return st.length>10 && !/T00:00:00/.test(st);
    });
    let mins = 0;
    evs.forEach(ev => {
      const st=new Date(ev.start_time), et=new Date(ev.end_time||ev.start_time);
      mins += Math.max(0, (et-st)/60000);
    });
    const hrs = mins/60;
    const hrsStr = Number.isInteger(hrs) ? hrs.toString() : hrs.toFixed(1);
    const el = document.getElementById('week-hours-line');
    if (el) el.textContent = `${hrsStr}h in meetings this week`;
  } catch(e) {}
}

// ── Resizable sidebar ──────────────────────────────────────────────────────────
(function(){
  const handle=document.getElementById('resize-handle');
  const sidebar=document.getElementById('sidebar');
  let dragging=false,startX=0,startW=0;
  handle.addEventListener('mousedown',e=>{
    dragging=true;startX=e.clientX;startW=sidebar.offsetWidth;
    handle.classList.add('dragging');
    document.body.style.cursor='ew-resize';
    document.body.style.userSelect='none';
    e.preventDefault();
  });
  document.addEventListener('mousemove',e=>{
    if(!dragging)return;
    const w=Math.max(160,Math.min(520,startW+(e.clientX-startX)));
    sidebar.style.width=w+'px';
  });
  document.addEventListener('mouseup',()=>{
    if(!dragging)return;
    dragging=false;handle.classList.remove('dragging');
    document.body.style.cursor='';document.body.style.userSelect='';
  });
})();

init();
</script>
</body>
</html>"""

# ─── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not ANTHROPIC_API_KEY:
        print("\n⚠  ANTHROPIC_API_KEY not set in .env\n")

    init_db()
    print(f"\n📧 Email Triage starting at http://localhost:{PORT}")

    _session_ready.wait(timeout=20)
    print("   Outlook MCP connected.")

    # Seed folder list immediately
    try:
        fdrs = call_tool("outlook_mail_list_folders", {})
        raw = fdrs.get("folders", fdrs.get("value", [])) if isinstance(fdrs, dict) else []
        meta_set("folders_raw", json.dumps(raw))
        efforts, other = _folder_lists(raw)
        print(f"   Folders loaded: {len(efforts)} Efforts, {len(other)} other")
    except Exception as e:
        print(f"   Warning: could not load folders: {e}")

    # Start background sync loop
    threading.Thread(target=_sync_loop, daemon=True).start()
    print("   Background sync started (every 5 min)\n")

    threading.Timer(2.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    app.run(port=PORT, debug=False, use_reloader=False)
