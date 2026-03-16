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
from datetime import datetime, timezone

import anthropic
from flask import Flask, jsonify, render_template_string, request
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# ─── Config ────────────────────────────────────────────────────────────────────

MCP_COMMAND     = "/opt/homebrew/bin/McpOutlookLocal"
DB_PATH         = os.path.join(os.path.dirname(__file__), "email_triage.db")
PORT            = 5001
SYNC_INTERVAL   = 300   # 5 minutes
INBOX_FETCH     = 100
ANALYSIS_MODEL  = "claude-haiku-4-5-20251001"
REPLY_MODEL     = "claude-sonnet-4-6"

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
    CREATE INDEX IF NOT EXISTS idx_emails_conv_key ON emails(conversation_key);
    CREATE INDEX IF NOT EXISTS idx_threads_updated  ON threads(updated_at);
    CREATE INDEX IF NOT EXISTS idx_threads_urgency  ON threads(urgency);
    """)
    db.commit()
    # Migrations: add columns if not present (idempotent)
    for migration in [
        "ALTER TABLE emails ADD COLUMN formatted_body TEXT",
        "ALTER TABLE threads ADD COLUMN is_flagged INTEGER DEFAULT 0",
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


def analyze_thread(emails: list, efforts_folders: list, other_folders: list) -> dict:
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
        f"{_clean(e.get('body_preview','(no preview)'), 500)}"
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
  "summary": "3-5 sentences: what this thread is about, who said what, current status, open action items. Use names and specific details.",
  "topic": "broad category label (e.g. Engineering, Product Planning, Finance, Incidents & Outages, Team & HR, Partnerships, FYI & Updates, Strategy & Leadership)",
  "action": "reply OR delete OR file OR read OR done",
  "urgency": "high OR medium OR low",
  "suggestedReply": "complete draft reply or empty string only if deleting",
  "suggestedFolder": "exact folder name or empty string"
}}"""

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


def _format_message_with_ai(msg: dict) -> list:
    """Format a single message into AI-annotated paragraphs with intent + fact-check."""
    body = msg.get("body") or msg.get("body_preview") or ""
    if not body.strip():
        return [{"text": "(no content)", "intent": "FYI", "emoji": "📭", "fact_concern": None}]

    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    date = (msg.get("received_date_time") or "")[:10]

    prompt = f"""You are an expert email analyst helping a senior tech leader understand an email.

FROM: {_clean(from_name, 80)}  |  DATE: {date}
EMAIL BODY:
{_clean(body, 4000)}

Break this email into its natural paragraphs. For each paragraph:
1. Provide the exact paragraph text (verbatim)
2. Classify the intent from EXACTLY one of: Status Update | Request | Decision | Question | Action Item | Context | FYI | Warning | Introduction | Closing
3. Choose an appropriate emoji for that intent
4. Fact-check: if the paragraph makes a specific claim that seems incorrect or worth verifying, provide a short concern string (1-2 sentences). Otherwise use null.

Return ONLY valid JSON (no markdown fences):
{{"paragraphs":[{{"text":"...","intent":"...","emoji":"...","fact_concern":null}}]}}"""

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
            raise ValueError(f"No JSON: {raw[:100]}")
        result = json.loads(m.group())
        return result.get("paragraphs", [])
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


def _do_sync():
    _sync_status.update({"phase": "fetching", "progress": "Fetching folder list…", "done": 0, "total": 0})
    efforts, other = _refresh_folders()

    _sync_status["progress"] = "Fetching inbox messages…"
    msgs_result = call_tool("outlook_mail_list_messages", {"folder": "Inbox", "top": INBOX_FETCH})
    emails = msgs_result.get("messages", []) if isinstance(msgs_result, dict) else []
    if not emails:
        return 0, 0

    _sync_status["progress"] = f"Checking {len(emails)} messages for new arrivals…"
    db = get_db()
    ids = [e["id"] for e in emails if e.get("id")]
    if not ids:
        return 0, 0
    placeholders = ",".join("?" * len(ids))
    existing_ids = {
        r[0] for r in db.execute(
            f"SELECT id FROM emails WHERE id IN ({placeholders})", ids
        ).fetchall()
    }
    new_emails = [e for e in emails if e.get("id") and e["id"] not in existing_ids]

    if not new_emails:
        _sync_status["progress"] = "No new messages."
        return 0, 0

    print(f"Sync: {len(new_emails)} new email(s)")
    _sync_status["progress"] = f"Inserting {len(new_emails)} new message(s)…"

    now = _utcnow()
    for e in new_emails:
        ck = _norm_subject(e.get("subject", ""))
        db.execute(
            "INSERT OR IGNORE INTO emails "
            "(id,subject,from_name,from_address,received_date_time,"
            " is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?)",
            (
                e["id"],
                e.get("subject", ""),
                e.get("from_name", ""),
                e.get("from_address", ""),
                e.get("received_date_time", ""),
                1 if e.get("is_read") else 0,
                _clean(e.get("body_preview", ""), 500),
                ck,
                json.dumps(e),
                now,
            )
        )
    db.commit()

    affected_keys = list({_norm_subject(e.get("subject", "")) for e in new_emails})
    threads_updated = 0
    total = len(affected_keys)
    _sync_status.update({"phase": "analyzing", "done": 0, "total": total})

    for idx, ck in enumerate(affected_keys):
        rows = db.execute(
            "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC",
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
        email_ids = [e["id"] for e in thread_emails]
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
                json.dumps(email_ids),
                latest["id"],
                len(thread_emails),
                1 if has_unread else 0,
                latest["received_date_time"],
                _utcnow(),
            )
        )
        db.commit()
        threads_updated += 1
        _sync_status["done"] = idx + 1
        print(f"  ✓ {display_subj!r} → {result.get('action')} [{result.get('urgency')}]")

    _sync_status.update({"phase": "done", "progress": f"Done — {threads_updated} thread(s) updated."})
    return len(new_emails), threads_updated


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

    return jsonify({
        "groups": result,
        "latestTs": latest_ts,
        "threadCount": len(threads),
        "emailCount": db.execute("SELECT COUNT(*) FROM emails").fetchone()[0],
        "syncStatus": {**_sync_status},
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
    return jsonify({
        "threads": threads,
        "latestTs": latest_ts,
        "syncStatus": {**_sync_status},
    })


@app.route("/api/sync_now", methods=["POST"])
def api_sync_now():
    if not _sync_status["running"]:
        threading.Thread(target=run_sync, daemon=True).start()
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


def _normalize_msg(m: dict) -> dict:
    from_name    = m.get("from_name") or ""
    from_address = m.get("from_address") or ""
    received     = m.get("received_date_time") or ""

    raw_html = m.get("body_content") or ""
    if raw_html:
        body_text = re.sub(r'<[^>]+>', ' ', raw_html)
        body_text = re.sub(r'&nbsp;', ' ', body_text)
        body_text = re.sub(r'&#\d+;|&[a-z]+;', ' ', body_text)
        body_text = re.sub(r'[ \t]{2,}', ' ', body_text)
        body_text = re.sub(r'\n{3,}', '\n\n', body_text).strip()
    else:
        body_text = m.get("body_preview") or ""

    to_recips  = _parse_recipients(m.get("to_recipients") or m.get("toRecipients"))
    cc_recips  = _parse_recipients(m.get("cc_recipients") or m.get("ccRecipients"))

    return {
        "id":                 m.get("id", ""),
        "subject":            m.get("subject", ""),
        "from_name":          from_name,
        "from_address":       from_address,
        "received_date_time": received,
        "is_read":            m.get("is_read"),
        "body":               body_text,
        "to_recipients":      to_recips,
        "cc_recipients":      cc_recips,
    }


@app.route("/api/thread_messages")
def api_thread_messages():
    ids = request.args.getlist("id")
    if not ids:
        return jsonify({"messages": []})
    db = get_db()
    rows = db.execute(
        "SELECT * FROM emails WHERE id IN ({})".format(",".join("?" * len(ids))),
        ids
    ).fetchall()
    db_msgs = {r["id"]: dict(r) for r in rows}

    result = []
    for msg_id in ids:
        fallback = db_msgs.get(msg_id, {"id": msg_id})
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
        result.append(msg)

    result.sort(key=lambda m: m.get("received_date_time", ""))
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
def api_markread():
    ids = request.json.get("ids", [])
    conv_key = request.json.get("conversationKey", "")
    try:
        call_tool("outlook_mail_mark_read", {"message_ids": ids, "is_read": True})
        db = get_db()
        id_ph = ",".join("?" * len(ids))
        db.execute(f"UPDATE emails SET is_read=1 WHERE id IN ({id_ph})", ids)
        db.execute("UPDATE threads SET has_unread=0 WHERE conversation_key=?", (conv_key,))
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
.sidebar{width:260px;flex-shrink:0;background:#0a1628;border-right:none;overflow-y:auto;display:flex;flex-direction:column;}
.resize-handle{width:5px;cursor:ew-resize;background:transparent;flex-shrink:0;transition:background .15s;}
.resize-handle:hover,.resize-handle.dragging{background:#58a6ff;}
.sidebar-hdr{padding:12px 14px 6px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.14em;color:#484f58;}
.topic-group{}
.tg-header{display:flex;align-items:center;gap:6px;padding:6px 14px;cursor:pointer;color:#7d8fa3;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;border-top:1px solid #0d2040;user-select:none;}
.tg-header:hover{color:#a8bccc;background:rgba(255,255,255,.03);}
.tg-name{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.tg-count{font-size:10px;background:#0d2040;color:#5a7a9e;border-radius:8px;padding:1px 6px;flex-shrink:0;}
.tg-chevron{font-size:9px;color:#484f58;transition:transform .2s;}
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
.ti-meta{font-size:10px;color:#484f58;margin-top:2px;display:flex;gap:5px;}

/* Right pane */
.right-pane{flex:1;overflow:hidden;display:flex;flex-direction:column;background:#0a1628;border-left:1px solid #1e3d6b;}
#first-load{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;}
.empty-pane{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;color:#484f58;}
.empty-pane .ep-icon{font-size:36px;opacity:.35;}
.empty-pane .ep-txt{font-size:12px;}

/* Thread detail */
.thread-detail{flex:1;display:flex;flex-direction:column;min-height:0;overflow:hidden;}
.thread-hdr{padding:18px 22px 14px;background:linear-gradient(160deg,#0d2040 0%,#122545 100%);border-bottom:1px solid #1e3d6b;flex-shrink:0;}
.th-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px;margin-bottom:8px;}
.th-badges{display:flex;align-items:center;gap:7px;margin-bottom:7px;}
.urg-pill{display:inline-flex;align-items:center;padding:2px 9px;border-radius:10px;font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;}
.urg-high{background:rgba(248,81,73,.15);color:#f85149;border:1px solid rgba(248,81,73,.3);}
.urg-medium{background:rgba(210,153,34,.15);color:#d29922;border:1px solid rgba(210,153,34,.3);}
.urg-low{background:rgba(63,185,80,.12);color:#3fb950;border:1px solid rgba(63,185,80,.25);}
.act-pill{display:inline-flex;align-items:center;padding:2px 9px;border-radius:10px;font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;background:#1a3252;color:#8b949e;border:1px solid #243f65;}
.act-reply{background:rgba(31,111,235,.15);color:#58a6ff;border-color:rgba(88,166,255,.25);}
.act-delete{background:rgba(248,81,73,.12);color:#f85149;border-color:rgba(248,81,73,.25);}
.act-file{background:rgba(63,185,80,.1);color:#3fb950;border-color:rgba(63,185,80,.2);}
.th-subject{font-size:16px;font-weight:700;color:#e6edf3;line-height:1.35;flex:1;letter-spacing:-.3px;}
.th-date{font-size:11px;color:#484f58;flex-shrink:0;margin-top:3px;}
.th-participants{display:flex;align-items:center;gap:8px;margin-bottom:10px;}
.avatars{display:flex;}
.avatar{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;font-size:8px;font-weight:700;color:#fff;border:2px solid #0a1628;margin-right:-4px;flex-shrink:0;}
.th-names{font-size:11px;color:#8b949e;margin-left:8px;}
.th-msgcount{font-size:10px;color:#8b949e;background:#1a3252;border-radius:7px;padding:2px 7px;}
.th-summary{background:rgba(88,166,255,.05);border:1px solid rgba(88,166,255,.12);border-radius:8px;padding:10px 13px;font-size:11.5px;color:#8b949e;line-height:1.7;margin-bottom:11px;}
.th-summary-lbl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#58a6ff;margin-bottom:5px;}
.th-actions{display:flex;gap:7px;flex-wrap:wrap;}

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
.msgs-section{flex:1;min-height:0;overflow-y:auto;padding:0 22px 24px;}
.msgs-label{display:flex;align-items:center;gap:8px;padding:12px 0 10px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.12em;color:#484f58;border-bottom:1px solid #21262d;margin-bottom:10px;}
.msgs-label span{color:#58a6ff;}

/* Message cards */
.msg-card{background:#0d2040;border:1px solid #1a3252;border-radius:9px;margin-bottom:7px;overflow:hidden;transition:border-color .15s;}
.msg-card:hover{border-color:#2a4d7a;}
.msg-card.open{border-color:#2a4d7a;}
.msg-hdr{display:flex;align-items:center;gap:9px;padding:9px 13px;cursor:pointer;user-select:none;}
.msg-hdr:hover{background:rgba(255,255,255,.03);}
.msg-from{font-size:12px;font-weight:600;color:#c9d1d9;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.msg-preview{font-size:11px;color:#484f58;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:2;max-width:260px;}
.msg-date{font-size:10.5px;color:#484f58;flex-shrink:0;}
.msg-chevron{color:#484f58;font-size:9px;flex-shrink:0;transition:transform .2s;}
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
.msg-ai-loading{display:flex;align-items:center;gap:8px;color:#484f58;font-size:11.5px;padding:6px 0;}

/* Sync */
.sync-status{display:flex;flex-direction:column;align-items:center;gap:3px;}
.sync-row{display:flex;align-items:center;gap:7px;font-size:11px;color:#8b949e;}
.sync-dot{width:7px;height:7px;border-radius:50%;background:#3fb950;flex-shrink:0;}
.sync-dot.syncing{background:#58a6ff;animation:pdot 1s infinite;}
.sync-dot.error{background:#f85149;}
@keyframes pdot{0%,100%{opacity:1;transform:scale(1);}50%{opacity:.5;transform:scale(.8);}}
.sync-txt{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:280px;}
.sync-bar-wrap{width:190px;height:3px;background:#1a3252;border-radius:2px;overflow:hidden;display:none;}
.sync-bar{height:100%;background:linear-gradient(90deg,#58a6ff,#bc8cff);border-radius:2px;transition:width .4s ease;}
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
.modal-lg{max-width:560px;}
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
.recip-empty{font-size:11px;color:#484f58;font-style:italic;padding:4px 2px;}
.reply-intent-hint{font-size:11px;color:#5a7a9e;margin-bottom:6px;}
.generating-overlay{display:flex;align-items:center;gap:8px;color:#58a6ff;font-size:12px;padding:8px 0;}
</style>
</head>
<body>
<div class="app">
  <header class="header">
    <div class="header-brand">
      <h1>✉ Email Triage</h1>
      <span id="email-count" class="badge-sm"></span>
    </div>
    <div class="header-center">
      <div class="sync-status">
        <div class="sync-row">
          <div class="sync-dot" id="sync-dot"></div>
          <span class="sync-txt" id="sync-txt">Connecting...</span>
          <span id="new-badge-wrap"></span>
        </div>
        <div class="sync-bar-wrap" id="sync-bar-wrap">
          <div class="sync-bar" id="sync-bar" style="width:0%"></div>
        </div>
      </div>
    </div>
    <div class="header-right">
      <button class="btn btn-ghost btn-sm" onclick="triggerSync()">⟳ Sync Now</button>
      <button class="btn btn-ghost btn-sm" onclick="reanalyzeAll()" id="reanalyze-btn">⚙ Re-analyze</button>
    </div>
  </header>
  <div class="body">
    <nav class="sidebar" id="sidebar">
      <div class="sidebar-hdr">Threads</div>
      <div id="topic-list"></div>
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
    </div>
  </div>
</div>

<!-- Reply modal (2-step) -->
<div class="modal-overlay" id="reply-modal">
  <div class="modal modal-lg">

    <!-- Step 1: Intent capture -->
    <div id="reply-step1">
      <h3>✏️ What do you want to say?</h3>
      <div class="modal-sub" id="reply-sub"></div>
      <div class="reply-intent-hint">Write a short note — your key point, decision, or ask. AI will shape it into a polished reply.</div>
      <textarea id="reply-intent" placeholder="e.g. 'Approve the budget increase, ask for monthly check-ins going forward'" style="min-height:90px;resize:vertical"></textarea>
      <div class="modal-footer">
        <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
        <button class="btn btn-reply" id="generate-btn" onclick="generateReply()">✨ Generate Reply</button>
      </div>
    </div>

    <!-- Step 2: Review + send -->
    <div id="reply-step2" style="display:none">
      <h3>↩ Reply</h3>
      <div class="modal-sub" id="reply-sub2"></div>
      <div class="recip-row">
        <span class="recip-label">TO</span>
        <div class="recip-field" id="reply-to-field"></div>
      </div>
      <div class="recip-row">
        <span class="recip-label">CC</span>
        <div class="recip-field" id="reply-cc-field"></div>
      </div>
      <textarea id="reply-body" style="margin-top:10px;min-height:220px;resize:vertical"></textarea>
      <div class="modal-footer">
        <button class="btn btn-ghost" onclick="backToStep1()">← Back</button>
        <button class="btn btn-ghost" onclick="closeModals()">Cancel</button>
        <button class="btn btn-reply" onclick="sendReply()">✉ Send Reply</button>
      </div>
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
  document.getElementById('empty-pane').style.display='flex';
  updateCounts(d.emailCount, Object.keys(state.threadMap).length);
  updateSyncStatus(d.syncStatus);
  renderSidebar();
  schedulePoll();
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
  document.getElementById('topic-list').innerHTML = state.groups.map(g=>{
    const collapsed = state.collapsedTopics.has(g.topic);
    const groupUnread = g.threads.some(t=>t.hasUnread);
    const groupFlagged = g.threads.some(t=>t.isFlagged);
    return `<div class="topic-group ${collapsed?'collapsed':''}">
      <div class="tg-header${groupUnread?' has-unread':''}${groupFlagged?' has-flagged':''}" onclick="toggleGroup('${esc(g.topic)}')">
        <span class="tg-name">${esc(g.topic)}</span>
        <span class="tg-count">${g.threads.length}</span>
        <span class="tg-chevron">▾</span>
      </div>
      <div class="tg-threads">${g.threads.map(t=>_threadItemHTML(t)).join('')}</div>
    </div>`;
  }).join('');
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
  renderSidebar();
  const t = state.threadMap[convKey];
  if (!t) return;
  document.getElementById('empty-pane').style.display='none';
  document.getElementById('thread-detail').style.display='flex';
  _renderThreadHdr(t);
  const sec = document.getElementById('msgs-section');
  sec.innerHTML=`<div class="msgs-label">Messages <span>${t.messageCount||0}</span></div><div class="msg-ai-loading"><div class="spinner spinner-sm"></div> Loading messages…</div>`;
  if (!t.emailIds||!t.emailIds.length){sec.innerHTML+='<div style="color:#484f58;font-size:12px;padding:10px 0">No messages found.</div>';return;}
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
  let fileBtnHtml=t.suggestedFolder
    ?`<button class="btn btn-file btn-sm" onclick="quickFile('${enc}','${esc(t.suggestedFolder)}')">📁 ${esc(t.suggestedFolder)}</button>`
    :`<button class="btn btn-file btn-sm" onclick="openFile('${enc}')">📁 File</button>`;
  document.getElementById('thread-hdr').innerHTML=`
    <div class="th-top">
      <div style="flex:1">
        <div class="th-badges">
          <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
          <span class="act-pill ${actCls}">${t.action||'read'}</span>
          ${t.hasUnread?'<span style="width:6px;height:6px;border-radius:50%;background:#58a6ff;display:inline-block"></span>':''}
        </div>
        <div class="th-subject">${esc(t.subject||'(No subject)')}</div>
      </div>
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
  if (!msgs.length){sec.innerHTML=`<div class="msgs-label">Messages</div><div style="color:#484f58;font-size:12px;padding:16px 0">No messages found.</div>`;return;}
  let html=`<div class="msgs-label">Messages <span>${msgs.length}</span></div>`;
  html+=msgs.map((m,i)=>_msgCardHTML(m,i)).join('');
  sec.innerHTML=html;
}

function _msgCardHTML(m, idx) {
  const from=m.from_name||m.from_address||'Unknown';
  const date=fmtDate((m.received_date_time||'').slice(0,19));
  const bodyText=String(m.body||m.body_preview||'').trim();
  const preview=bodyText.slice(0,100).replace(/\n+/g,' ');
  const isOpen=state.expandedMsgs.has(idx);
  return `<div class="msg-card${isOpen?' open':''}" id="mc-${idx}">
    <div class="msg-hdr" onclick="toggleMsg(${idx})">
      <span class="avatar" style="background:${avColor(from)};width:24px;height:24px;font-size:8.5px;border:2px solid #0a1628;flex-shrink:0">${initials(from)}</span>
      <span class="msg-from">${esc(from)}</span>
      <span class="msg-preview">${esc(preview)}</span>
      <span class="msg-date">${esc(date)}</span>
      <span class="msg-chevron">▾</span>
    </div>
    <div class="msg-body" id="mb-${idx}">${isOpen?_bodyContent(idx):''}
    </div>
  </div>`;
}

function _bodyContent(idx) {
  const m=state.currentMsgs[idx];
  if (!m) return '';
  if (state.formatCache[m.id]) return _renderParas(state.formatCache[m.id]);
  setTimeout(()=>loadFormatted(idx),0);
  return `<div class="msg-ai-loading"><div class="spinner spinner-sm"></div> Formatting with AI…</div>`;
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

async function loadFormatted(idx) {
  const m=state.currentMsgs[idx];
  if (!m||!state.expandedMsgs.has(idx)) return;
  const body=document.getElementById('mb-'+idx);
  try {
    const d=await fetch(`/api/format_message?id=${encodeURIComponent(m.id)}`).then(r=>r.json());
    state.formatCache[m.id]=d.paragraphs||[];
    if (body&&state.expandedMsgs.has(idx)) body.innerHTML=_renderParas(state.formatCache[m.id]);
  } catch(e) {
    const fallback=String(m.body||m.body_preview||'').trim();
    if (body&&state.expandedMsgs.has(idx)) body.innerHTML=`<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(fallback)}</div>`;
  }
}

// intent → CSS class suffix
const INTENT_CLS = {
  'Status Update':'status-update','Request':'request','Decision':'decision',
  'Question':'question','Action Item':'action-item','Context':'context',
  'FYI':'fyi','Warning':'warning','Introduction':'introduction','Closing':'closing'
};

function _renderParas(paras) {
  if (!paras||!paras.length) return '<div style="color:#484f58;font-size:12px">(no content)</div>';
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
    wrap.style.display='none';
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
  const el=document.getElementById('email-count');
  el.textContent=emailCount!==null?`${emailCount} emails · ${threadCount} threads`:`${threadCount} threads`;
}
async function triggerSync() {
  const d=await fetch('/api/sync_now',{method:'POST'}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
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

function openReply(enc) {
  _replyState.thread = decodeThread(enc);
  _activeThread = _replyState.thread;
  const t = _replyState.thread;

  // Seed To/CC from the most-recent message (newest-first = index 0)
  const latest = state.currentMsgs[0];
  _replyState.to = [];
  _replyState.cc = [];
  if (latest) {
    // Sender of latest message always goes into To
    if (latest.from_address) {
      _replyState.to.push({name: latest.from_name||latest.from_address, address: latest.from_address});
    }
    // Add any To recipients not already listed
    for (const r of (latest.to_recipients||[])) {
      if (!_replyState.to.find(x=>x.address===r.address)) _replyState.to.push(r);
    }
    _replyState.cc = (latest.cc_recipients||[]).slice();
  }

  document.getElementById('reply-sub').textContent = `Re: ${t.subject||''}`;
  document.getElementById('reply-intent').value = '';
  document.getElementById('reply-step1').style.display = '';
  document.getElementById('reply-step2').style.display = 'none';
  document.getElementById('reply-modal').classList.add('open');
  setTimeout(()=>document.getElementById('reply-intent').focus(), 50);
}

function backToStep1() {
  document.getElementById('reply-step1').style.display = '';
  document.getElementById('reply-step2').style.display = 'none';
}

async function generateReply() {
  const intent = document.getElementById('reply-intent').value.trim();
  if (!intent) { document.getElementById('reply-intent').focus(); return; }
  const btn = document.getElementById('generate-btn');
  btn.disabled = true;
  btn.innerHTML = '<div class="spinner spinner-sm"></div> Generating…';
  try {
    const t = _replyState.thread;
    const d = await fetch('/api/generate_reply', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: t.conversationKey, userPrompt: intent})
    }).then(r=>r.json());
    if (d.error) { alert('Error: '+d.error); return; }
    document.getElementById('reply-sub2').textContent = `Re: ${t.subject||''}`;
    document.getElementById('reply-body').value = d.reply || '';
    _renderRecipFields();
    document.getElementById('reply-step1').style.display = 'none';
    document.getElementById('reply-step2').style.display = '';
    setTimeout(()=>document.getElementById('reply-body').focus(), 50);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '✨ Generate Reply';
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
    `<span class="recip-tag">${esc(r.name||r.address)}<span class="rm" onclick="removeRecip('${field}','${esc(r.address)}')">×</span></span>`
  ).join('');
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
function openDelete(enc) {
  _activeThread=decodeThread(enc);
  document.getElementById('del-subj').textContent=_activeThread.subject||'(No subject)';
  document.getElementById('del-count').textContent=_activeThread.messageCount||(_activeThread.emailIds||[]).length;
  document.getElementById('delete-modal').classList.add('open');
}
function closeModals() {
  document.querySelectorAll('.modal-overlay').forEach(m=>m.classList.remove('open'));
  _activeThread=null;
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
  delete state.threadMap[convKey];
  for (const g of state.groups) g.threads=g.threads.filter(t=>t.conversationKey!==convKey);
  state.groups=state.groups.filter(g=>g.threads.length>0);
  if (state.selectedKey===convKey) {
    state.selectedKey=null;
    document.getElementById('thread-detail').style.display='none';
    document.getElementById('empty-pane').style.display='flex';
  }
  renderSidebar();
  updateCounts(null,Object.keys(state.threadMap).length);
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
function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

document.querySelectorAll('.modal-overlay').forEach(m=>
  m.addEventListener('click',e=>{if(e.target===m)closeModals();}));

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
