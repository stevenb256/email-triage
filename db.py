"""
db.py — Database layer for Clanker email triage app.
"""
import json
import sqlite3
import threading

from config import DB_PATH

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
