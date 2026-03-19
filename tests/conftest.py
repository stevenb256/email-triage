"""
conftest.py — Shared pytest fixtures for the email triage app test suite.

All external I/O (MCP subprocess, Anthropic API, keychain) is mocked.
The DB uses an in-memory SQLite connection so no files are written to disk.

IMPORTANT: This file is loaded by pytest before any test module. The top-level
code here sets up sys.modules stubs so that mcp_client can be imported without
spawning the real McpOutlookLocal subprocess.
"""
import json
import sqlite3
import sys
import os
import threading
import types
import unittest.mock as _um
from unittest.mock import MagicMock, patch

import pytest

# ---------------------------------------------------------------------------
# Make the app root importable when running `pytest tests/` from any cwd
# ---------------------------------------------------------------------------
APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)


# ---------------------------------------------------------------------------
# Stub out the `mcp` package before mcp_client is imported.
# mcp_client.py does:
#   from mcp import ClientSession, StdioServerParameters
#   from mcp.client.stdio import stdio_client
# We provide fake versions of these so the import succeeds without the binary.
# ---------------------------------------------------------------------------
_mcp_pkg = types.ModuleType("mcp")
_mcp_pkg.ClientSession = MagicMock()
_mcp_pkg.StdioServerParameters = MagicMock()

_mcp_client_pkg = types.ModuleType("mcp.client")
_mcp_client_stdio_pkg = types.ModuleType("mcp.client.stdio")
_mcp_client_stdio_pkg.stdio_client = MagicMock()

sys.modules.setdefault("mcp", _mcp_pkg)
sys.modules.setdefault("mcp.client", _mcp_client_pkg)
sys.modules.setdefault("mcp.client.stdio", _mcp_client_stdio_pkg)

# ---------------------------------------------------------------------------
# Import mcp_client with the background thread and asyncio coroutine patched
# out so they never fire during tests.
# ---------------------------------------------------------------------------
_mock_session_ready = threading.Event()
_mock_session_ready.set()  # always "ready" in tests

with _um.patch("threading.Thread"), \
     _um.patch("asyncio.run_coroutine_threadsafe"), \
     _um.patch("asyncio.new_event_loop", return_value=MagicMock()):
    import mcp_client as _mcp_module

# Replace module-level singletons with test-safe stubs
_mcp_module._session_ready = _mock_session_ready
_mcp_module.call_tool = MagicMock(return_value={})


# ---------------------------------------------------------------------------
# In-memory DB schema + helpers
# ---------------------------------------------------------------------------

_SCHEMA_SQL = """
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
    formatted_body      TEXT,
    folder              TEXT,
    body_html           TEXT
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
    updated_at          TEXT,
    is_flagged          INTEGER DEFAULT 0
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
CREATE TABLE IF NOT EXISTS contacts (
    email       TEXT PRIMARY KEY,
    name        TEXT,
    frequency   INTEGER DEFAULT 0,
    last_seen   TEXT
);
CREATE INDEX IF NOT EXISTS idx_emails_conv_key ON emails(conversation_key);
CREATE INDEX IF NOT EXISTS idx_threads_updated  ON threads(updated_at);
CREATE INDEX IF NOT EXISTS idx_calendar_start ON calendar_events(start_time);
"""


def _make_in_memory_db() -> sqlite3.Connection:
    """Create and initialise a fresh in-memory SQLite DB."""
    conn = sqlite3.connect(":memory:", check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.executescript(_SCHEMA_SQL)
    conn.commit()
    return conn


def _seed_db(conn: sqlite3.Connection):
    """Insert sample data: 5 emails in 2 threads, 1 calendar event."""
    emails = [
        # Thread 1 — conversation_key = "project alpha update"
        ("email-1", "Project Alpha Update", "Alice Smith", "alice@example.com",
         "2026-03-15T10:00:00Z", 0, "Alpha is on track for Q2 launch.",
         "project alpha update", "{}", "2026-03-15T10:00:00Z", None, "Inbox", None),
        ("email-2", "RE: Project Alpha Update", "Bob Jones", "bob@example.com",
         "2026-03-15T11:00:00Z", 1, "Agreed, let's sync Friday.",
         "project alpha update", "{}", "2026-03-15T11:00:00Z", None, "Inbox", None),
        ("email-3", "RE: Project Alpha Update", "Alice Smith", "alice@example.com",
         "2026-03-16T09:00:00Z", 0, "Friday works. Sending invite now.",
         "project alpha update", "{}", "2026-03-16T09:00:00Z", None, "Inbox", None),
        # Thread 2 — conversation_key = "budget review"
        ("email-4", "Budget Review Q2", "Carol White", "carol@example.com",
         "2026-03-17T08:00:00Z", 0, "Please review the Q2 budget attached.",
         "budget review", "{}", "2026-03-17T08:00:00Z", None, "Inbox", None),
        ("email-5", "RE: Budget Review Q2", "Dan Brown", "dan@example.com",
         "2026-03-17T14:00:00Z", 1, "Reviewed. Looks good to me.",
         "budget review q2", "{}", "2026-03-17T14:00:00Z", None, "Inbox", None),
    ]
    conn.executemany(
        "INSERT OR IGNORE INTO emails "
        "(id,subject,from_name,from_address,received_date_time,is_read,"
        "body_preview,conversation_key,raw_json,synced_at,formatted_body,folder,body_html) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
        emails,
    )

    threads = [
        ("project alpha update", "Project Alpha Update", "Projects", "reply", "high",
         "Alpha on track||BREAK||None||BREAK||Reply to Alice by EOD",
         "Sounds good, see you Friday!", "Efforts/Alpha",
         json.dumps(["Alice Smith", "Bob Jones"]),
         json.dumps(["email-1", "email-2", "email-3"]),
         "email-3", 3, 1, "2026-03-16T09:00:00Z", "2026-03-16T09:00:00Z"),
        ("budget review", "Budget Review Q2", "Finance", "read", "medium",
         "Budget review for Q2||BREAK||None||BREAK||File for reference",
         "", "",
         json.dumps(["Carol White", "Dan Brown"]),
         json.dumps(["email-4"]),
         "email-4", 1, 1, "2026-03-17T08:00:00Z", "2026-03-17T08:00:00Z"),
    ]
    conn.executemany(
        "INSERT OR IGNORE INTO threads "
        "(conversation_key,subject,topic,action,urgency,summary,"
        "suggested_reply,suggested_folder,participants,email_ids,"
        "latest_id,message_count,has_unread,latest_received,updated_at) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        threads,
    )

    # 1 calendar event — starts in the "future" relative to test date 2026-03-18
    conn.execute(
        "INSERT OR IGNORE INTO calendar_events "
        "(id,subject,start_time,end_time,location,attendees,raw_json,synced_at) "
        "VALUES(?,?,?,?,?,?,?,?)",
        (
            "cal-1", "Weekly Sync",
            "2026-03-19T14:00:00", "2026-03-19T15:00:00",
            "Teams", json.dumps(["Alice Smith", "Bob Jones"]),
            "{}", "2026-03-18T00:00:00Z",
        ),
    )

    # meta entries
    conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('my_email','me@example.com')")
    conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('efforts_subfolders','[\"Efforts/Alpha\"]')")
    conn.execute("INSERT OR REPLACE INTO meta(key,value) VALUES('other_folders','[\"Partners\"]')")
    conn.commit()


# ---------------------------------------------------------------------------
# Core pytest fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="function")
def in_memory_conn():
    """Fresh in-memory SQLite connection (schema only, no seed data)."""
    conn = _make_in_memory_db()
    yield conn
    conn.close()


@pytest.fixture(scope="function")
def db(in_memory_conn):
    """Seeded in-memory DB: 5 emails, 2 threads, 1 calendar event."""
    _seed_db(in_memory_conn)
    return in_memory_conn


@pytest.fixture(scope="function")
def app(db):
    """
    Flask app configured for testing with the DB patched to use in-memory SQLite.

    We patch db.get_db as well as the imported name in each route module so that
    all DB access within a test goes to the in-memory connection.
    """
    import db as db_module

    # Patch get_db globally and in each blueprint
    with patch.object(db_module, "get_db", return_value=db), \
         patch("routes.mail.get_db", return_value=db), \
         patch("routes.triage.get_db", return_value=db), \
         patch("routes.calendar.get_db", return_value=db):

        # Import the Flask app — blueprints are already registered at module level
        from app import app as flask_app
        flask_app.config["TESTING"] = True
        flask_app.config["SECRET_KEY"] = "test-secret"

        with flask_app.app_context():
            yield flask_app


@pytest.fixture(scope="function")
def client(app):
    """Flask test client."""
    return app.test_client()


# ---------------------------------------------------------------------------
# MCP response builders (used by individual test modules)
# ---------------------------------------------------------------------------

def make_mcp_message(
    msg_id="msg-1",
    subject="Test Subject",
    from_name="Alice Smith",
    from_address="alice@example.com",
    body_content="<html><body><p>Hello, this is a test email body.</p></body></html>",
    is_read=False,
    received="2026-03-15T10:00:00Z",
    to_recipients=None,
    cc_recipients=None,
    body_content_type="HTML",
):
    """Build a realistic MCP message dict."""
    return {
        "id": msg_id,
        "subject": subject,
        "from_name": from_name,
        "from_address": from_address,
        "body_content": body_content,
        "body_content_type": body_content_type,
        "is_read": is_read,
        "received_date_time": received,
        "to_recipients": to_recipients or [
            {"name": "Me", "address": "me@example.com"}
        ],
        "cc_recipients": cc_recipients or [],
    }


def make_mcp_draft_response(draft_id="draft-abc123"):
    """Build a realistic MCP draft response."""
    return {"draft_id": draft_id, "id": draft_id}


def make_mcp_send_response():
    """Build a realistic MCP send response."""
    return {"ok": True}
