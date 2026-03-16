"""
Comprehensive test suite for Email Triage application.
Tests database layer, helpers, AI analysis, sync engine, and all Flask API routes.
"""

import json
import os
import sqlite3
import tempfile
import threading
import time
from datetime import datetime, timezone
from unittest.mock import MagicMock, patch, PropertyMock

import pytest

# Patch MCP and background threads BEFORE importing app
# This prevents the module-level MCP session and background loop from starting
with patch("mcp.client.stdio.stdio_client"), \
     patch("mcp.ClientSession"), \
     patch("threading.Thread") as _mock_thread:
    _mock_thread.return_value.start = MagicMock()
    import app as email_app


# ─── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def temp_db(tmp_path, monkeypatch):
    """Use a fresh temp database for every test."""
    db_path = str(tmp_path / "test.db")
    monkeypatch.setattr(email_app, "DB_PATH", db_path)
    # Clear thread-local connection so get_db creates a new one
    if hasattr(email_app._thread_local, "conn"):
        delattr(email_app._thread_local, "conn")
    email_app.init_db()
    yield db_path
    if hasattr(email_app._thread_local, "conn"):
        try:
            email_app._thread_local.conn.close()
        except Exception:
            pass
        delattr(email_app._thread_local, "conn")


@pytest.fixture
def client():
    """Flask test client."""
    email_app.app.config["TESTING"] = True
    with email_app.app.test_client() as c:
        yield c


@pytest.fixture
def db():
    """Get a fresh DB connection for assertions."""
    return email_app.get_db()


@pytest.fixture
def sample_emails():
    """Sample email data for testing."""
    return [
        {
            "id": "msg-001",
            "subject": "RE: Project Update",
            "from_name": "Alice Smith",
            "from_address": "alice@example.com",
            "received_date_time": "2024-01-15T10:00:00Z",
            "is_read": False,
            "body_preview": "Here is the latest update on the project.",
        },
        {
            "id": "msg-002",
            "subject": "RE: Project Update",
            "from_name": "Bob Jones",
            "from_address": "bob@example.com",
            "received_date_time": "2024-01-15T11:00:00Z",
            "is_read": True,
            "body_preview": "Thanks Alice, looks good. I have a question about the timeline.",
        },
    ]


@pytest.fixture
def populated_db(db, sample_emails):
    """DB with sample emails and a thread inserted."""
    conv_key = "project update"
    now = email_app._utcnow()
    for e in sample_emails:
        db.execute(
            "INSERT INTO emails (id,subject,from_name,from_address,received_date_time,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
            (e["id"], e["subject"], e["from_name"], e["from_address"],
             e["received_date_time"], 1 if e["is_read"] else 0,
             e["body_preview"], conv_key, json.dumps(e), now),
        )
    db.execute(
        "INSERT INTO threads (conversation_key,subject,topic,action,urgency,summary,"
        "suggested_reply,suggested_folder,participants,email_ids,latest_id,"
        "message_count,has_unread,latest_received,updated_at,is_flagged) "
        "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (conv_key, "RE: Project Update", "Engineering", "reply", "medium",
         "Project update discussion", "Looks good, thanks!", "",
         json.dumps(["Alice Smith", "Bob Jones"]),
         json.dumps(["msg-001", "msg-002"]),
         "msg-002", 2, 1, "2024-01-15T11:00:00Z", now, 0),
    )
    db.commit()
    return conv_key


# ─── Database Layer Tests ─────────────────────────────────────────────────────

class TestDatabase:
    def test_get_db_returns_connection(self):
        conn = email_app.get_db()
        assert isinstance(conn, sqlite3.Connection)
        assert conn.row_factory == sqlite3.Row

    def test_get_db_returns_same_connection(self):
        conn1 = email_app.get_db()
        conn2 = email_app.get_db()
        assert conn1 is conn2

    def test_init_db_creates_tables(self, db):
        tables = {r[0] for r in db.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
        assert "emails" in tables
        assert "threads" in tables
        assert "meta" in tables

    def test_init_db_creates_indexes(self, db):
        indexes = {r[0] for r in db.execute(
            "SELECT name FROM sqlite_master WHERE type='index'"
        ).fetchall()}
        assert "idx_emails_conv_key" in indexes
        assert "idx_threads_updated" in indexes
        assert "idx_threads_urgency" in indexes

    def test_init_db_idempotent(self, db):
        # Calling init_db again should not error
        email_app.init_db()
        tables = {r[0] for r in db.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
        assert "emails" in tables

    def test_meta_get_default(self):
        assert email_app.meta_get("nonexistent") is None
        assert email_app.meta_get("nonexistent", "fallback") == "fallback"

    def test_meta_set_and_get(self):
        email_app.meta_set("test_key", "test_value")
        assert email_app.meta_get("test_key") == "test_value"

    def test_meta_set_overwrites(self):
        email_app.meta_set("key", "value1")
        email_app.meta_set("key", "value2")
        assert email_app.meta_get("key") == "value2"

    def test_thread_to_dict_basic(self, db, populated_db):
        row = db.execute("SELECT * FROM threads WHERE conversation_key=?", (populated_db,)).fetchone()
        result = email_app._thread_to_dict(row)
        assert result["conversationKey"] == populated_db
        assert result["subject"] == "RE: Project Update"
        assert result["topic"] == "Engineering"
        assert result["action"] == "reply"
        assert result["urgency"] == "medium"
        assert result["messageCount"] == 2
        assert result["hasUnread"] is True
        assert result["isFlagged"] is False
        assert isinstance(result["participants"], list)
        assert "Alice Smith" in result["participants"]
        assert isinstance(result["emailIds"], list)
        assert "msg-001" in result["emailIds"]

    def test_thread_to_dict_defaults(self, db):
        db.execute(
            "INSERT INTO threads (conversation_key) VALUES(?)", ("empty-key",)
        )
        db.commit()
        row = db.execute("SELECT * FROM threads WHERE conversation_key='empty-key'").fetchone()
        result = email_app._thread_to_dict(row)
        assert result["subject"] == ""
        assert result["topic"] == "General"
        assert result["action"] == "read"
        assert result["urgency"] == "low"
        assert result["participants"] == []
        assert result["emailIds"] == []

    def test_thread_to_dict_invalid_json(self, db):
        db.execute(
            "INSERT INTO threads (conversation_key, participants, email_ids) VALUES(?,?,?)",
            ("bad-json", "not-json", "not-json"),
        )
        db.commit()
        row = db.execute("SELECT * FROM threads WHERE conversation_key='bad-json'").fetchone()
        result = email_app._thread_to_dict(row)
        assert result["participants"] == []
        assert result["emailIds"] == []

    def test_remove_thread(self, db, populated_db):
        email_app.remove_thread(populated_db)
        assert db.execute("SELECT COUNT(*) FROM emails WHERE conversation_key=?", (populated_db,)).fetchone()[0] == 0
        assert db.execute("SELECT COUNT(*) FROM threads WHERE conversation_key=?", (populated_db,)).fetchone()[0] == 0

    def test_remove_thread_nonexistent(self, db):
        # Should not error
        email_app.remove_thread("nonexistent-key")


# ─── Helper Functions Tests ───────────────────────────────────────────────────

class TestHelpers:
    # _norm_subject
    def test_norm_subject_removes_re(self):
        assert email_app._norm_subject("RE: Hello") == "hello"

    def test_norm_subject_removes_fw(self):
        assert email_app._norm_subject("FW: Hello") == "hello"

    def test_norm_subject_removes_fwd(self):
        assert email_app._norm_subject("FWD: Hello") == "hello"

    def test_norm_subject_removes_aw(self):
        assert email_app._norm_subject("AW: Hallo") == "hallo"

    def test_norm_subject_case_insensitive(self):
        assert email_app._norm_subject("re: Test") == "test"
        assert email_app._norm_subject("Re: Test") == "test"
        assert email_app._norm_subject("RE: Test") == "test"

    def test_norm_subject_empty(self):
        assert email_app._norm_subject("") == "no-subject"
        assert email_app._norm_subject(None) == "no-subject"

    def test_norm_subject_no_prefix(self):
        assert email_app._norm_subject("Hello World") == "hello world"

    def test_norm_subject_strips_whitespace(self):
        assert email_app._norm_subject("RE:  Hello  ") == "hello"

    # _clean
    def test_clean_removes_control_chars(self):
        assert email_app._clean("hello\x00world") == "helloworld"
        assert email_app._clean("test\x0b\x0c\x0e") == "test"

    def test_clean_truncates(self):
        assert email_app._clean("hello world", 5) == "hello"

    def test_clean_no_truncate(self):
        assert email_app._clean("hello") == "hello"

    def test_clean_none_input(self):
        assert email_app._clean(None) == ""
        assert email_app._clean("") == ""

    # _utcnow
    def test_utcnow_format(self):
        result = email_app._utcnow()
        # Should match YYYY-MM-DDTHH:MM:SSZ
        assert len(result) == 20
        assert result.endswith("Z")
        assert result[4] == "-"
        assert result[10] == "T"

    # _folder_lists
    def test_folder_lists_basic(self):
        folders = [
            {"display_name": "Efforts/Project A"},
            {"display_name": "Archive"},
            {"display_name": "Drafts"},
            {"display_name": "Sent Items"},
        ]
        efforts, other = email_app._folder_lists(folders)
        assert "Efforts/Project A" in efforts
        assert "Archive" in other
        assert "Drafts" not in efforts and "Drafts" not in other

    def test_folder_lists_skips_system(self):
        folders = [
            {"display_name": "Drafts"},
            {"display_name": "Sent Items"},
            {"display_name": "Outbox"},
            {"display_name": "Deleted Items"},
            {"display_name": "Junk Email"},
        ]
        efforts, other = email_app._folder_lists(folders)
        assert efforts == []
        assert other == []

    def test_folder_lists_empty(self):
        efforts, other = email_app._folder_lists([])
        assert efforts == []
        assert other == []

    def test_folder_lists_displayName_key(self):
        folders = [{"displayName": "Efforts/Test"}, {"displayName": "Projects"}]
        efforts, other = email_app._folder_lists(folders)
        assert "Efforts/Test" in efforts
        assert "Projects" in other


# ─── Topic Normalization Tests ────────────────────────────────────────────────

class TestTopicNormalization:
    def test_exact_match(self):
        assert email_app._normalize_topic("Engineering") == "Engineering"
        assert email_app._normalize_topic("Finance") == "Finance"

    def test_case_insensitive_exact(self):
        assert email_app._normalize_topic("engineering") == "Engineering"
        assert email_app._normalize_topic("FINANCE") == "Finance"

    def test_keyword_match_incident(self):
        assert email_app._normalize_topic("Server outage report") == "Incidents & Outages"

    def test_keyword_match_finance(self):
        assert email_app._normalize_topic("Q4 budget review") == "Finance"

    def test_keyword_match_legal(self):
        assert email_app._normalize_topic("GDPR compliance update") == "Legal & Compliance"

    def test_keyword_match_travel(self):
        assert email_app._normalize_topic("Team offsite planning") == "Events & Travel"

    def test_keyword_match_partnerships(self):
        assert email_app._normalize_topic("Customer engagement plan") == "Partnerships"

    def test_keyword_match_hr(self):
        assert email_app._normalize_topic("New hiring plan") == "Team & HR"

    def test_keyword_match_customer(self):
        assert email_app._normalize_topic("Customer issue escalation") == "Customer Issues"

    def test_keyword_match_architecture(self):
        assert email_app._normalize_topic("System design review") == "Architecture & Design"

    def test_keyword_match_strategy(self):
        assert email_app._normalize_topic("OKR planning session") == "Strategy & Leadership"

    def test_keyword_match_external(self):
        assert email_app._normalize_topic("Press announcement draft") == "External Communications"

    def test_keyword_match_product(self):
        assert email_app._normalize_topic("Product roadmap update") == "Product Planning"

    def test_keyword_match_engineering(self):
        assert email_app._normalize_topic("Infrastructure migration") == "Engineering"

    def test_keyword_match_fyi(self):
        assert email_app._normalize_topic("Weekly status update") == "FYI & Updates"

    def test_fallback(self):
        assert email_app._normalize_topic("Random topic xyz") == "FYI & Updates"

    def test_empty_string(self):
        assert email_app._normalize_topic("") == "FYI & Updates"

    def test_priority_order(self):
        # "incident" keywords should match before "engineering" keywords
        assert email_app._normalize_topic("incident deploy failure") == "Incidents & Outages"


# ─── Parse Recipients Tests ──────────────────────────────────────────────────

class TestParseRecipients:
    def test_basic(self):
        raw = [{"name": "Alice", "address": "alice@example.com"}]
        result = email_app._parse_recipients(raw)
        assert len(result) == 1
        assert result[0]["name"] == "Alice"
        assert result[0]["address"] == "alice@example.com"

    def test_outlook_format(self):
        raw = [{"emailAddress": {"name": "Bob", "address": "bob@example.com"}}]
        result = email_app._parse_recipients(raw)
        assert len(result) == 1
        assert result[0]["name"] == "Bob"
        assert result[0]["address"] == "bob@example.com"

    def test_display_name_key(self):
        raw = [{"display_name": "Charlie", "email": "charlie@example.com"}]
        result = email_app._parse_recipients(raw)
        assert result[0]["name"] == "Charlie"
        assert result[0]["address"] == "charlie@example.com"

    def test_empty_list(self):
        assert email_app._parse_recipients([]) == []
        assert email_app._parse_recipients(None) == []

    def test_skips_non_dict(self):
        raw = ["not-a-dict", 42, None]
        assert email_app._parse_recipients(raw) == []

    def test_skips_no_address(self):
        raw = [{"name": "NoEmail"}]
        assert email_app._parse_recipients(raw) == []

    def test_strips_whitespace(self):
        raw = [{"name": "  Alice  ", "address": "  alice@example.com  "}]
        result = email_app._parse_recipients(raw)
        assert result[0]["name"] == "Alice"
        assert result[0]["address"] == "alice@example.com"


# ─── Normalize Message Tests ─────────────────────────────────────────────────

class TestNormalizeMsg:
    def test_basic_message(self):
        msg = {
            "id": "m1",
            "subject": "Test",
            "from_name": "Alice",
            "from_address": "alice@example.com",
            "received_date_time": "2024-01-15T10:00:00Z",
            "is_read": True,
            "body_preview": "Hello world",
        }
        result = email_app._normalize_msg(msg)
        assert result["id"] == "m1"
        assert result["subject"] == "Test"
        assert result["from_name"] == "Alice"
        assert result["body"] == "Hello world"

    def test_html_body(self):
        msg = {
            "body_content": "<p>Hello</p>&nbsp;<b>World</b>",
            "body_preview": "fallback",
        }
        result = email_app._normalize_msg(msg)
        assert "<p>" not in result["body"]
        assert "<b>" not in result["body"]
        assert "Hello" in result["body"]
        assert "World" in result["body"]

    def test_falls_back_to_preview(self):
        msg = {"body_preview": "Preview text"}
        result = email_app._normalize_msg(msg)
        assert result["body"] == "Preview text"

    def test_empty_msg(self):
        result = email_app._normalize_msg({})
        assert result["id"] == ""
        assert result["from_name"] == ""
        assert result["body"] == ""

    def test_recipients_parsing(self):
        msg = {
            "to_recipients": [{"name": "Bob", "address": "bob@example.com"}],
            "cc_recipients": [{"name": "Carol", "address": "carol@example.com"}],
        }
        result = email_app._normalize_msg(msg)
        assert len(result["to_recipients"]) == 1
        assert len(result["cc_recipients"]) == 1
        assert result["to_recipients"][0]["name"] == "Bob"

    def test_camelCase_recipients(self):
        msg = {
            "toRecipients": [{"name": "Bob", "address": "bob@example.com"}],
            "ccRecipients": [{"name": "Carol", "address": "carol@example.com"}],
        }
        result = email_app._normalize_msg(msg)
        assert len(result["to_recipients"]) == 1
        assert len(result["cc_recipients"]) == 1


# ─── AI Analysis Tests (mocked) ──────────────────────────────────────────────

class TestAnalyzeThread:
    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_success(self, mock_get_ai, sample_emails):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text=json.dumps({
            "summary": "Project update discussion between Alice and Bob.",
            "topic": "Engineering",
            "action": "reply",
            "urgency": "medium",
            "suggestedReply": "Thanks for the update.",
            "suggestedFolder": "",
        }))]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        result = email_app.analyze_thread(sample_emails, ["Efforts/Proj"], ["Archive"])
        assert result["topic"] == "Engineering"
        assert result["action"] == "reply"
        assert result["urgency"] == "medium"
        assert result["suggestedReply"] == "Thanks for the update."

    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_with_markdown_fences(self, mock_get_ai, sample_emails):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text='```json\n{"summary":"test","topic":"Finance","action":"read","urgency":"low","suggestedReply":"","suggestedFolder":""}\n```')]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        result = email_app.analyze_thread(sample_emails, [], [])
        assert result["topic"] == "Finance"
        assert result["action"] == "read"

    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_normalizes_topic(self, mock_get_ai, sample_emails):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text=json.dumps({
            "summary": "test", "topic": "budget review",
            "action": "read", "urgency": "low",
            "suggestedReply": "", "suggestedFolder": "",
        }))]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        result = email_app.analyze_thread(sample_emails, [], [])
        assert result["topic"] == "Finance"

    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_error_returns_fallback(self, mock_get_ai, sample_emails):
        mock_get_ai.return_value.messages.create.side_effect = Exception("API error")

        result = email_app.analyze_thread(sample_emails, [], [])
        assert result["topic"] == "FYI & Updates"
        assert result["action"] == "read"
        assert result["urgency"] == "low"
        assert "Could not analyze" in result["summary"]

    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_no_json_returns_fallback(self, mock_get_ai, sample_emails):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text="I cannot help with that.")]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        result = email_app.analyze_thread(sample_emails, [], [])
        assert "Could not analyze" in result["summary"]

    @patch.object(email_app, "_get_ai")
    def test_analyze_thread_sorts_by_date(self, mock_get_ai):
        emails = [
            {"id": "2", "from_name": "B", "received_date_time": "2024-01-02", "body_preview": "second", "subject": "Test"},
            {"id": "1", "from_name": "A", "received_date_time": "2024-01-01", "body_preview": "first", "subject": "Test"},
        ]
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text=json.dumps({
            "summary": "test", "topic": "Engineering", "action": "read",
            "urgency": "low", "suggestedReply": "", "suggestedFolder": "",
        }))]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        email_app.analyze_thread(emails, [], [])
        call_args = mock_get_ai.return_value.messages.create.call_args
        prompt = call_args[1]["messages"][0]["content"]
        # "first" should appear before "second" in the prompt since sorted by date
        assert prompt.index("first") < prompt.index("second")


class TestFormatMessageWithAI:
    @patch.object(email_app, "_get_ai")
    def test_format_success(self, mock_get_ai):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text=json.dumps({
            "paragraphs": [
                {"text": "Hello", "intent": "Introduction", "emoji": "👋", "fact_concern": None},
                {"text": "Please review.", "intent": "Request", "emoji": "📋", "fact_concern": None},
            ]
        }))]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        result = email_app._format_message_with_ai({
            "body": "Hello\n\nPlease review.",
            "from_name": "Alice",
            "received_date_time": "2024-01-15T10:00:00Z",
        })
        assert len(result) == 2
        assert result[0]["intent"] == "Introduction"

    @patch.object(email_app, "_get_ai")
    def test_format_empty_body(self, mock_get_ai):
        result = email_app._format_message_with_ai({"body": "", "from_name": "X"})
        assert len(result) == 1
        assert result[0]["text"] == "(no content)"

    @patch.object(email_app, "_get_ai")
    def test_format_error_fallback(self, mock_get_ai):
        mock_get_ai.return_value.messages.create.side_effect = Exception("fail")

        result = email_app._format_message_with_ai({
            "body": "Para one\n\nPara two",
            "from_name": "Alice",
            "received_date_time": "2024-01-15",
        })
        assert len(result) == 2
        assert result[0]["text"] == "Para one"
        assert result[1]["text"] == "Para two"
        assert result[0]["intent"] == "FYI"


# ─── Flask API Route Tests ───────────────────────────────────────────────────

class TestAPIThreads:
    def test_threads_empty(self, client):
        resp = client.get("/api/threads")
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["groups"] == []
        assert data["threadCount"] == 0

    def test_threads_with_data(self, client, populated_db):
        resp = client.get("/api/threads")
        data = resp.get_json()
        assert data["threadCount"] == 1
        assert len(data["groups"]) == 1
        assert data["groups"][0]["topic"] == "Engineering"
        assert len(data["groups"][0]["threads"]) == 1

    def test_threads_grouped_by_topic(self, client, db):
        now = email_app._utcnow()
        for i, topic in enumerate(["Engineering", "Finance", "Engineering"]):
            key = f"thread-{i}"
            db.execute(
                "INSERT INTO threads (conversation_key,subject,topic,action,urgency,"
                "participants,email_ids,latest_received,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
                (key, f"Subject {i}", topic, "read", "low", "[]", "[]",
                 f"2024-01-{15+i}T10:00:00Z", now),
            )
        db.commit()
        resp = client.get("/api/threads")
        data = resp.get_json()
        assert data["threadCount"] == 3
        topics = [g["topic"] for g in data["groups"]]
        assert "Engineering" in topics
        assert "Finance" in topics


class TestAPIUpdates:
    def test_updates_empty(self, client):
        resp = client.get("/api/updates?since=2024-01-01T00:00:00Z")
        data = resp.get_json()
        assert data["threads"] == []

    def test_updates_returns_new(self, client, db):
        now = email_app._utcnow()
        db.execute(
            "INSERT INTO threads (conversation_key,subject,topic,action,urgency,"
            "participants,email_ids,latest_received,updated_at) VALUES(?,?,?,?,?,?,?,?,?)",
            ("k1", "Test", "Engineering", "read", "low", "[]", "[]", now, now),
        )
        db.commit()
        resp = client.get("/api/updates?since=2020-01-01T00:00:00Z")
        data = resp.get_json()
        assert len(data["threads"]) == 1

    def test_updates_no_since(self, client):
        resp = client.get("/api/updates")
        data = resp.get_json()
        assert "threads" in data


class TestAPISyncNow:
    @patch.object(email_app, "run_sync")
    def test_sync_now(self, mock_sync, client):
        email_app._sync_status["running"] = False
        resp = client.post("/api/sync_now")
        data = resp.get_json()
        assert data["ok"] is True

    def test_sync_now_already_running(self, client):
        email_app._sync_status["running"] = True
        resp = client.post("/api/sync_now")
        data = resp.get_json()
        assert data["ok"] is True
        email_app._sync_status["running"] = False


class TestAPIReanalyzeAll:
    def test_reanalyze_when_running(self, client):
        email_app._sync_status["running"] = True
        resp = client.post("/api/reanalyze_all")
        data = resp.get_json()
        assert data["ok"] is False
        assert "already running" in data["error"]
        email_app._sync_status["running"] = False

    @patch.object(email_app, "analyze_thread")
    def test_reanalyze_starts(self, mock_analyze, client):
        email_app._sync_status["running"] = False
        resp = client.post("/api/reanalyze_all")
        data = resp.get_json()
        assert data["ok"] is True


class TestAPIFolders:
    def test_folders_empty(self, client):
        resp = client.get("/api/folders")
        data = resp.get_json()
        assert "folders" in data
        assert "effortsFolders" in data

    def test_folders_with_data(self, client):
        email_app.meta_set("efforts_subfolders", json.dumps(["Efforts/A", "Efforts/B"]))
        email_app.meta_set("other_folders", json.dumps(["Archive"]))
        resp = client.get("/api/folders")
        data = resp.get_json()
        assert "Efforts/A" in data["folders"]
        assert "Efforts/B" in data["effortsFolders"]
        assert "Archive" in data["folders"]


class TestAPIThreadMessages:
    def test_no_ids(self, client):
        resp = client.get("/api/thread_messages")
        data = resp.get_json()
        assert data["messages"] == []

    @patch.object(email_app, "call_tool")
    def test_with_ids(self, mock_call, client, populated_db):
        mock_call.return_value = {
            "id": "msg-001",
            "from_name": "Alice",
            "from_address": "alice@example.com",
            "subject": "RE: Project Update",
            "received_date_time": "2024-01-15T10:00:00Z",
            "body_content": "<p>Hello from Outlook</p>",
        }
        resp = client.get("/api/thread_messages?id=msg-001")
        data = resp.get_json()
        assert len(data["messages"]) == 1
        assert data["messages"][0]["from_name"] == "Alice"

    @patch.object(email_app, "call_tool")
    def test_mcp_failure_uses_db_fallback(self, mock_call, client, populated_db):
        mock_call.side_effect = Exception("MCP error")
        resp = client.get("/api/thread_messages?id=msg-001")
        data = resp.get_json()
        assert len(data["messages"]) == 1
        # Should fall back to DB data
        assert data["messages"][0]["from_name"] == "Alice Smith"


class TestAPIFormatMessage:
    @patch.object(email_app, "call_tool")
    @patch.object(email_app, "_format_message_with_ai")
    def test_format_message(self, mock_format, mock_call, client, populated_db):
        mock_call.return_value = {
            "id": "msg-001", "from_name": "Alice",
            "body_content": "<p>Hello</p>",
            "received_date_time": "2024-01-15",
        }
        mock_format.return_value = [
            {"text": "Hello", "intent": "FYI", "emoji": "📄", "fact_concern": None}
        ]
        resp = client.get("/api/format_message?id=msg-001")
        data = resp.get_json()
        assert len(data["paragraphs"]) == 1

    def test_format_message_cached(self, client, db):
        cached = json.dumps([{"text": "Cached", "intent": "FYI", "emoji": "📄", "fact_concern": None}])
        db.execute(
            "INSERT INTO emails (id, conversation_key, formatted_body) VALUES(?,?,?)",
            ("cached-msg", "k", cached),
        )
        db.commit()
        resp = client.get("/api/format_message?id=cached-msg")
        data = resp.get_json()
        assert data["cached"] is True
        assert data["paragraphs"][0]["text"] == "Cached"


class TestAPIGenerateReply:
    @patch.object(email_app, "_get_ai")
    def test_generate_reply_success(self, mock_get_ai, client, populated_db):
        ai_response = MagicMock()
        ai_response.content = [MagicMock(text="Thanks for the update, team!")]
        mock_get_ai.return_value.messages.create.return_value = ai_response

        resp = client.post("/api/generate_reply", json={
            "conversationKey": populated_db,
            "userPrompt": "Acknowledge the update",
        })
        data = resp.get_json()
        assert "reply" in data
        assert data["reply"] == "Thanks for the update, team!"

    def test_generate_reply_no_prompt(self, client):
        resp = client.post("/api/generate_reply", json={
            "conversationKey": "k",
            "userPrompt": "",
        })
        assert resp.status_code == 400
        data = resp.get_json()
        assert "error" in data

    @patch.object(email_app, "_get_ai")
    def test_generate_reply_api_error(self, mock_get_ai, client, populated_db):
        mock_get_ai.return_value.messages.create.side_effect = Exception("API down")
        resp = client.post("/api/generate_reply", json={
            "conversationKey": populated_db,
            "userPrompt": "Test",
        })
        assert resp.status_code == 500


class TestAPIReply:
    @patch.object(email_app, "call_tool")
    def test_reply_success(self, mock_call, client, populated_db):
        mock_call.side_effect = [
            {"draft_id": "draft-123"},  # draft
            {},  # send
        ]
        resp = client.post("/api/reply/msg-002", json={
            "body": "Thanks!",
            "conversationKey": populated_db,
        })
        data = resp.get_json()
        assert data["ok"] is True
        # Thread should be removed
        db = email_app.get_db()
        assert db.execute("SELECT COUNT(*) FROM threads WHERE conversation_key=?",
                         (populated_db,)).fetchone()[0] == 0

    @patch.object(email_app, "call_tool")
    def test_reply_no_draft_id(self, mock_call, client):
        mock_call.return_value = {}
        resp = client.post("/api/reply/msg-002", json={"body": "Hi"})
        assert resp.status_code == 500

    @patch.object(email_app, "call_tool")
    def test_reply_with_to_cc(self, mock_call, client):
        mock_call.side_effect = [
            {"draft_id": "d1"},
            {},
        ]
        resp = client.post("/api/reply/msg-002", json={
            "body": "Hi",
            "to": ["alice@example.com"],
            "cc": ["bob@example.com"],
        })
        data = resp.get_json()
        assert data["ok"] is True
        draft_call = mock_call.call_args_list[0]
        assert draft_call[0][1]["to"] == ["alice@example.com"]
        assert draft_call[0][1]["cc"] == ["bob@example.com"]

    @patch.object(email_app, "call_tool")
    def test_reply_error(self, mock_call, client):
        mock_call.side_effect = Exception("MCP error")
        resp = client.post("/api/reply/msg-002", json={"body": "Hi"})
        assert resp.status_code == 500


class TestAPIDelete:
    @patch.object(email_app, "call_tool")
    def test_delete_success(self, mock_call, client, populated_db):
        mock_call.return_value = {}
        resp = client.post("/api/delete", json={
            "ids": ["msg-001", "msg-002"],
            "conversationKey": populated_db,
        })
        data = resp.get_json()
        assert data["ok"] is True
        db = email_app.get_db()
        assert db.execute("SELECT COUNT(*) FROM threads WHERE conversation_key=?",
                         (populated_db,)).fetchone()[0] == 0

    @patch.object(email_app, "call_tool")
    def test_delete_mcp_error_still_removes(self, mock_call, client, populated_db):
        mock_call.side_effect = Exception("MCP 404")
        resp = client.post("/api/delete", json={
            "ids": ["msg-001"],
            "conversationKey": populated_db,
        })
        data = resp.get_json()
        assert data["ok"] is True

    @patch.object(email_app, "call_tool")
    def test_delete_no_conv_key(self, mock_call, client):
        mock_call.return_value = {}
        resp = client.post("/api/delete", json={"ids": ["msg-001"]})
        data = resp.get_json()
        assert data["ok"] is True


class TestAPIMove:
    @patch.object(email_app, "call_tool")
    def test_move_success(self, mock_call, client, populated_db):
        mock_call.return_value = {}
        resp = client.post("/api/move", json={
            "ids": ["msg-001", "msg-002"],
            "folder": "Archive",
            "conversationKey": populated_db,
        })
        data = resp.get_json()
        assert data["ok"] is True
        db = email_app.get_db()
        assert db.execute("SELECT COUNT(*) FROM threads WHERE conversation_key=?",
                         (populated_db,)).fetchone()[0] == 0

    @patch.object(email_app, "call_tool")
    def test_move_error(self, mock_call, client):
        mock_call.side_effect = Exception("fail")
        resp = client.post("/api/move", json={
            "ids": ["msg-001"],
            "folder": "Archive",
        })
        data = resp.get_json()
        assert data["ok"] is False


class TestAPIMarkRead:
    @patch.object(email_app, "call_tool")
    def test_markread_success(self, mock_call, client, populated_db):
        mock_call.return_value = {}
        resp = client.post("/api/markread", json={
            "ids": ["msg-001", "msg-002"],
            "conversationKey": populated_db,
        })
        data = resp.get_json()
        assert data["ok"] is True
        db = email_app.get_db()
        thread = db.execute("SELECT * FROM threads WHERE conversation_key=?",
                           (populated_db,)).fetchone()
        assert thread["has_unread"] == 0
        email_row = db.execute("SELECT * FROM emails WHERE id='msg-001'").fetchone()
        assert email_row["is_read"] == 1

    @patch.object(email_app, "call_tool")
    def test_markread_error(self, mock_call, client):
        mock_call.side_effect = Exception("fail")
        resp = client.post("/api/markread", json={"ids": ["msg-001"], "conversationKey": "k"})
        assert resp.status_code == 500


class TestAPIFlag:
    def test_flag_thread(self, client, populated_db):
        resp = client.post("/api/flag", json={
            "conversationKey": populated_db,
            "flagged": True,
        })
        data = resp.get_json()
        assert data["ok"] is True
        assert data["isFlagged"] is True
        db = email_app.get_db()
        row = db.execute("SELECT is_flagged FROM threads WHERE conversation_key=?",
                        (populated_db,)).fetchone()
        assert row["is_flagged"] == 1

    def test_unflag_thread(self, client, populated_db):
        db = email_app.get_db()
        db.execute("UPDATE threads SET is_flagged=1 WHERE conversation_key=?", (populated_db,))
        db.commit()
        resp = client.post("/api/flag", json={
            "conversationKey": populated_db,
            "flagged": False,
        })
        data = resp.get_json()
        assert data["ok"] is True
        assert data["isFlagged"] is False

    def test_flag_missing_key(self, client):
        resp = client.post("/api/flag", json={"conversationKey": "", "flagged": True})
        assert resp.status_code == 400


class TestIndexRoute:
    def test_index_returns_html(self, client):
        resp = client.get("/")
        assert resp.status_code == 200
        assert b"Email Triage" in resp.data


# ─── Sync Engine Tests ────────────────────────────────────────────────────────

class TestSyncEngine:
    @patch.object(email_app, "call_tool")
    @patch.object(email_app, "analyze_thread")
    def test_do_sync_new_emails(self, mock_analyze, mock_call, db):
        mock_call.side_effect = [
            # _refresh_folders - list_folders
            {"folders": [{"display_name": "Inbox"}, {"display_name": "Archive"}]},
            # _do_sync - list_messages
            {"messages": [
                {
                    "id": "new-msg-1",
                    "subject": "New Email",
                    "from_name": "Dave",
                    "from_address": "dave@example.com",
                    "received_date_time": "2024-01-20T10:00:00Z",
                    "is_read": False,
                    "body_preview": "Hello!",
                },
            ]},
        ]
        mock_analyze.return_value = {
            "summary": "New email from Dave",
            "topic": "Engineering",
            "action": "reply",
            "urgency": "high",
            "suggestedReply": "Hi Dave!",
            "suggestedFolder": "",
        }

        added, updated = email_app._do_sync()
        assert added == 1
        assert updated == 1
        assert mock_analyze.called

    @patch.object(email_app, "call_tool")
    def test_do_sync_no_new_emails(self, mock_call, db, populated_db):
        mock_call.side_effect = [
            {"folders": []},  # folders
            {"messages": [
                {"id": "msg-001", "subject": "RE: Project Update"},
                {"id": "msg-002", "subject": "RE: Project Update"},
            ]},  # existing messages
        ]
        added, updated = email_app._do_sync()
        assert added == 0
        assert updated == 0

    @patch.object(email_app, "call_tool")
    def test_do_sync_empty_inbox(self, mock_call, db):
        mock_call.side_effect = [
            {"folders": []},
            {"messages": []},
        ]
        added, updated = email_app._do_sync()
        assert added == 0
        assert updated == 0

    @patch.object(email_app, "_do_sync")
    def test_run_sync_updates_status(self, mock_do_sync):
        mock_do_sync.return_value = (5, 3)
        email_app._sync_lock = threading.Lock()
        email_app.run_sync()
        assert email_app._sync_status["running"] is False
        assert email_app._sync_status["emailsAdded"] == 5
        assert email_app._sync_status["threadsUpdated"] == 3
        assert email_app._sync_status["lastSync"] is not None

    @patch.object(email_app, "_do_sync")
    def test_run_sync_handles_error(self, mock_do_sync):
        mock_do_sync.side_effect = Exception("sync failed")
        email_app._sync_lock = threading.Lock()
        email_app.run_sync()
        assert email_app._sync_status["running"] is False
        assert email_app._sync_status["lastError"] == "sync failed"

    @patch.object(email_app, "_do_sync")
    def test_run_sync_lock(self, mock_do_sync):
        """Test that concurrent syncs are prevented by the lock."""
        mock_do_sync.return_value = (0, 0)
        email_app._sync_lock = threading.Lock()
        email_app._sync_lock.acquire()
        # run_sync should return immediately without calling _do_sync
        email_app.run_sync()
        assert not mock_do_sync.called
        email_app._sync_lock.release()


class TestRefreshFolders:
    @patch.object(email_app, "call_tool")
    def test_refresh_with_efforts(self, mock_call):
        mock_call.side_effect = [
            {"folders": [
                {"display_name": "Inbox", "id": "id-inbox"},
                {"display_name": "Efforts", "id": "id-efforts"},
                {"display_name": "Archive", "id": "id-archive"},
                {"display_name": "Drafts", "id": "id-drafts"},
            ]},
            {"folders": [
                {"display_name": "Efforts/ProjectA"},
                {"display_name": "Efforts/ProjectB"},
            ]},
        ]
        efforts, other = email_app._refresh_folders()
        assert "Efforts/ProjectA" in efforts
        assert "Efforts/ProjectB" in efforts
        assert "Archive" in other
        assert "Inbox" not in other
        assert "Drafts" not in other

    @patch.object(email_app, "call_tool")
    def test_refresh_no_efforts(self, mock_call):
        mock_call.return_value = {"folders": [
            {"display_name": "Inbox"},
            {"display_name": "Archive"},
        ]}
        efforts, other = email_app._refresh_folders()
        assert efforts == []
        assert "Archive" in other

    @patch.object(email_app, "call_tool")
    def test_refresh_fallback_on_error(self, mock_call):
        mock_call.side_effect = Exception("MCP error")
        email_app.meta_set("efforts_subfolders", json.dumps(["cached-effort"]))
        email_app.meta_set("other_folders", json.dumps(["cached-other"]))
        efforts, other = email_app._refresh_folders()
        assert efforts == ["cached-effort"]
        assert other == ["cached-other"]


# ─── MCP call_tool Tests ─────────────────────────────────────────────────────

class TestCallTool:
    @patch.object(email_app, "_session_ready")
    def test_call_tool_timeout(self, mock_ready):
        mock_ready.wait.return_value = False
        with pytest.raises(RuntimeError, match="MCP session not ready"):
            email_app.call_tool("test_tool", {})

    @patch.object(email_app, "_session_ready")
    @patch.object(email_app, "_session")
    @patch.object(email_app, "_loop")
    def test_call_tool_error_result(self, mock_loop, mock_session, mock_ready):
        mock_ready.wait.return_value = True
        future = MagicMock()
        result = MagicMock()
        result.isError = True
        result.content = [MagicMock(text="something failed")]
        future.result.return_value = result

        with patch("asyncio.run_coroutine_threadsafe", return_value=future):
            with pytest.raises(RuntimeError, match="MCP error"):
                email_app.call_tool("test_tool", {})

    @patch.object(email_app, "_session_ready")
    @patch.object(email_app, "_session")
    @patch.object(email_app, "_loop")
    def test_call_tool_structured_content(self, mock_loop, mock_session, mock_ready):
        mock_ready.wait.return_value = True
        future = MagicMock()
        result = MagicMock()
        result.isError = False
        result.structuredContent = {"key": "value"}
        result.content = None
        future.result.return_value = result

        with patch("asyncio.run_coroutine_threadsafe", return_value=future):
            out = email_app.call_tool("test_tool", {})
            assert out == {"key": "value"}

    @patch.object(email_app, "_session_ready")
    @patch.object(email_app, "_session")
    @patch.object(email_app, "_loop")
    def test_call_tool_json_content(self, mock_loop, mock_session, mock_ready):
        mock_ready.wait.return_value = True
        future = MagicMock()
        result = MagicMock()
        result.isError = False
        result.structuredContent = None
        result.content = [MagicMock(text='{"messages": []}')]
        future.result.return_value = result

        with patch("asyncio.run_coroutine_threadsafe", return_value=future):
            out = email_app.call_tool("test_tool", {})
            assert out == {"messages": []}

    @patch.object(email_app, "_session_ready")
    @patch.object(email_app, "_session")
    @patch.object(email_app, "_loop")
    def test_call_tool_plain_text_content(self, mock_loop, mock_session, mock_ready):
        mock_ready.wait.return_value = True
        future = MagicMock()
        result = MagicMock()
        result.isError = False
        result.structuredContent = None
        result.content = [MagicMock(text="plain text response")]
        future.result.return_value = result

        with patch("asyncio.run_coroutine_threadsafe", return_value=future):
            out = email_app.call_tool("test_tool", {})
            assert out == "plain text response"

    @patch.object(email_app, "_session_ready")
    @patch.object(email_app, "_session")
    @patch.object(email_app, "_loop")
    def test_call_tool_no_content(self, mock_loop, mock_session, mock_ready):
        mock_ready.wait.return_value = True
        future = MagicMock()
        result = MagicMock()
        result.isError = False
        result.structuredContent = None
        result.content = None
        future.result.return_value = result

        with patch("asyncio.run_coroutine_threadsafe", return_value=future):
            out = email_app.call_tool("test_tool", {})
            assert out is None
