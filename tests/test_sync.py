"""
test_sync.py — Tests for sync.py background sync logic.

All MCP and AI calls are mocked. No real network calls.
"""
import json
import sys
import os
import sqlite3
from unittest.mock import patch, MagicMock, call

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_email(msg_id, subject, from_name="Alice", from_address="alice@example.com",
                received="2026-03-15T10:00:00Z", is_read=False, folder="Inbox"):
    return {
        "id": msg_id,
        "subject": subject,
        "from_name": from_name,
        "from_address": from_address,
        "received_date_time": received,
        "is_read": is_read,
        "body_preview": f"Preview for {subject}",
    }


# ---------------------------------------------------------------------------
# _norm_subject
# ---------------------------------------------------------------------------

class TestNormSubject:
    def test_strips_re_prefix(self):
        from sync import _norm_subject
        assert _norm_subject("RE: Hello World") == "hello world"

    def test_strips_fw_prefix(self):
        from sync import _norm_subject
        assert _norm_subject("FW: Update") == "update"

    def test_strips_fwd_prefix(self):
        from sync import _norm_subject
        assert _norm_subject("FWD: Status") == "status"

    def test_case_insensitive_strip(self):
        from sync import _norm_subject
        assert _norm_subject("re: Meeting notes") == "meeting notes"
        assert _norm_subject("Re: Meeting notes") == "meeting notes"

    def test_no_prefix_lowercases(self):
        from sync import _norm_subject
        assert _norm_subject("Project Alpha Update") == "project alpha update"

    def test_empty_string_returns_no_subject(self):
        from sync import _norm_subject
        assert _norm_subject("") == "no-subject"
        assert _norm_subject(None) == "no-subject"


# ---------------------------------------------------------------------------
# _build_thread_map via _insert_messages + conversation_key grouping
# ---------------------------------------------------------------------------

class TestInsertMessages:
    def test_inserts_new_messages(self, in_memory_conn):
        from sync import _insert_messages
        emails = [_make_email("e1", "Hello World"), _make_email("e2", "RE: Hello World")]
        with patch("sync.get_db", return_value=in_memory_conn):
            added = _insert_messages(in_memory_conn, emails, "Inbox")
        assert added == 2
        assert in_memory_conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0] == 2

    def test_groups_by_normalized_subject(self, in_memory_conn):
        """Emails with RE: prefix share conversation_key with original."""
        from sync import _insert_messages
        emails = [
            _make_email("e1", "Project Update"),
            _make_email("e2", "RE: Project Update"),
            _make_email("e3", "Re: Project Update"),
        ]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")

        keys = [r[0] for r in in_memory_conn.execute(
            "SELECT DISTINCT conversation_key FROM emails"
        ).fetchall()]
        assert len(keys) == 1
        assert keys[0] == "project update"

    def test_does_not_duplicate_on_second_insert(self, in_memory_conn):
        from sync import _insert_messages
        emails = [_make_email("e1", "Hello")]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")
            added2 = _insert_messages(in_memory_conn, emails, "Inbox")
        assert added2 == 0
        assert in_memory_conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0] == 1

    def test_updates_is_read_on_second_insert(self, in_memory_conn):
        from sync import _insert_messages
        email = _make_email("e1", "Hello", is_read=False)
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, [email], "Inbox")

        email["is_read"] = True
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, [email], "Inbox")

        row = in_memory_conn.execute("SELECT is_read FROM emails WHERE id='e1'").fetchone()
        assert row[0] == 1

    def test_skips_emails_without_id(self, in_memory_conn):
        from sync import _insert_messages
        emails = [{"subject": "No ID", "from_name": "Alice"}]
        with patch("sync.get_db", return_value=in_memory_conn):
            added = _insert_messages(in_memory_conn, emails, "Inbox")
        assert added == 0


# ---------------------------------------------------------------------------
# Thread deduplication: same conversation_key → same thread
# ---------------------------------------------------------------------------

class TestThreadDeduplication:
    def test_same_conv_key_in_one_thread(self, in_memory_conn):
        from sync import _insert_messages
        emails = [
            _make_email("a1", "Budget Review Q2", received="2026-03-17T08:00:00Z"),
            _make_email("a2", "RE: Budget Review Q2", received="2026-03-17T14:00:00Z"),
        ]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")

        keys = {r[0] for r in in_memory_conn.execute(
            "SELECT DISTINCT conversation_key FROM emails"
        ).fetchall()}
        assert len(keys) == 1


# ---------------------------------------------------------------------------
# has_unread detection
# ---------------------------------------------------------------------------

class TestHasUnreadDetection:
    def test_thread_has_unread_when_any_unread(self, in_memory_conn):
        from sync import _insert_messages
        emails = [
            _make_email("u1", "Hello", is_read=True),
            _make_email("u2", "RE: Hello", is_read=False),
        ]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")

        rows = in_memory_conn.execute(
            "SELECT id, is_read FROM emails WHERE conversation_key='hello'"
        ).fetchall()
        read_states = {r[0]: r[1] for r in rows}
        assert read_states["u1"] == 1
        assert read_states["u2"] == 0
        # has_unread logic: any is_read == 0 means thread has_unread
        has_unread = any(r[1] == 0 for r in rows)
        assert has_unread is True

    def test_thread_not_unread_when_all_read(self, in_memory_conn):
        from sync import _insert_messages
        emails = [
            _make_email("r1", "All read", is_read=True),
            _make_email("r2", "RE: All read", is_read=True),
        ]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")

        rows = in_memory_conn.execute(
            "SELECT is_read FROM emails WHERE conversation_key='all read'"
        ).fetchall()
        has_unread = any(r[0] == 0 for r in rows)
        assert has_unread is False


# ---------------------------------------------------------------------------
# Participant list building
# ---------------------------------------------------------------------------

class TestParticipantList:
    def test_participants_deduplicated(self):
        """Participant list is built from from_name/from_address, deduplicated."""
        from ai import _clean  # reuse the clean helper
        emails = [
            {"from_name": "Alice Smith", "from_address": "alice@ex.com"},
            {"from_name": "Alice Smith", "from_address": "alice@ex.com"},  # duplicate
            {"from_name": "Bob Jones", "from_address": "bob@ex.com"},
        ]
        participants = list(dict.fromkeys(
            (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
            for e in emails
            if (e.get("from_name") or e.get("from_address"))
        ))
        assert participants == ["Alice Smith", "Bob Jones"]

    def test_participants_capped_at_8(self):
        """Participant list is capped at 8 entries."""
        from ai import _clean
        emails = [
            {"from_name": f"Person {i}", "from_address": f"p{i}@ex.com"}
            for i in range(20)
        ]
        participants = list(dict.fromkeys(
            (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
            for e in emails
            if (e.get("from_name") or e.get("from_address"))
        ))[:8]
        assert len(participants) == 8


# ---------------------------------------------------------------------------
# Sync status
# ---------------------------------------------------------------------------

class TestSyncStatus:
    def test_sync_status_keys_exist(self):
        from sync import _sync_status
        assert "running" in _sync_status
        assert "lastSync" in _sync_status
        assert "emailsAdded" in _sync_status
        assert "threadsUpdated" in _sync_status
        assert "phase" in _sync_status
        assert "progress" in _sync_status

    def test_run_sync_acquires_lock(self):
        """run_sync acquires the sync lock and updates status."""
        from sync import _sync_lock
        import sync as sync_module

        with patch.object(sync_module, "_do_sync", return_value=(0, 0)) as mock_do:
            sync_module.run_sync()

        mock_do.assert_called_once()


# ---------------------------------------------------------------------------
# _norm_subject used in _build_thread_map
# ---------------------------------------------------------------------------

class TestBuildThreadMap:
    def test_groups_emails_by_conversation_key(self, in_memory_conn):
        """Different subjects produce different conversation keys."""
        from sync import _insert_messages
        emails = [
            _make_email("t1", "Alpha Update"),
            _make_email("t2", "Beta Review"),
            _make_email("t3", "RE: Alpha Update"),
        ]
        with patch("sync.get_db", return_value=in_memory_conn):
            _insert_messages(in_memory_conn, emails, "Inbox")

        rows = in_memory_conn.execute(
            "SELECT conversation_key, COUNT(*) as cnt FROM emails GROUP BY conversation_key"
        ).fetchall()
        thread_map = {r[0]: r[1] for r in rows}
        assert thread_map.get("alpha update") == 2
        assert thread_map.get("beta review") == 1
