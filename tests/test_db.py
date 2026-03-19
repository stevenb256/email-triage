"""
test_db.py — Tests for db.py database layer.

Each test uses an in-memory SQLite DB so no files are written to disk.
"""
import json
import sqlite3
import sys
import os
from unittest.mock import patch, MagicMock

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)


# ---------------------------------------------------------------------------
# init_db
# ---------------------------------------------------------------------------

class TestInitDb:
    def test_creates_all_tables(self, in_memory_conn):
        """init_db creates all expected tables."""
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            db.init_db()

        tables = {
            row[0] for row in in_memory_conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        assert "emails" in tables
        assert "threads" in tables
        assert "meta" in tables
        assert "calendar_events" in tables
        assert "contacts" in tables

    def test_creates_indexes(self, in_memory_conn):
        """init_db creates the expected indexes."""
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            db.init_db()

        indexes = {
            row[0] for row in in_memory_conn.execute(
                "SELECT name FROM sqlite_master WHERE type='index'"
            ).fetchall()
        }
        assert "idx_emails_conv_key" in indexes
        assert "idx_threads_updated" in indexes

    def test_idempotent_migrations(self, in_memory_conn):
        """Calling init_db twice does not raise (migrations are idempotent)."""
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            db.init_db()
            db.init_db()  # second call should not raise


# ---------------------------------------------------------------------------
# meta_get / meta_set
# ---------------------------------------------------------------------------

class TestMeta:
    def test_round_trip(self, in_memory_conn):
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            db.meta_set("test_key", "test_value")
            assert db.meta_get("test_key") == "test_value"

    def test_default_when_missing(self, in_memory_conn):
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            result = db.meta_get("nonexistent_key", default="fallback")
            assert result == "fallback"

    def test_none_default_when_missing(self, in_memory_conn):
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            result = db.meta_get("nonexistent_key")
            assert result is None

    def test_overwrite_existing(self, in_memory_conn):
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            db.meta_set("k", "v1")
            db.meta_set("k", "v2")
            assert db.meta_get("k") == "v2"

    def test_stores_json(self, in_memory_conn):
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            data = {"foo": [1, 2, 3]}
            db.meta_set("json_key", json.dumps(data))
            loaded = json.loads(db.meta_get("json_key"))
            assert loaded == data


# ---------------------------------------------------------------------------
# get_my_email
# ---------------------------------------------------------------------------

class TestGetMyEmail:
    def test_returns_cached_meta(self, in_memory_conn):
        """get_my_email returns the cached value from meta table."""
        in_memory_conn.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('my_email','cached@example.com')"
        )
        in_memory_conn.commit()
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            assert db.get_my_email() == "cached@example.com"

    def test_falls_back_to_sent_items(self, in_memory_conn):
        """get_my_email falls back to Sent Items query when meta is empty."""
        in_memory_conn.execute(
            "INSERT INTO emails (id,subject,from_name,from_address,received_date_time,"
            "is_read,body_preview,conversation_key,raw_json,synced_at,folder) "
            "VALUES('s1','Sent','Me','sent@example.com','2026-01-01T00:00:00Z',"
            "1,'','key','{}','2026-01-01T00:00:00Z','Sent Items')"
        )
        in_memory_conn.commit()
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            # Clear any cached my_email
            in_memory_conn.execute("DELETE FROM meta WHERE key='my_email'")
            in_memory_conn.commit()
            result = db.get_my_email()
            assert result == "sent@example.com"

    def test_returns_empty_when_no_data(self, in_memory_conn):
        """get_my_email returns '' when meta is empty and no Sent Items."""
        with patch("db.get_db", return_value=in_memory_conn):
            import db
            in_memory_conn.execute("DELETE FROM meta WHERE key='my_email'")
            in_memory_conn.commit()
            result = db.get_my_email()
            assert result == ""


# ---------------------------------------------------------------------------
# _thread_to_dict
# ---------------------------------------------------------------------------

class TestThreadToDict:
    def test_maps_all_fields(self, db):
        """_thread_to_dict correctly maps all fields from a thread row."""
        import db as db_module
        row = db.execute(
            "SELECT * FROM threads WHERE conversation_key='project alpha update'"
        ).fetchone()
        assert row is not None
        result = db_module._thread_to_dict(row)

        assert result["conversationKey"] == "project alpha update"
        assert result["subject"] == "Project Alpha Update"
        assert result["topic"] == "Projects"
        assert result["action"] == "reply"
        assert result["urgency"] == "high"
        assert isinstance(result["participants"], list)
        assert len(result["participants"]) == 2
        assert isinstance(result["emailIds"], list)
        assert len(result["emailIds"]) == 3
        assert result["latestId"] == "email-3"
        assert result["messageCount"] == 3
        assert result["hasUnread"] is True
        assert isinstance(result["isFlagged"], bool)
        assert result["suggestedReply"] == "Sounds good, see you Friday!"

    def test_handles_missing_participants_gracefully(self, in_memory_conn):
        """_thread_to_dict handles null participants JSON without raising."""
        in_memory_conn.execute(
            "INSERT INTO threads (conversation_key,subject,participants,email_ids) "
            "VALUES('test-key','Test',NULL,NULL)"
        )
        in_memory_conn.commit()
        import db as db_module
        row = in_memory_conn.execute(
            "SELECT * FROM threads WHERE conversation_key='test-key'"
        ).fetchone()
        result = db_module._thread_to_dict(row)
        assert result["participants"] == []
        assert result["emailIds"] == []

    def test_defaults_for_missing_fields(self, in_memory_conn):
        """_thread_to_dict applies sensible defaults for null fields."""
        in_memory_conn.execute(
            "INSERT INTO threads (conversation_key) VALUES('bare-key')"
        )
        in_memory_conn.commit()
        import db as db_module
        row = in_memory_conn.execute(
            "SELECT * FROM threads WHERE conversation_key='bare-key'"
        ).fetchone()
        result = db_module._thread_to_dict(row)
        assert result["topic"] == "General"
        assert result["action"] == "read"
        assert result["urgency"] == "low"
        assert result["summary"] == ""
        assert result["hasUnread"] is False


# ---------------------------------------------------------------------------
# rebuild_contacts
# ---------------------------------------------------------------------------

class TestRebuildContacts:
    def test_basic_count(self, in_memory_conn):
        """rebuild_contacts returns count of unique contacts."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("e1", "Alice", "alice@example.com", "2026-01-01T00:00:00Z", "Sub"),
                ("e2", "Bob", "bob@example.com", "2026-01-02T00:00:00Z", "Sub"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 2

    def test_deduplicates_by_email(self, in_memory_conn):
        """rebuild_contacts groups same email address together."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("e1", "Alice Smith", "alice@example.com", "2026-01-01T00:00:00Z", "A"),
                ("e2", "Alice Smith", "alice@example.com", "2026-01-02T00:00:00Z", "B"),
                ("e3", "Alice Smith", "alice@example.com", "2026-01-03T00:00:00Z", "C"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 1
        row = in_memory_conn.execute(
            "SELECT * FROM contacts WHERE email='alice@example.com'"
        ).fetchone()
        assert row is not None
        assert row["frequency"] == 3

    def test_deduplicates_case_insensitively(self, in_memory_conn):
        """rebuild_contacts treats alice@EXAMPLE.COM and alice@example.com as the same."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("e1", "Alice", "alice@EXAMPLE.COM", "2026-01-01T00:00:00Z", "A"),
                ("e2", "Alice", "alice@example.com", "2026-01-02T00:00:00Z", "B"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 1

    def test_picks_most_frequent_name(self, in_memory_conn):
        """rebuild_contacts picks the display name used most often for an address."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("e1", "Al Smith", "al@example.com", "2026-01-01T00:00:00Z", "A"),
                ("e2", "Al Smith", "al@example.com", "2026-01-02T00:00:00Z", "B"),
                ("e3", "Albert Smith", "al@example.com", "2026-01-03T00:00:00Z", "C"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            db_module.rebuild_contacts()
        row = in_memory_conn.execute(
            "SELECT name FROM contacts WHERE email='al@example.com'"
        ).fetchone()
        # "Al Smith" appears twice, "Albert Smith" once — pick "Al Smith"
        assert row["name"] == "Al Smith"

    def test_excludes_my_email(self, in_memory_conn):
        """rebuild_contacts excludes the current user's own email address."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("e1", "Me", "me@example.com", "2026-01-01T00:00:00Z", "A"),
                ("e2", "Other", "other@example.com", "2026-01-02T00:00:00Z", "B"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts(my_email="me@example.com")
        assert count == 1
        assert in_memory_conn.execute(
            "SELECT COUNT(*) FROM contacts WHERE email='me@example.com'"
        ).fetchone()[0] == 0

    def test_zero_emails_returns_zero(self, in_memory_conn):
        """rebuild_contacts with no emails returns 0."""
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 0

    def test_idempotent(self, in_memory_conn):
        """Calling rebuild_contacts twice gives same result."""
        in_memory_conn.execute(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES('e1','Alice','alice@example.com','2026-01-01T00:00:00Z','A',"
            "0,'','key','{}','2026-01-01T00:00:00Z')"
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count1 = db_module.rebuild_contacts()
            count2 = db_module.rebuild_contacts()
        assert count1 == count2 == 1


# ---------------------------------------------------------------------------
# remove_thread
# ---------------------------------------------------------------------------

class TestRemoveThread:
    def test_deletes_from_emails_and_threads(self, db):
        """remove_thread deletes rows from both emails and threads tables."""
        import db as db_module
        with patch("db.get_db", return_value=db):
            db_module.remove_thread("project alpha update")

        assert db.execute(
            "SELECT COUNT(*) FROM emails WHERE conversation_key='project alpha update'"
        ).fetchone()[0] == 0
        assert db.execute(
            "SELECT COUNT(*) FROM threads WHERE conversation_key='project alpha update'"
        ).fetchone()[0] == 0

    def test_non_existent_key_does_not_raise(self, db):
        """remove_thread with a non-existent key does not raise."""
        import db as db_module
        with patch("db.get_db", return_value=db):
            db_module.remove_thread("does-not-exist")  # should not raise
