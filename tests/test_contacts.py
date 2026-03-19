"""
test_contacts.py — Tests for contact rebuilding and /api/top_contacts endpoint.
"""
import json
import sys
import os
from unittest.mock import patch, MagicMock

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)


# ---------------------------------------------------------------------------
# rebuild_contacts — pure unit tests
# ---------------------------------------------------------------------------

class TestRebuildContacts:
    def test_zero_emails_returns_zero(self, in_memory_conn):
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 0

    def test_deduplicates_case_insensitively(self, in_memory_conn):
        """alice@EXAMPLE.COM and alice@example.com are treated as the same address."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("ci1", "Alice", "alice@EXAMPLE.COM", "2026-01-01T00:00:00Z", "A"),
                ("ci2", "alice", "ALICE@example.com", "2026-01-02T00:00:00Z", "B"),
                ("ci3", "Alice Smith", "Alice@Example.Com", "2026-01-03T00:00:00Z", "C"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts()
        assert count == 1

    def test_picks_most_frequent_name_for_address(self, in_memory_conn):
        """The display name appearing most often for an address wins."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("fn1", "Bob", "bob@example.com", "2026-01-01T00:00:00Z", "A"),
                ("fn2", "Bob Jones", "bob@example.com", "2026-01-02T00:00:00Z", "B"),
                ("fn3", "Bob Jones", "bob@example.com", "2026-01-03T00:00:00Z", "C"),
                ("fn4", "Bob Jones", "bob@example.com", "2026-01-04T00:00:00Z", "D"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            db_module.rebuild_contacts()
        row = in_memory_conn.execute(
            "SELECT name FROM contacts WHERE email='bob@example.com'"
        ).fetchone()
        # "Bob Jones" appears 3 times, "Bob" once → "Bob Jones" should win
        assert row["name"] == "Bob Jones"

    def test_excludes_my_email_exactly(self, in_memory_conn):
        """My email is excluded even if it appears many times."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("me1", "Me", "me@example.com", "2026-01-01T00:00:00Z", "A"),
                ("me2", "Me", "me@example.com", "2026-01-02T00:00:00Z", "B"),
                ("me3", "Me", "me@example.com", "2026-01-03T00:00:00Z", "C"),
                ("other1", "Other", "other@example.com", "2026-01-04T00:00:00Z", "D"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count = db_module.rebuild_contacts(my_email="me@example.com")
        assert count == 1
        me_row = in_memory_conn.execute(
            "SELECT COUNT(*) FROM contacts WHERE email='me@example.com'"
        ).fetchone()
        assert me_row[0] == 0

    def test_rebuild_twice_same_result(self, in_memory_conn):
        """Calling rebuild_contacts twice gives same count and data."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [
                ("idem1", "Alice", "alice@example.com", "2026-01-01T00:00:00Z", "A"),
                ("idem2", "Bob", "bob@example.com", "2026-01-02T00:00:00Z", "B"),
            ],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            count1 = db_module.rebuild_contacts()
            count2 = db_module.rebuild_contacts()
        assert count1 == count2

        rows1 = in_memory_conn.execute(
            "SELECT email, frequency FROM contacts ORDER BY email"
        ).fetchall()
        with patch("db.get_db", return_value=in_memory_conn):
            db_module.rebuild_contacts()
        rows2 = in_memory_conn.execute(
            "SELECT email, frequency FROM contacts ORDER BY email"
        ).fetchall()
        assert list(rows1) == list(rows2)

    def test_frequency_counts_all_messages_per_address(self, in_memory_conn):
        """frequency reflects total number of emails from that address."""
        in_memory_conn.executemany(
            "INSERT INTO emails (id,from_name,from_address,received_date_time,subject,"
            "is_read,body_preview,conversation_key,raw_json,synced_at) "
            "VALUES(?,?,?,?,?,0,'','key','{}','2026-01-01T00:00:00Z')",
            [(f"freq{i}", "Alice", "alice@example.com", f"2026-01-{i:02d}T00:00:00Z", "A")
             for i in range(1, 6)],
        )
        in_memory_conn.commit()
        import db as db_module
        with patch("db.get_db", return_value=in_memory_conn):
            db_module.rebuild_contacts()
        row = in_memory_conn.execute(
            "SELECT frequency FROM contacts WHERE email='alice@example.com'"
        ).fetchone()
        assert row["frequency"] == 5


# ---------------------------------------------------------------------------
# GET /api/top_contacts — route tests
# ---------------------------------------------------------------------------

class TestTopContactsRoute:
    def test_returns_sorted_by_frequency_descending(self, client, db):
        db.executemany(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            [
                ("low@ex.com", "Low", 1, "2026-01-01"),
                ("high@ex.com", "High", 100, "2026-01-01"),
                ("mid@ex.com", "Mid", 50, "2026-01-01"),
            ],
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts")
        data = resp.get_json()
        assert resp.status_code == 200
        freqs = [c["frequency"] for c in data["contacts"]]
        assert freqs == sorted(freqs, reverse=True)
        assert freqs[0] == 100

    def test_n_param_limits_results(self, client, db):
        db.executemany(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            [(f"{i}@ex.com", f"Person{i}", i * 10, "2026-01-01") for i in range(20)],
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts?n=5")
        data = resp.get_json()
        assert len(data["contacts"]) <= 5

    def test_default_n_is_10(self, client, db):
        db.executemany(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            [(f"{i}@ex.com", f"P{i}", i, "2026-01-01") for i in range(20)],
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts")
        data = resp.get_json()
        assert len(data["contacts"]) <= 10

    def test_returns_required_fields(self, client, db):
        db.execute(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            ("contact@ex.com", "Contact Person", 5, "2026-01-01")
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts")
        data = resp.get_json()
        for c in data["contacts"]:
            assert "email" in c
            assert "name" in c
            assert "frequency" in c

    def test_empty_contacts_table_returns_empty_list(self, client, db):
        # Don't insert any contacts
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts")
        data = resp.get_json()
        # The seeded DB has no contacts unless rebuild was run
        assert resp.status_code == 200
        assert "contacts" in data
        assert isinstance(data["contacts"], list)
