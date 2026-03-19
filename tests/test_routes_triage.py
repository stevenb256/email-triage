"""
test_routes_triage.py — Tests for routes/triage.py using Flask test client.

All MCP and AI calls are mocked.
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
# GET /api/threads
# ---------------------------------------------------------------------------

class TestApiThreads:
    def test_returns_groups_threads_sync_status(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "groups" in data
        assert "syncStatus" in data
        assert "threadCount" in data
        assert "emailCount" in data

    def test_groups_threads_by_topic(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        # Each group has a topic and list of threads
        for group in data["groups"]:
            assert "topic" in group
            assert "threads" in group
            assert isinstance(group["threads"], list)

    def test_includes_latest_ts(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert "latestTs" in data

    def test_includes_next_meeting(self, client, db):
        nm_json = json.dumps({"id": "cal-1", "subject": "Weekly Sync"})
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value=nm_json):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert "nextMeeting" in data

    def test_empty_db_returns_empty_groups(self, client, db):
        """Works correctly when no threads exist (delete seeded threads)."""
        db.execute("DELETE FROM threads")
        db.execute("DELETE FROM emails")
        db.commit()
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["groups"] == []
        assert data["threadCount"] == 0


# ---------------------------------------------------------------------------
# GET /api/updates
# ---------------------------------------------------------------------------

class TestApiUpdates:
    def test_returns_threads_updated_after_since(self, client, db):
        since = "2026-03-15T00:00:00Z"
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get(f"/api/updates?since={since}")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "threads" in data
        assert "latestTs" in data
        assert "syncStatus" in data

    def test_returns_only_threads_after_since(self, client, db):
        # Use a future since — no threads should match
        since = "2099-01-01T00:00:00Z"
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get(f"/api/updates?since={since}")
        data = resp.get_json()
        assert data["threads"] == []

    def test_no_since_defaults_to_now(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/updates")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "threads" in data


# ---------------------------------------------------------------------------
# POST /api/suggested_reply
# ---------------------------------------------------------------------------

class TestSuggestedReply:
    def _make_analyze_result(self):
        return {
            "summary": "Facts||BREAK||None||BREAK||Reply by EOD",
            "topic": "Projects",
            "action": "reply",
            "urgency": "high",
            "suggestedReply": "Thanks for the update, Alice!",
            "suggestedFolder": "",
        }

    def test_returns_reply_for_valid_thread(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="[]"), \
             patch("routes.triage.analyze_thread", return_value=self._make_analyze_result()):
            resp = client.post(
                "/api/suggested_reply",
                json={"conversationKey": "project alpha update"},
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert "reply" in data
        assert data["reply"] == "Thanks for the update, Alice!"

    def test_caches_analysis_in_db(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="[]"), \
             patch("routes.triage.analyze_thread", return_value=self._make_analyze_result()):
            client.post(
                "/api/suggested_reply",
                json={"conversationKey": "project alpha update"},
            )
        # Verify the thread was updated in DB
        row = db.execute(
            "SELECT suggested_reply FROM threads WHERE conversation_key='project alpha update'"
        ).fetchone()
        assert row is not None
        assert row[0] == "Thanks for the update, Alice!"

    def test_returns_404_when_no_messages(self, client, db):
        with patch("routes.triage.get_db", return_value=db):
            resp = client.post(
                "/api/suggested_reply",
                json={"conversationKey": "nonexistent-thread"},
            )
        assert resp.status_code == 404

    def test_requires_conversation_key(self, client, db):
        with patch("routes.triage.get_db", return_value=db):
            resp = client.post("/api/suggested_reply", json={})
        assert resp.status_code == 400

    def test_passes_context_to_analyze(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="[]"), \
             patch("routes.triage.analyze_thread", return_value=self._make_analyze_result()) as mock_at:
            client.post(
                "/api/suggested_reply",
                json={"conversationKey": "project alpha update", "context": "User note"},
            )
        call_kwargs = mock_at.call_args
        assert call_kwargs[1].get("reply_context") == "User note" or \
               (len(call_kwargs[0]) > 3 and call_kwargs[0][3] == "User note")


# ---------------------------------------------------------------------------
# POST /api/resync_thread
# ---------------------------------------------------------------------------

class TestResyncThread:
    def test_returns_ok_and_starts_thread(self, client, db):
        import sync as sync_module
        original = dict(sync_module._sync_status)
        sync_module._sync_status["running"] = False
        try:
            with patch("routes.triage.get_db", return_value=db), \
                 patch("threading.Thread") as mock_thread:
                mock_thread.return_value.start = MagicMock()
                resp = client.post(
                    "/api/resync_thread",
                    json={"conversationKey": "project alpha update"},
                )
        finally:
            sync_module._sync_status.update(original)
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True

    def test_requires_conversation_key(self, client, db):
        with patch("routes.triage.get_db", return_value=db):
            resp = client.post("/api/resync_thread", json={})
        assert resp.status_code == 400

    def test_returns_409_when_sync_running(self, client, db):
        import sync as sync_module
        original = sync_module._sync_status["running"]
        sync_module._sync_status["running"] = True
        try:
            with patch("routes.triage.get_db", return_value=db):
                resp = client.post(
                    "/api/resync_thread",
                    json={"conversationKey": "project alpha update"},
                )
        finally:
            sync_module._sync_status["running"] = original
        assert resp.status_code == 409


# ---------------------------------------------------------------------------
# GET /api/folders  (served by routes/mail.py but logically triage context)
# ---------------------------------------------------------------------------

class TestApiFolders:
    def test_returns_folder_lists(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.meta_get") as mock_mg:
            mock_mg.side_effect = lambda k, d="": {
                "efforts_subfolders": '["Efforts/Alpha","Efforts/Beta"]',
                "other_folders": '["Partners","Inbox"]',
            }.get(k, d)
            resp = client.get("/api/folders")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "folders" in data
        assert "effortsFolders" in data
