"""
test_routes_mail.py — Tests for routes/mail.py using Flask test client.

All MCP and AI calls are mocked. No real network or subprocess calls.
"""
import json
import sys
import os
from unittest.mock import patch, MagicMock

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)

from tests.conftest import make_mcp_message, make_mcp_draft_response


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ai_para_response(text="Hello world"):
    mock_resp = MagicMock()
    mock_content = MagicMock()
    mock_content.text = json.dumps({
        "paragraphs": [{"text": text, "intent": "FYI", "emoji": "📄", "fact_concern": None}]
    })
    mock_resp.content = [mock_content]
    return mock_resp


def _ai_summary_response(text="Summary text"):
    mock_resp = MagicMock()
    mock_content = MagicMock()
    mock_content.text = text
    mock_resp.content = [mock_content]
    return mock_resp


# ---------------------------------------------------------------------------
# GET /api/thread_messages
# ---------------------------------------------------------------------------

class TestThreadMessages:
    def test_returns_messages_by_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?id=email-1&id=email-2")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "messages" in data
        ids = {m["id"] for m in data["messages"]}
        assert "email-1" in ids
        assert "email-2" in ids

    def test_returns_messages_by_conversation_key(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?conversationKey=project+alpha+update")
        data = resp.get_json()
        assert resp.status_code == 200
        assert len(data["messages"]) == 3

    def test_returns_empty_when_no_ids(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["messages"] == []

    def test_messages_sorted_by_date_descending(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?conversationKey=project+alpha+update")
        data = resp.get_json()
        dates = [m["received_date_time"] for m in data["messages"]]
        assert dates == sorted(dates, reverse=True)


# ---------------------------------------------------------------------------
# GET /api/format_message
# ---------------------------------------------------------------------------

class TestFormatMessage:
    def test_returns_cached_formatted_body(self, client, db):
        """Returns cached formatted_body without calling AI."""
        paras = [{"text": "Cached para", "intent": "FYI", "emoji": "📄", "fact_concern": None}]
        db.execute(
            "UPDATE emails SET formatted_body=? WHERE id='email-1'",
            (json.dumps(paras),)
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            resp = client.get("/api/format_message?id=email-1")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["cached"] is True
        assert data["paragraphs"][0]["text"] == "Cached para"
        mock_ct.assert_not_called()

    def test_calls_ai_when_no_cache(self, client, db):
        """Calls AI and returns paragraphs when no cached formatted_body."""
        mock_resp = MagicMock()
        mock_content = MagicMock()
        mock_content.text = json.dumps({"paragraphs": [
            {"text": "AI para", "intent": "FYI", "emoji": "📄", "fact_concern": None}
        ]})
        mock_resp.content = [mock_content]

        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=make_mcp_message("email-1")), \
             patch("ai._get_ai") as mock_ai_fn:
            mock_ai_fn.return_value.messages.create.return_value = mock_resp
            resp = client.get("/api/format_message?id=email-1")

        data = resp.get_json()
        assert resp.status_code == 200
        assert "paragraphs" in data


# ---------------------------------------------------------------------------
# GET /api/message_recipients
# ---------------------------------------------------------------------------

class TestMessageRecipients:
    def test_returns_to_and_cc(self, client, db):
        mcp_resp = make_mcp_message(
            "email-1",
            to_recipients=[{"name": "Alice", "address": "alice@example.com"}],
            cc_recipients=[{"name": "Bob", "address": "bob@example.com"}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-1")
        data = resp.get_json()
        assert resp.status_code == 200
        assert len(data["to"]) == 1
        assert data["to"][0]["address"] == "alice@example.com"
        assert len(data["cc"]) == 1
        assert data["cc"][0]["address"] == "bob@example.com"

    def test_returns_empty_when_no_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/message_recipients")
        data = resp.get_json()
        assert data["to"] == []
        assert data["cc"] == []

    def test_handles_mcp_error_gracefully(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("MCP down")):
            resp = client.get("/api/message_recipients?id=email-1")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "error" in data

    def test_messages_wrapper_format(self, client, db):
        """Handles {'messages': [msg]} wrapper format from MCP."""
        mcp_resp = {"messages": [make_mcp_message(
            "email-1",
            to_recipients=[{"name": "Me", "address": "me@example.com"}],
        )]}
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-1")
        data = resp.get_json()
        assert len(data["to"]) >= 1


# ---------------------------------------------------------------------------
# GET /api/summarize_message
# ---------------------------------------------------------------------------

class TestSummarizeMessage:
    def test_returns_summary(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.summarize_message_ai", return_value="Short summary"):
            resp = client.get("/api/summarize_message?id=email-1")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["summary"] == "Short summary"

    def test_requires_id(self, client, db):
        resp = client.get("/api/summarize_message")
        assert resp.status_code == 400


# ---------------------------------------------------------------------------
# GET /api/top_contacts
# ---------------------------------------------------------------------------

class TestTopContacts:
    def test_returns_contacts_sorted_by_frequency(self, client, db):
        db.executemany(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            [
                ("a@ex.com", "Alice", 10, "2026-03-01"),
                ("b@ex.com", "Bob", 5, "2026-03-01"),
                ("c@ex.com", "Carol", 20, "2026-03-01"),
            ],
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts")
        data = resp.get_json()
        assert resp.status_code == 200
        freqs = [c["frequency"] for c in data["contacts"]]
        assert freqs == sorted(freqs, reverse=True)

    def test_n_parameter_limits_results(self, client, db):
        db.executemany(
            "INSERT INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
            [(f"{i}@ex.com", f"Person{i}", i, "2026-03-01") for i in range(10)],
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts?n=3")
        data = resp.get_json()
        assert len(data["contacts"]) <= 3


# ---------------------------------------------------------------------------
# POST /api/reply/<id>
# ---------------------------------------------------------------------------

class TestReply:
    def test_successful_reply(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread") as mock_rm:
            mock_ct.side_effect = [
                make_mcp_draft_response("draft-1"),  # draft call
                {"ok": True},                         # send call
            ]
            resp = client.post(
                "/api/reply/email-3",
                json={
                    "body": "Great, see you then!",
                    "conversationKey": "project alpha update",
                    "to": ["alice@example.com"],
                    "cc": [],
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True
        mock_rm.assert_called_once_with("project alpha update")

    def test_reply_returns_error_when_no_draft_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={}):
            resp = client.post(
                "/api/reply/email-3",
                json={"body": "Hello", "conversationKey": "project alpha update"},
            )
        data = resp.get_json()
        assert resp.status_code == 500
        assert "error" in data

    def test_reply_mcp_exception_returns_500(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("MCP error")):
            resp = client.post(
                "/api/reply/email-3",
                json={"body": "Hello", "conversationKey": "project alpha update"},
            )
        assert resp.status_code == 500


# ---------------------------------------------------------------------------
# POST /api/send_new
# ---------------------------------------------------------------------------

class TestSendNew:
    def test_successful_send(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = [
                make_mcp_draft_response("new-draft-1"),
                {"ok": True},
            ]
            resp = client.post(
                "/api/send_new",
                json={
                    "to": ["alice@example.com"],
                    "subject": "Hello",
                    "body": "This is a new email.",
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True

    def test_send_new_returns_error_when_no_draft_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={}):
            resp = client.post(
                "/api/send_new",
                json={"to": ["x@x.com"], "subject": "S", "body": "B"},
            )
        assert resp.status_code == 500


# ---------------------------------------------------------------------------
# POST /api/delete
# ---------------------------------------------------------------------------

class TestDelete:
    def test_deletes_and_removes_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            resp = client.post(
                "/api/delete",
                json={
                    "ids": ["email-3"],
                    "conversationKey": "project alpha update",
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True
        mock_rm.assert_called_once_with("project alpha update")

    def test_skips_inaccessible_messages(self, client, db):
        """MCP errors per-message are silently skipped (logged only)."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("inaccessible")), \
             patch("routes.mail.remove_thread"):
            resp = client.post(
                "/api/delete",
                json={"ids": ["email-1"], "conversationKey": "project alpha update"},
            )
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True


# ---------------------------------------------------------------------------
# POST /api/move
# ---------------------------------------------------------------------------

class TestMove:
    def test_moves_and_removes_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            resp = client.post(
                "/api/move",
                json={
                    "ids": ["email-3"],
                    "folder": "Efforts/Alpha",
                    "conversationKey": "project alpha update",
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        mock_rm.assert_called_once_with("project alpha update")


# ---------------------------------------------------------------------------
# POST /api/markread
# ---------------------------------------------------------------------------

class TestMarkRead:
    def test_marks_emails_read(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            resp = client.post(
                "/api/markread",
                json={
                    "ids": ["email-1"],
                    "conversationKey": "project alpha update",
                    "read": True,
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True

    def test_marks_read_by_conv_key_when_no_ids(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            resp = client.post(
                "/api/markread",
                json={"conversationKey": "project alpha update", "read": True},
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True

    def test_returns_ok_when_no_ids_and_no_key(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/markread", json={})
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True


# ---------------------------------------------------------------------------
# POST /api/flag
# ---------------------------------------------------------------------------

class TestFlag:
    def test_flags_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post(
                "/api/flag",
                json={"conversationKey": "project alpha update", "flagged": True},
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True
        assert data["isFlagged"] is True

    def test_unflags_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post(
                "/api/flag",
                json={"conversationKey": "project alpha update", "flagged": False},
            )
        data = resp.get_json()
        assert data["isFlagged"] is False

    def test_requires_conversation_key(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/flag", json={"flagged": True})
        assert resp.status_code == 400


# ---------------------------------------------------------------------------
# GET /api/my_email
# ---------------------------------------------------------------------------

class TestMyEmail:
    def test_returns_email_from_meta(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/my_email")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["email"] == "me@example.com"


# ---------------------------------------------------------------------------
# GET /api/people
# ---------------------------------------------------------------------------

class TestPeople:
    def test_returns_people_list(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "people" in data
        assert isinstance(data["people"], list)

    def test_excludes_my_email(self, client, db):
        # Add "me" as a from_address in the DB
        db.execute(
            "INSERT INTO emails (id,subject,from_name,from_address,received_date_time,"
            "is_read,body_preview,conversation_key,raw_json,synced_at,folder) "
            "VALUES('self-1','Test','Me','me@example.com','2026-01-01T00:00:00Z',"
            "1,'','key','{}','2026-01-01T00:00:00Z','Inbox')"
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people")
        data = resp.get_json()
        addresses = [p["address"] for p in data["people"]]
        assert "me@example.com" not in addresses

    def test_filters_by_q_param(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people?q=alice")
        data = resp.get_json()
        for p in data["people"]:
            assert "alice" in p["address"].lower() or "alice" in (p.get("name") or "").lower()


# ---------------------------------------------------------------------------
# GET /api/mailbox/folders
# ---------------------------------------------------------------------------

class TestMailboxFolders:
    def test_returns_folder_list(self, client, db):
        db.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('folders_raw', ?)",
            (json.dumps([
                {"display_name": "Inbox"},
                {"display_name": "Sent Items"},
                {"display_name": "Efforts"},
            ]),)
        )
        db.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('efforts_subfolders', ?)",
            (json.dumps(["Alpha", "Beta"]),)
        )
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.meta_get") as mock_mg:
            def _meta(key, default=""):
                if key == "folders_raw":
                    return json.dumps([
                        {"display_name": "Inbox"},
                        {"display_name": "Efforts"},
                    ])
                if key == "efforts_subfolders":
                    return json.dumps(["Alpha"])
                return default
            mock_mg.side_effect = _meta
            resp = client.get("/api/mailbox/folders")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "folders" in data


# ---------------------------------------------------------------------------
# GET /api/mailbox/folder
# ---------------------------------------------------------------------------

class TestMailboxFolder:
    def test_returns_threads_for_folder(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/mailbox/folder?folder=Inbox")
        data = resp.get_json()
        assert resp.status_code == 200
        assert "threads" in data
        assert "total" in data

    def test_empty_folder_param_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/mailbox/folder")
        data = resp.get_json()
        assert data["threads"] == []


# ---------------------------------------------------------------------------
# POST /api/rebuild_contacts
# ---------------------------------------------------------------------------

class TestRebuildContacts:
    def test_triggers_rebuild(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"), \
             patch("routes.mail.rebuild_contacts", return_value=3) as mock_rb:
            resp = client.post("/api/rebuild_contacts")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True
        assert data["count"] == 3
        mock_rb.assert_called_once_with("me@example.com")


# ---------------------------------------------------------------------------
# _embed_cid_images
# ---------------------------------------------------------------------------

class TestEmbedCidImages:
    def test_no_cid_no_external_returns_unchanged(self):
        from routes.mail import _embed_cid_images
        html = "<html><body><p>Hello</p></body></html>"
        result = _embed_cid_images(html)
        assert result == html

    def test_replaces_cid_with_blank_gif(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = '<html><body><img src="cid:abc123"></body></html>'
        result = _embed_cid_images(html)
        assert _BLANK_GIF in result
        assert "cid:" not in result.lower()

    def test_replaces_external_http_src_with_blank_gif(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = '<html><body><img src="https://example.com/track.png"></body></html>'
        result = _embed_cid_images(html)
        assert _BLANK_GIF in result
        assert "https://example.com" not in result

    def test_preserves_data_uri_src(self):
        from routes.mail import _embed_cid_images
        data_uri = "data:image/png;base64,abc123=="
        html = f'<html><body><img src="{data_uri}"></body></html>'
        result = _embed_cid_images(html)
        assert data_uri in result
