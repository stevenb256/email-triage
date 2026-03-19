"""
test_ux_flows.py — Comprehensive UX flow tests for Outlook Express.

Covers every user-facing API interaction: reply dialog, compose, delete, move,
flag, search, thread loading, format stream, suggested reply, triage, and
mailbox navigation.

IMPORTANT: all call_tool usages are mocked — no real emails are sent.
"""
import json
import sys
import os
from unittest.mock import patch, MagicMock, call as mock_call

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)

from tests.conftest import make_mcp_message, make_mcp_draft_response


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _draft_then_send(draft_id="draft-001"):
    """Side-effect list: draft succeeds, send succeeds."""
    return [make_mcp_draft_response(draft_id), {"ok": True}]


def _get_reply_draft_args(mock_ct):
    """Return the kwargs dict passed to the draft call_tool invocation."""
    return mock_ct.call_args_list[0][0][1]


# ══════════════════════════════════════════════════════════════════════════════
# REPLY DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class TestReplyDialogRecipients:
    """Tests for /api/message_recipients — the endpoint that populates
    the reply modal's To/CC fields."""

    def test_empty_id_returns_empty_lists(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/message_recipients")
        data = resp.get_json()
        assert resp.status_code == 200
        assert data == {"to": [], "cc": []}

    def test_returns_to_and_cc(self, client, db):
        mcp_resp = make_mcp_message(
            "email-3",
            to_recipients=[{"name": "Alice", "address": "alice@example.com"}],
            cc_recipients=[{"name": "Bob", "address": "bob@example.com"}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-3")
        data = resp.get_json()
        assert [r["address"] for r in data["to"]] == ["alice@example.com"]
        assert [r["address"] for r in data["cc"]] == ["bob@example.com"]

    def test_mcp_exception_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("network error")):
            resp = client.get("/api/message_recipients?id=email-3")
        data = resp.get_json()
        assert resp.status_code == 200          # soft failure — not 500
        assert data["to"] == []
        assert data["cc"] == []
        assert "error" in data

    def test_messages_wrapper_format(self, client, db):
        """MCP wraps message in {'messages': [...]}."""
        inner = make_mcp_message(
            "email-3",
            to_recipients=[{"name": "X", "address": "x@example.com"}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"messages": [inner]}):
            resp = client.get("/api/message_recipients?id=email-3")
        assert resp.get_json()["to"][0]["address"] == "x@example.com"

    def test_nested_emailaddress_format(self, client, db):
        """Microsoft Graph nested emailAddress dict format."""
        mcp_resp = make_mcp_message(
            "email-3",
            to_recipients=[{"emailAddress": {"name": "Y", "address": "y@example.com"}}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-3")
        assert resp.get_json()["to"][0]["address"] == "y@example.com"

    def test_empty_messages_list_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"messages": []}):
            resp = client.get("/api/message_recipients?id=email-3")
        assert resp.get_json() == {"to": [], "cc": []}

    def test_preserves_recipient_display_names(self, client, db):
        mcp_resp = make_mcp_message(
            "email-3",
            to_recipients=[{"name": "Full Name", "address": "full@example.com"}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-3")
        recip = resp.get_json()["to"][0]
        assert recip["name"] == "Full Name"
        assert recip["address"] == "full@example.com"


class TestReplyDialogSend:
    """Tests for POST /api/reply/<id> — the actual send step."""

    def test_basic_reply_sends_to_recipients(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send("d1")
            resp = client.post("/api/reply/email-3", json={
                "body": "Thanks!",
                "conversationKey": "project alpha update",
                "to": ["alice@example.com"],
                "cc": [],
            })
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True

    def test_reply_calls_draft_with_correct_body(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": "My exact reply", "conversationKey": "c"})
        assert _get_reply_draft_args(mock_ct)["bodyText"] == "My exact reply"

    def test_reply_uses_replyall_operation(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": "Reply", "conversationKey": "c"})
        assert _get_reply_draft_args(mock_ct)["operation"] == "ReplyAll"

    def test_reply_uses_source_message_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c"})
        assert _get_reply_draft_args(mock_ct)["source_message_id"] == "email-3"

    def test_reply_includes_to_when_provided(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={
                "body": "r", "conversationKey": "c",
                "to": ["alice@example.com", "bob@example.com"],
            })
        args = _get_reply_draft_args(mock_ct)
        assert args["to"] == ["alice@example.com", "bob@example.com"]

    def test_reply_includes_cc_when_provided(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={
                "body": "r", "conversationKey": "c",
                "cc": ["carol@example.com"],
            })
        assert _get_reply_draft_args(mock_ct).get("cc") == ["carol@example.com"]

    def test_reply_omits_to_field_when_empty(self, client, db):
        """Empty to list → no explicit 'to' in draft args → Outlook uses ReplyAll logic."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c", "to": []})
        assert "to" not in _get_reply_draft_args(mock_ct)

    def test_reply_removes_thread_from_db_on_success(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread") as mock_rm:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={
                "body": "r", "conversationKey": "project alpha update",
            })
        mock_rm.assert_called_once_with("project alpha update")

    def test_reply_does_not_remove_thread_without_conv_key(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread") as mock_rm:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": "r"})
        mock_rm.assert_not_called()

    def test_reply_draft_id_from_id_field(self, client, db):
        """Draft response uses 'id' not 'draft_id'."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = [{"id": "d-from-id"}, {"ok": True}]
            resp = client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c"})
        assert resp.status_code == 200

    def test_reply_draft_id_from_widget_state(self, client, db):
        """Draft response nests draftId inside widgetState."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = [{"widgetState": {"draftId": "d-widget"}}, {"ok": True}]
            resp = client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c"})
        assert resp.status_code == 200

    def test_reply_returns_500_when_no_draft_id(self, client, db):
        """MCP returns something but no recognisable draft ID."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"status": "queued"}):
            resp = client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c"})
        assert resp.status_code == 500
        assert "error" in resp.get_json()

    def test_reply_returns_500_on_draft_exception(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("draft failed")):
            resp = client.post("/api/reply/email-3", json={"body": "r", "conversationKey": "c"})
        assert resp.status_code == 500

    def test_reply_returns_500_on_send_exception(self, client, db):
        """Draft succeeds but send raises — should 500 and NOT remove thread."""
        def _ct(tool, args):
            if tool == "outlook_mail_draft_message":
                return make_mcp_draft_response("d1")
            raise Exception("send failed")

        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=_ct), \
             patch("routes.mail.remove_thread") as mock_rm:
            resp = client.post("/api/reply/email-3", json={
                "body": "r", "conversationKey": "project alpha update",
            })
        assert resp.status_code == 500
        mock_rm.assert_not_called()  # thread must NOT be removed on send failure

    def test_reply_not_sent_when_body_empty(self, client, db):
        """Frontend prevents empty sends; backend should draft an empty body anyway."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            resp = client.post("/api/reply/email-3", json={"body": "", "conversationKey": "c"})
        # Backend accepts empty body (validation is frontend's job)
        assert resp.status_code == 200
        assert _get_reply_draft_args(mock_ct)["bodyText"] == ""

    def test_reply_with_multiline_body(self, client, db):
        body = "Line 1\n\nLine 2\n\nLine 3"
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": body, "conversationKey": "c"})
        assert _get_reply_draft_args(mock_ct)["bodyText"] == body

    def test_reply_thread_not_removed_on_draft_failure(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("fail")), \
             patch("routes.mail.remove_thread") as mock_rm:
            client.post("/api/reply/email-3", json={
                "body": "r", "conversationKey": "project alpha update",
            })
        mock_rm.assert_not_called()

    def test_two_independent_replies_to_different_threads(self, client, db):
        """Sequential replies to two different threads don't interfere."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = (
                _draft_then_send("d1") +
                _draft_then_send("d2")
            )
            r1 = client.post("/api/reply/email-3", json={"body": "Reply A", "conversationKey": "conv-a"})
            r2 = client.post("/api/reply/email-4", json={"body": "Reply B", "conversationKey": "conv-b"})
        assert r1.get_json()["ok"] is True
        assert r2.get_json()["ok"] is True
        # First draft call used email-3, second used email-4
        assert mock_ct.call_args_list[0][0][1]["source_message_id"] == "email-3"
        assert mock_ct.call_args_list[2][0][1]["source_message_id"] == "email-4"

    def test_reply_with_special_chars_in_body(self, client, db):
        body = 'Hi "Alice" & <Bob> — see you at 5pm!'
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/reply/email-3", json={"body": body, "conversationKey": "c"})
        assert _get_reply_draft_args(mock_ct)["bodyText"] == body


# ══════════════════════════════════════════════════════════════════════════════
# COMPOSE NEW MESSAGE
# ══════════════════════════════════════════════════════════════════════════════

class TestComposeNewMessage:
    def test_basic_send_succeeds(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            resp = client.post("/api/send_new", json={
                "to": ["alice@example.com"],
                "subject": "Hello",
                "body": "Hi there",
            })
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True

    def test_send_uses_new_operation(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/send_new", json={"to": ["a@b.com"], "subject": "S", "body": "B"})
        assert mock_ct.call_args_list[0][0][1]["operation"] == "New"

    def test_send_passes_subject_and_body(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/send_new", json={"to": ["x@y.com"], "subject": "My Subject", "body": "My Body"})
        args = mock_ct.call_args_list[0][0][1]
        assert args["subject"] == "My Subject"
        assert args["bodyText"] == "My Body"

    def test_send_with_cc(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/send_new", json={
                "to": ["a@b.com"], "cc": ["c@d.com"], "subject": "S", "body": "B",
            })
        assert mock_ct.call_args_list[0][0][1].get("cc") == ["c@d.com"]

    def test_send_without_cc_omits_cc_field(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            client.post("/api/send_new", json={"to": ["a@b.com"], "subject": "S", "body": "B"})
        assert "cc" not in mock_ct.call_args_list[0][0][1]

    def test_send_fails_when_no_draft_id(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"status": "unknown"}):
            resp = client.post("/api/send_new", json={"to": ["a@b.com"], "subject": "S", "body": "B"})
        assert resp.status_code == 500

    def test_send_fails_on_mcp_exception(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("MCP down")):
            resp = client.post("/api/send_new", json={"to": ["a@b.com"], "subject": "S", "body": "B"})
        assert resp.status_code == 500
        assert "error" in resp.get_json()

    def test_send_with_empty_to_list(self, client, db):
        """Empty to list is passed through; MCP / Outlook handles validation."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            mock_ct.side_effect = _draft_then_send()
            resp = client.post("/api/send_new", json={"to": [], "subject": "S", "body": "B"})
        # 'to' should NOT be included when empty (consistent with reply behaviour)
        assert "to" not in mock_ct.call_args_list[0][0][1]


# ══════════════════════════════════════════════════════════════════════════════
# DELETE FLOW
# ══════════════════════════════════════════════════════════════════════════════

class TestDeleteFlow:
    def test_delete_single_message_calls_mcp(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.return_value = {"ok": True}
            resp = client.post("/api/delete", json={
                "ids": ["email-3"],
                "conversationKey": "project alpha update",
            })
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True
        mock_ct.assert_called_once_with(
            "outlook_mail_move_message",
            {"message_id": "email-3", "destination_folder": "Deleted Items"},
        )

    def test_delete_multiple_messages_calls_mcp_for_each(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.return_value = {"ok": True}
            client.post("/api/delete", json={
                "ids": ["email-1", "email-2", "email-3"],
                "conversationKey": "project alpha update",
            })
        assert mock_ct.call_count == 3
        moved_ids = [c[0][1]["message_id"] for c in mock_ct.call_args_list]
        assert set(moved_ids) == {"email-1", "email-2", "email-3"}

    def test_delete_removes_thread_from_db(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            client.post("/api/delete", json={
                "ids": ["email-3"],
                "conversationKey": "project alpha update",
            })
        mock_rm.assert_called_once_with("project alpha update")

    def test_delete_without_conv_key_does_not_remove_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            client.post("/api/delete", json={"ids": ["email-3"]})
        mock_rm.assert_not_called()

    def test_delete_continues_after_mcp_failure(self, client, db):
        """One message failing to delete should not prevent others from being deleted."""
        call_results = [Exception("inaccessible"), {"ok": True}]
        def _ct(tool, args):
            result = call_results.pop(0)
            if isinstance(result, Exception):
                raise result
            return result

        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=_ct), \
             patch("routes.mail.remove_thread") as mock_rm:
            resp = client.post("/api/delete", json={
                "ids": ["email-bad", "email-3"],
                "conversationKey": "project alpha update",
            })
        # Returns ok regardless of individual failures
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True
        mock_rm.assert_called_once_with("project alpha update")

    def test_delete_empty_ids_list(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            resp = client.post("/api/delete", json={"ids": [], "conversationKey": "c"})
        assert resp.status_code == 200
        mock_ct.assert_not_called()


# ══════════════════════════════════════════════════════════════════════════════
# MOVE / FILE FLOW
# ══════════════════════════════════════════════════════════════════════════════

class TestMoveFlow:
    def test_move_single_message(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.return_value = {"ok": True}
            resp = client.post("/api/move", json={
                "ids": ["email-3"],
                "folder": "Efforts/Alpha",
                "conversationKey": "project alpha update",
            })
        assert resp.status_code == 200
        assert resp.get_json()["ok"] is True
        mock_ct.assert_called_once_with(
            "outlook_mail_move_message",
            {"message_id": "email-3", "destination_folder": "Efforts/Alpha"},
        )

    def test_move_multiple_messages(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.return_value = {"ok": True}
            client.post("/api/move", json={
                "ids": ["email-1", "email-2"],
                "folder": "Archive",
                "conversationKey": "c",
            })
        assert mock_ct.call_count == 2

    def test_move_removes_thread_from_db(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            client.post("/api/move", json={
                "ids": ["email-3"],
                "folder": "Archive",
                "conversationKey": "project alpha update",
            })
        mock_rm.assert_called_once_with("project alpha update")

    def test_move_returns_not_ok_on_all_failures(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("fail")), \
             patch("routes.mail.remove_thread"):
            resp = client.post("/api/move", json={
                "ids": ["email-3"],
                "folder": "Archive",
                "conversationKey": "c",
            })
        # When all moves fail, ok should be False
        assert resp.get_json()["ok"] is False

    def test_move_without_conv_key_does_not_remove_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}), \
             patch("routes.mail.remove_thread") as mock_rm:
            client.post("/api/move", json={"ids": ["email-3"], "folder": "Archive"})
        mock_rm.assert_not_called()


# ══════════════════════════════════════════════════════════════════════════════
# FLAG OPERATIONS
# ══════════════════════════════════════════════════════════════════════════════

class TestFlagOperations:
    def test_flag_thread(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/flag", json={
                "conversationKey": "project alpha update",
                "flagged": True,
            })
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["ok"] is True
        assert data["isFlagged"] is True
        row = db.execute("SELECT is_flagged FROM threads WHERE conversation_key=?",
                         ("project alpha update",)).fetchone()
        assert row["is_flagged"] == 1

    def test_unflag_thread(self, client, db):
        db.execute("UPDATE threads SET is_flagged=1 WHERE conversation_key=?",
                   ("project alpha update",))
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/flag", json={
                "conversationKey": "project alpha update",
                "flagged": False,
            })
        assert resp.get_json()["isFlagged"] is False
        row = db.execute("SELECT is_flagged FROM threads WHERE conversation_key=?",
                         ("project alpha update",)).fetchone()
        assert row["is_flagged"] == 0

    def test_flag_missing_conv_key_returns_400(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/flag", json={"flagged": True})
        assert resp.status_code == 400

    def test_flag_reflected_in_threads_list(self, client, db):
        """After flagging, GET /api/threads shows isFlagged=True."""
        db.execute("UPDATE threads SET is_flagged=1 WHERE conversation_key=?",
                   ("project alpha update",))
        db.commit()
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        threads = [t for g in resp.get_json()["groups"] for t in g["threads"]]
        alpha = next(t for t in threads if t["conversationKey"] == "project alpha update")
        assert alpha["isFlagged"] is True


# ══════════════════════════════════════════════════════════════════════════════
# MARK READ / UNREAD
# ══════════════════════════════════════════════════════════════════════════════

class TestMarkRead:
    def test_mark_read_by_ids(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            resp = client.post("/api/markread", json={
                "ids": ["email-1"],
                "conversationKey": "project alpha update",
                "read": True,
            })
        assert resp.status_code == 200
        row = db.execute("SELECT is_read FROM emails WHERE id='email-1'").fetchone()
        assert row["is_read"] == 1

    def test_mark_read_by_conv_key(self, client, db):
        """No ids provided — looks up by conversationKey."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            resp = client.post("/api/markread", json={
                "conversationKey": "project alpha update",
                "read": True,
            })
        assert resp.status_code == 200

    def test_mark_unread(self, client, db):
        db.execute("UPDATE emails SET is_read=1 WHERE id='email-1'")
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            client.post("/api/markread", json={"ids": ["email-1"], "read": False})
        row = db.execute("SELECT is_read FROM emails WHERE id='email-1'").fetchone()
        assert row["is_read"] == 0

    def test_mark_read_updates_thread_has_unread(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"ok": True}):
            client.post("/api/markread", json={
                "ids": ["email-1"],
                "conversationKey": "project alpha update",
                "read": True,
            })
        row = db.execute("SELECT has_unread FROM threads WHERE conversation_key=?",
                         ("project alpha update",)).fetchone()
        assert row["has_unread"] == 0

    def test_mark_read_no_ids_returns_ok(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.post("/api/markread", json={})
        assert resp.get_json()["ok"] is True


# ══════════════════════════════════════════════════════════════════════════════
# SEARCH
# ══════════════════════════════════════════════════════════════════════════════

class TestSearch:
    def test_short_query_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=a")
        data = resp.get_json()
        assert data["results"] == []
        assert data["count"] == 0

    def test_search_by_subject(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=Alpha")
        data = resp.get_json()
        assert data["count"] > 0
        assert all("alpha" in r["subject"].lower() for r in data["results"])

    def test_search_by_sender_name(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=Alice")
        data = resp.get_json()
        assert data["count"] > 0
        assert all("alice" in (r["from_name"] or "").lower() for r in data["results"])

    def test_search_by_email_address(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=carol@example.com")
        data = resp.get_json()
        assert data["count"] >= 1

    def test_search_by_body_preview(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=on+track")
        data = resp.get_json()
        assert data["count"] >= 1

    def test_search_returns_correct_fields(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=Budget")
        results = resp.get_json()["results"]
        assert len(results) > 0
        r = results[0]
        assert "id" in r
        assert "subject" in r
        assert "from_name" in r
        assert "received_date_time" in r

    def test_search_no_match_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=ZZZNOMATCH9999")
        assert resp.get_json()["count"] == 0

    def test_search_sorted_by_recency(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/search?q=Project+Alpha")
        results = resp.get_json()["results"]
        dates = [r["received_date_time"] for r in results]
        assert dates == sorted(dates, reverse=True)


# ══════════════════════════════════════════════════════════════════════════════
# THREAD MESSAGES
# ══════════════════════════════════════════════════════════════════════════════

class TestThreadMessages:
    def test_load_by_ids(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?id=email-1&id=email-2")
        data = resp.get_json()
        assert len(data["messages"]) == 2
        ids = {m["id"] for m in data["messages"]}
        assert ids == {"email-1", "email-2"}

    def test_load_by_conversation_key(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?conversationKey=project+alpha+update")
        data = resp.get_json()
        assert len(data["messages"]) == 3

    def test_no_params_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages")
        assert resp.get_json() == {"messages": []}

    def test_messages_sorted_newest_first(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?id=email-1&id=email-2&id=email-3")
        msgs = resp.get_json()["messages"]
        dates = [m["received_date_time"] for m in msgs]
        assert dates == sorted(dates, reverse=True)

    def test_messages_include_standard_fields(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?id=email-1")
        msg = resp.get_json()["messages"][0]
        for field in ["id", "subject", "from_name", "from_address", "received_date_time"]:
            assert field in msg

    def test_unknown_id_returns_stub(self, client, db):
        """An ID not in DB returns a stub entry (not a 404)."""
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/thread_messages?id=nonexistent")
        data = resp.get_json()
        assert len(data["messages"]) == 1
        assert data["messages"][0]["id"] == "nonexistent"

    def test_ids_take_precedence_over_conv_key(self, client, db):
        """When both id= and conversationKey= are provided, id= wins."""
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get(
                "/api/thread_messages?id=email-4&conversationKey=project+alpha+update"
            )
        data = resp.get_json()
        ids = {m["id"] for m in data["messages"]}
        assert ids == {"email-4"}


# ══════════════════════════════════════════════════════════════════════════════
# SUGGESTED REPLY
# ══════════════════════════════════════════════════════════════════════════════

class TestSuggestedReply:
    def test_missing_conv_key_returns_400(self, client, db):
        with patch("routes.triage.get_db", return_value=db):
            resp = client.post("/api/suggested_reply", json={})
        assert resp.status_code == 400

    def test_unknown_conv_key_returns_404(self, client, db):
        with patch("routes.triage.get_db", return_value=db):
            resp = client.post("/api/suggested_reply",
                               json={"conversationKey": "does-not-exist"})
        assert resp.status_code == 404

    def test_returns_reply_for_known_thread(self, client, db):
        mock_result = {
            "summary": "Summary",
            "topic": "Projects",
            "action": "reply",
            "urgency": "high",
            "suggestedReply": "Sure, see you Friday!",
            "suggestedFolder": "",
        }
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.analyze_thread", return_value=mock_result), \
             patch("routes.triage.meta_get", return_value="[]"):
            resp = client.post("/api/suggested_reply",
                               json={"conversationKey": "project alpha update"})
        assert resp.status_code == 200
        assert resp.get_json()["reply"] == "Sure, see you Friday!"

    def test_caches_reply_in_threads_table(self, client, db):
        mock_result = {
            "summary": "s", "topic": "T", "action": "reply", "urgency": "low",
            "suggestedReply": "Cached reply text", "suggestedFolder": "",
        }
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.analyze_thread", return_value=mock_result), \
             patch("routes.triage.meta_get", return_value="[]"):
            client.post("/api/suggested_reply",
                        json={"conversationKey": "project alpha update"})
        row = db.execute(
            "SELECT suggested_reply FROM threads WHERE conversation_key=?",
            ("project alpha update",)
        ).fetchone()
        assert row["suggested_reply"] == "Cached reply text"

    def test_reply_with_context_passed_to_ai(self, client, db):
        mock_result = {
            "summary": "s", "topic": "T", "action": "reply", "urgency": "low",
            "suggestedReply": "Context-aware reply", "suggestedFolder": "",
        }
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.analyze_thread", return_value=mock_result) as mock_at, \
             patch("routes.triage.meta_get", return_value="[]"):
            client.post("/api/suggested_reply", json={
                "conversationKey": "project alpha update",
                "context": "Please keep it brief",
            })
        call_kwargs = mock_at.call_args
        assert call_kwargs[1].get("reply_context") == "Please keep it brief"

    def test_ai_exception_returns_500(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.analyze_thread", side_effect=Exception("AI error")), \
             patch("routes.triage.meta_get", return_value="[]"):
            resp = client.post("/api/suggested_reply",
                               json={"conversationKey": "project alpha update"})
        assert resp.status_code == 500


# ══════════════════════════════════════════════════════════════════════════════
# FORMAT MESSAGE STREAM (SSE)
# ══════════════════════════════════════════════════════════════════════════════

class TestFormatMessageStream:
    def _read_sse(self, response):
        """Collect all SSE events from a stream response into a list of dicts."""
        events = []
        for line in response.data.decode().splitlines():
            if line.startswith("data: "):
                try:
                    events.append(json.loads(line[6:]))
                except Exception:
                    pass
        return events

    def test_cache_hit_returns_done_immediately(self, client, db):
        paras = [{"text": "Hello", "intent": "FYI", "emoji": "👋", "fact_concern": None}]
        db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                   (json.dumps(paras), "email-1"))
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/format_message_stream?id=email-1")
        events = self._read_sse(resp)
        assert any(e.get("type") == "done" for e in events)
        done = next(e for e in events if e.get("type") == "done")
        assert done["paragraphs"] == paras

    def test_cache_hit_does_not_call_mcp(self, client, db):
        paras = [{"text": "Cached", "intent": "FYI", "emoji": "📄", "fact_concern": None}]
        db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                   (json.dumps(paras), "email-1"))
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct:
            client.get("/api/format_message_stream?id=email-1")
        mock_ct.assert_not_called()

    def test_empty_body_returns_no_content_paragraph(self, client, db):
        mcp_resp = make_mcp_message("email-1", body_content="")
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/format_message_stream?id=email-1")
        events = self._read_sse(resp)
        done = next((e for e in events if e.get("type") == "done"), None)
        assert done is not None
        assert any("no content" in p.get("text", "").lower() for p in done["paragraphs"])

    def test_stream_returns_body_html(self, client, db):
        html = "<html><body><p>Test email body</p></body></html>"
        mcp_resp = make_mcp_message("email-1", body_content=html)
        mock_ai = MagicMock()
        mock_stream = MagicMock()
        mock_stream.__enter__ = MagicMock(return_value=mock_stream)
        mock_stream.__exit__ = MagicMock(return_value=False)
        mock_stream.text_stream = iter([])
        mock_ai.messages.stream.return_value = mock_stream

        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp), \
             patch("routes.mail._get_ai", return_value=mock_ai), \
             patch("routes.mail.format_message_ai", return_value=[]):
            resp = client.get("/api/format_message_stream?id=email-1")
        events = self._read_sse(resp)
        done = next((e for e in events if e.get("type") == "done"), None)
        assert done is not None
        # body_html should be present in the done event
        assert "body_html" in done

    def test_cid_images_replaced_in_stream(self, client, db):
        html = '<html><body><img src="cid:abc123"></body></html>'
        mcp_resp = make_mcp_message("email-1", body_content=html)
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp), \
             patch("routes.mail.format_message_ai", return_value=[]):
            resp = client.get("/api/format_message_stream?id=email-1")
        events = self._read_sse(resp)
        done = next((e for e in events if e.get("type") == "done"), None)
        assert done is not None
        assert "cid:" not in (done.get("body_html") or "").lower()

    def test_external_images_replaced_in_stream(self, client, db):
        html = '<html><body><img src="https://tracker.example.com/t.gif"></body></html>'
        mcp_resp = make_mcp_message("email-1", body_content=html)
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp), \
             patch("routes.mail.format_message_ai", return_value=[]):
            resp = client.get("/api/format_message_stream?id=email-1")
        events = self._read_sse(resp)
        done = next((e for e in events if e.get("type") == "done"), None)
        body_html = done.get("body_html") or ""
        assert "tracker.example.com" not in body_html


# ══════════════════════════════════════════════════════════════════════════════
# TRIAGE SHEET — /api/threads and /api/updates
# ══════════════════════════════════════════════════════════════════════════════

class TestTriageSheet:
    def test_threads_returns_groups(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert "groups" in data
        assert len(data["groups"]) >= 1

    def test_threads_grouped_by_topic(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        groups = resp.get_json()["groups"]
        topics = [g["topic"] for g in groups]
        assert "Projects" in topics

    def test_threads_includes_counts(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        data = resp.get_json()
        assert "threadCount" in data
        assert "emailCount" in data
        assert data["threadCount"] >= 2

    def test_threads_include_required_fields(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        thread = resp.get_json()["groups"][0]["threads"][0]
        for field in ["conversationKey", "subject", "urgency", "action", "summary", "latestId"]:
            assert field in thread

    def test_updates_returns_only_newer_threads(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/updates?since=2026-03-17T00:00:00Z")
        data = resp.get_json()
        # All returned threads should have updatedAt > since
        for t in data["threads"]:
            assert t["updatedAt"] > "2026-03-17T00:00:00Z"

    def test_updates_returns_empty_when_no_new(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/updates?since=2099-01-01T00:00:00Z")
        assert resp.get_json()["threads"] == []

    def test_threads_sorted_by_recency(self, client, db):
        with patch("routes.triage.get_db", return_value=db), \
             patch("routes.triage.meta_get", return_value="{}"):
            resp = client.get("/api/threads")
        groups = resp.get_json()["groups"]
        all_threads = [t for g in groups for t in g["threads"]]
        dates = [t["latestReceived"] for t in all_threads if t["latestReceived"]]
        assert dates == sorted(dates, reverse=True)


# ══════════════════════════════════════════════════════════════════════════════
# MAILBOX — folders and folder content
# ══════════════════════════════════════════════════════════════════════════════

class TestMailboxFolders:
    def test_mailbox_folders_returns_list(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.meta_get", side_effect=lambda k, d="": {
                 "folders_raw": json.dumps([{"display_name": "Inbox"}, {"display_name": "Efforts"}]),
                 "efforts_subfolders": '["Alpha"]',
             }.get(k, d)):
            resp = client.get("/api/mailbox/folders")
        assert resp.status_code == 200
        assert "folders" in resp.get_json()

    def test_mailbox_folder_returns_threads(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/mailbox/folder?folder=Inbox")
        assert resp.status_code == 200
        data = resp.get_json()
        assert "threads" in data
        assert data["folder"] == "Inbox"

    def test_mailbox_folder_empty_folder_returns_empty(self, client, db):
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/mailbox/folder?folder=")
        data = resp.get_json()
        assert data["threads"] == []


# ══════════════════════════════════════════════════════════════════════════════
# PEOPLE / CONTACTS
# ══════════════════════════════════════════════════════════════════════════════

class TestPeopleEndpoints:
    def test_people_returns_list(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people")
        data = resp.get_json()
        assert "people" in data
        assert len(data["people"]) > 0

    def test_people_excludes_own_email(self, client, db):
        db.execute("INSERT OR IGNORE INTO emails (id,subject,from_name,from_address,"
                   "received_date_time,is_read,body_preview,conversation_key,raw_json,synced_at) "
                   "VALUES('e-me','S','Me','me@example.com','2026-01-01',1,'preview','c','{}','2026-01-01')")
        db.commit()
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people")
        addresses = [p["address"] for p in resp.get_json()["people"]]
        assert "me@example.com" not in addresses

    def test_people_filtered_by_query(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people?q=alice")
        people = resp.get_json()["people"]
        assert all("alice" in (p["name"] or "").lower() or "alice" in p["address"].lower()
                   for p in people)

    def test_people_sorted_alphabetically(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/people")
        names = [(p["name"] or p["address"]).lower() for p in resp.get_json()["people"]]
        assert names == sorted(names)

    def test_my_email_returns_address(self, client, db):
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/my_email")
        assert resp.get_json()["email"] == "me@example.com"

    def test_top_contacts_returns_sorted_by_frequency(self, client, db):
        db.execute("INSERT OR REPLACE INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
                   ("frequent@example.com", "Frequent Person", 100, "2026-03-01"))
        db.execute("INSERT OR REPLACE INTO contacts(email,name,frequency,last_seen) VALUES(?,?,?,?)",
                   ("rare@example.com", "Rare Person", 1, "2026-03-01"))
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts?n=2")
        contacts = resp.get_json()["contacts"]
        assert contacts[0]["frequency"] >= contacts[1]["frequency"]

    def test_top_contacts_respects_n_limit(self, client, db):
        for i in range(5):
            db.execute("INSERT OR REPLACE INTO contacts(email,name,frequency,last_seen) "
                       "VALUES(?,?,?,?)", (f"c{i}@x.com", f"Contact {i}", i+1, "2026-01-01"))
        db.commit()
        with patch("routes.mail.get_db", return_value=db):
            resp = client.get("/api/top_contacts?n=3")
        assert len(resp.get_json()["contacts"]) <= 3


# ══════════════════════════════════════════════════════════════════════════════
# IMAGE EMBEDDING (security / privacy)
# ══════════════════════════════════════════════════════════════════════════════

class TestImageEmbedding:
    def test_cid_replaced_with_blank_gif(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = '<p><img src="cid:img001@domain.com"></p>'
        result = _embed_cid_images(html)
        assert _BLANK_GIF in result
        assert "cid:" not in result.lower()

    def test_https_tracking_pixel_replaced(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = '<img src="https://open.tracker.com/t/abc.gif">'
        result = _embed_cid_images(html)
        assert _BLANK_GIF in result
        assert "tracker.com" not in result

    def test_http_image_replaced(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = '<img src="http://insecure.example.com/img.png">'
        result = _embed_cid_images(html)
        assert _BLANK_GIF in result

    def test_data_uri_preserved(self):
        from routes.mail import _embed_cid_images
        data_uri = "data:image/png;base64,abc123"
        html = f'<img src="{data_uri}">'
        result = _embed_cid_images(html)
        assert data_uri in result

    def test_plain_text_unchanged(self):
        from routes.mail import _embed_cid_images
        html = "<p>No images here.</p>"
        assert _embed_cid_images(html) == html

    def test_multiple_images_all_replaced(self):
        from routes.mail import _embed_cid_images, _BLANK_GIF
        html = ('<img src="cid:a">'
                '<img src="https://x.com/b.png">'
                '<img src="http://y.com/c.png">')
        result = _embed_cid_images(html)
        assert result.count(_BLANK_GIF) == 3
        assert "cid:" not in result.lower()
        assert "x.com" not in result
        assert "y.com" not in result
