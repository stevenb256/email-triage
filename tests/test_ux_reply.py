"""
test_ux_reply.py — End-to-end backend tests for the reply UX flow.

Tests the full reply lifecycle: fetching recipients, composing a reply,
reply-all logic, and edge cases (empty body, missing latestId).
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
# POST /api/message_recipients — recipient parsing
# ---------------------------------------------------------------------------

class TestMessageRecipientsParsing:
    def test_returns_correct_to_and_cc(self, client, db):
        """Returns correct to and cc arrays when MCP returns full message data."""
        mcp_resp = make_mcp_message(
            "email-3",
            to_recipients=[
                {"name": "Me", "address": "me@example.com"},
                {"name": "Charlie", "address": "charlie@example.com"},
            ],
            cc_recipients=[
                {"name": "Dan", "address": "dan@example.com"},
            ],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-3")

        data = resp.get_json()
        assert resp.status_code == 200
        to_addrs = {r["address"] for r in data["to"]}
        cc_addrs = {r["address"] for r in data["cc"]}
        assert "me@example.com" in to_addrs
        assert "charlie@example.com" in to_addrs
        assert "dan@example.com" in cc_addrs

    def test_handles_messages_wrapper(self, client, db):
        """Handles {'messages': [msg]} wrapper from MCP."""
        inner = make_mcp_message(
            "email-3",
            to_recipients=[{"name": "Alice", "address": "alice@example.com"}],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"messages": [inner]}):
            resp = client.get("/api/message_recipients?id=email-3")
        data = resp.get_json()
        assert len(data["to"]) == 1
        assert data["to"][0]["address"] == "alice@example.com"

    def test_emailaddress_nested_format(self, client, db):
        """Parses the nested emailAddress dict format from Microsoft Graph API."""
        mcp_resp = make_mcp_message(
            "email-3",
            to_recipients=[
                {
                    "emailAddress": {"name": "Alice", "address": "alice@example.com"}
                }
            ],
        )
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value=mcp_resp):
            resp = client.get("/api/message_recipients?id=email-3")
        data = resp.get_json()
        assert len(data["to"]) == 1
        assert data["to"][0]["address"] == "alice@example.com"


# ---------------------------------------------------------------------------
# POST /api/reply — sending reply to correct recipients
# ---------------------------------------------------------------------------

class TestReplyToCorrectRecipients:
    def test_reply_sends_to_specified_recipients(self, client, db):
        """reply endpoint calls MCP draft with the provided to/cc lists."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = [
                make_mcp_draft_response("draft-xyz"),
                {"ok": True},
            ]
            resp = client.post(
                "/api/reply/email-3",
                json={
                    "body": "Sounds great!",
                    "conversationKey": "project alpha update",
                    "to": ["alice@example.com", "bob@example.com"],
                    "cc": ["carol@example.com"],
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        assert data["ok"] is True

        # Verify draft call args
        draft_call = mock_ct.call_args_list[0]
        args = draft_call[0][1]  # second positional arg = kwargs dict to call_tool
        assert "outlook_mail_draft_message" == mock_ct.call_args_list[0][0][0]
        assert args["bodyText"] == "Sounds great!"
        assert args["to"] == ["alice@example.com", "bob@example.com"]
        assert args["cc"] == ["carol@example.com"]

    def test_reply_removes_thread_from_db(self, client, db):
        """Thread is removed from local DB after successful reply."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread") as mock_rm:
            mock_ct.side_effect = [
                make_mcp_draft_response("d1"),
                {"ok": True},
            ]
            client.post(
                "/api/reply/email-3",
                json={
                    "body": "Reply text",
                    "conversationKey": "project alpha update",
                },
            )
        mock_rm.assert_called_once_with("project alpha update")

    def test_reply_without_conv_key_does_not_call_remove(self, client, db):
        """If no conversationKey, remove_thread is not called."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread") as mock_rm:
            mock_ct.side_effect = [
                make_mcp_draft_response("d1"),
                {"ok": True},
            ]
            client.post(
                "/api/reply/email-3",
                json={"body": "Hello"},
            )
        mock_rm.assert_not_called()


# ---------------------------------------------------------------------------
# Reply-all: should exclude current user's own email
# ---------------------------------------------------------------------------

class TestReplyAllExcludesMyEmail:
    def test_reply_all_excludes_own_email_in_to(self, client, db):
        """
        When composing reply-all, the caller (frontend) is responsible for
        filtering out my_email from the to list. We verify the route passes
        through the to/cc lists exactly as provided.
        """
        # Simulate frontend already having filtered out me@example.com
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = [
                make_mcp_draft_response("d1"),
                {"ok": True},
            ]
            resp = client.post(
                "/api/reply/email-3",
                json={
                    "body": "My reply",
                    "conversationKey": "project alpha update",
                    "to": ["alice@example.com"],  # me@example.com excluded by frontend
                    "cc": [],
                },
            )
        data = resp.get_json()
        assert resp.status_code == 200
        # Verify me@example.com is NOT in the draft to list
        draft_args = mock_ct.call_args_list[0][0][1]
        assert "me@example.com" not in draft_args.get("to", [])

    def test_get_my_email_endpoint(self, client, db):
        """Frontend can query /api/my_email to know which addresses to exclude."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.get_my_email", return_value="me@example.com"):
            resp = client.get("/api/my_email")
        data = resp.get_json()
        assert data["email"] == "me@example.com"


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestReplyEdgeCases:
    def test_reply_with_empty_body_drafts_anyway(self, client, db):
        """
        Route does not validate body content — an empty body is passed to MCP.
        The MCP call should still be attempted (empty body is valid in Outlook).
        """
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool") as mock_ct, \
             patch("routes.mail.remove_thread"):
            mock_ct.side_effect = [
                make_mcp_draft_response("d1"),
                {"ok": True},
            ]
            resp = client.post(
                "/api/reply/email-3",
                json={
                    "body": "",
                    "conversationKey": "project alpha update",
                },
            )
        # With empty body, MCP still called (draft endpoint accepts empty body)
        assert resp.status_code == 200
        draft_args = mock_ct.call_args_list[0][0][1]
        assert draft_args["bodyText"] == ""

    def test_reply_with_missing_latest_id_in_mcp_returns_error(self, client, db):
        """Reply to a non-existent message ID: MCP raises an exception → 500."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=Exception("Message not found")):
            resp = client.post(
                "/api/reply/nonexistent-id",
                json={
                    "body": "Hello",
                    "conversationKey": "project alpha update",
                },
            )
        assert resp.status_code == 500
        data = resp.get_json()
        assert "error" in data

    def test_reply_draft_without_draft_id_returns_500(self, client, db):
        """When MCP returns no draft_id, reply endpoint returns 500 with error."""
        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", return_value={"status": "ok"}):
            resp = client.post(
                "/api/reply/email-3",
                json={"body": "Hello", "conversationKey": "project alpha update"},
            )
        assert resp.status_code == 500
        data = resp.get_json()
        assert "error" in data

    def test_reply_mcp_send_failure_returns_500(self, client, db):
        """When the send MCP call fails after drafting, returns 500."""
        def _mock_ct(tool_name, args):
            if tool_name == "outlook_mail_draft_message":
                return make_mcp_draft_response("d1")
            raise Exception("Send failed")

        with patch("routes.mail.get_db", return_value=db), \
             patch("routes.mail.call_tool", side_effect=_mock_ct):
            resp = client.post(
                "/api/reply/email-3",
                json={"body": "Hello", "conversationKey": "project alpha update"},
            )
        assert resp.status_code == 500
