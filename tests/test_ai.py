"""
test_ai.py — Tests for ai.py AI functions.

All Anthropic API calls are mocked. No real network calls are made.
"""
import json
import sys
import os
from unittest.mock import MagicMock, patch

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if APP_ROOT not in sys.path:
    sys.path.insert(0, APP_ROOT)


def _make_ai_response(text: str):
    """Build a mock Anthropic response with the given text."""
    mock_resp = MagicMock()
    mock_content = MagicMock()
    mock_content.text = text
    mock_resp.content = [mock_content]
    return mock_resp


# ---------------------------------------------------------------------------
# _normalize_topic
# ---------------------------------------------------------------------------

class TestNormalizeTopic:
    def test_trims_whitespace(self):
        from ai import _normalize_topic
        assert _normalize_topic("  Hello World  ") == "Hello World"

    def test_caps_at_50_chars(self):
        from ai import _normalize_topic
        long = "A" * 60
        result = _normalize_topic(long)
        assert len(result) == 50

    def test_collapses_internal_spaces(self):
        from ai import _normalize_topic
        result = _normalize_topic("Hello   World")
        assert result == "Hello World"

    def test_defaults_to_general_for_empty(self):
        from ai import _normalize_topic
        assert _normalize_topic("") == "General"
        assert _normalize_topic("   ") == "General"
        assert _normalize_topic(None) == "General"

    def test_does_not_alter_valid_topic(self):
        from ai import _normalize_topic
        assert _normalize_topic("Knowledge Anchors") == "Knowledge Anchors"

    def test_exact_50_chars_preserved(self):
        from ai import _normalize_topic
        t = "A" * 50
        assert _normalize_topic(t) == t


# ---------------------------------------------------------------------------
# _clean
# ---------------------------------------------------------------------------

class TestClean:
    def test_strips_control_chars(self):
        from ai import _clean
        raw = "Hello\x00\x01\x07World"
        result = _clean(raw)
        assert "\x00" not in result
        assert "Hello" in result
        assert "World" in result

    def test_respects_length_limit(self):
        from ai import _clean
        result = _clean("Hello World", n=5)
        assert result == "Hello"

    def test_no_limit_returns_full(self):
        from ai import _clean
        s = "Hello World"
        assert _clean(s) == s

    def test_handles_none(self):
        from ai import _clean
        assert _clean(None) == ""

    def test_preserves_tabs_and_newlines(self):
        from ai import _clean
        # \t and \n are NOT in the control char range being stripped
        result = _clean("Hello\tWorld\nFoo")
        assert "\t" in result
        assert "\n" in result


# ---------------------------------------------------------------------------
# _get_full_body
# ---------------------------------------------------------------------------

class TestGetFullBody:
    def test_formatted_body_takes_priority(self):
        """formatted_body JSON paragraphs are used first when present and substantial."""
        from ai import _get_full_body
        paras = [{"text": "A" * 200, "intent": "FYI"}]
        email = {
            "id": "msg-1",
            "formatted_body": json.dumps(paras),
            "body_html": "<html><body>" + "B" * 500 + "</body></html>",
            "body_preview": "C" * 400,
        }
        result = _get_full_body(email)
        assert "A" * 100 in result

    def test_body_html_used_when_no_formatted_body(self):
        """body_html is used when formatted_body is absent."""
        from ai import _get_full_body
        long_html = "<html><body>" + "<p>Hello world content here.</p>" * 20 + "</body></html>"
        email = {
            "id": "msg-1",
            "body_html": long_html,
            "body_preview": "short",
        }
        result = _get_full_body(email)
        assert "Hello world content here" in result

    def test_body_preview_fallback_when_substantial(self):
        """body_preview > 300 chars is used as fallback (skips MCP fetch)."""
        from ai import _get_full_body
        long_preview = "X" * 400
        email = {
            "id": "msg-1",
            "body_preview": long_preview,
        }
        # Should NOT call _call_tool since preview is substantial
        with patch("ai._call_tool") as mock_ct:
            result = _get_full_body(email)
            mock_ct.assert_not_called()
        assert long_preview[:100] in result

    def test_mcp_fetch_when_no_other_body(self):
        """Falls back to MCP outlook_mail_get_message when no body available."""
        from ai import _get_full_body
        email = {"id": "msg-xyz", "body_preview": ""}
        mcp_body = "<p>" + "MCP content. " * 20 + "</p>"
        with patch("ai._call_tool") as mock_ct, \
             patch("db.get_db") as mock_db:
            mock_ct.return_value = {"body_content": mcp_body}
            mock_db.return_value = MagicMock()
            result = _get_full_body(email)
            mock_ct.assert_called_once_with("outlook_mail_get_message", {"message_id": "msg-xyz"})
        assert "MCP content" in result

    def test_short_body_preview_falls_through_to_mcp(self):
        """Short body_preview (<= 300 chars) triggers MCP fetch."""
        from ai import _get_full_body
        email = {"id": "msg-xyz", "body_preview": "Short preview."}
        with patch("ai._call_tool") as mock_ct, \
             patch("db.get_db") as mock_db:
            mock_ct.return_value = {}
            mock_db.return_value = MagicMock()
            # Falls through to body_preview fallback
            result = _get_full_body(email)
        # Should return the body_preview as last resort
        assert result != ""


# ---------------------------------------------------------------------------
# analyze_thread
# ---------------------------------------------------------------------------

class TestAnalyzeThread:
    def _sample_emails(self):
        return [
            {
                "id": "e1",
                "subject": "Project Update",
                "from_name": "Alice",
                "from_address": "alice@example.com",
                "received_date_time": "2026-03-15T10:00:00Z",
                "body_preview": "Here is the project update for Q2.",
            }
        ]

    def test_parses_valid_json_response(self):
        from ai import analyze_thread
        response_json = json.dumps({
            "summary": "Facts||BREAK||None||BREAK||Reply by EOD",
            "topic": "Project Alpha",
            "action": "reply",
            "urgency": "high",
            "suggestedReply": "Thanks for the update!",
            "suggestedFolder": "",
        })
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(response_json)

            result = analyze_thread(self._sample_emails(), [], [])

        assert result["action"] == "reply"
        assert result["urgency"] == "high"
        assert result["topic"] == "Project Alpha"
        assert result["suggestedReply"] == "Thanks for the update!"

    def test_handles_json_with_literal_newlines(self):
        """Handles malformed JSON with literal newlines inside string values."""
        from ai import analyze_thread
        # Simulate Claude emitting literal \n inside a JSON string value
        bad_json = '{"summary": "line1\nline2", "topic": "Test", "action": "read", "urgency": "low", "suggestedReply": "", "suggestedFolder": ""}'
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(bad_json)
            result = analyze_thread(self._sample_emails(), [], [])
        # Should not raise, and action should be parsed
        assert result["action"] in ("read", "reply", "delete", "file", "done")

    def test_regex_fallback_for_unparseable_json(self):
        """Uses regex field extraction when JSON is completely malformed."""
        from ai import analyze_thread
        totally_broken = '{"summary": "broken json, "action": "file", "urgency": "low"'
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(totally_broken)
            result = analyze_thread(self._sample_emails(), [], [])
        # Should return a dict with at least action and urgency
        assert "action" in result
        assert "urgency" in result

    def test_sanitizes_placeholder_suggested_folder(self):
        """Rejects placeholder-like suggestedFolder values."""
        from ai import analyze_thread
        for bad_folder in ["Select Folder", "None", "n/a", "TBD", "unknown", "folder", ""]:
            response_json = json.dumps({
                "summary": "...",
                "topic": "Test",
                "action": "file",
                "urgency": "low",
                "suggestedReply": "",
                "suggestedFolder": bad_folder,
            })
            with patch("ai._get_ai") as mock_get_ai, \
                 patch("ai._call_tool", return_value={}):
                mock_client = MagicMock()
                mock_get_ai.return_value = mock_client
                mock_client.messages.create.return_value = _make_ai_response(response_json)
                result = analyze_thread(self._sample_emails(), [], [])
            assert result["suggestedFolder"] == "", f"Expected '' for bad_folder={bad_folder!r}, got {result['suggestedFolder']!r}"

    def test_returns_error_dict_on_exception(self):
        """Returns safe default dict when AI call fails."""
        from ai import analyze_thread
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.side_effect = Exception("API error")
            result = analyze_thread(self._sample_emails(), [], [])
        assert "action" in result
        assert result["action"] == "read"

    def test_normalizes_topic(self):
        """Topic is normalized via _normalize_topic after parsing."""
        from ai import analyze_thread
        response_json = json.dumps({
            "summary": "...",
            "topic": "  knowledge anchors  ",
            "action": "read",
            "urgency": "low",
            "suggestedReply": "",
            "suggestedFolder": "",
        })
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(response_json)
            result = analyze_thread(self._sample_emails(), [], [])
        assert result["topic"] == "knowledge anchors"  # normalized (trimmed)

    def test_valid_folder_not_sanitized(self):
        """A real folder name like 'Efforts/Alpha' is kept as-is."""
        from ai import analyze_thread
        response_json = json.dumps({
            "summary": "...",
            "topic": "Alpha",
            "action": "file",
            "urgency": "low",
            "suggestedReply": "",
            "suggestedFolder": "Efforts/Alpha",
        })
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(response_json)
            result = analyze_thread(self._sample_emails(), [], [])
        assert result["suggestedFolder"] == "Efforts/Alpha"


# ---------------------------------------------------------------------------
# format_message_ai
# ---------------------------------------------------------------------------

class TestFormatMessageAi:
    def test_returns_paragraph_list(self):
        from ai import format_message_ai
        para_json = json.dumps({
            "paragraphs": [
                {"text": "Hello world", "intent": "FYI", "emoji": "📄", "fact_concern": None}
            ]
        })
        with patch("ai._get_ai") as mock_get_ai:
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response(para_json)
            msg = {"id": "m1", "body": "Hello world", "from_name": "Alice",
                   "received_date_time": "2026-03-15T10:00:00Z"}
            result = format_message_ai(msg)
        assert isinstance(result, list)
        assert len(result) == 1
        assert result[0]["text"] == "Hello world"

    def test_handles_empty_body(self):
        from ai import format_message_ai
        msg = {"id": "m1", "body": "", "body_preview": "", "from_name": "Alice",
               "received_date_time": "2026-03-15T10:00:00Z"}
        result = format_message_ai(msg)
        assert isinstance(result, list)
        assert len(result) == 1
        assert result[0]["text"] == "(no content)"

    def test_falls_back_on_ai_error(self):
        from ai import format_message_ai
        with patch("ai._get_ai") as mock_get_ai:
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.side_effect = Exception("fail")
            msg = {"id": "m1", "body": "Para1\n\nPara2", "from_name": "Alice",
                   "received_date_time": "2026-03-15T10:00:00Z"}
            result = format_message_ai(msg)
        assert isinstance(result, list)
        assert len(result) >= 1


# ---------------------------------------------------------------------------
# summarize_message_ai
# ---------------------------------------------------------------------------

class TestSummarizeMessageAi:
    def test_returns_non_empty_for_valid_message(self):
        from ai import summarize_message_ai
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response("Project Alpha is on track for Q2")
            msg = {
                "id": "m1",
                "body_preview": "X" * 400,  # substantial preview
                "from_name": "Alice",
                "from_address": "alice@example.com",
            }
            result = summarize_message_ai(msg)
        assert isinstance(result, str)
        assert len(result) > 0

    def test_returns_empty_for_empty_body(self):
        from ai import summarize_message_ai
        with patch("ai._call_tool", return_value={}):
            msg = {
                "id": "m1",
                "body_preview": "",
                "from_name": "Alice",
            }
            result = summarize_message_ai(msg)
        assert result == ""

    def test_strips_trailing_period(self):
        from ai import summarize_message_ai
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response("Summary sentence.")
            msg = {"id": "m1", "body_preview": "X" * 400, "from_name": "Alice"}
            result = summarize_message_ai(msg)
        assert not result.endswith(".")

    def test_returns_empty_on_exception(self):
        from ai import summarize_message_ai
        with patch("ai._get_ai") as mock_get_ai, \
             patch("ai._call_tool", return_value={}):
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.side_effect = Exception("network error")
            msg = {"id": "m1", "body_preview": "X" * 400, "from_name": "Alice"}
            result = summarize_message_ai(msg)
        assert result == ""


# ---------------------------------------------------------------------------
# generate_reply_ai
# ---------------------------------------------------------------------------

class TestGenerateReplyAi:
    def test_returns_string_reply(self):
        from ai import generate_reply_ai
        with patch("ai._get_ai") as mock_get_ai:
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response("Here is my reply text.")
            result = generate_reply_ai(
                subject="Test Subject",
                msgs_text="From: Alice | 2026-03-15\nHello there.",
                user_prompt="Thank Alice for the update",
            )
        assert isinstance(result, str)
        assert len(result) > 0

    def test_passes_subject_and_context(self):
        from ai import generate_reply_ai
        with patch("ai._get_ai") as mock_get_ai:
            mock_client = MagicMock()
            mock_get_ai.return_value = mock_client
            mock_client.messages.create.return_value = _make_ai_response("Reply body")
            generate_reply_ai("My Subject", "Thread context", "User intent")
            call_kwargs = mock_client.messages.create.call_args
            prompt_content = call_kwargs[1]["messages"][0]["content"]
            assert "My Subject" in prompt_content
            assert "User intent" in prompt_content
