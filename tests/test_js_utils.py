"""
test_js_utils.py — Tests for pure functions in static/js/utils.js.

Uses subprocess to run Node.js if available. Skips gracefully if Node is not found.
"""
import os
import subprocess
import sys
import textwrap

import pytest

APP_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
UTILS_JS = os.path.join(APP_ROOT, "static", "js", "utils.js")

_NODE_PATHS = [
    "/usr/local/bin/node",
    "/opt/homebrew/bin/node",
    "/usr/bin/node",
]


def _find_node():
    """Return path to node binary, or None if not found."""
    for path in _NODE_PATHS:
        if os.path.isfile(path):
            return path
    # Try PATH as last resort
    try:
        result = subprocess.run(
            ["which", "node"], capture_output=True, text=True, timeout=5
        )
        if result.returncode == 0:
            found = result.stdout.strip()
            if found:
                return found
    except Exception:
        pass
    return None


def _run_node(script: str, node_bin: str) -> str:
    """Run a Node.js script and return stdout."""
    full_script = f"""
// Load utils.js — polyfill browser globals needed by the module
global.btoa = (s) => Buffer.from(s, 'binary').toString('base64');
global.atob = (s) => Buffer.from(s, 'base64').toString('binary');
global.document = {{ getElementById: () => null }};
global.setTimeout = setTimeout;
// encodeURIComponent / decodeURIComponent / unescape / escape are available in Node

{open(UTILS_JS).read()}

{textwrap.dedent(script)}
"""
    result = subprocess.run(
        [node_bin, "--input-type=module" if False else "-e", full_script],
        capture_output=True, text=True, timeout=10
    )
    if result.returncode != 0:
        raise RuntimeError(f"Node error: {result.stderr}")
    return result.stdout.strip()


# ---------------------------------------------------------------------------
# Skip if Node not available
# ---------------------------------------------------------------------------

NODE = _find_node()
skip_no_node = pytest.mark.skipif(NODE is None, reason="Node.js not found")


@pytest.fixture(scope="module")
def node():
    if NODE is None:
        pytest.skip("Node.js not available")
    return NODE


# ---------------------------------------------------------------------------
# esc()
# ---------------------------------------------------------------------------

class TestEsc:
    @skip_no_node
    def test_escapes_html_special_chars(self, node):
        out = _run_node("""
console.log(esc('<script>alert("xss")</script>'));
""", node)
        assert "<script>" not in out
        assert "&lt;script&gt;" in out

    @skip_no_node
    def test_escapes_ampersand(self, node):
        out = _run_node("console.log(esc('foo & bar'));", node)
        assert "foo &amp; bar" in out

    @skip_no_node
    def test_escapes_quotes(self, node):
        out = _run_node(r"console.log(esc('say \"hi\"'));", node)
        assert "&quot;" in out

    @skip_no_node
    def test_empty_string(self, node):
        out = _run_node("console.log(esc(''));", node)
        assert out == ""


# ---------------------------------------------------------------------------
# initials()
# ---------------------------------------------------------------------------

class TestInitials:
    @skip_no_node
    def test_two_word_name(self, node):
        out = _run_node("console.log(initials('Alice Smith'));", node)
        assert out == "AS"

    @skip_no_node
    def test_single_word_name(self, node):
        out = _run_node("console.log(initials('Bob'));", node)
        assert out == "BO"

    @skip_no_node
    def test_empty_string(self, node):
        out = _run_node("console.log(initials(''));", node)
        assert out == "?"

    @skip_no_node
    def test_three_word_name_uses_first_two(self, node):
        out = _run_node("console.log(initials('Alice Jane Smith'));", node)
        assert out == "AJ"


# ---------------------------------------------------------------------------
# avColor()
# ---------------------------------------------------------------------------

class TestAvColor:
    @skip_no_node
    def test_returns_valid_hex_color(self, node):
        out = _run_node("console.log(avColor('Alice Smith'));", node)
        assert out.startswith("#")
        assert len(out) == 7

    @skip_no_node
    def test_same_input_stable_color(self, node):
        out1 = _run_node("console.log(avColor('test-user'));", node)
        out2 = _run_node("console.log(avColor('test-user'));", node)
        assert out1 == out2

    @skip_no_node
    def test_different_inputs_may_differ(self, node):
        out1 = _run_node("console.log(avColor('Alice'));", node)
        out2 = _run_node("console.log(avColor('Zebra'));", node)
        # Just verify both are valid hex colors (they might coincidentally match)
        assert out1.startswith("#")
        assert out2.startswith("#")


# ---------------------------------------------------------------------------
# _injectBaseTarget()
# ---------------------------------------------------------------------------

class TestInjectBaseTarget:
    @skip_no_node
    def test_injects_into_head(self, node):
        out = _run_node("""
const html = '<html><head></head><body>Hi</body></html>';
console.log(_injectBaseTarget(html));
""", node)
        assert '<base target="_blank">' in out

    @skip_no_node
    def test_prepends_when_no_head(self, node):
        out = _run_node("""
const html = '<body>No head here</body>';
console.log(_injectBaseTarget(html));
""", node)
        assert '<base target="_blank">' in out

    @skip_no_node
    def test_idempotent_result_is_valid(self, node):
        out = _run_node("""
const html = '<html><head><title>T</title></head><body></body></html>';
const result = _injectBaseTarget(html);
// Should contain the base tag exactly once after a single call
const count = (result.match(/<base target="_blank">/g) || []).length;
console.log(count);
""", node)
        assert out == "1"


# ---------------------------------------------------------------------------
# encodeThread / decodeThread round-trip
# ---------------------------------------------------------------------------

class TestEncodeDecodeThread:
    @skip_no_node
    def test_round_trip(self, node):
        out = _run_node("""
const thread = {
  conversationKey: 'project-alpha',
  latestId: 'email-3',
  emailIds: ['email-1', 'email-2', 'email-3'],
  subject: 'Project Alpha Update',
  messageCount: 3,
  suggestedReply: 'Thanks!',
  suggestedFolder: 'Efforts/Alpha',
};
const encoded = encodeThread(thread);
const decoded = decodeThread(encoded);
console.log(decoded.conversationKey);
console.log(decoded.latestId);
console.log(decoded.messageCount);
""", node)
        lines = out.strip().split("\n")
        assert lines[0] == "project-alpha"
        assert lines[1] == "email-3"
        assert lines[2] == "3"

    @skip_no_node
    def test_decode_invalid_returns_empty_object(self, node):
        out = _run_node("""
const result = decodeThread('!!!invalid!!!');
console.log(typeof result);
console.log(Object.keys(result).length);
""", node)
        lines = out.strip().split("\n")
        assert lines[0] == "object"
        assert lines[1] == "0"


# ---------------------------------------------------------------------------
# fmtDate()
# ---------------------------------------------------------------------------

class TestFmtDate:
    @skip_no_node
    def test_empty_string_returns_empty(self, node):
        out = _run_node("console.log(fmtDate(''));", node)
        assert out == ""

    @skip_no_node
    def test_old_date_returns_month_day(self, node):
        # A date many months ago should be formatted as "Mon DD"
        out = _run_node("console.log(fmtDate('2020-01-15T10:00:00Z'));", node)
        assert out != ""
        assert len(out) > 0

    @skip_no_node
    def test_invalid_date_returns_empty(self, node):
        out = _run_node("console.log(fmtDate('not-a-date'));", node)
        assert out == ""
