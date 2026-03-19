"""
config.py — Constants for Outlook Express email triage app.
"""
import os

MCP_COMMAND     = "/opt/homebrew/bin/McpOutlookLocal"
DB_PATH         = os.path.join(os.path.dirname(__file__), "email_triage.db")
PORT            = 5002
SYNC_INTERVAL   = 300   # 5 minutes
INBOX_FETCH     = 100
FOLDER_FETCH    = 100   # messages per non-inbox folder per sync cycle
ANALYSIS_MODEL  = "claude-haiku-4-5-20251001"
REPLY_MODEL     = "claude-sonnet-4-6"

SKIP_SYNC_FOLDERS = {
    "Drafts", "Outbox", "Junk Email",
    "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled", "Inbox", "",
}

_env_path = os.path.join(os.path.dirname(__file__), ".env")
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
