#!/usr/bin/env python3
"""
app.py — Thin entry point for Clanker email triage app.
Imports modules, registers blueprints, starts sync thread, serves index.html.
"""
import threading
import webbrowser

from flask import Flask, render_template

from config import PORT
from db import init_db
from mcp_client import call_tool, _session_ready
from sync import _sync_loop
from db import meta_set
import json

from routes.triage import bp as triage_bp
from routes.mail import bp as mail_bp
from routes.calendar import bp as calendar_bp

app = Flask(__name__)

# Register blueprints (all serve under /api/ via route decorators in each bp)
app.register_blueprint(triage_bp)
app.register_blueprint(mail_bp)
app.register_blueprint(calendar_bp)


@app.route("/")
def index():
    return render_template("index.html")


def _folder_lists(folders_raw: list):
    skip = {"Drafts", "Sent Items", "Outbox", "Deleted Items", "Junk Email", ""}
    names = [
        f.get("display_name") or f.get("displayName", "")
        for f in folders_raw
        if (f.get("display_name") or f.get("displayName", "")) not in skip
    ]
    efforts = [n for n in names if n.startswith("Efforts")]
    other   = [n for n in names if not n.startswith("Efforts")]
    return efforts, other


if __name__ == "__main__":
    from config import ANTHROPIC_API_KEY
    if not ANTHROPIC_API_KEY:
        print("\n⚠  ANTHROPIC_API_KEY not set in .env\n")

    init_db()
    print(f"\n📧 Email Triage starting at http://localhost:{PORT}")

    _session_ready.wait(timeout=20)
    print("   Outlook MCP connected.")

    # Seed folder list immediately
    try:
        fdrs = call_tool("outlook_mail_list_folders", {})
        raw = fdrs.get("folders", fdrs.get("value", [])) if isinstance(fdrs, dict) else []
        meta_set("folders_raw", json.dumps(raw))
        efforts, other = _folder_lists(raw)
        print(f"   Folders loaded: {len(efforts)} Efforts, {len(other)} other")
    except Exception as e:
        print(f"   Warning: could not load folders: {e}")

    # Start background sync loop
    threading.Thread(target=_sync_loop, daemon=True).start()
    print("   Background sync started (every 5 min)\n")

    threading.Timer(2.0, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    app.run(port=PORT, debug=False, use_reloader=False)
