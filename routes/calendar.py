"""
routes/calendar.py — Calendar API routes for Clanker.
"""
import json
import re
from datetime import datetime, timedelta

from flask import Blueprint, jsonify, request

from db import get_db, meta_get
from ai import summarize_thread_ai

bp = Blueprint("calendar", __name__)


@bp.route("/api/calendar")
def api_calendar():
    db = get_db()
    start_str = request.args.get("start")
    end_str = request.args.get("end")
    if not start_str:
        start_str = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%S")
    if not end_str:
        end_str = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%S")
    rows = db.execute(
        "SELECT id,subject,start_time,end_time,location,attendees FROM calendar_events "
        "WHERE start_time >= ? AND start_time <= ? ORDER BY start_time ASC",
        (start_str, end_str)
    ).fetchall()
    events = [{"id": r[0], "subject": r[1], "start_time": r[2], "end_time": r[3],
               "location": r[4], "attendees": json.loads(r[5] or "[]")} for r in rows]
    return jsonify({"events": events})


@bp.route("/api/meeting_prep", methods=["POST"])
def api_meeting_prep():
    data = request.json or {}
    subject    = data.get("subject", "(No title)")
    attendees  = data.get("attendees", [])
    start_time = data.get("start_time", "")
    end_time   = data.get("end_time", "")
    location   = data.get("location", "")

    names = ", ".join(
        a.get("name") or a.get("email", "") for a in (attendees or [])[:12]
    ) or "not listed"

    try:
        st = datetime.fromisoformat(start_time) if start_time else None
        time_str = st.strftime("%A %b %d at %-I:%M %p") if st else ""
    except Exception:
        time_str = start_time

    try:
        result = summarize_thread_ai(subject, names, time_str, location)
        return jsonify(result)
    except Exception as e:
        print(f"meeting_prep error: {e}")
    return jsonify({"ok": False, "headsup": "", "topics": []})
