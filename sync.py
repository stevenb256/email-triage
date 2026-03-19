"""
sync.py — Background sync thread for Outlook Express email triage app.
"""
import json
import re
import threading
import time
from datetime import datetime, timezone, timedelta

from config import INBOX_FETCH, FOLDER_FETCH, SYNC_INTERVAL, SKIP_SYNC_FOLDERS
from db import get_db, meta_get, meta_set
from mcp_client import call_tool
from ai import analyze_thread, format_message_ai, _normalize_topic

# ─── Sync status ───────────────────────────────────────────────────────────────

_sync_status = {
    "running": False,
    "lastSync": None,
    "lastError": None,
    "emailsAdded": 0,
    "threadsUpdated": 0,
    "phase": "",
    "progress": "",
    "done": 0,
    "total": 0,
}
_sync_lock = threading.Lock()


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _norm_subject(subj: str) -> str:
    s = re.sub(r'^(RE|FW|FWD|AW|R|RES|SV)[\s:]+', '', str(subj or ''), flags=re.IGNORECASE)
    return s.strip().lower() or "no-subject"


def _clean(s, n=None) -> str:
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s or ''))
    return s[:n] if n else s


def _utcnow() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


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


# ─── Refresh helpers ───────────────────────────────────────────────────────────

def _refresh_folders() -> tuple:
    try:
        fdrs = call_tool("outlook_mail_list_folders", {"top": 50})
        top_level = fdrs.get("folders", fdrs.get("value", [])) if isinstance(fdrs, dict) else []

        efforts_id = None
        for f in top_level:
            name = f.get("display_name") or f.get("displayName", "")
            if name.lower() == "efforts":
                efforts_id = f.get("id") or f.get("folderId")
                break

        efforts_subfolders = []
        if efforts_id:
            sub = call_tool("outlook_mail_list_folders", {"parent_folder_id": efforts_id, "top": 50})
            sub_list = sub.get("folders", sub.get("value", [])) if isinstance(sub, dict) else []
            efforts_subfolders = [
                f.get("display_name") or f.get("displayName", "")
                for f in sub_list
                if (f.get("display_name") or f.get("displayName", ""))
            ]

        skip = {"Drafts", "Sent Items", "Outbox", "Deleted Items", "Junk Email",
                "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled", "Inbox", ""}
        other = [
            f.get("display_name") or f.get("displayName", "")
            for f in top_level
            if (f.get("display_name") or f.get("displayName", "")) not in skip
            and (f.get("display_name") or f.get("displayName", "")).lower() != "efforts"
        ]

        meta_set("efforts_subfolders", json.dumps(efforts_subfolders))
        meta_set("other_folders", json.dumps(other))
        meta_set("folders_raw", json.dumps(top_level))
        print(f"  Folders: {len(efforts_subfolders)} Efforts subfolders, {len(other)} other")
        return efforts_subfolders, other
    except Exception as e:
        print(f"Warning: could not refresh folders: {e}")
        return json.loads(meta_get("efforts_subfolders", "[]")), json.loads(meta_get("other_folders", "[]"))


def _refresh_calendar():
    """Fetch upcoming calendar events and store them in the local DB. Stores a meta key 'next_meeting' with JSON for the soonest future event."""
    try:
        import time as _time
        tz_name = _time.tzname[0] if _time.tzname else "UTC"
        # Try to get a proper IANA timezone name
        try:
            import subprocess
            result = subprocess.run(["readlink", "/etc/localtime"], capture_output=True, text=True)
            if result.returncode == 0 and "zoneinfo/" in result.stdout:
                tz_name = result.stdout.strip().split("zoneinfo/")[-1]
        except Exception:
            pass

        now = datetime.now(timezone.utc)
        # Format without Z suffix as required by the tool
        start = (now - timedelta(days=7)).strftime("%Y-%m-%dT%H:%M:%S")
        end = (now + timedelta(days=60)).strftime("%Y-%m-%dT%H:%M:%S")
        resp = None
        try:
            print(f"  Calling outlook_calendar_list_events: start={start}, end={end}, tz={tz_name}")
            resp = call_tool("outlook_calendar_list_events", {
                "start_datetime": start,
                "end_datetime": end,
                "time_zone": tz_name,
            })
            print(f"  Calendar resp type={type(resp).__name__}: {str(resp)[:300]}")
        except Exception as ex:
            print(f"  calendar tool failed: {ex}")
            resp = None
        events = []
        if isinstance(resp, dict):
            if "events" in resp:
                events = resp.get("events")
            elif "value" in resp and isinstance(resp.get("value"), list):
                events = resp.get("value")
            else:
                # maybe the resp itself is a single event or list
                if isinstance(resp.get("items"), list):
                    events = resp.get("items")
        elif isinstance(resp, list):
            events = resp

        db = get_db()
        now_iso = _utcnow()
        def _ev_get(ev, *keys):
            """Case-insensitive multi-key lookup for event dicts."""
            for k in keys:
                v = ev.get(k) or ev.get(k.lower()) or ev.get(k[0].upper() + k[1:]) or ev.get(k.upper())
                if v:
                    return v
            return None

        def _get_dt(o):
            if isinstance(o, dict):
                return (_ev_get(o, "dateTime", "DateTime") or
                        _ev_get(o, "date", "Date") or
                        _ev_get(o, "date_time") or '')
            return str(o or '')

        for ev in events:
            ev_id = _ev_get(ev, "id", "Id", "eventId", "uid") or ''
            subj = _ev_get(ev, "subject", "Subject", "title", "Title") or ''
            start_obj = _ev_get(ev, "start", "Start", "startDateTime", "start_time") or {}
            end_obj = _ev_get(ev, "end", "End", "endDateTime", "end_time") or {}
            start_time = _get_dt(start_obj)
            end_time = _get_dt(end_obj)
            loc_raw = _ev_get(ev, "location", "Location", "place")
            location = ''
            if isinstance(loc_raw, dict):
                location = (_ev_get(loc_raw, "displayName", "DisplayName") or
                            _ev_get(loc_raw, "address", "Address") or '')
            else:
                location = str(loc_raw or '')
            attendees = []
            for a in (_ev_get(ev, "attendees", "Attendees", "participants") or []):
                if isinstance(a, dict):
                    name = (_ev_get(a, "displayName", "DisplayName") or
                            _ev_get(a, "name", "Name", "email", "EmailAddress", "address") or '')
                    attendees.append(name)
                else:
                    attendees.append(str(a))
            db.execute(
                "INSERT OR REPLACE INTO calendar_events (id,subject,start_time,end_time,location,attendees,raw_json,synced_at) VALUES(?,?,?,?,?,?,?,?)",
                (ev_id, subj, start_time, end_time, location, json.dumps(attendees), json.dumps(ev), now_iso)
            )
        db.commit()

        # Compute next meeting (compare in local time since MCP returns local-tz datetimes)
        now_local = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
        row = db.execute("SELECT id,subject,start_time,end_time,location,attendees FROM calendar_events WHERE start_time > ? ORDER BY start_time ASC LIMIT 1", (now_local,)).fetchone()
        if row:
            nm = {"id": row[0], "subject": row[1], "start_time": row[2], "end_time": row[3], "location": row[4], "attendees": json.loads(row[5] or "[]")}
            meta_set("next_meeting", json.dumps(nm))
            return nm
        else:
            meta_set("next_meeting", json.dumps({}))
            return None
    except Exception as e:
        print(f"Warning: could not refresh calendar: {e}")
        try:
            return json.loads(meta_get("next_meeting", "{}") or "{}")
        except Exception:
            return None


def _insert_messages(db, emails: list, folder: str) -> int:
    """Upsert a list of raw message dicts into the emails table.
    New rows are inserted; existing rows have mutable fields (is_read, folder, raw_json)
    updated so the DB stays in sync with Outlook.
    Returns count of newly inserted rows."""
    now = _utcnow()
    added = 0
    for e in emails:
        if not e.get("id"):
            continue
        is_read = 1 if e.get("is_read") else 0
        cur = db.execute(
            "INSERT OR IGNORE INTO emails "
            "(id,subject,from_name,from_address,received_date_time,"
            " is_read,body_preview,conversation_key,raw_json,synced_at,folder) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?)",
            (
                e["id"],
                e.get("subject", ""),
                e.get("from_name", ""),
                e.get("from_address", ""),
                e.get("received_date_time", ""),
                is_read,
                _clean(e.get("body_preview", ""), 500),
                _norm_subject(e.get("subject", "")),
                json.dumps(e),
                now,
                folder,
            ),
        )
        if cur.rowcount:
            added += 1
        else:
            # Row exists — update mutable fields without touching cached formatted_body
            db.execute(
                "UPDATE emails SET is_read=?, folder=?, raw_json=?, synced_at=? WHERE id=?",
                (is_read, folder, json.dumps(e), now, e["id"]),
            )
    db.commit()
    return added


def _do_sync():
    _sync_status.update({"phase": "fetching", "progress": "Fetching folder list…", "done": 0, "total": 0})
    efforts, other = _refresh_folders()
    try:
        _sync_status["progress"] = "Refreshing calendar…"
        nm = _refresh_calendar()
        if nm:
            _sync_status["nextMeeting"] = nm
    except Exception as e:
        print(f"  Calendar refresh failed: {e}")

    db = get_db()

    # ── Phase 1: Inbox — store + AI analyze ────────────────────────────────────
    _sync_status["progress"] = "Fetching inbox…"
    inbox_emails = []
    inbox_before = None
    while True:
        _args = {"folder": "Inbox", "top": INBOX_FETCH}
        if inbox_before:
            _args["received_before"] = inbox_before
        _r = call_tool("outlook_mail_list_messages", _args)
        _page = _r.get("messages", []) if isinstance(_r, dict) else []
        if not _page:
            break
        inbox_emails.extend(_page)
        if len(_page) < INBOX_FETCH:
            break
        _oldest = min((e.get("received_date_time","") for e in _page if e.get("received_date_time")), default="")
        if not _oldest or _oldest == inbox_before:
            break
        inbox_before = _oldest
        _sync_status["progress"] = f"Fetching inbox… ({len(inbox_emails)} so far)"

    # Determine which inbox messages are new before inserting
    inbox_ids = [e["id"] for e in inbox_emails if e.get("id")]
    existing_inbox = set()
    if inbox_ids:
        placeholders = ",".join("?" * len(inbox_ids))
        existing_inbox = {
            r[0] for r in db.execute(
                f"SELECT id FROM emails WHERE id IN ({placeholders})", inbox_ids
            ).fetchall()
        }
    new_inbox = [e for e in inbox_emails if e.get("id") and e["id"] not in existing_inbox]

    if inbox_emails:
        _insert_messages(db, inbox_emails, "Inbox")

    # Purge inbox emails that are no longer in Outlook inbox (moved/deleted)
    # inbox_ids is the complete current inbox from Outlook (fully paginated)
    if inbox_ids:
        db_inbox_ids = {r[0] for r in db.execute(
            "SELECT id FROM emails WHERE folder='Inbox'"
        ).fetchall()}
        stale = db_inbox_ids - set(inbox_ids)
        if stale:
            ph = ",".join("?" * len(stale))
            db.execute(f"DELETE FROM emails WHERE id IN ({ph})", list(stale))
            # Remove thread records that no longer have any inbox emails
            db.execute(
                "DELETE FROM threads WHERE conversation_key NOT IN "
                "(SELECT DISTINCT conversation_key FROM emails WHERE folder='Inbox')"
            )
            db.commit()
            print(f"Sync: purged {len(stale)} stale inbox email(s) no longer in Outlook")

    # AI analyze new inbox threads
    threads_updated = 0
    if new_inbox:
        print(f"Sync: {len(new_inbox)} new inbox email(s)")
        affected_keys = list({_norm_subject(e.get("subject", "")) for e in new_inbox})
        total = len(affected_keys)
        _sync_status.update({"phase": "analyzing", "done": 0, "total": total})

        # Seed from already-analyzed threads, then accumulate new ones in-memory
        existing_topics: list[str] = [r[0] for r in db.execute(
            "SELECT DISTINCT topic FROM threads WHERE topic != ''"
        ).fetchall()]

        for idx, ck in enumerate(affected_keys):
            # Only consider inbox emails for thread analysis
            rows = db.execute(
                "SELECT * FROM emails WHERE conversation_key=? AND folder='Inbox' "
                "ORDER BY received_date_time ASC",
                (ck,)
            ).fetchall()
            if not rows:
                continue
            thread_emails = [dict(r) for r in rows]
            display_subj = _clean(thread_emails[-1].get("subject", ck), 55)
            _sync_status["progress"] = f"Analyzing {idx+1}/{total}: \"{display_subj}\""

            try:
                result = analyze_thread(thread_emails, efforts, other, existing_topics=existing_topics)
            except Exception as ex:
                print(f"  Failed to analyze {ck!r}: {ex}")
                _sync_status["done"] = idx + 1
                continue

            latest = thread_emails[-1]
            participants = list(dict.fromkeys(
                (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
                for e in thread_emails
                if (e.get("from_name") or e.get("from_address"))
            ))[:8]
            has_unread = any(not e.get("is_read") for e in thread_emails)

            db.execute(
                "INSERT OR REPLACE INTO threads "
                "(conversation_key,subject,topic,action,urgency,summary,"
                " suggested_reply,suggested_folder,participants,email_ids,"
                " latest_id,message_count,has_unread,latest_received,updated_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    ck,
                    latest["subject"],
                    result.get("topic", "General"),
                    result.get("action", "read"),
                    result.get("urgency", "low"),
                    result.get("summary", ""),
                    result.get("suggestedReply", ""),
                    result.get("suggestedFolder", ""),
                    json.dumps(participants),
                    json.dumps([e["id"] for e in thread_emails]),
                    latest["id"],
                    len(thread_emails),
                    1 if has_unread else 0,
                    latest["received_date_time"],
                    _utcnow(),
                ),
            )
            db.commit()
            threads_updated += 1
            new_topic = _normalize_topic(result.get("topic", ""))
            if new_topic and new_topic not in existing_topics:
                existing_topics.append(new_topic)
            _sync_status["done"] = idx + 1
            print(f"  ✓ {display_subj!r} → {result.get('action')} [{result.get('urgency')}] [{new_topic}]")

    # Recompute has_unread for all threads based on current is_read state in emails table.
    # This keeps threads in sync even when emails are read in native Outlook/mobile.
    db.executescript("""
        UPDATE threads SET has_unread = (
            SELECT CASE WHEN COUNT(*) > 0 THEN 1 ELSE 0 END
            FROM emails
            WHERE emails.conversation_key = threads.conversation_key
              AND emails.folder = 'Inbox'
              AND emails.is_read = 0
        );
    """)
    db.commit()

    # ── Phase 2: All other folders — store only, no AI ─────────────────────────
    top_level = json.loads(meta_get("folders_raw", "[]"))
    sync_folders = []
    for f in top_level:
        name = f.get("display_name") or f.get("displayName", "")
        if name not in SKIP_SYNC_FOLDERS:
            sync_folders.append(name)
    for sub in efforts:
        sync_folders.append(f"Efforts/{sub}")

    total_folders = len(sync_folders)
    _sync_status.update({"phase": "fetching", "done": 0, "total": total_folders})
    for fi, folder_name in enumerate(sync_folders):
        _sync_status["progress"] = f"Syncing {fi+1}/{total_folders}: {folder_name}…"
        try:
            # Paginate backwards using received_before until we get no new messages
            folder_total_added = 0
            before = None
            while True:
                args = {"folder": folder_name, "top": FOLDER_FETCH}
                if before:
                    args["received_before"] = before
                result = call_tool("outlook_mail_list_messages", args)
                folder_emails = result.get("messages", []) if isinstance(result, dict) else []
                if not folder_emails:
                    break
                added = _insert_messages(db, folder_emails, folder_name)
                folder_total_added += added
                # If we added nothing new this page, the rest will also be known — stop
                if added == 0:
                    break
                # If fewer results than requested, we've hit the end
                if len(folder_emails) < FOLDER_FETCH:
                    break
                # Advance cursor to oldest message in this batch
                oldest = min(e.get("received_date_time", "") for e in folder_emails if e.get("received_date_time"))
                if not oldest or oldest == before:
                    break
                before = oldest
            if folder_total_added:
                print(f"  {folder_name}: {folder_total_added} new message(s)")
        except Exception as ex:
            print(f"  Warning: could not sync folder '{folder_name}': {ex}")
        _sync_status["done"] = fi + 1

    _sync_status.update({"phase": "done", "progress": f"Done — {threads_updated} thread(s) updated."})
    return len(new_inbox), threads_updated


def run_sync():
    if not _sync_lock.acquire(blocking=False):
        return
    _sync_status.update({"running": True, "lastError": None})
    try:
        added, updated = _do_sync()
        _sync_status.update({
            "running": False,
            "lastSync": _utcnow(),
            "emailsAdded": added,
            "threadsUpdated": updated,
        })
    except Exception as e:
        _sync_status.update({"running": False, "lastError": str(e)})
        print(f"Sync error: {e}")
    finally:
        _sync_lock.release()


def _sync_loop():
    from mcp_client import _session_ready
    _session_ready.wait(timeout=30)
    while True:
        run_sync()
        time.sleep(SYNC_INTERVAL)
