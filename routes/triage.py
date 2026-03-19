"""
routes/triage.py — Triage-related API routes for Outlook Express.
"""
import json
import threading
from datetime import datetime, timezone

from flask import Blueprint, jsonify, request

from db import get_db, meta_get, _thread_to_dict
from sync import _sync_status, _sync_lock, run_sync
from ai import analyze_thread, _normalize_topic
from mcp_client import call_tool

bp = Blueprint("triage", __name__)


def _utcnow() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _clean(s, n=None) -> str:
    import re
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s or ''))
    return s[:n] if n else s


def _norm_subject(subj: str) -> str:
    import re
    s = re.sub(r'^(RE|FW|FWD|AW|R|RES|SV)[\s:]+', '', str(subj or ''), flags=re.IGNORECASE)
    return s.strip().lower() or "no-subject"


def _normalize_msg(m: dict) -> dict:
    import re
    from_name    = m.get("from_name") or ""
    from_address = m.get("from_address") or ""
    received     = m.get("received_date_time") or ""

    raw_html = m.get("body_content") or ""
    if raw_html:
        # Strip quoted/forwarded history before any other processing
        raw_html = _strip_quoted_html(raw_html)
        # Replace block-level elements with newlines to preserve paragraph structure
        raw_html = re.sub(r'<br\s*/?>', '\n', raw_html, flags=re.IGNORECASE)
        raw_html = re.sub(r'</?(?:div|p|tr|li|blockquote|hr)[^>]*>', '\n', raw_html, flags=re.IGNORECASE)
        # Strip all remaining tags
        body_text = re.sub(r'<[^>]+>', '', raw_html)
        body_text = re.sub(r'&nbsp;', ' ', body_text)
        body_text = re.sub(r'&#\d+;|&[a-z]+;', ' ', body_text)
        body_text = re.sub(r'[ \t]{2,}', ' ', body_text)
        body_text = re.sub(r'\n{3,}', '\n\n', body_text).strip()
        # Fallback text-level stripping for non-HTML quote patterns
        text_cut = re.search(
            r'\n[Ff]rom:\s.{3,120}\n\s*(?:[Ss]ent|[Dd]ate|[Tt]o|[Cc]c)\s*:',
            body_text
        )
        if text_cut:
            body_text = body_text[:text_cut.start()].strip()
    else:
        body_text = m.get("body_preview") or ""

    to_recips  = _parse_recipients(m.get("to_recipients") or m.get("toRecipients"))
    cc_recips  = _parse_recipients(m.get("cc_recipients") or m.get("ccRecipients"))

    # Sanitize raw HTML for safe iframe rendering
    body_html = ""
    raw_content = m.get("body_content") or ""
    if raw_content and (m.get("body_content_type", "").upper() == "HTML" or raw_content.lstrip().startswith("<")):
        h = raw_content
        h = re.sub(r'<script\b[^>]*>.*?</script>', '', h, flags=re.IGNORECASE | re.DOTALL)
        h = re.sub(r'<style\b[^>]*>.*?</style>', lambda mo: mo.group(), h, flags=re.IGNORECASE | re.DOTALL)  # keep styles
        h = re.sub(r'\s+on\w+="[^"]*"', '', h, flags=re.IGNORECASE)
        h = re.sub(r"\s+on\w+='[^']*'", '', h, flags=re.IGNORECASE)
        # Neutralize CID inline images — show placeholder
        h = re.sub(r'src=["\']cid:[^"\']*["\']',
                   'src="" alt="[inline image — not available]" style="display:inline-block;padding:4px 8px;background:#eee;color:#666;font-size:11px;border-radius:3px"',
                   h, flags=re.IGNORECASE)
        body_html = h

    return {
        "id":                 m.get("id", ""),
        "subject":            m.get("subject", ""),
        "from_name":          from_name,
        "from_address":       from_address,
        "received_date_time": received,
        "is_read":            m.get("is_read"),
        "body":               body_text,
        "body_html":          body_html,
        "to_recipients":      to_recips,
        "cc_recipients":      cc_recips,
    }


def _strip_quoted_html(html: str) -> str:
    markers = [
        'id="mail-editor-reference-message-container"',
        'id="divRplyFwdMsg"',
        'id="appendonsend"',
        'class="gmail_quote"',
        'id="divTaggedContent"',
    ]
    lower = html.lower()
    cut = len(html)
    for marker in markers:
        idx = lower.find(marker.lower())
        if 0 < idx < cut:
            tag_start = html.rfind('<', 0, idx)
            if tag_start != -1:
                cut = tag_start
    return html[:cut]


def _parse_recipients(raw) -> list:
    result = []
    for r in (raw or []):
        if not isinstance(r, dict):
            continue
        name = (r.get("name") or r.get("display_name") or r.get("displayName") or
                r.get("emailAddress", {}).get("name") or "")
        addr = (r.get("address") or r.get("email") or
                r.get("emailAddress", {}).get("address") or "")
        if addr:
            result.append({"name": name.strip(), "address": addr.strip()})
    return result


@bp.route("/api/threads")
def api_threads():
    db = get_db()
    # Sort by recency — groups and threads within groups will both be newest-first
    rows = db.execute("""
        SELECT * FROM threads ORDER BY latest_received DESC
    """).fetchall()
    threads = [_thread_to_dict(r) for r in rows]

    # Build groups preserving insertion order (first occurrence of topic = most recent thread)
    groups: dict[str, list] = {}
    order: list[str] = []
    for t in threads:
        topic = t["topic"]
        if topic not in groups:
            groups[topic] = []
            order.append(topic)
        groups[topic].append(t)

    # order already reflects most-recent-thread-first per topic
    result = [{"topic": k, "threads": groups[k]} for k in order]
    latest_ts = max((t["updatedAt"] for t in threads), default="")

    try:
        nm = json.loads(meta_get("next_meeting", "{}") or "{}")
    except Exception:
        nm = {}
    return jsonify({
        "groups": result,
        "latestTs": latest_ts,
        "threadCount": len(threads),
        "emailCount": db.execute("SELECT COUNT(*) FROM emails").fetchone()[0],
        "syncStatus": {**_sync_status},
        "nextMeeting": nm,
    })


@bp.route("/api/updates")
def api_updates():
    since = request.args.get("since", "")
    if not since:
        since = _utcnow()
    db = get_db()
    rows = db.execute(
        "SELECT * FROM threads WHERE updated_at > ? ORDER BY updated_at ASC",
        (since,)
    ).fetchall()
    threads = [_thread_to_dict(r) for r in rows]
    latest_ts = threads[-1]["updatedAt"] if threads else since
    try:
        nm = json.loads(meta_get("next_meeting", "{}") or "{}")
    except Exception:
        nm = {}
    return jsonify({
        "threads": threads,
        "latestTs": latest_ts,
        "syncStatus": {**_sync_status},
        "nextMeeting": nm,
    })


@bp.route("/api/suggested_reply", methods=["POST"])
def api_suggested_reply():
    payload = request.json or {}
    conv_key = payload.get("conversationKey", "")
    reply_context = payload.get("context", "")
    if not conv_key:
        return jsonify({"error": "conversationKey required"}), 400
    db = get_db()
    rows = db.execute("SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)).fetchall()
    if not rows:
        return jsonify({"error": "No messages found for thread."}), 404
    emails = [dict(r) for r in rows]
    efforts = json.loads(meta_get("efforts_subfolders", "[]"))
    other = json.loads(meta_get("other_folders", "[]"))
    try:
        result = analyze_thread(emails, efforts, other, reply_context=reply_context)
        latest = emails[-1]
        participants = list(dict.fromkeys(
            (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
            for e in emails if (e.get("from_name") or e.get("from_address"))
        ))[:8]
        email_ids = [e["id"] for e in emails]
        has_unread = any(not e.get("is_read") for e in emails)
        db.execute(
            "INSERT OR REPLACE INTO threads "
            "(conversation_key,subject,topic,action,urgency,summary,"
            " suggested_reply,suggested_folder,participants,email_ids,"
            " latest_id,message_count,has_unread,latest_received,updated_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                conv_key,
                latest.get("subject", ""),
                result.get("topic", "General"),
                result.get("action", "read"),
                result.get("urgency", "low"),
                result.get("summary", ""),
                result.get("suggestedReply", ""),
                result.get("suggestedFolder", ""),
                json.dumps(participants),
                json.dumps(email_ids),
                latest.get("id", ""),
                len(emails),
                1 if has_unread else 0,
                latest.get("received_date_time", ""),
                _utcnow(),
            )
        )
        db.commit()
        return jsonify({"reply": result.get("suggestedReply", "")})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/generate_reply", methods=["POST"])
def api_generate_reply():
    from ai import generate_reply_ai
    conv_key    = request.json.get("conversationKey", "")
    user_prompt = request.json.get("userPrompt", "").strip()
    if not user_prompt:
        return jsonify({"error": "userPrompt required"}), 400

    db = get_db()
    thread_row = db.execute("SELECT * FROM threads WHERE conversation_key=?", (conv_key,)).fetchone()
    email_rows = db.execute(
        "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)
    ).fetchall()
    thread = dict(thread_row) if thread_row else {}
    emails = [dict(r) for r in email_rows]

    subject = thread.get("subject") or (emails[-1].get("subject", "") if emails else "")
    context_emails = emails[-6:]
    msgs_text = "\n\n".join(
        f"From: {_clean(e.get('from_name') or e.get('from_address','Unknown'),50)} | {(e.get('received_date_time',''))[:10]}\n"
        f"{_clean(e.get('body_preview','(no preview)'), 600)}"
        for e in context_emails
    )

    try:
        reply = generate_reply_ai(subject, msgs_text, user_prompt)
        return jsonify({"reply": reply})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/reanalyze_all", methods=["POST"])
def api_reanalyze_all():
    if _sync_status["running"]:
        return jsonify({"ok": False, "error": "Sync already running"})

    def _do_reanalyze():
        efforts = json.loads(meta_get("efforts_subfolders", "[]"))
        other   = json.loads(meta_get("other_folders", "[]"))
        db = get_db()
        keys = [r[0] for r in db.execute("SELECT conversation_key FROM threads").fetchall()]
        total = len(keys)
        _sync_status.update({"running": True, "phase": "analyzing", "done": 0, "total": total,
                              "progress": f"Re-analyzing {total} threads…"})
        updated = 0
        for idx, ck in enumerate(keys):
            rows = db.execute(
                "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (ck,)
            ).fetchall()
            if not rows:
                _sync_status["done"] = idx + 1
                continue
            thread_emails = [dict(r) for r in rows]
            display_subj = _clean(thread_emails[-1].get("subject", ck), 55)
            _sync_status["progress"] = f"Re-analyzing {idx+1}/{total}: \"{display_subj}\""
            try:
                # No existing_topics — each thread picks a fresh specific label independently.
                # Similar threads will naturally land on the same project name.
                result = analyze_thread(thread_emails, efforts, other)
                db.execute(
                    "UPDATE threads SET topic=?,action=?,urgency=?,summary=?,"
                    "suggested_reply=?,suggested_folder=?,updated_at=? WHERE conversation_key=?",
                    (_normalize_topic(result.get("topic","")), result.get("action","read"),
                     result.get("urgency","low"), result.get("summary",""),
                     result.get("suggestedReply",""), result.get("suggestedFolder",""),
                     _utcnow(), ck)
                )
                db.commit()
                updated += 1
                print(f"  topic: {result.get('topic','?')!r} — {display_subj[:45]}")
            except Exception as ex:
                print(f"  Re-analyze error for {ck}: {ex}")
            _sync_status["done"] = idx + 1
        _sync_status.update({"running": False, "lastSync": _utcnow(), "threadsUpdated": updated,
                              "phase": "done", "progress": f"Re-analyzed {updated}/{total} threads."})

    threading.Thread(target=_do_reanalyze, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})


@bp.route("/api/resync_thread", methods=["POST"])
def api_resync_thread():
    conv_key = (request.json or {}).get("conversationKey", "")
    if not conv_key:
        return jsonify({"ok": False, "error": "conversationKey required"}), 400
    if _sync_status["running"]:
        return jsonify({"ok": False, "error": "Sync already running"}), 409

    def _do_resync():
        db = get_db()
        try:
            # ── Step 1: Collect known IDs from the thread record (authoritative source)
            thread_row = db.execute(
                "SELECT * FROM threads WHERE conversation_key=?", (conv_key,)
            ).fetchone()
            stored_ids = []
            if thread_row:
                try:
                    stored_ids = json.loads(dict(thread_row).get("email_ids") or "[]")
                except Exception:
                    pass

            db_ids = [r[0] for r in db.execute(
                "SELECT id FROM emails WHERE conversation_key=?", (conv_key,)
            ).fetchall()]
            all_ids = list(dict.fromkeys(stored_ids + db_ids))  # deduplicated, stored_ids first

            if not all_ids:
                _sync_status.update({"running": False, "lastError": "No messages found for thread."})
                return

            total_steps = len(all_ids) * 2 + 1  # fetch + format + analyze
            _sync_status.update({"phase": "fetching", "done": 0, "total": total_steps,
                                  "progress": f"Blowing away cached data for {len(all_ids)} message(s)…"})

            # ── Step 2: Nuke all existing data for this thread
            db.execute("DELETE FROM emails WHERE conversation_key=?", (conv_key,))
            db.execute("DELETE FROM threads WHERE conversation_key=?", (conv_key,))
            db.commit()

            # ── Step 3: Re-fetch every message fresh from Outlook
            _sync_status["progress"] = f"Re-fetching {len(all_ids)} messages from Outlook…"
            fresh_msgs = []   # list of normalized msg dicts with full body
            now = _utcnow()
            for i, msg_id in enumerate(all_ids):
                _sync_status["done"] = i
                _sync_status["progress"] = f"Fetching message {i+1}/{len(all_ids)}…"
                try:
                    resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
                    if isinstance(resp, dict) and resp:
                        raw = resp["messages"][0] if isinstance(resp.get("messages"), list) and resp["messages"] else resp
                    else:
                        print(f"  Resync: empty response for {msg_id}, skipping")
                        continue
                    nm = _normalize_msg(raw)
                    fresh_msgs.append((msg_id, raw, nm))

                    # Re-insert with FULL quote-stripped body stored (no truncation limit)
                    ck = _norm_subject(raw.get("subject", "") or nm.get("subject", ""))
                    full_body = nm.get("body", "")
                    db.execute(
                        "INSERT OR REPLACE INTO emails "
                        "(id,subject,from_name,from_address,received_date_time,"
                        " is_read,body_preview,conversation_key,raw_json,synced_at,formatted_body) "
                        "VALUES(?,?,?,?,?,?,?,?,?,?,NULL)",
                        (
                            msg_id,
                            raw.get("subject", nm.get("subject", "")),
                            nm.get("from_name", ""),
                            nm.get("from_address", ""),
                            nm.get("received_date_time", ""),
                            1 if raw.get("is_read") else 0,
                            full_body,   # store full quote-stripped body, no char limit
                            conv_key,    # keep original conv_key so thread stays together
                            json.dumps(raw),
                            now,
                        )
                    )
                    db.commit()
                except Exception as ex:
                    print(f"  Resync: failed to re-fetch {msg_id}: {ex}")

            if not fresh_msgs:
                _sync_status.update({"running": False, "lastError": "Could not re-fetch any messages."})
                return

            # ── Step 4: Re-analyze the thread (now with full bodies in body_preview)
            thread_emails = [dict(r) for r in db.execute(
                "SELECT * FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC", (conv_key,)
            ).fetchall()]
            display_subj = _clean(thread_emails[-1].get("subject", conv_key), 55)
            _sync_status.update({"done": len(all_ids), "progress": f"Re-analyzing \"{display_subj}\"…"})

            efforts = json.loads(meta_get("efforts_subfolders", "[]"))
            other   = json.loads(meta_get("other_folders", "[]"))
            existing_topics = [r[0] for r in db.execute(
                "SELECT DISTINCT topic FROM threads WHERE topic != '' AND conversation_key != ?", (conv_key,)
            ).fetchall()]
            result  = analyze_thread(thread_emails, efforts, other, existing_topics=existing_topics)

            latest       = thread_emails[-1]
            participants = list(dict.fromkeys(
                (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
                for e in thread_emails if (e.get("from_name") or e.get("from_address"))
            ))[:8]
            fetched_ids  = [e["id"] for e in thread_emails]
            has_unread   = any(not e.get("is_read") for e in thread_emails)

            db.execute(
                "INSERT OR REPLACE INTO threads "
                "(conversation_key,subject,topic,action,urgency,summary,"
                " suggested_reply,suggested_folder,participants,email_ids,"
                " latest_id,message_count,has_unread,latest_received,updated_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    conv_key,
                    latest["subject"],
                    result.get("topic", "General"),
                    result.get("action", "read"),
                    result.get("urgency", "low"),
                    result.get("summary", ""),
                    result.get("suggestedReply", ""),
                    result.get("suggestedFolder", ""),
                    json.dumps(participants),
                    json.dumps(fetched_ids),
                    latest["id"],
                    len(thread_emails),
                    1 if has_unread else 0,
                    latest["received_date_time"],
                    _utcnow(),
                )
            )
            db.commit()

            # ── Step 5: Re-run AI formatting on every message (warm cache with clean content)
            _sync_status["progress"] = f"Re-formatting {len(fresh_msgs)} messages with AI…"
            for i, (msg_id, _raw, nm) in enumerate(fresh_msgs):
                _sync_status["done"] = len(all_ids) + 1 + i
                try:
                    paras = format_message_ai(nm)
                    db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                               (json.dumps(paras), msg_id))
                    db.commit()
                except Exception as ex:
                    print(f"  Resync format error for {msg_id}: {ex}")

            _sync_status.update({
                "running": False, "threadsUpdated": 1, "lastSync": _utcnow(),
                "phase": "done",
                "progress": f"Thread fully resynced — {len(thread_emails)} message(s) rebuilt.",
            })
        except Exception as ex:
            _sync_status.update({"running": False, "lastError": str(ex)})
            print(f"Resync thread error: {ex}")

    def _run_resync():
        if not _sync_lock.acquire(blocking=False):
            return
        _sync_status.update({"running": True, "lastError": None})
        try:
            _do_resync()
        finally:
            _sync_lock.release()

    if not _sync_status["running"]:
        threading.Thread(target=_run_resync, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})


@bp.route("/api/sync_now", methods=["POST"])
def api_sync_now():
    if not _sync_status["running"]:
        threading.Thread(target=run_sync, daemon=True).start()
    return jsonify({"ok": True, "syncStatus": {**_sync_status}})
