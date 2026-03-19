"""
routes/mail.py — Mail-related API routes for Outlook Express.
"""
import json
import re
from datetime import datetime, timezone

from flask import Blueprint, Response, jsonify, request, stream_with_context

from db import get_db, meta_get, remove_thread, get_my_email, rebuild_contacts
from mcp_client import call_tool
from ai import format_message_ai, generate_reply_ai, summarize_message_ai, _format_prompt, _parse_format_response, _get_ai
from config import ANALYSIS_MODEL

bp = Blueprint("mail", __name__)

_BLANK_GIF = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"


def _embed_cid_images(html: str) -> str:
    """Replace cid: and external http(s): image src refs with placeholders."""
    # Replace unresolved cid: refs with blank gif
    html = re.sub(r'src=(["\'])cid:[^"\']*\1', lambda m: f'src={m.group(1)}{_BLANK_GIF}{m.group(1)}', html, flags=re.IGNORECASE)
    # Replace external http(s) image srcs with blank gif (no credentials needed)
    html = re.sub(r'src=(["\'])https?://[^"\']*\1', lambda m: f'src={m.group(1)}{_BLANK_GIF}{m.group(1)}', html, flags=re.IGNORECASE)
    return html


def _utcnow() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _clean(s, n=None) -> str:
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s or ''))
    return s[:n] if n else s


def _norm_subject(subj: str) -> str:
    s = re.sub(r'^(RE|FW|FWD|AW|R|RES|SV)[\s:]+', '', str(subj or ''), flags=re.IGNORECASE)
    return s.strip().lower() or "no-subject"


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


def _normalize_msg(m: dict) -> dict:
    from_name    = m.get("from_name") or ""
    from_address = m.get("from_address") or ""
    received     = m.get("received_date_time") or ""

    raw_html = m.get("body_content") or ""
    if raw_html:
        raw_html = _strip_quoted_html(raw_html)
        raw_html = re.sub(r'<br\s*/?>', '\n', raw_html, flags=re.IGNORECASE)
        raw_html = re.sub(r'</?(?:div|p|tr|li|blockquote|hr)[^>]*>', '\n', raw_html, flags=re.IGNORECASE)
        body_text = re.sub(r'<[^>]+>', '', raw_html)
        body_text = re.sub(r'&nbsp;', ' ', body_text)
        body_text = re.sub(r'&#\d+;|&[a-z]+;', ' ', body_text)
        body_text = re.sub(r'[ \t]{2,}', ' ', body_text)
        body_text = re.sub(r'\n{3,}', '\n\n', body_text).strip()
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

    body_html = ""
    raw_content = m.get("body_content") or ""
    if raw_content and (m.get("body_content_type", "").upper() == "HTML" or raw_content.lstrip().startswith("<")):
        h = raw_content
        h = re.sub(r'<script\b[^>]*>.*?</script>', '', h, flags=re.IGNORECASE | re.DOTALL)
        h = re.sub(r'<style\b[^>]*>.*?</style>', lambda mo: mo.group(), h, flags=re.IGNORECASE | re.DOTALL)
        h = re.sub(r'\s+on\w+="[^"]*"', '', h, flags=re.IGNORECASE)
        h = re.sub(r"\s+on\w+='[^']*'", '', h, flags=re.IGNORECASE)
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


_FOLDER_ICONS = {
    "Inbox": "📥", "Sent Items": "📤", "Archive": "🗄️",
    "Drafts": "📝", "Deleted Items": "🗑️", "Junk Email": "🚫",
}
_FOLDERS_SKIP_DISPLAY = {
    "Drafts", "Outbox", "Junk Email",
    "Conversation History", "RSS Feeds", "Sync Issues", "Scheduled",
}


@bp.route("/api/thread_messages")
def api_thread_messages():
    ids = request.args.getlist("id")
    conv_key = request.args.get("conversationKey")
    if not ids and conv_key:
        db = get_db()
        rows = db.execute(
            "SELECT id FROM emails WHERE conversation_key=? ORDER BY received_date_time ASC",
            (conv_key,)
        ).fetchall()
        ids = [r["id"] for r in rows]
    if not ids:
        return jsonify({"messages": []})
    db = get_db()
    rows = db.execute(
        "SELECT * FROM emails WHERE id IN ({})".format(",".join("?" * len(ids))),
        ids
    ).fetchall()
    db_msgs = {r["id"]: dict(r) for r in rows}

    result = []
    for msg_id in ids:
        db_row = db_msgs.get(msg_id, {"id": msg_id})
        msg = _normalize_msg(db_row)
        # Include cached body_html from DB if the stream endpoint has already fetched it
        if not msg.get("body_html") and isinstance(db_row, dict) and db_row.get("body_html"):
            msg["body_html"] = db_row["body_html"]
        result.append(msg)
    result.sort(key=lambda m: m.get("received_date_time", ""), reverse=True)
    return jsonify({"messages": result})


@bp.route("/api/format_message")
def api_format_message():
    msg_id = request.args.get("id", "")
    db = get_db()
    row = db.execute("SELECT * FROM emails WHERE id=?", (msg_id,)).fetchone()

    # Return cached formatted body if available
    if row and row["formatted_body"]:
        try:
            return jsonify({"paragraphs": json.loads(row["formatted_body"]), "cached": True})
        except Exception:
            pass

    fallback = dict(row) if row else {"id": msg_id}
    try:
        resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
        if isinstance(resp, dict) and "messages" in resp:
            raw = resp["messages"][0] if resp["messages"] else fallback
        else:
            raw = fallback
        msg = _normalize_msg(raw)
    except Exception:
        msg = _normalize_msg(fallback)
    paragraphs = format_message_ai(msg)

    # Persist to DB so future opens are instant
    if row:
        try:
            db.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                       (json.dumps(paragraphs), msg_id))
            db.commit()
        except Exception:
            pass

    return jsonify({"paragraphs": paragraphs})


@bp.route("/api/format_message_stream")
def api_format_message_stream():
    msg_id = request.args.get("id", "")
    db = get_db()
    row = db.execute("SELECT * FROM emails WHERE id=?", (msg_id,)).fetchone()

    # Serve from cache immediately as a single done event (include body_html if cached)
    if row and row["formatted_body"]:
        try:
            paras = json.loads(row["formatted_body"])
            cached_html = row["body_html"] or ""
            # Replace any cid:/external image refs with placeholders
            if cached_html:
                cached_html = _embed_cid_images(cached_html)
                try:
                    db.execute("UPDATE emails SET body_html=? WHERE id=?", (cached_html, msg_id))
                    db.commit()
                except Exception:
                    pass
            def _cached():
                yield f"data: {json.dumps({'type':'done','paragraphs':paras,'body_html':cached_html})}\n\n"
            return Response(stream_with_context(_cached()), mimetype="text/event-stream",
                            headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})
        except Exception:
            pass

    # Fetch full message body from Outlook
    fallback = dict(row) if row else {"id": msg_id}
    try:
        resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
        if isinstance(resp, dict) and "messages" in resp:
            raw = resp["messages"][0] if resp["messages"] else fallback
        elif isinstance(resp, dict) and resp:
            raw = resp
        else:
            raw = fallback
        msg = _normalize_msg(raw)
        _bc = raw.get("body_content","") if isinstance(raw,dict) else ""
        print(f"DEBUG body_content[0:200]: {_bc[:200]!r}")
    except Exception:
        msg = _normalize_msg(fallback)

    body_html = msg.get("body_html") or ""
    print(f"DEBUG body_html[0:200]: {body_html[:200]!r}")
    # Embed CID inline images as base64 data URIs
    if body_html:
        body_html = _embed_cid_images(body_html)
    print(f"DEBUG after embed body_html[0:200]: {body_html[:200]!r}")
    body = (msg.get("body") or msg.get("body_preview") or "").strip()
    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    date = (msg.get("received_date_time") or "")[:10]

    # Persist body_html to DB so it's available immediately on future opens
    if row and body_html:
        try:
            db.execute("UPDATE emails SET body_html=? WHERE id=?", (body_html, msg_id))
            db.commit()
        except Exception:
            pass

    if not body:
        def _empty():
            paras = [{"text": "(no content)", "intent": "FYI", "emoji": "📭", "fact_concern": None}]
            yield f"data: {json.dumps({'type':'done','paragraphs':paras,'body_html':body_html})}\n\n"
        return Response(stream_with_context(_empty()), mimetype="text/event-stream",
                        headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})

    prompt = _format_prompt(body, from_name, date)

    @stream_with_context
    def _stream():
        full_text = ""
        try:
            with _get_ai().messages.stream(
                model=ANALYSIS_MODEL,
                max_tokens=3000,
                messages=[{"role": "user", "content": prompt}],
            ) as stream:
                for chunk in stream.text_stream:
                    full_text += chunk
                    yield f"data: {json.dumps({'type':'token','text':chunk})}\n\n"
        except Exception as ex:
            print(f"  Stream format error: {ex}")
            paras = [{"text": p.strip(), "intent": "FYI", "emoji": "📄", "fact_concern": None}
                     for p in body.split('\n\n') if p.strip()][:20]
            yield f"data: {json.dumps({'type':'done','paragraphs':paras,'body_html':body_html})}\n\n"
            return

        try:
            paras = _parse_format_response(full_text, body)
        except Exception:
            paras = [{"text": p.strip(), "intent": "FYI", "emoji": "📄", "fact_concern": None}
                     for p in body.split('\n\n') if p.strip()][:20]

        if row:
            try:
                db2 = get_db()
                db2.execute("UPDATE emails SET formatted_body=? WHERE id=?",
                            (json.dumps(paras), msg_id))
                db2.commit()
            except Exception:
                pass

        yield f"data: {json.dumps({'type':'done','paragraphs':paras,'body_html':body_html})}\n\n"

    return Response(_stream(), mimetype="text/event-stream",
                    headers={"Cache-Control":"no-cache","X-Accel-Buffering":"no"})


@bp.route("/api/summarize_message")
def api_summarize_message():
    msg_id = request.args.get("id", "")
    if not msg_id:
        return jsonify({"error": "id required"}), 400
    db = get_db()
    row = db.execute("SELECT * FROM emails WHERE id=?", (msg_id,)).fetchone()
    msg = dict(row) if row else {"id": msg_id}
    summary = summarize_message_ai(msg)
    return jsonify({"summary": summary})


@bp.route("/api/message_recipients")
def api_message_recipients():
    """Fetch to/cc recipients for a specific message from Outlook MCP."""
    msg_id = request.args.get("id", "")
    if not msg_id:
        return jsonify({"to": [], "cc": []})
    try:
        resp = call_tool("outlook_mail_get_message", {"message_id": msg_id})
        raw = None
        if isinstance(resp, dict) and resp.get("messages"):
            raw = resp["messages"][0]
        elif isinstance(resp, dict):
            raw = resp
        if not raw:
            return jsonify({"to": [], "cc": []})
        msg = _normalize_msg(raw)
        return jsonify({"to": msg.get("to_recipients", []), "cc": msg.get("cc_recipients", [])})
    except Exception as e:
        return jsonify({"to": [], "cc": [], "error": str(e)})


@bp.route("/api/top_contacts")
def api_top_contacts():
    n = int(request.args.get("n", 10))
    db = get_db()
    rows = db.execute(
        "SELECT email, name, frequency FROM contacts ORDER BY frequency DESC LIMIT ?", (n,)
    ).fetchall()
    return jsonify({"contacts": [dict(r) for r in rows]})


@bp.route("/api/rebuild_contacts", methods=["POST"])
def api_rebuild_contacts():
    my_email = get_my_email()
    count = rebuild_contacts(my_email)
    return jsonify({"ok": True, "count": count})


@bp.route("/api/reply/<latest_id>", methods=["POST"])
def api_reply(latest_id):
    body     = request.json.get("body", "")
    conv_key = request.json.get("conversationKey", "")
    to_list  = request.json.get("to", [])   # list of email address strings
    cc_list  = request.json.get("cc", [])
    try:
        draft_args = {
            "source_message_id": latest_id,
            "operation": "ReplyAll",
            "bodyText": body,
        }
        if to_list:
            draft_args["to"] = to_list
        if cc_list:
            draft_args["cc"] = cc_list
        draft = call_tool("outlook_mail_draft_message", draft_args)
        draft_id = None
        if isinstance(draft, dict):
            draft_id = (draft.get("draft_id") or draft.get("id") or
                        (draft.get("widgetState") or {}).get("draftId"))
        if not draft_id:
            return jsonify({"error": f"Could not get draft ID: {draft}"}), 500
        call_tool("outlook_mail_send_message", {"draft_id": draft_id})
        if conv_key:
            remove_thread(conv_key)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/send_new", methods=["POST"])
def api_send_new():
    payload = request.json or {}
    to_list = payload.get("to", [])
    cc_list = payload.get("cc", [])
    subject = payload.get("subject", "")
    body = payload.get("body", "")
    try:
        draft_args = {
            "operation": "New",
            "subject": subject,
            "bodyText": body,
        }
        if to_list:
            draft_args["to"] = to_list
        if cc_list:
            draft_args["cc"] = cc_list
        draft = call_tool("outlook_mail_draft_message", draft_args)
        draft_id = None
        if isinstance(draft, dict):
            draft_id = (draft.get("draft_id") or draft.get("id") or
                        (draft.get("widgetState") or {}).get("draftId"))
        if not draft_id:
            return jsonify({"error": f"Could not get draft ID: {draft}"}), 500
        call_tool("outlook_mail_send_message", {"draft_id": draft_id})
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/delete", methods=["POST"])
def api_delete():
    ids = request.json.get("ids", [])
    conv_key = request.json.get("conversationKey", "")
    for msg_id in ids:
        try:
            print(f"  Moving to Deleted Items: {msg_id[:40]}…")
            call_tool("outlook_mail_move_message", {
                "message_id": msg_id,
                "destination_folder": "Deleted Items",
            })
            print(f"  ✓ Moved: {msg_id[:40]}")
        except Exception as e:
            print(f"  ~ Skipping inaccessible message ({e}): {msg_id[:40]}")
    if conv_key:
        remove_thread(conv_key)
    return jsonify({"ok": True})


@bp.route("/api/move", methods=["POST"])
def api_move():
    ids = request.json.get("ids", [])
    folder = request.json.get("folder", "")
    conv_key = request.json.get("conversationKey", "")
    errors = []
    for msg_id in ids:
        try:
            call_tool("outlook_mail_move_message", {"message_id": msg_id, "destination_folder": folder})
        except Exception as e:
            errors.append(str(e))
    if conv_key:
        remove_thread(conv_key)
    return jsonify({"ok": not errors})


@bp.route("/api/markread", methods=["POST"])
@bp.route("/api/mark_read", methods=["POST"])
def api_markread():
    payload = request.json or {}
    ids = payload.get("ids", [])
    conv_key = payload.get("conversationKey", "")
    read = payload.get("read", True)
    # If no ids provided, look up by conv_key
    if not ids and conv_key:
        db = get_db()
        rows = db.execute("SELECT id FROM emails WHERE conversation_key=?", (conv_key,)).fetchall()
        ids = [r["id"] for r in rows]
    if not ids:
        return jsonify({"ok": True})
    try:
        call_tool("outlook_mail_mark_read", {"message_ids": ids, "is_read": read})
        db = get_db()
        id_ph = ",".join("?" * len(ids))
        db.execute(f"UPDATE emails SET is_read=? WHERE id IN ({id_ph})", [1 if read else 0] + ids)
        db.execute("UPDATE threads SET has_unread=? WHERE conversation_key=?", (0 if read else 1, conv_key))
        db.commit()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@bp.route("/api/flag", methods=["POST"])
def api_flag():
    conv_key = request.json.get("conversationKey", "")
    flagged   = request.json.get("flagged", True)   # True = flag, False = unflag
    if not conv_key:
        return jsonify({"error": "missing conversationKey"}), 400
    db = get_db()
    db.execute("UPDATE threads SET is_flagged=? WHERE conversation_key=?",
               (1 if flagged else 0, conv_key))
    db.commit()
    return jsonify({"ok": True, "isFlagged": flagged})


@bp.route("/api/people")
def api_people():
    q = request.args.get("q", "").lower().strip()
    my = get_my_email().lower()
    db = get_db()
    rows = db.execute(
        "SELECT from_name, from_address FROM emails WHERE from_address != '' ORDER BY from_name"
    ).fetchall()
    seen = {}
    for r in rows:
        addr = (r["from_address"] or "").strip().lower()
        name = (r["from_name"] or "").strip()
        if not addr or addr == my:
            continue
        if addr not in seen:
            seen[addr] = name or r["from_address"]
    result = [{"name": name, "address": addr} for addr, name in seen.items()]
    if q:
        result = [p for p in result if q in p["address"].lower() or q in (p["name"] or "").lower()]
    result.sort(key=lambda p: (p["name"] or p["address"]).lower())
    return jsonify({"people": result[:60]})


@bp.route("/api/my_email")
def api_my_email():
    return jsonify({"email": get_my_email()})


@bp.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify({"results": [], "query": q, "count": 0})
    like = f"%{q}%"
    db = get_db()
    rows = db.execute("""
        SELECT id, subject, from_name, from_address, received_date_time,
               body_preview, folder, is_read, conversation_key
        FROM emails
        WHERE subject LIKE ? OR from_name LIKE ? OR from_address LIKE ? OR body_preview LIKE ?
        ORDER BY received_date_time DESC
        LIMIT 100
    """, (like, like, like, like)).fetchall()
    return jsonify({"results": [dict(r) for r in rows], "query": q, "count": len(rows)})


@bp.route("/api/mailbox/folders")
def api_mailbox_folders():
    top_level = json.loads(meta_get("folders_raw", "[]"))
    efforts_subs = json.loads(meta_get("efforts_subfolders", "[]"))
    db = get_db()
    # Get message counts per folder
    counts = {r["folder"]: r["cnt"] for r in db.execute(
        "SELECT folder, COUNT(*) as cnt FROM emails GROUP BY folder"
    ).fetchall()}
    def subfolder_count(path):
        return counts.get(path, 0)
    folder_map = {}
    for f in top_level:
        name = f.get("display_name") or f.get("displayName", "")
        if not name or name in _FOLDERS_SKIP_DISPLAY:
            continue
        if name.lower() == "efforts":
            children = []
            for s in sorted(efforts_subs):
                path = f"Efforts/{s}"
                children.append({"name": s, "path": path, "icon": "📂", "count": subfolder_count(path)})
            folder_map["Efforts"] = {"name": "Efforts", "icon": "📁", "count": 0, "children": children}
        else:
            folder_map[name] = {"name": name, "icon": _FOLDER_ICONS.get(name, "📁"), "count": counts.get(name, 0)}

    # Fixed display order
    ORDER = ["Inbox", "Efforts", "Partners", "Deleted Items", "Sent Items"]
    folders = []
    for name in ORDER:
        if name in folder_map:
            folders.append(folder_map.pop(name))
    # Append any remaining folders not in ORDER
    folders.extend(folder_map.values())
    return jsonify({"folders": folders})


@bp.route("/api/mailbox/folder")
def api_mailbox_folder():
    folder = request.args.get("folder", "").strip()
    if not folder:
        return jsonify({"threads": [], "total": 0, "folder": ""})
    offset = int(request.args.get("offset", 0))
    limit = 100
    db = get_db()
    # Latest email per conversation_key in this folder
    rows = db.execute("""
        SELECT e.id, e.subject, e.from_name, e.from_address,
               e.received_date_time, e.body_preview, e.is_read, e.conversation_key,
               g.cnt, g.unread
        FROM emails e
        JOIN (
            SELECT conversation_key,
                   MAX(received_date_time) AS latest,
                   COUNT(*) AS cnt,
                   SUM(CASE WHEN is_read=0 THEN 1 ELSE 0 END) AS unread
            FROM emails WHERE folder=?
            GROUP BY conversation_key
        ) g ON e.conversation_key = g.conversation_key
              AND e.received_date_time = g.latest
              AND e.folder = ?
        GROUP BY e.conversation_key
        ORDER BY e.received_date_time DESC
        LIMIT ? OFFSET ?
    """, (folder, folder, limit, offset)).fetchall()
    total = db.execute(
        "SELECT COUNT(DISTINCT conversation_key) FROM emails WHERE folder=?", (folder,)
    ).fetchone()[0]
    threads = [{
        "id": r["id"], "subject": r["subject"] or "(No subject)",
        "fromName": r["from_name"] or "", "fromAddress": r["from_address"] or "",
        "date": r["received_date_time"] or "", "preview": r["body_preview"] or "",
        "isRead": bool(r["is_read"]), "conversationKey": r["conversation_key"],
        "messageCount": r["cnt"], "unreadCount": r["unread"],
    } for r in rows]
    return jsonify({"threads": threads, "total": total, "folder": folder})


@bp.route("/api/folders")
def api_folders():
    efforts = json.loads(meta_get("efforts_subfolders", "[]"))
    other   = json.loads(meta_get("other_folders", "[]"))
    return jsonify({"folders": efforts + other, "effortsFolders": efforts})
