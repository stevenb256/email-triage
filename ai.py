"""
ai.py — All Claude AI calls for Outlook Express email triage app.
"""
import json
import re

import anthropic

from config import ANALYSIS_MODEL, REPLY_MODEL, ANTHROPIC_API_KEY

# Lazy import to avoid circular dependency — only used in analyze_thread
def _call_tool(name, args):
    from mcp_client import call_tool
    return call_tool(name, args)

# ─── Topic sanitization ────────────────────────────────────────────────────────

def _normalize_topic(raw: str) -> str:
    """Sanitize a free-form topic label — trim whitespace, cap length, title-case."""
    t = re.sub(r'\s+', ' ', str(raw or "").strip())
    return t[:50] if t else "General"


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _clean(s, n=None) -> str:
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s or ''))
    return s[:n] if n else s


def _get_full_body(email: dict) -> str:
    """
    Return the best available body text for an email, fetching from Outlook if needed.
    Priority: formatted_body (cached AI parse) > full body from MCP > body_preview fallback.
    Strips HTML tags for clean plain text.
    """
    # 1. formatted_body is a JSON array of {text, intent, ...} — already plain text paragraphs
    fb = email.get("formatted_body")
    if fb:
        try:
            paras = json.loads(fb) if isinstance(fb, str) else fb
            if isinstance(paras, list) and paras:
                text = "\n\n".join(p.get("text", "") for p in paras if p.get("text"))
                if len(text) > 100:
                    return _clean(text, 8000)
        except Exception:
            pass

    # 2. body_html — strip tags for plain text
    bh = email.get("body_html", "")
    if bh and len(bh) > 200:
        plain = re.sub(r'<[^>]+>', ' ', bh)
        plain = re.sub(r'[ \t]{2,}', ' ', plain)
        plain = re.sub(r'\n{3,}', '\n\n', plain).strip()
        if len(plain) > 100:
            return _clean(plain, 8000)

    # 2.5 body_preview — if it's substantial, it's a previously cached full body (skip MCP)
    bp = email.get("body_preview", "")
    if bp and len(bp) > 300:
        return _clean(bp, 8000)

    # 3. Fetch full message from Outlook MCP — returns body_content (HTML string) directly
    msg_id = email.get("id", "")
    if msg_id:
        try:
            resp = _call_tool("outlook_mail_get_message", {"message_id": msg_id})
            raw = None
            if isinstance(resp, dict):
                raw = resp["messages"][0] if resp.get("messages") else resp

            if raw:
                # MCP returns body_content (HTML) as a top-level field
                content = raw.get("body_content") or ""
                if not content:
                    # also check nested body dict just in case
                    body = raw.get("body", {})
                    if isinstance(body, dict):
                        content = body.get("content", "")

                if content:
                    plain = re.sub(r'<[^>]+>', ' ', content)
                    plain = re.sub(r'[ \t]{2,}', ' ', plain)
                    plain = re.sub(r'\n{3,}', '\n\n', plain).strip()
                    if len(plain) > 50:
                        full = _clean(plain, 8000)
                        # Cache full body back to DB so re-analyze doesn't need MCP
                        try:
                            from db import get_db
                            _db = get_db()
                            _db.execute("UPDATE emails SET body_preview=? WHERE id=?", (full, msg_id))
                            _db.commit()
                        except Exception:
                            pass
                        return full
        except Exception as ex:
            print(f"  _get_full_body fetch failed for {msg_id[:20]}: {ex}")

    # 4. body_preview is all we have (short Outlook preview, ~255 chars)
    return _clean(email.get("body_preview", "(no content)"), 2000)


# ─── Anthropic client ──────────────────────────────────────────────────────────

_ai = None


def _get_ai():
    global _ai
    if _ai is None:
        _ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    return _ai


def _extract_fields_regex(json_str: str) -> dict:
    """Last-resort field extraction when JSON is unparseable."""
    def _grab(key):
        # Match "key": "value" with any content including newlines
        m = re.search(rf'"{key}"\s*:\s*"((?:[^"\\]|\\.)*)"', json_str, re.DOTALL)
        return m.group(1).replace('\\n', '\n').replace('\\"', '"') if m else ""
    return {
        "summary":       _grab("summary"),
        "topic":         _grab("topic"),
        "action":        _grab("action") or "read",
        "urgency":       _grab("urgency") or "low",
        "suggestedReply":_grab("suggestedReply"),
        "suggestedFolder":_grab("suggestedFolder"),
    }


# ─── AI functions ──────────────────────────────────────────────────────────────

def analyze_thread(emails: list, efforts_folders: list, other_folders: list, reply_context: str = "", existing_topics: list = None) -> dict:
    emails = sorted(emails, key=lambda e: e.get("received_date_time", ""))
    participants = list(dict.fromkeys(
        (_clean(e.get("from_name") or e.get("from_address", ""), 50)).strip()
        for e in emails
        if (e.get("from_name") or e.get("from_address"))
    ))[:8]

    context = emails[-8:]
    msg_parts = []
    for e in context:
        body = _get_full_body(e)
        msg_parts.append(
            f"From: {_clean(e.get('from_name') or e.get('from_address','Unknown'), 50)} "
            f"| {(e.get('received_date_time',''))[:10]}\n"
            f"{body}"
        )
    msgs_text = "\n\n---\n\n".join(msg_parts)

    if efforts_folders:
        efforts_list = ", ".join(efforts_folders[:15])
        folder_guidance = (
            f"\nEFFORTS SUBFOLDERS (pick one of these for filing — use the exact name): {efforts_list}"
        )
        if other_folders:
            folder_guidance += f"\nOther folders: {', '.join(other_folders[:6])}"
    else:
        folder_guidance = f"\nAvailable folders: {', '.join(other_folders[:10])}" if other_folders else ""

    subject = _clean(emails[-1].get("subject", "(no subject)"), 100)

    # Build topic guidance — existing topics let Claude reuse groups instead of fragmenting
    _topic_base = (
        " Use the product or project name only — strip out the specific activity or issue."
        " Good examples: 'Knowledge Anchors', 'ODSP', 'FaceAI', 'SPAI', 'NGIA', 'OneDrive',"
        " 'SharePoint Copilot', 'Lists', 'DocLab'. Bad (too specific): 'Knowledge Anchors Scenario"
        " Ownership', 'ODSP Experimentation Pipelines'. Bad (too generic): 'Engineering',"
        " 'FYI & Updates', 'Strategy'. Aim for the level where 3-6 related threads would share"
        " the same label. 1-3 words, title case."
    )
    if existing_topics:
        existing_list = ", ".join(f'"{t}"' for t in sorted(set(existing_topics))[:30])
        topic_guidance = (
            f"\nEXISTING TOPICS (reuse if this thread belongs to the same product/project): {existing_list}"
            f"\nPrefer reusing an existing topic over creating a new one.{_topic_base}"
        )
    else:
        topic_guidance = f"\nCreate a product/project-level topic.{_topic_base}"

    prompt = f"""You are a world-class executive communication assistant for a senior engineering/product leader at a large tech company.
Analyze this email thread and return ONLY valid JSON.

SUBJECT: {subject}
PARTICIPANTS: {', '.join(participants)}
TOTAL MESSAGES: {len(emails)}{folder_guidance}{topic_guidance}

MESSAGES (chronological, most recent last):
{msgs_text}

INSTRUCTIONS:

1. Determine the best action: reply | delete | file | read | done
   - reply: thread is waiting on the leader, a question was asked, action or decision required
   - delete: spam, automated notification with zero value, or social/marketing
   - file: substantive content worth keeping for reference
   - read: informational FYI, team-wide broadcast, no response needed
   - done: already fully resolved

2. Write "suggestedReply" — a complete, send-ready reply in FIRST PERSON.
   ALWAYS write a reply UNLESS action is "delete". Even for status updates and read-only threads,
   draft a reply that a leader would find useful to send.

   TONE: Direct, warm, confident. Like a senior leader who respects people's time and genuinely cares
   about the team. Never sycophantic, never hollow. No "I hope this finds you well", no "Thanks for sharing".

   TAILOR THE REPLY TO THE THREAD TYPE:

   A) STATUS UPDATES / PROGRESS REPORTS (action=read or file):
      — Acknowledge the specific work done. Name actual people. Reference specific metrics, milestones,
        or decisions mentioned in the thread.
      — Show you actually read it: reference a detail that proves it ("The latency drop from X to Y is great to see.")
      — Express genuine appreciation for the effort, not just the result.
      — If you spot anything worth a follow-up question or push, add it naturally.
      — Length: 2-4 sentences.

   B) QUESTIONS / REQUESTS WAITING ON YOU (action=reply):
      — Answer the question or fulfill the request directly and completely.
      — If you need more information first, ask exactly what you need — be specific.
      — State any decision you're making and the reason in one sentence.
      — Length: 3-5 sentences.

   C) VAGUE, UNCLEAR, OR MISSING-CONTEXT THREADS:
      — Do NOT pretend to understand. Push for the specific clarity needed.
      — Ask 1-3 sharp, targeted questions: What's the current state? What's the ask? What's the timeline?
        What decision needs to be made and by whom?
      — Be direct: "I want to engage on this but need a bit more context first…"
      — Length: 2-4 sentences.

   D) INCIDENT / OUTAGE / HIGH-URGENCY THREADS (urgency=high):
      — Acknowledge you're aware. State your immediate priority or what you're unblocking.
      — Offer your help or decision clearly: "I can free up [name] from X to help on this."
      — Ask the one most critical follow-up question if resolution is unclear.
      — Length: 2-4 sentences.

   E) CROSS-TEAM / PARTNERSHIP / EXTERNAL THREADS:
      — Professional but warm. Align on next steps. Name the right owner if it's not you.
      — Length: 3-5 sentences.

   BE SPECIFIC: Use actual names, numbers, and details from the thread. Generic replies are useless.
   Only use empty string "" for "suggestedReply" if action is "delete".

3. For "suggestedFolder": REQUIRED when action=file. Pick the single best name from the EFFORTS SUBFOLDERS
   list above using the exact name. Leave "" if action is not file.

Return ONLY this JSON (no markdown fences, no explanation):
{{
  "summary": "Write exactly 3 sections separated by the literal token ||BREAK||. No actual newlines inside this string. Section 1 — FACTS: 2-4 sentences, verifiable facts only. Name every person by full name. State exact numbers, dates, systems, metrics, decisions, blockers. Section 2 — OPEN QUESTIONS: 1-3 short bullet items (use • prefix), or 'None'. Section 3 — NEXT ACTION: 1 imperative sentence, e.g. 'Reply to Jane Bailey by EOD approving the $2.4M budget' or 'No action needed, file for reference'.",
  "topic": "product/project name only, 1-3 words — e.g. 'Knowledge Anchors', 'ODSP', 'FaceAI', 'SPAI'. Strip the specific activity. Reuse an existing topic if this thread belongs to the same product.",
  "action": "reply OR delete OR file OR read OR done",
  "urgency": "high OR medium OR low",
  "suggestedReply": "complete draft reply or empty string only if deleting",
  "suggestedFolder": "exact folder name or empty string"
}}"""

    if reply_context:
        prompt += f"\n\nNOTE: The user has provided the following context/notes for the reply. Incorporate this into your suggestedReply:\n{reply_context}"

    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r'^```[a-z]*\n?', '', raw)
        raw = re.sub(r'\n?```$', '', raw.strip())
        m = re.search(r'\{[\s\S]*\}', raw)
        if not m:
            raise ValueError(f"No JSON found: {raw[:200]}")
        json_str = m.group()
        try:
            result = json.loads(json_str)
        except json.JSONDecodeError:
            # Haiku sometimes emits literal newlines inside string values — fix and retry
            fixed = re.sub(r'\n', ' ', json_str)
            try:
                result = json.loads(fixed)
            except json.JSONDecodeError:
                # Last resort: extract each field with regex
                result = _extract_fields_regex(json_str)
        raw_topic = result.get("topic", "")
        result["topic"] = _normalize_topic(raw_topic)
        print(f"  [topic] raw={raw_topic!r} → {result['topic']!r}")
        # Sanitize suggestedFolder — reject placeholder-like values
        sf = result.get("suggestedFolder") or ""
        if re.match(r'^(select(\s+folder)?|choose|pick|none|n\/a|folder|tbd|unknown|\s*)$', sf.strip(), re.IGNORECASE):
            result["suggestedFolder"] = ""
        return result
    except Exception as ex:
        print(f"  Analysis error: {ex}")
        return {
            "summary": f"Could not analyze thread: {ex}",
            "topic": "FYI & Updates",
            "action": "read",
            "urgency": "low",
            "suggestedReply": "",
            "suggestedFolder": "",
        }


def _format_prompt(body: str, from_name: str, date: str) -> str:
    return f"""You are an expert email analyst helping a senior tech leader understand an email.

FROM: {_clean(from_name, 80)}  |  DATE: {date}
EMAIL BODY:
{_clean(body, 8000)}

Break this email into its natural paragraphs. For each paragraph:
1. Provide the exact paragraph text (verbatim)
2. Classify the intent from EXACTLY one of: Status Update | Request | Decision | Question | Action Item | Context | FYI | Warning | Introduction | Closing
3. Choose an appropriate emoji for that intent
4. Fact-check: if the paragraph makes a specific claim that seems incorrect or worth verifying, provide a short concern string (1-2 sentences). Otherwise use null.

Return ONLY valid JSON (no markdown fences):
{{"paragraphs":[{{"text":"...","intent":"...","emoji":"...","fact_concern":null}}]}}"""


def _parse_format_response(raw: str, body: str) -> list:
    raw = re.sub(r'^```[a-z]*\n?', '', raw.strip())
    raw = re.sub(r'\n?```$', '', raw.strip())
    m = re.search(r'\{[\s\S]*\}', raw)
    if not m:
        raise ValueError(f"No JSON: {raw[:100]}")
    result = json.loads(m.group())
    return result.get("paragraphs", [])


def format_message_ai(msg: dict) -> list:
    """Format a single message into AI-annotated paragraphs with intent + fact-check."""
    body = msg.get("body") or msg.get("body_preview") or ""
    if not body.strip():
        return [{"text": "(no content)", "intent": "FYI", "emoji": "📭", "fact_concern": None}]
    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    date = (msg.get("received_date_time") or "")[:10]
    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=3000,
            messages=[{"role": "user", "content": _format_prompt(body, from_name, date)}],
        )
        return _parse_format_response(resp.content[0].text, body)
    except Exception as ex:
        print(f"  Format error: {ex}")
        paras = [p.strip() for p in body.split('\n\n') if p.strip()][:20]
        return [{"text": p, "intent": "FYI", "emoji": "📄", "fact_concern": None} for p in paras]


def summarize_message_ai(msg: dict) -> str:
    """Return a short 1-sentence summary of the message body for triage view."""
    body = _get_full_body(msg)
    from_name = msg.get("from_name") or msg.get("from_address") or "Unknown"
    if not body.strip() or body == "(no content)":
        return ""
    prompt = (
        f"From: {from_name}\n\n{body[:3000]}\n\n"
        "Write a single concise sentence (max 20 words) summarising the key point or ask of this email. "
        "Be specific — mention names, numbers, or decisions if relevant. Reply with ONLY the sentence, no punctuation at the end."
    )
    try:
        resp = _get_ai().messages.create(
            model=ANALYSIS_MODEL,
            max_tokens=80,
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.content[0].text.strip().rstrip(".")
    except Exception:
        return ""


def generate_reply_ai(subject: str, msgs_text: str, user_prompt: str) -> str:
    """Generate a polished reply based on thread context and user's core message."""
    prompt = f"""You are helping a senior tech leader craft a professional email reply.

THREAD SUBJECT: {subject}
THREAD CONTEXT (oldest first, most recent last):
{msgs_text}

THE LEADER'S CORE MESSAGE (what they want to say — stay true to this):
"{user_prompt}"

Write a polished, professional reply that:
1. Leads with and stays grounded in the leader's core intent — this is non-negotiable
2. Uses specific names, decisions, and details from the thread to make it feel personal and grounded
3. Is warm but direct — no filler phrases, no corporate speak, no "I hope this finds you well"
4. Uses 1-3 emojis placed naturally (not forced) to add energy and approachability
5. Has clear paragraph breaks for readability
6. Ends with clear next steps, a question, or a crisp closing — whichever fits
7. Length: match the complexity. Simple acknowledgement = 2-3 sentences. Complex topic = 4-7 sentences.

Return ONLY the reply body text. No subject line, no "From:", no markdown fences."""

    resp = _get_ai().messages.create(
        model=REPLY_MODEL,
        max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text.strip()


def summarize_thread_ai(subject: str, attendees: str, time_str: str, location: str) -> dict:
    """Generate meeting prep (headsup + topics) for a calendar event."""
    prompt = (
        f"You are preparing a senior tech leader at Microsoft for an upcoming meeting.\n\n"
        f"Meeting: {subject}\n"
        f"Time: {time_str}\n"
        f"Location: {location or 'Not specified'}\n"
        f"Attendees: {attendees}\n\n"
        f"Provide:\n"
        f"1. A 1-2 sentence heads-up: what this meeting is likely about and what the leader should be ready for.\n"
        f"2. Exactly 3 concise, specific topics or questions worth raising or keeping in mind.\n\n"
        f'Respond ONLY with valid JSON: {{"headsup": "...", "topics": ["...", "...", "..."]}}'
    )
    resp = _get_ai().messages.create(
        model=ANALYSIS_MODEL,
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}],
    )
    text = resp.content[0].text.strip()
    import re as _re
    m = _re.search(r'\{.*\}', text, _re.DOTALL)
    if m:
        result = json.loads(m.group())
        return {"ok": True, "headsup": result.get("headsup", ""), "topics": result.get("topics", [])}
    return {"ok": False, "headsup": "", "topics": []}
