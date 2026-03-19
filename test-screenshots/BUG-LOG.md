# Outlook Express UX Bug Log
## Found via Playwright MCP testing — 2026-03-18

---

### BUG-1: [CRITICAL] Reply Cancel from Triage leaves user stuck on thread detail
- **Severity**: Critical
- **Status**: FIXED
- **Steps**: Triage view → Click Reply on any thread → Reply dialog opens → Click Cancel
- **Expected**: Return to triage sheet view
- **Actual**: Reply closes but user is stuck on thread detail view with "Select a folder" text visible in the footer. Escape key does NOT navigate back to triage either.
- **Fix**: Added `fromTriage` flag to `_replyState`. `triageOpenReply()` sets it before opening reply. `closeModals()` checks the flag and calls `openTriageSheet()` to restore triage view. Flag cleared on send to avoid restoring triage after successful send.
- **Files changed**: `static/js/reply.js`, `static/js/triage.js`, `static/js/actions.js`

### BUG-2: [HIGH] HTML entities leak into message preview text
- **Severity**: High
- **Status**: FIXED
- **Steps**: Open any thread with HTML email content (inbox or triage)
- **Expected**: Preview text shows clean readable text
- **Actual**: Raw HTML entities like `&nbsp;`, `&quot;`, `&lt;` appear in message preview snippets throughout the app (inbox thread list, triage summaries, thread detail message rows, search results)
- **Fix**: Added `decodeEntities()` utility function that uses a textarea element to decode HTML entities. Applied to all body_preview display sites: thread.js (message cards, fallback text), triage.js (message body), sidebar.js (mailbox rows), search.js (search results).
- **Files changed**: `static/js/utils.js`, `static/js/thread.js`, `static/js/triage.js`, `static/js/sidebar.js`, `static/js/search.js`

### BUG-3: [MEDIUM] Escape key does not close compose modal
- **Severity**: Medium
- **Status**: FIXED
- **Steps**: Click "New Message" → Compose modal opens → Press Escape
- **Expected**: Compose modal closes
- **Actual**: Nothing happens. Must click Cancel button to close.
- **Fix**: Added global `keydown` listener for Escape in `main.js` that checks for any open `.modal-overlay.open` and calls `closeModals()`. Skips if a people dropdown is open (handled separately in compose.js).
- **Files changed**: `static/js/main.js`

### BUG-4: [MEDIUM] Double "Re:" prefix in reply subject line
- **Severity**: Medium
- **Status**: FIXED
- **Steps**: Open a thread whose subject already starts with "Re:" → Click Reply
- **Expected**: Reply subject: "Re: Riffing with Zach about plan mode..."
- **Actual**: Reply subject: "Re: Re: Riffing with Zach about plan mode..."
- **Fix**: Added regex check `/^re:\s/i.test(subj)` before prepending "Re:" — only adds prefix if not already present (case-insensitive).
- **Files changed**: `static/js/reply.js`

### BUG-5: [LOW] Missing favicon.ico (404)
- **Severity**: Low
- **Status**: FIXED
- **Steps**: Load any page
- **Expected**: No console errors
- **Actual**: Console error: "Failed to load resource: 404 (NOT FOUND) @ /favicon.ico"
- **Fix**: Added inline SVG favicon (`data:image/svg+xml` with email emoji) to `<head>` in index.html.
- **Files changed**: `templates/index.html`

### BUG-6: [MEDIUM] Escape key does not go back from thread detail to folder list
- **Severity**: Medium
- **Status**: FIXED (by BUG-1 + BUG-3 fixes)
- **Steps**: In triage context, Escape does nothing after reply cancel. The keyboard shortcut hints show "Esc back" but it doesn't work in post-reply-cancel state.
- **Fix**: Now Escape closes the reply modal AND restores triage view (BUG-1 fix). The global Escape handler (BUG-3 fix) also ensures modals close before keyboard nav handlers fire.

### BUG-7: [LOW] Reply Cancel triggers re-analysis / re-sync of thread
- **Severity**: Low
- **Status**: OPEN — low priority, cosmetic only
- **Steps**: Click Reply on triage thread → Cancel
- **Actual**: Thread gets re-analyzed ("Analyzing 1/1...") because `triageOpenReply` calls `selectThread` which triggers a resync. Not harmful but causes unnecessary API calls and visual flicker.

---

## Working Features (verified via Playwright)
- Splash screen loads correctly with ASCII art
- Sidebar navigation (all folders load)
- Efforts subfolder expansion (26 subfolders visible, clickable)
- Inbox thread list with pagination
- Thread detail view from inbox (with Close button)
- Triage sheet with topic grouping and AI summaries
- Reply dialog opens with AI-suggested reply and correct recipients
- Compose modal opens with TO/CC/Subject/Body fields
- Cancel button works on compose and reply modals
- Escape key closes all modals (compose, reply, file)
- Reply Cancel from triage restores triage view
- Search returns results (tested "budget" → 100 results)
- Calendar day view with meetings and time slots
- Calendar week view with multi-day events
- Calendar day navigation (prev/next arrows)
- Top Collaborators panel
- Calendar sidebar widget
- Sent Items, Deleted Items, Archive folders all load
- Thread detail shows email body in iframe with correct rendering
- Message previews display clean text without HTML entities
- Favicon loads without 404 error
