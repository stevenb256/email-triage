# Clanker Refactor Plan

## Backend Structure (`clanker/`)

```
clanker/
├── app.py              # Entry point: Flask app creation, startup
├── config.py           # Constants (ports, intervals, model names)
├── db.py               # Database layer (schema, get_db, meta_get/set)
├── mcp_client.py       # Outlook MCP wrapper (call_tool, session mgmt)
├── ai.py               # All Claude calls (analyze_thread, format_message, generate_reply)
├── sync.py             # Background sync thread + stale purge logic
└── routes/
    ├── triage.py       # /api/threads, /api/updates, /api/suggested_reply, /api/generate_reply, /api/reanalyze_all
    ├── mail.py         # /api/thread_messages, /api/format_message*, /api/reply, /api/send_new,
    │                   # /api/delete, /api/move, /api/markread, /api/flag, /api/resync_thread,
    │                   # /api/sync_now, /api/people, /api/my_email, /api/search, /api/mailbox/*
    └── calendar.py     # /api/calendar, /api/meeting_prep
```

## Frontend Structure (`clanker/static/`)

```
clanker/static/
├── js/
│   ├── main.js         # init(), state, schedulePoll/pollUpdates, switchTab, showToast, fmtDate utils
│   ├── titlebar.js     # triggerSync, resyncThread, reanalyzeAll — controls in <header>
│   ├── sidebar.js      # renderSidebar, initMailbox folder tree, updateWeekHours, today-cal-list
│   ├── triage.js       # openTriageSheet, renderTriageSheet, triageMark, executeAllActions,
│   │                   # _triageRowHTML, _triagePaneClick, triage keyboard nav (_triageKeydown)
│   ├── thread.js       # _renderThreadHdr, _renderMsgs, _msgCardHTML, _bodyContent,
│   │                   # toggleMsg, loadFormatted, toggleFormatView, _renderParas
│   ├── reply.js        # openReply, regenerateReply, sendReply, _renderRecipFields,
│   │                   # _renderTags, removeRecip, sendInlineReply, regenerateInlineReply
│   ├── compose.js      # openCompose, peopleSuggest, peopleSuggestKey, sendNewMessage,
│   │                   # _addComposeRecip, _renderComposeTags, focusComposeInput
│   ├── folder.js       # selectMailboxFolder, openMailboxThread, backToMailboxList,
│   │                   # mboxQuickReply, mboxQuickDelete, _mboxToggleRead, toggleFlag
│   ├── folder-nav.js   # _mboxKeydown, _mboxRegisterKeys, _mboxUnregisterKeys,
│   │                   # _mboxSetFocus, _mboxGetRows, _mboxFocusIdx
│   ├── search.js       # doSearch, clearSearch, openSearchResult
│   ├── calendar.js     # renderCalendar, calMove, calGoToday, calSetView, calGetWeekStart,
│   │                   # updateWeekHours (calendar-specific)
│   ├── actions.js      # openFile, openDelete, confirmDelete, fileThread, quickFile,
│   │                   # _act, _showActSpinner, _hideActSpinner
│   └── utils.js        # esc, encodeThread, decodeThread, avColor, initials, highlightEmails,
│                       # linkify, updateSyncStatus, _showNewBadge, updateCounts, AV_COLORS
├── css/
│   ├── base.css        # Reset, body, fonts, dark theme variables
│   ├── layout.css      # Header, sidebar, main pane, resize handle
│   ├── triage.css      # Triage sheet rows, topic groups, queue pills, KB hints
│   ├── thread.css      # Thread header, message cards, avatar, intent pills, body views
│   ├── modals.css      # Reply modal, compose modal, file modal, delete modal, toast
│   ├── folder.css      # Mailbox list rows, unread dot, hover actions
│   └── calendar.css    # Cal grid, day/week view, event tiles, time slots
└── html/
    ├── layout.html     # Shell: <header>, <nav class="sidebar">, main content area
    ├── triage.html     # #triage-pane structure
    ├── thread.html     # #thread-detail, #thread-hdr fragments
    ├── modals.html     # Reply, compose, file, delete modals
    └── calendar.html   # #calendar-pane structure
```

## Component Name Reference

| Name | What it is |
|------|-----------|
| **titlebar** | `<header>` — search, New Message, Sync Now, Resync, Re-analyze buttons |
| **sidebar** | Left nav — Triage Sheet btn, folder tree, calendar schedule, week hours, status bar |
| **triage** | Triage Sheet view — topic groups, thread rows, action queue, Execute All |
| **thread** | Thread detail — header (pills, subject, avatars, AI summary, action buttons) + message list |
| **folder** | Mailbox folder view — thread list with unread dots, hover actions |
| **folder-nav** | Keyboard navigation within folder/thread (j/k, r/d/f/u, Esc) |
| **thread-body** | Per-message card — AI/HTML/Original view toggle, intent pills, OWA link |
| **reply** | Reply modal — to/cc tags, body textarea, Regenerate button |
| **compose** | New Message modal — to/cc people picker, subject, body |
| **actions** | File modal, delete modal, spinner overlay, `_act()` executor |
| **search** | Search results pane |
| **calendar** | Calendar day/week view |
