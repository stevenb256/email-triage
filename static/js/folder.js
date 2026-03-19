// ── Mailbox thread view ─────────────────────────────────────────────────────────
let _mboxLoadSeq = 0;

async function openMailboxThread(convKey, folder) {
  const seq = ++_mboxLoadSeq;
  document.querySelectorAll('.mbox-row.active').forEach(e=>e.classList.remove('active'));
  try { document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`)?.classList.add('active'); } catch(e){}
  state.selectedKey = convKey;
  state.mailboxContext = true;
  state.expandedMsgs = new Set();
  state.currentMsgs = [];
  const hdrEl = document.getElementById('thread-hdr');
  const msgsEl = document.getElementById('msgs-section');
  // Hide everything else, show thread-detail
  ['empty-pane','triage-pane','mailbox-pane','search-pane','calendar-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  document.getElementById('thread-detail').style.display = 'flex';
  document.getElementById('thread-detail').dataset.loaded = '1';
  hdrEl.innerHTML = '<div style="padding:20px"><div class="spinner"></div></div>';
  msgsEl.innerHTML = '';
  const r = await fetch(`/api/thread_messages?conversationKey=${encodeURIComponent(convKey)}`).then(r=>r.json()).catch(()=>null);
  // Stale response — another thread was clicked while this was loading
  if (seq !== _mboxLoadSeq) return;
  if (!r) {
    hdrEl.innerHTML = `<div style="padding:14px 22px 10px;display:flex;justify-content:space-between;align-items:center">
      <button class="mbox-back" onclick="backToMailboxList()">✕ Close</button>
      <span style="color:#8b949e">Error loading messages</span></div>`;
    return;
  }
  state.currentMsgs = r.messages || [];
  // newest-first; latest for latestId is index 0
  const latestMsg = state.currentMsgs[0] || {};
  const thread = state.threadMap[convKey];
  if (thread) {
    _renderThreadHdr(thread);
  } else {
    // Build a minimal thread object so we always show action buttons
    const syntheticThread = {
      conversationKey: convKey,
      subject: latestMsg.subject || '(No subject)',
      latestId: latestMsg.id || '',
      emailIds: state.currentMsgs.map(m=>m.id),
      messageCount: state.currentMsgs.length,
      participants: [...new Set(state.currentMsgs.map(m=>m.from_name||m.from_address).filter(Boolean))],
      urgency: 'low', action: 'read',
      summary: '', suggestedReply: '', suggestedFolder: folder,
      hasUnread: state.currentMsgs.some(m=>!m.is_read),
    };
    state.threadMap[convKey] = syntheticThread;
    _renderThreadHdr(syntheticThread);
    // Trigger analysis in background so summary populates
    fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: convKey})}).then(r=>r.json()).then(d=>{
      if (d && d.reply) {
        const t = state.threadMap[convKey];
        if (t) { t.suggestedReply = d.reply; }
      }
    }).catch(()=>{});
  }
  const msgs = state.currentMsgs;
  msgsEl.innerHTML = msgs.map((m,i)=>_msgCardHTML(m,i)).join('');
}

function backToMailboxList() {
  state.mailboxContext = false;
  document.getElementById('thread-detail').style.display = 'none';
  document.getElementById('thread-detail').dataset.loaded = '';
  document.getElementById('mailbox-pane').style.display = 'flex';
  // Restore keyboard focus to the row that was open
  if (state.selectedKey) {
    const rows = _mboxGetRows();
    const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
    _mboxSetFocus(idx >= 0 ? idx : 0);
  }
  _mboxRegisterKeys();
}

async function _mboxToggleRead(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  const markRead = thread.hasUnread;
  await fetch('/api/mark_read', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({conversationKey: convKey, read: markRead})}).catch(()=>{});
  thread.hasUnread = !markRead;
  if (markRead) thread.isRead = true;
  // Update unread dot on the row
  const row = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  if (row) {
    row.classList.toggle('unread', !markRead);
    const dot = row.querySelector('.mbox-dot, .mbox-dot-empty');
    if (dot) { dot.className = !markRead ? 'mbox-dot' : 'mbox-dot-empty'; }
  }
}

async function mboxQuickReply(convKey, folder) {
  await openMailboxThread(convKey, folder);
  const thread = state.threadMap[convKey];
  if (thread) openReply(encodeThread(thread));
}

async function mboxQuickDelete(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  // Find next row before removing
  const row = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  const nextRow = row?.nextElementSibling;
  // Animate out
  if (row) { row.style.opacity='0'; row.style.transition='opacity .15s'; }
  _showActSpinner('Deleting…');
  await fetch('/api/delete', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({ids: thread.emailIds, conversationKey: convKey})});
  _hideActSpinner();
  if (row) row.remove();
  delete state.threadMap[convKey];
  for (const g of state.groups) g.threads = g.threads.filter(t=>t.conversationKey!==convKey);
  state.groups = state.groups.filter(g=>g.threads.length>0);
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
  // Update folder count
  const countEl = document.getElementById('mailbox-folder-count');
  if (countEl) {
    const cur = parseInt(countEl.textContent) || 0;
    countEl.textContent = `${Math.max(0,cur-1)} threads`;
  }
  // If we were viewing this thread, open the next one or go back
  if (state.selectedKey === convKey) {
    if (nextRow && nextRow.dataset.key) {
      openMailboxThread(nextRow.dataset.key, nextRow.dataset.folder || mailboxCurrentFolder);
    } else {
      backToMailboxList();
    }
  }
}
