// ── Triage sheet ───────────────────────────────────────────────────────────────
function openTriageSheet() {
  state.triageView = true;
  state.triageFocusIdx = -1;
  state.selectedKey = null;
  document.getElementById('empty-pane').style.display='none';
  document.getElementById('thread-detail').style.display='none';
  const pane = document.getElementById('triage-pane');
  pane.style.display='flex';
  // Attach delegated click handler once
  if (!pane._triageDelegated) {
    pane._triageDelegated = true;
    pane.addEventListener('click', _triagePaneClick);
  }
  renderSidebar();
  renderTriageSheet();
  initMailbox(); // ensure folder tree is populated
  document.addEventListener('keydown', _triageKeydown);
}

function _triagePaneClick(e) {
  // Triage action buttons
  const btn = e.target.closest('[data-triage-action]');
  if (btn) {
    e.stopPropagation();
    const row = btn.closest('[data-convkey]');
    if (!row) return;
    const convKey = row.dataset.convkey;
    const action  = btn.dataset.triageAction;
    if      (action==='reply')  triageOpenReply(convKey);
    else if (action==='file')   triageMark(convKey,'file');
    else if (action==='delete') triageMark(convKey,'delete');
    else if (action==='clear')  triageMark(convKey,null);
    return;
  }
  // Row expand/collapse (summary area)
  const summary = e.target.closest('[data-triage-expand]');
  if (summary) {
    const row = summary.closest('[data-convkey]');
    if (row) triageToggleExpand(row.dataset.convkey);
    return;
  }
  // Inline message row toggle
  const msgRow = e.target.closest('[data-triage-msg]');
  if (msgRow) {
    const row = msgRow.closest('[data-convkey]');
    if (row) triageToggleMsgBody(row.dataset.convkey, parseInt(msgRow.dataset.triageMsg));
    return;
  }
  // Topic header collapse/expand
  const topicHdr = e.target.closest('[data-triage-topic]');
  if (topicHdr) {
    toggleTriageTopic(topicHdr.dataset.topic);
  }
}

function _triageRowHTML(t) {
  const convKey = t.conversationKey;
  const action = state.triageActions[convKey];
  const actionCls = action ? ' ts-'+action.type : '';
  const expanded = state.expandedTriageRows.has(convKey);
  const urgCls = {high:'urg-high',medium:'urg-medium',low:'urg-low'}[t.urgency]||'urg-low';
  const rec = ACTION_REC[t.action] || ACTION_REC.read;
  const statusLbl = action ? (action.type==='delete'?'🗑 Queued':action.type==='file'?'📁 Queued':'') : '';
  const msgsHtml = expanded ? _triageMsgsHTML(convKey) : '';
  return `<div class="triage-row${actionCls}${expanded?' expanded':''}" id="triage-row-${esc(convKey)}" data-convkey="${esc(convKey)}">
    <div class="triage-row-summary" data-triage-expand="1">
      <div class="triage-row-top">
        <span class="triage-row-expand-chevron">▶</span>
        <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
        <span class="triage-subj">${esc(t.subject||'(No subject)')}</span>
        <span class="action-rec ${rec.cls}">${rec.icon} ${rec.label}</span>
      </div>
      ${t.summary?`<div class="triage-sum" style="padding-left:17px">${_renderSummary(t.summary)}</div>`:''}
    </div>
    <div class="triage-msgs" id="triage-msgs-${esc(convKey)}" style="${expanded?'':'display:none'}">${msgsHtml}</div>
    <div class="triage-btns">
      <button class="btn btn-reply btn-sm" data-triage-action="reply">↩ Reply</button>
      <button class="btn btn-ghost btn-sm btn-ts-file${action&&action.type==='file'?' active':''}" data-triage-action="file">📁 File</button>
      <button class="btn btn-ghost btn-sm btn-ts-del${action&&action.type==='delete'?' active':''}" data-triage-action="delete">🗑 Delete</button>
      ${action?`<button class="btn btn-ghost btn-sm" data-triage-action="clear">✕</button>`:''}
      <span class="triage-qlbl">${esc(statusLbl)}</span>
    </div>
  </div>`;
}

function _triageMsgsHTML(convKey) {
  const msgs = state.triageMsgCache[convKey];
  if (!msgs) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block"></div> Loading…</div>';
  if (!msgs.length) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf">No messages</div>';
  return [...msgs].reverse().map((m,i) => {
    const from = m.from_name||m.from_address||'?';
    const preview = String(m.body||m.body_preview||'').replace(/\s+/g,' ').trim().slice(0,80);
    const date = fmtDate((m.received_date_time||'').slice(0,19));
    return `<div class="triage-msg-row" id="tmr-${esc(convKey)}-${i}" data-triage-msg="${i}">
      <span class="triage-msg-chev">▶</span>
      <span class="triage-msg-from">${esc(from)}</span>
      <span class="triage-msg-prev">${esc(preview)}</span>
      <span class="triage-msg-date">${esc(date)}</span>
    </div>
    <div class="triage-msg-body" id="tmb-${esc(convKey)}-${i}" style="display:none">${esc(String(m.body||m.body_preview||'').trim())}</div>`;
  }).join('');
}

async function triageToggleExpand(convKey) {
  const row = document.getElementById('triage-row-'+convKey);
  const msgsEl = document.getElementById('triage-msgs-'+convKey);
  if (!row || !msgsEl) return;
  const expanding = !state.expandedTriageRows.has(convKey);
  if (expanding) {
    state.expandedTriageRows.add(convKey);
    row.classList.add('expanded');
    msgsEl.style.display = '';
    if (!state.triageMsgCache[convKey]) {
      msgsEl.innerHTML = '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block;margin-right:6px"></div>Loading…</div>';
      const r = await fetch(`/api/thread_messages?conversationKey=${encodeURIComponent(convKey)}`).then(r=>r.json()).catch(()=>null);
      const msgs = (r&&r.messages||[]).slice().sort((a,b)=>(b.received_date_time||'')>(a.received_date_time||'')?1:-1);
      state.triageMsgCache[convKey] = msgs;
      msgsEl.innerHTML = _triageMsgsHTML(convKey);
    }
  } else {
    state.expandedTriageRows.delete(convKey);
    row.classList.remove('expanded');
    msgsEl.style.display = 'none';
  }
}

function triageToggleMsgBody(convKey, idx) {
  const row = document.getElementById(`tmr-${convKey}-${idx}`);
  const body = document.getElementById(`tmb-${convKey}-${idx}`);
  if (!body) return;
  const open = body.style.display !== 'none';
  body.style.display = open ? 'none' : '';
  if (row) row.classList.toggle('open', !open);
}

function triageOpenReply(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  closeTriageSheet();
  // Navigate to thread then open reply modal
  selectThread(convKey);
  setTimeout(() => openReply(encodeThread(thread)), 500);
}

function toggleTriageTopic(topic) {
  if (state.collapsedTriageTopics.has(topic)) state.collapsedTriageTopics.delete(topic);
  else state.collapsedTriageTopics.add(topic);
  const grp = document.getElementById('ttg-'+btoa(unescape(encodeURIComponent(topic))).replace(/[^a-zA-Z0-9]/g,''));
  if (!grp) { renderTriageSheet(); return; }
  const rows = grp.querySelector('.triage-topic-rows');
  const hdr  = grp.querySelector('.triage-topic-hdr');
  const collapsed = state.collapsedTriageTopics.has(topic);
  if (rows) rows.style.display = collapsed ? 'none' : '';
  if (hdr)  hdr.classList.toggle('open', !collapsed);
  _triageUpdateFocus();
}

function renderTriageSheet() {
  const pane = document.getElementById('triage-pane');
  const queuedCount = Object.keys(state.triageActions).length;
  // Sort threads within each group newest-first, then sort groups by their latest thread
  const sortedGroups = state.groups.map(g => ({
    ...g,
    threads: [...g.threads].sort((a,b) => (b.latestReceived||'').localeCompare(a.latestReceived||''))
  })).sort((a,b) => {
    const aLat = a.threads[0]?.latestReceived || '';
    const bLat = b.threads[0]?.latestReceived || '';
    return bLat.localeCompare(aLat);
  });
  const groupsHtml = sortedGroups.map(g => {
    const topic = g.topic || 'Uncategorized';
    const safeId = 'ttg-'+btoa(unescape(encodeURIComponent(topic))).replace(/[^a-zA-Z0-9]/g,'');
    const collapsed = state.collapsedTriageTopics.has(topic);
    return `<div class="triage-topic-group" id="${safeId}">
      <div class="triage-topic-hdr${collapsed?'':' open'}" data-topic="${esc(topic)}" data-triage-topic="1">
        <span class="triage-topic-chevron">▶</span>
        <span class="triage-topic-label">${esc(topic)}</span>
        <span class="triage-topic-badge">${g.threads.length}</span>
      </div>
      <div class="triage-topic-rows" style="${collapsed?'display:none':''}">
        ${g.threads.map(_triageRowHTML).join('')}
      </div>
    </div>`;
  }).join('');
  pane.innerHTML = `<div class="triage-hdr">
    <span class="triage-title">📋 Triage Sheet</span>
    <span class="triage-queue-count" id="triage-queue-count">${queuedCount} queued</span>
    <button class="btn btn-reply btn-sm" id="triage-execute-btn" onclick="executeAllActions()"${queuedCount===0?' disabled':''}>⚡ Execute All</button>
  </div>
  <div class="triage-rows">${groupsHtml}</div>
  <div class="mbox-kb-hint">
    <span><kbd>j</kbd><kbd>k</kbd> navigate</span>
    <span><kbd>Enter</kbd> expand</span>
    <span><kbd>r</kbd> reply</span>
    <span><kbd>d</kbd> delete</span>
    <span><kbd>f</kbd> file</span>
    <span><kbd>x</kbd> clear</span>
    <span><kbd>Esc</kbd> back</span>
  </div>`;
}

// ── Triage keyboard navigation ──────────────────────────────────────────────
function _triageNavList() {
  const list = [];
  for (const g of state.groups) {
    const topic = g.topic || 'Uncategorized';
    list.push({type:'topic', topic});
    if (!state.collapsedTriageTopics.has(topic)) {
      for (const t of g.threads) list.push({type:'thread', convKey:t.conversationKey});
    }
  }
  return list;
}

function _triageUpdateFocus() {
  document.querySelectorAll('.triage-kb-focus').forEach(el=>el.classList.remove('triage-kb-focus'));
  const nav = _triageNavList();
  const item = nav[state.triageFocusIdx];
  if (!item) return;
  let el;
  if (item.type==='topic') {
    el = document.querySelector(`.triage-topic-hdr[data-topic="${CSS.escape(item.topic)}"]`);
  } else {
    el = document.getElementById('triage-row-'+item.convKey);
  }
  if (el) { el.classList.add('triage-kb-focus'); el.scrollIntoView({block:'nearest',behavior:'smooth'}); }
}

function _triageKeydown(e) {
  if (!state.triageView) return;
  // Don't intercept if user is typing in an input
  if (e.target.tagName==='INPUT'||e.target.tagName==='TEXTAREA') return;
  const nav = _triageNavList();
  let idx = state.triageFocusIdx;

  if (e.key==='ArrowDown') {
    e.preventDefault();
    idx = idx < nav.length-1 ? idx+1 : idx;
  } else if (e.key==='ArrowUp') {
    e.preventDefault();
    idx = idx > 0 ? idx-1 : 0;
  } else if (e.key==='ArrowRight'||e.key==='ArrowLeft') {
    e.preventDefault();
    const item = nav[idx];
    if (item&&item.type==='topic') {
      if (e.key==='ArrowRight') state.collapsedTriageTopics.delete(item.topic);
      else state.collapsedTriageTopics.add(item.topic);
      toggleTriageTopic(item.topic);
    } else if (item&&item.type==='thread'&&e.key==='ArrowRight') {
      triageToggleExpand(item.convKey);
    } else if (item&&item.type==='thread'&&e.key==='ArrowLeft') {
      state.expandedTriageRows.delete(item.convKey);
      const row=document.getElementById('triage-row-'+item.convKey);
      const msgsEl=document.getElementById('triage-msgs-'+item.convKey);
      if(row){row.classList.remove('expanded');}
      if(msgsEl){msgsEl.style.display='none';}
    }
    state.triageFocusIdx = idx;
    _triageUpdateFocus(); return;
  } else if (e.key==='Enter'||e.key===' ') {
    e.preventDefault();
    const item = nav[idx];
    if (!item) { idx=0; }
    else if (item.type==='topic') toggleTriageTopic(item.topic);
    else triageToggleExpand(item.convKey);
  } else if (e.key==='r'||e.key==='R') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageOpenReply(item.convKey); }
    return;
  } else if (e.key==='d'||e.key==='D') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, state.triageActions[item.convKey]?.type==='delete'?null:'delete'); }
    return;
  } else if (e.key==='f'||e.key==='F') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, state.triageActions[item.convKey]?.type==='file'?null:'file'); }
    return;
  } else if (e.key==='x'||e.key==='X') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageMark(item.convKey, null); }
    return;
  } else if (e.key==='Escape') {
    e.preventDefault(); closeTriageSheet(); return;
  } else return;

  state.triageFocusIdx = idx;
  _triageUpdateFocus();
}

function closeTriageSheet() {
  state.triageView = false;
  state.triageFocusIdx = -1;
  document.removeEventListener('keydown', _triageKeydown);
  switchTab('mailbox');
}

function triageMark(convKey, type) {
  if (type === null) delete state.triageActions[convKey];
  else state.triageActions[convKey] = {type};
  // Update row visual
  const row = document.getElementById('triage-row-'+convKey);
  if (row) {
    const expanded = state.expandedTriageRows.has(convKey);
    row.className = 'triage-row' + (type?' ts-'+type:'') + (expanded?' expanded':'');
    const qlbl = row.querySelector('.triage-qlbl');
    if (qlbl) qlbl.textContent = type==='delete'?'🗑 Queued':type==='file'?'📁 Queued':'';
    row.querySelectorAll('.btn-ts-del,.btn-ts-file').forEach(b=>b.classList.remove('active'));
    if (type==='delete'){const b=row.querySelector('.btn-ts-del');if(b)b.classList.add('active');}
    else if (type==='file'){const b=row.querySelector('.btn-ts-file');if(b)b.classList.add('active');}
    // Rebuild clear button
    const btns = row.querySelector('.triage-btns');
    if (btns) {
      let clr = btns.querySelector('.btn-ts-clr');
      if (type && !clr) {
        const b=document.createElement('button');b.className='btn btn-ghost btn-sm btn-ts-clr';
        b.textContent='✕';b.onclick=()=>triageMark(convKey,null);
        btns.insertBefore(b, btns.querySelector('.triage-qlbl'));
      } else if (!type && clr) clr.remove();
    }
  }
  const queuedCount = Object.keys(state.triageActions).length;
  const countEl = document.getElementById('triage-queue-count');
  if (countEl) countEl.textContent = queuedCount+' queued';
  const execBtn = document.getElementById('triage-execute-btn');
  if (execBtn) execBtn.disabled = queuedCount === 0;
}

async function executeAllActions() {
  const entries = Object.entries(state.triageActions);
  if (!entries.length) return;
  const execBtn = document.getElementById('triage-execute-btn');
  let done = 0;
  const total = entries.length;
  for (const [convKey, action] of entries) {
    if (execBtn) execBtn.textContent = `Executing ${done+1}/${total}...`;
    const thread = state.threadMap[convKey];
    if (!thread) { done++; continue; }
    try {
      if (action.type === 'send') {
        // Open reply modal for this thread so user can compose
        closeTriageSheet();
        selectThread(convKey);
        const enc = encodeThread(thread);
        setTimeout(()=>openReply(enc), 400);
        break; // handle one reply at a time
      } else if (action.type === 'delete') {
        await fetch('/api/delete', {
          method: 'POST',
          headers: {'Content-Type':'application/json'},
          body: JSON.stringify({ids: thread.emailIds, conversationKey: convKey})
        });
      } else if (action.type === 'file') {
        await fetch('/api/move', {
          method: 'POST',
          headers: {'Content-Type':'application/json'},
          body: JSON.stringify({ids: thread.emailIds, folder: thread.suggestedFolder||'', conversationKey: convKey})
        });
      }
    } catch(e) {
      console.error('Execute action error for '+convKey, e);
    }
    // Mark row as done
    const row = document.getElementById('triage-row-'+convKey);
    if (row) {
      row.className = 'triage-row ts-done';
      const qlbl = row.querySelector('.triage-qlbl');
      if (qlbl) qlbl.textContent = '✓ Done';
    }
    // Remove from state
    delete state.triageActions[convKey];
    delete state.threadMap[convKey];
    for (const g of state.groups) g.threads = g.threads.filter(t=>t.conversationKey!==convKey);
    state.groups = state.groups.filter(g=>g.threads.length>0);
    done++;
  }
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
  // Re-render triage sheet so completed items are removed
  renderTriageSheet();
  const execBtn2 = document.getElementById('triage-execute-btn');
  if (execBtn2) { execBtn2.textContent = `✓ ${done} action${done!==1?'s':''} done`; execBtn2.disabled = true; }
}
