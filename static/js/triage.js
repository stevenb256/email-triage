// ── Triage sheet ───────────────────────────────────────────────────────────────
function openTriageSheet() {
  state.triageView = true;
  state.triageFocusIdx = -1;
  state.selectedKey = null;
  document.getElementById('empty-pane').style.display='none';
  document.getElementById('thread-detail').style.display='none';
  document.querySelectorAll('.folder-item.active').forEach(e=>e.classList.remove('active'));
  const navTriage = document.getElementById('nav-triage'); if (navTriage) navTriage.classList.add('active');
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
    if      (action==='reply')        triageOpenReply(convKey, 'all');
    else if (action==='reply-sender') triageOpenReply(convKey, 'sender');
    else if (action==='forward')      triageOpenForward(convKey);
    else if (action==='file')         triageFileNow(convKey);
    else if (action==='delete')       triageDeleteNow(convKey);
    else if (action==='view')         triageViewMessage(convKey);
    return;
  }
  // Row expand/collapse (summary area)
  const summary = e.target.closest('[data-triage-expand]');
  if (summary) {
    const row = summary.closest('[data-convkey]');
    if (row) triageToggleExpand(row.dataset.convkey);
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
  const expanded = state.expandedTriageRows.has(convKey);
  const urgCls = {high:'urg-high',medium:'urg-medium',low:'urg-low'}[t.urgency]||'urg-low';
  const rec = ACTION_REC[t.action] || ACTION_REC.read;
  const folder = t.suggestedFolder || '';
  const fileLbl = folder ? `📁 ${folder}` : '📁 File';
  const msgsHtml = expanded ? _triageMsgsHTML(convKey) : '';
  // Recipients row (populated async)
  const recipId = `triage-recips-${esc(convKey)}`;
  return `<div class="triage-row${expanded?' expanded':''}" id="triage-row-${esc(convKey)}" data-convkey="${esc(convKey)}">
    <div class="triage-row-summary" data-triage-expand="1">
      <div class="triage-row-top">
        <span class="triage-row-expand-chevron">▶</span>
        <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
        <span class="triage-subj">${esc(t.subject||'(No subject)')}</span>
        <span class="action-rec ${rec.cls}">${rec.icon} ${rec.label}</span>
        <button class="btn btn-ghost btn-xs triage-view-btn" data-triage-action="view" title="View full message">👁 View</button>
      </div>
      <div class="triage-row-recips" id="${recipId}"></div>
      ${t.summary?`<div class="triage-sum" style="padding-left:17px">${_renderSummary(t.summary)}</div>`:''}
    </div>
    <div class="triage-msgs" id="triage-msgs-${esc(convKey)}" style="${expanded?'':'display:none'}">${msgsHtml}</div>
    <div class="triage-btns">
      <button class="btn btn-reply btn-sm" data-triage-action="reply">↩ Reply All</button>
      <button class="btn btn-ghost btn-sm" data-triage-action="reply-sender">↩ Reply</button>
      <button class="btn btn-ghost btn-sm" data-triage-action="forward">⤳ Fwd</button>
      <button class="btn btn-ghost btn-sm btn-ts-file" data-triage-action="file">${esc(fileLbl)}</button>
      <button class="btn btn-ghost btn-sm btn-ts-del" data-triage-action="delete">🗑 Delete</button>
    </div>
  </div>`;
}

function _loadTriageRecips(convKey) {
  const t = state.threadMap[convKey];
  if (!t || !t.latestId) return;
  const el = document.getElementById('triage-recips-'+convKey);
  if (!el || el.dataset.loaded) return;
  el.dataset.loaded = '1';
  fetch(`/api/message_recipients?id=${encodeURIComponent(t.latestId)}`).then(r=>r.json()).then(rd=>{
    if (!rd.to?.length && !rd.cc?.length) return;
    let h = '';
    if (rd.to?.length) h += `<span class="msg-recip-lbl">To:</span> <span class="triage-recip-list">${rd.to.map(r=>esc(r.name||r.address)).join(', ')}</span>`;
    if (rd.cc?.length) h += `${rd.to?.length?' &nbsp; ':''}<span class="msg-recip-lbl">CC:</span> <span class="triage-recip-list">${rd.cc.map(r=>esc(r.name||r.address)).join(', ')}</span>`;
    el.innerHTML = h;
  }).catch(()=>{});
}

function triageViewMessage(convKey) {
  const t = state.threadMap[convKey];
  if (!t || !t.latestId) return;
  const overlay = document.getElementById('view-msg-overlay');
  if (!overlay) return;
  // Set header info
  document.getElementById('view-msg-subject').textContent = t.subject || '(No subject)';
  document.getElementById('view-msg-to').textContent = 'Loading…';
  document.getElementById('view-msg-cc').textContent = '';
  document.getElementById('view-msg-body').innerHTML = '<div style="padding:20px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block;margin-right:8px"></div>Loading message…</div>';
  overlay.classList.add('open');
  // Fetch recipients
  fetch(`/api/message_recipients?id=${encodeURIComponent(t.latestId)}`).then(r=>r.json()).then(rd=>{
    document.getElementById('view-msg-to').textContent = (rd.to||[]).map(r=>r.name||r.address).join(', ') || '(none)';
    const ccEl = document.getElementById('view-msg-cc');
    const ccTxt = (rd.cc||[]).map(r=>r.name||r.address).join(', ');
    ccEl.textContent = ccTxt;
    ccEl.parentElement.style.display = ccTxt ? '' : 'none';
  }).catch(()=>{});
  // Fetch full HTML body
  fetch(`/api/format_message_stream?id=${encodeURIComponent(t.latestId)}`).then(r=>r.text()).then(txt=>{
    let bodyHtml = '';
    for (const line of txt.split('\n')) {
      if (!line.startsWith('data: ')) continue;
      try {
        const data = JSON.parse(line.slice(6));
        if (data.type === 'done' && data.body_html) bodyHtml = data.body_html;
      } catch(e) {}
    }
    const container = document.getElementById('view-msg-body');
    if (bodyHtml) {
      const safe = _injectBaseTarget(bodyHtml).replace(/"/g, '&quot;');
      container.innerHTML = `<iframe sandbox="allow-same-origin allow-popups" srcdoc="${safe}"
        style="width:100%;border:none;min-height:400px;display:block;background:#fff;border-radius:4px;"
        onload="this.style.height=Math.min(800,this.contentDocument.body.scrollHeight+20)+'px'"></iframe>`;
    } else {
      container.innerHTML = '<div style="padding:20px;color:#8b949e">No HTML body available</div>';
    }
  }).catch(()=>{
    document.getElementById('view-msg-body').innerHTML = '<div style="padding:20px;color:#f85149">Failed to load message</div>';
  });
}

function _triageMsgsHTML(convKey) {
  const msgs = state.triageMsgCache[convKey];
  if (!msgs) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf"><div class="spinner spinner-sm" style="display:inline-block"></div> Loading…</div>';
  if (!msgs.length) return '<div style="padding:8px 14px;font-size:11px;color:#5ba4cf">No messages</div>';
  return msgs.map((m,i) => _triageMsgCard(convKey, m, i)).join('');
}

function _triageMsgCard(convKey, m, idx) {
  const from = m.from_name||m.from_address||'Unknown';
  const date = fmtDate(m.received_date_time||'');
  const isSent = m.folder&&(m.folder.toLowerCase()==='sent items'||m.folder.toLowerCase()==='sent');
  if (!state.triageExpandedMsgs[convKey]) state.triageExpandedMsgs[convKey] = new Set();
  const isOpen = state.triageExpandedMsgs[convKey].has(idx);
  const toList = (m.to_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const ccList = (m.cc_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const recipRow = (toList||ccList)?`<div class="msg-recips">`
    +(toList?`<span><span class="msg-recip-lbl">To:</span>${toList}</span>`:'')
    +(ccList?`<span><span class="msg-recip-lbl">CC:</span>${ccList}</span>`:'')
    +`</div>`:'';
  const ck = esc(convKey);
  const sentBadge = isSent?'<span class="msg-sent-badge">Sent</span>':'';
  const summaryVal = state.triageMsgSummaries[m.id];
  const summaryHtml = summaryVal == null
    ? `<span class="msg-preview msg-summary-loading" id="tms-${esc(m.id)}"><span class="spinner spinner-sm" style="display:inline-block;width:8px;height:8px;margin-right:4px"></span></span>`
    : summaryVal
      ? `<span class="msg-preview" id="tms-${esc(m.id)}">${esc(summaryVal)}</span>`
      : `<span class="msg-preview" id="tms-${esc(m.id)}" style="color:#4a6080;font-style:italic">No summary</span>`;
  return `<div class="msg-card${isOpen?' open':''}${isSent?' msg-sent':''}" id="tmc-${ck}-${idx}">
    <div class="msg-hdr" data-tmc-key="${ck}" data-tmc-idx="${idx}" onclick="triageToggleMsg(this.dataset.tmcKey,+this.dataset.tmcIdx)">
      ${avatarHTML(from, m.from_address||'', 24, '')}
      <span class="msg-from-wrap"><span class="msg-from">${esc(from)}${sentBadge}</span>${summaryHtml}</span>
      <span class="msg-date">${esc(date)}</span>
      <a class="msg-owa-link" href="${'https://outlook.office.com/owa/?ItemID='+encodeURIComponent(m.id)+'&exvsurl=1&viewmodel=ReadMessageItem'}" target="_blank" rel="noopener" title="Open in Outlook Web" onclick="event.stopPropagation()">📎</a>
      <a class="msg-json-link" href="#" title="View raw JSON" onclick="event.stopPropagation();event.preventDefault();showTriageMsgJson('${ck}',${idx})">{ }</a>
      <span class="msg-chevron">▾</span>
    </div>
    ${recipRow}
    <div class="msg-body" id="tmb-${ck}-${idx}">${isOpen?_triageMsgBody(convKey,m,idx):''}</div>
  </div>`;
}

function _triageMsgBody(convKey, m, idx) {
  if (!m) return '';
  if (!state.showOriginal) state.showOriginal = {};
  const showOrig = state.showOriginal[m.id];
  const domId = `trewrite-${convKey}-${idx}`;
  const toggleBtn = `<div class="msg-body-toggle"><button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();triageToggleBodyView('${convKey}',${idx})">${showOrig ? '🤖 Smart View' : '📄 Original'}</button></div>`;

  if (showOrig) {
    // Original HTML view
    if (m.body_html) {
      const safe = _injectBaseTarget(m.body_html).replace(/"/g, '&quot;');
      return toggleBtn + `<iframe sandbox="allow-same-origin allow-popups" srcdoc="${safe}"
        style="width:100%;border:none;min-height:200px;display:block;background:#fff;border-radius:4px;"
        onload="this.style.height=Math.min(700,this.contentDocument.body.scrollHeight+20)+'px'"></iframe>`;
    }
    setTimeout(() => loadTriageMsgHtml(convKey, idx), 0);
    const plain = decodeEntities(String(m.body || m.body_preview || '')).trim();
    return toggleBtn + (plain
      ? `<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(plain)}</div>`
      : `<div style="padding:12px;color:#5ba4cf;font-size:11px"><div class="spinner spinner-sm" style="display:inline-block;margin-right:6px"></div>Loading…</div>`);
  }
  // Smart rewrite view (default)
  return toggleBtn + _rewriteBodyContent(m.id, domId);
}

function triageToggleBodyView(convKey, idx) {
  const msgs = state.triageMsgCache[convKey] || [];
  const m = msgs[idx];
  if (!m) return;
  if (!state.showOriginal) state.showOriginal = {};
  state.showOriginal[m.id] = !state.showOriginal[m.id];
  const bodyEl = document.getElementById(`tmb-${convKey}-${idx}`);
  if (bodyEl && state.triageExpandedMsgs[convKey]?.has(idx)) {
    bodyEl.innerHTML = _triageMsgBody(convKey, m, idx);
    // If switching to original and no HTML yet, fetch it
    if (state.showOriginal[m.id] && !m.body_html) loadTriageMsgHtml(convKey, idx);
  }
}

function triageToggleMsg(convKey, idx) {
  const card = document.getElementById(`tmc-${convKey}-${idx}`);
  const body = document.getElementById(`tmb-${convKey}-${idx}`);
  if (!card || !body) return;
  if (!state.triageExpandedMsgs[convKey]) state.triageExpandedMsgs[convKey] = new Set();
  const expanded = state.triageExpandedMsgs[convKey];
  const msgs = state.triageMsgCache[convKey] || [];
  const m = msgs[idx];
  if (expanded.has(idx)) {
    expanded.delete(idx);
    card.classList.remove('open');
    body.innerHTML = '';
  } else {
    expanded.add(idx);
    card.classList.add('open');
    body.innerHTML = _triageMsgBody(convKey, m, idx);
  }
}

async function loadTriageMsgSummary(msgId) {
  if (state.triageMsgSummaries[msgId] !== null && state.triageMsgSummaries[msgId] !== undefined) return;
  state.triageMsgSummaries[msgId] = null; // mark as in-flight
  try {
    const r = await fetch(`/api/summarize_message?id=${encodeURIComponent(msgId)}`).then(r=>r.json());
    state.triageMsgSummaries[msgId] = r.summary || '';
  } catch(e) {
    state.triageMsgSummaries[msgId] = '';
  }
  // Update the summary span if it's still in the DOM
  const el = document.getElementById('tms-'+msgId);
  if (el) {
    const s = state.triageMsgSummaries[msgId];
    el.className = 'msg-preview';
    el.innerHTML = '';
    if (s) el.textContent = s;
    else { el.style.color = '#4a6080'; el.style.fontStyle = 'italic'; el.textContent = 'No summary'; }
  }
}

function loadTriageMsgHtml(convKey, idx) {
  const msgs = state.triageMsgCache[convKey] || [];
  const m = msgs[idx];
  if (!m || m.body_html) return;
  const es = new EventSource(`/api/format_message_stream?id=${encodeURIComponent(m.id)}`);
  es.onmessage = (evt) => {
    const data = JSON.parse(evt.data);
    if (data.type === 'done') {
      es.close();
      if (data.body_html) {
        m.body_html = data.body_html;
        const bodyEl = document.getElementById(`tmb-${convKey}-${idx}`);
        if (bodyEl && state.triageExpandedMsgs[convKey]?.has(idx)) bodyEl.innerHTML = _triageMsgBody(convKey, m, idx);
      }
    }
  };
  es.onerror = () => es.close();
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
      const t = state.threadMap[convKey];
      const ids = t && t.emailIds && t.emailIds.length ? t.emailIds : [];
      const url = '/api/thread_messages?' + ids.map(id=>`id=${encodeURIComponent(id)}`).join('&') + '&conversationKey=' + encodeURIComponent(convKey);
      const r = await fetch(url).then(r=>r.json()).catch(()=>null);
      const msgs = (r&&r.messages||[]).slice().sort((a,b)=>(b.received_date_time||'')>(a.received_date_time||'')?1:-1);
      state.triageMsgCache[convKey] = msgs;
    }
    msgsEl.innerHTML = _triageMsgsHTML(convKey);
    // Load profile images for senders
    const senderEmails = [...new Set((state.triageMsgCache[convKey]||[]).map(m=>m.from_address).filter(Boolean))];
    if (senderEmails.length) loadProfileImages(senderEmails);
    // Kick off AI summary for any messages not yet summarised
    for (const m of (state.triageMsgCache[convKey] || [])) {
      if (m.id && state.triageMsgSummaries[m.id] === undefined) {
        loadTriageMsgSummary(m.id);
      }
    }
  } else {
    state.expandedTriageRows.delete(convKey);
    row.classList.remove('expanded');
    msgsEl.style.display = 'none';
  }
}


async function triageOpenReply(convKey, mode) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  state.returnToTriage = true;
  document.removeEventListener('keydown', _triageKeydown);
  await selectThread(convKey);  // hides all panes, shows thread-detail
  openReply(encodeThread(thread), mode || 'all');
}

async function triageOpenForward(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  state.returnToTriage = true;
  document.removeEventListener('keydown', _triageKeydown);
  await selectThread(convKey);  // hides all panes, shows thread-detail
  openForward(encodeThread(thread));
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
  pane.innerHTML = `<div class="triage-rows">${groupsHtml}</div>
  <div class="mbox-kb-hint">
    <span><kbd>j</kbd><kbd>k</kbd> navigate</span>
    <span><kbd>Enter</kbd> expand</span>
    <span><kbd>r</kbd> reply</span>
    <span><kbd>d</kbd> delete</span>
    <span><kbd>f</kbd> file</span>
    <span><kbd>Esc</kbd> back</span>
  </div>`;
  // Load recipients for all visible triage rows
  for (const g of sortedGroups) {
    if (!state.collapsedTriageTopics.has(g.topic || 'Uncategorized')) {
      for (const t of g.threads) _loadTriageRecips(t.conversationKey);
    }
  }
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
    if (item&&item.type==='thread') { e.preventDefault(); triageDeleteNow(item.convKey); }
    return;
  } else if (e.key==='f'||e.key==='F') {
    const item = nav[idx];
    if (item&&item.type==='thread') { e.preventDefault(); triageFileNow(item.convKey); }
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

async function triageDeleteNow(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  await _act('/api/delete', {ids: thread.emailIds, conversationKey: convKey}, convKey);
  openTriageSheet();
}

function showTriageMsgJson(convKey, idx) {
  const msgs = state.triageMsgCache[convKey] || [];
  const m = msgs[idx];
  if (!m) return;
  const overlay = document.getElementById('json-debug-overlay');
  const pre = document.getElementById('json-debug-content');
  if (!overlay || !pre) return;
  let obj = m;
  if (m.raw_json) {
    try { obj = JSON.parse(m.raw_json); } catch(e) { obj = m; }
  }
  pre.textContent = JSON.stringify(obj, null, 2);
  overlay.classList.add('open');
}

function triageFileNow(convKey) {
  const thread = state.threadMap[convKey];
  if (!thread) return;
  state.returnToTriage = true;
  openFile(encodeThread(thread));
}
