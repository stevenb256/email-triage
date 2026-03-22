// ── Select thread ──────────────────────────────────────────────────────────────
async function selectThread(convKey) {
  state.selectedKey = convKey;
  state.expandedMsgs = new Set();
  state.currentMsgs = [];
  state.triageView = false;
  state.mailboxContext = false;
  const rb=document.getElementById('resync-thread-btn');
  if(rb){rb.disabled=false;rb.textContent='↺ Resync Thread';}
  renderSidebar();
  const t = state.threadMap[convKey];
  if (!t) return;
  // Hide ALL right-pane views before showing thread detail
  ['empty-pane','triage-pane','mailbox-pane','search-pane','calendar-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  const navTriage = document.getElementById('nav-triage'); if (navTriage) navTriage.classList.remove('active');
  document.getElementById('thread-detail').style.display='flex';
  _renderThreadHdr(t);
  const sec = document.getElementById('msgs-section');
  sec.innerHTML=`<div class="msg-ai-loading"><div class="spinner spinner-sm"></div> Loading messages…</div>`;
  if (!t.emailIds||!t.emailIds.length){sec.innerHTML+='<div style="color:#5ba4cf;font-size:12px;padding:10px 0">No messages found.</div>';return;}
  // Pass both explicit IDs and conversationKey so backend can merge in sent items
  const params=(t.emailIds||[]).map(id=>`id=${encodeURIComponent(id)}`).join('&');
  const d=await fetch('/api/thread_messages?'+params+'&conversationKey='+encodeURIComponent(convKey)).then(r=>r.json()).catch(()=>({messages:[]}));
  let msgs=d.messages||[];
  msgs=msgs.slice().sort((a,b)=>(b.received_date_time||'')>(a.received_date_time||'')?1:-1);
  state.currentMsgs=msgs;
  _renderMsgs(msgs,t);
  // Load profile images for all senders
  const senderEmails = [...new Set(msgs.map(m=>m.from_address).filter(Boolean))];
  if (senderEmails.length) loadProfileImages(senderEmails);
  if (msgs.length>0 && !state.returnToTriage) toggleMsg(0);
  // Auto mark-read
  if (t.hasUnread && t.emailIds && t.emailIds.length) {
    fetch('/api/markread',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({ids:t.emailIds,conversationKey:convKey})}).then(()=>{
      t.hasUnread=false;
      renderSidebar();
    }).catch(()=>{});
  }
}

function _renderThreadHdr(t) {
  const enc=encodeThread(t);
  const urgCls={high:'urg-high',medium:'urg-medium',low:'urg-low'}[t.urgency]||'urg-low';
  const actCls={reply:'act-reply',delete:'act-delete',file:'act-file'}[t.action]||'';
  const parts=t.participants||[];
  const avHTML=parts.slice(0,5).map(p=>{const email=(t._participantEmails&&t._participantEmails[p])||'';return avatarHTML(p, email, 20, '');}).join('');
  const dateStr=fmtDate(t.latestReceived||'');
  const fileLabel=t.suggestedFolder?`📁 ${esc(t.suggestedFolder)}`:'📁 File';
  let fileBtnHtml=`<button class="btn btn-file btn-sm" onclick="openFile('${enc}')">${fileLabel}</button>`;
  document.getElementById('thread-hdr').innerHTML=`
    <div class="th-top">
      ${state.mailboxContext?`<button class="mbox-back" onclick="backToMailboxList()">✕ Close</button>`:''}
      <span class="urg-pill ${urgCls}">${(t.urgency||'low').toUpperCase()}</span>
      <span class="act-pill ${actCls}">${t.action||'read'}</span>
      ${t.hasUnread?'<span style="width:6px;height:6px;border-radius:50%;background:#58a6ff;display:inline-block;flex-shrink:0"></span>':''}
      <div class="th-subject">${esc(t.subject||'(No subject)')}</div>
      <div class="th-date">${esc(dateStr)}</div>
    </div>
    <div class="th-participants">
      <div class="avatars">${avHTML}</div>
      <span class="th-names">${esc(parts.slice(0,4).join(', '))}${parts.length>4?' +'+(parts.length-4):''}</span>
      <span class="th-msgcount">${t.messageCount||0} msg${(t.messageCount||0)!==1?'s':''}</span>
    </div>
    <div class="th-recips" id="th-recips"></div>
    ${t.summary?`<div class="th-summary"><div class="th-summary-lbl">🤖 AI Summary</div><div class="th-summary-body">${_renderSummary(t.summary)}</div></div>`:''}
    <div class="th-actions">
      <button class="btn btn-reply btn-sm" onclick="openReply('${enc}','all')">↩ Reply All</button>
      <button class="btn btn-ghost btn-sm" onclick="openReply('${enc}','sender')">↩ Reply</button>
      <button class="btn btn-ghost btn-sm" onclick="openForward('${enc}')">⤳ Forward</button>
      ${fileBtnHtml}
      <button class="btn btn-flag btn-sm${t.isFlagged?' flagged':''}" id="flag-btn-${esc(t.conversationKey)}" onclick="toggleFlag('${enc}')">${t.isFlagged?'🚩 Flagged':'🚩 Flag'}</button>
      <button class="btn btn-delete btn-sm" onclick="openDelete('${enc}')">🗑 Delete</button>
    </div>`;
  // Fetch and render recipients for this thread
  if (t.latestId) {
    fetch(`/api/message_recipients?id=${encodeURIComponent(t.latestId)}`).then(r=>r.json()).then(rd=>{
      const el = document.getElementById('th-recips');
      if (!el || (!rd.to?.length && !rd.cc?.length)) return;
      let h = '';
      if (rd.to?.length) h += `<span class="th-recip-lbl">To:</span> <span class="th-recip-list">${rd.to.map(r=>esc(r.name||r.address)).join(', ')}</span>`;
      if (rd.cc?.length) h += `${rd.to?.length?' &nbsp; ':''}<span class="th-recip-lbl">CC:</span> <span class="th-recip-list">${rd.cc.map(r=>esc(r.name||r.address)).join(', ')}</span>`;
      el.innerHTML = h;
    }).catch(()=>{});
  }
}

function _renderMsgs(msgs, t) {
  const sec=document.getElementById('msgs-section');
  if (!msgs.length){sec.innerHTML=`<div style="color:#5ba4cf;font-size:12px;padding:16px 0">No messages found.</div>`;return;}
  let html=`<div class="msgs-label"></div>`;
  html+=msgs.map((m,i)=>_msgCardHTML(m,i)).join('');
  sec.innerHTML=html;
}

function _msgCardHTML(m, idx) {
  const from=m.from_name||m.from_address||'Unknown';
  const date=fmtDate(m.received_date_time||'');
  const preview=(m.body_preview||'').replace(/\n+/g,' ');
  const isOpen=state.expandedMsgs.has(idx);
  const isSent=m.folder&&(m.folder.toLowerCase()==='sent items'||m.folder.toLowerCase()==='sent');
  const toList=(m.to_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const ccList=(m.cc_recipients||[]).map(r=>esc(r.name||r.address)).join(', ');
  const recipRow=(toList||ccList)?`<div class="msg-recips">`
    +(toList?`<span><span class="msg-recip-lbl">To:</span>${toList}</span>`:'')
    +(ccList?`<span><span class="msg-recip-lbl">CC:</span>${ccList}</span>`:'')
    +`</div>`:'';
  const sentBadge=isSent?'<span class="msg-sent-badge">Sent</span>':'';
  return `<div class="msg-card${isOpen?' open':''}${isSent?' msg-sent':''}" id="mc-${idx}">
    <div class="msg-hdr" onclick="toggleMsg(${idx})">
      ${avatarHTML(from, m.from_address||'', 24, '')}
      <span class="msg-from-wrap"><span class="msg-from">${esc(from)}${sentBadge}</span><span class="msg-preview">${esc(preview)}</span></span>
      <span class="msg-date">${esc(date)}</span>
      <a class="msg-owa-link" href="${'https://outlook.office.com/owa/?ItemID='+encodeURIComponent(m.id)+'&exvsurl=1&viewmodel=ReadMessageItem'}" target="_blank" rel="noopener" title="Open in Outlook Web" onclick="event.stopPropagation()">📎</a>
      <a class="msg-json-link" href="#" title="View raw JSON" onclick="event.stopPropagation();event.preventDefault();showMsgJson(${idx})">{ }</a>
      <span class="msg-chevron">▾</span>
    </div>
    ${recipRow}
    <div class="msg-body" id="mb-${idx}">${isOpen?_bodyContent(idx):''}</div>
  </div>`;
}

function showMsgJson(idx) {
  const m = state.currentMsgs[idx];
  if (!m) return;
  const overlay = document.getElementById('json-debug-overlay');
  const pre = document.getElementById('json-debug-content');
  if (!overlay || !pre) return;
  // Show the message object; prefer raw_json if available
  let obj = m;
  if (m.raw_json) {
    try { obj = JSON.parse(m.raw_json); } catch(e) { obj = m; }
  }
  pre.textContent = JSON.stringify(obj, null, 2);
  overlay.classList.add('open');
}

// LLM rewrite: clear bullet/action format for expanded messages
// domId is a unique DOM element id for the rewrite container (supports both thread and triage views)
function _rewriteBodyContent(msgId, domId) {
  if (!msgId) return '';
  // Check cache first
  if (state.rewriteCache[msgId]) {
    return `<div class="msg-rewrite">${state.rewriteCache[msgId]}</div>`;
  }
  // Start streaming rewrite
  if (!state.rewriteInFlight[msgId]) {
    state.rewriteInFlight[msgId] = true;
    _streamRewrite(msgId, domId);
  }
  return `<div class="msg-rewrite" id="${domId}"><div style="padding:12px;color:#5ba4cf;font-size:12px"><div class="spinner spinner-sm" style="display:inline-block;margin-right:6px"></div>Rewriting for clarity…</div></div>`;
}

function _streamRewrite(msgId, domId) {
  let full = '';
  const es = new EventSource(`/api/rewrite_message_stream?id=${encodeURIComponent(msgId)}`);
  es.onmessage = (evt) => {
    const data = JSON.parse(evt.data);
    if (data.type === 'token') {
      full += data.text;
      const el = document.getElementById(domId);
      if (el) el.innerHTML = `<div class="msg-rewrite-content">${full}</div>`;
    } else if (data.type === 'done') {
      es.close();
      state.rewriteCache[msgId] = data.html || full;
      delete state.rewriteInFlight[msgId];
      const el = document.getElementById(domId);
      if (el) el.innerHTML = `<div class="msg-rewrite-content">${state.rewriteCache[msgId]}</div>`;
    }
  };
  es.onerror = () => { es.close(); delete state.rewriteInFlight[msgId]; };
}

function _bodyContent(idx) {
  const m = state.currentMsgs[idx];
  if (!m) return '';
  // Default to rewrite view; toggle with button
  if (!state.showOriginal) state.showOriginal = {};
  const showOrig = state.showOriginal[m.id];
  const toggleBtn = `<div class="msg-body-toggle"><button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();toggleBodyView(${idx})">${showOrig ? '🤖 Smart View' : '📄 Original'}</button></div>`;

  if (showOrig) {
    // Original HTML view
    if (m.body_html) {
      const safe = _injectBaseTarget(m.body_html).replace(/"/g, '&quot;');
      return toggleBtn + `<iframe sandbox="allow-same-origin allow-popups" srcdoc="${safe}"
        style="width:100%;border:none;min-height:200px;display:block;background:#fff;border-radius:4px;"
        onload="this.style.height=Math.min(700,this.contentDocument.body.scrollHeight+20)+'px'"></iframe>`;
    }
    setTimeout(() => loadMsgHtml(idx), 0);
    const plain = decodeEntities(String(m.body || m.body_preview || '')).trim();
    return toggleBtn + (plain
      ? `<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(plain)}</div>`
      : `<div style="padding:12px;color:#5ba4cf;font-size:11px"><div class="spinner spinner-sm" style="display:inline-block;margin-right:6px"></div>Loading…</div>`);
  }
  // Smart rewrite view (default)
  return toggleBtn + _rewriteBodyContent(m.id, 'rewrite-'+idx);
}

function toggleBodyView(idx) {
  const m = state.currentMsgs[idx];
  if (!m) return;
  if (!state.showOriginal) state.showOriginal = {};
  state.showOriginal[m.id] = !state.showOriginal[m.id];
  const bodyEl = document.getElementById('mb-'+idx);
  if (bodyEl && state.expandedMsgs.has(idx)) {
    bodyEl.innerHTML = _bodyContent(idx);
    // If switching to original and no HTML yet, fetch it
    if (state.showOriginal[m.id] && !m.body_html) loadMsgHtml(idx);
  }
}

function loadMsgHtml(idx) {
  const m = state.currentMsgs[idx];
  if (!m || m.body_html) return;
  const es = new EventSource(`/api/format_message_stream?id=${encodeURIComponent(m.id)}`);
  es.onmessage = (evt) => {
    const data = JSON.parse(evt.data);
    if (data.type === 'done') {
      es.close();
      if (data.body_html) {
        m.body_html = data.body_html;
        const bodyEl = document.getElementById('mb-'+idx);
        if (bodyEl && state.expandedMsgs.has(idx)) bodyEl.innerHTML = _bodyContent(idx);
      }
    }
  };
  es.onerror = () => es.close();
}

function toggleMsg(idx) {
  const card=document.getElementById('mc-'+idx);
  const body=document.getElementById('mb-'+idx);
  if (!card||!body) return;
  if (state.expandedMsgs.has(idx)) {
    state.expandedMsgs.delete(idx);
    card.classList.remove('open');
    body.innerHTML='';
  } else {
    state.expandedMsgs.add(idx);
    card.classList.add('open');
    body.innerHTML=_bodyContent(idx);
  }
}

function loadFormatted(idx) {
  const m=state.currentMsgs[idx];
  if (!m||!state.expandedMsgs.has(idx)) return;
  const bodyEl=document.getElementById('mb-'+idx);
  if (!bodyEl) return;

  // Cached — render immediately, no stream needed
  if (state.formatCache[m.id]) {
    bodyEl.innerHTML=_renderParas(state.formatCache[m.id]);
    return;
  }

  // Set up streaming container
  bodyEl.innerHTML='<div class="stream-wrap" id="sw-'+idx+'"></div>';
  const wrap=document.getElementById('sw-'+idx);
  let shownCount=0;
  let accumulated='';

  const es=new EventSource(`/api/format_message_stream?id=${encodeURIComponent(m.id)}`);

  es.onmessage=(evt)=>{
    const data=JSON.parse(evt.data);
    if (data.type==='token') {
      accumulated+=data.text;
      if (!wrap||!state.expandedMsgs.has(idx)) {es.close();return;}
      // Extract completed "text":"..." values as paragraphs become available
      const matches=[...accumulated.matchAll(/"text":\s*"((?:[^"\\]|\\.)*)"/g)];
      for (let i=shownCount;i<matches.length;i++) {
        const txt=matches[i][1].replace(/\\n/g,'\n').replace(/\\"/g,'"').replace(/\\\\/g,'\\');
        const div=document.createElement('div');
        div.className='stream-para';
        div.textContent=txt;
        wrap.appendChild(div);
      }
      shownCount=matches.length;
      // Keep cursor at end
      let cur=wrap.querySelector('.stream-cursor');
      if (!cur){cur=document.createElement('span');cur.className='stream-cursor';wrap.appendChild(cur);}
      else wrap.appendChild(cur); // move to end
    } else if (data.type==='done') {
      es.close();
      state.formatCache[m.id]=data.paragraphs||[];
      // Store body_html on the message object so HTML view works
      if (data.body_html) {
        m.body_html = data.body_html;
      }
      if (bodyEl && state.expandedMsgs.has(idx)) {
        bodyEl.innerHTML = _bodyContent(idx);
      }
    }
  };

  es.onerror=()=>{
    es.close();
    if (bodyEl&&state.expandedMsgs.has(idx)) {
      const fallback=decodeEntities(String(m.body||m.body_preview||'')).trim();
      bodyEl.innerHTML=`<div style="font-size:12px;color:#c9d1d9;line-height:1.8;white-space:pre-wrap">${esc(fallback)}</div>`;
    }
  };
}

function toggleFormatView(idx) {
  const m = state.currentMsgs[idx];
  if (!m) return;
  // Effective current state (mirrors _bodyContent default logic)
  const hasHtml = !!(m.body_html);
  const current = state.showOriginal[m.id] !== undefined ? !!state.showOriginal[m.id] : hasHtml;
  state.showOriginal[m.id] = !current;
  const bodyEl = document.getElementById('mb-'+idx);
  const btn = document.getElementById('fmt-btn-'+idx);
  if (btn) btn.textContent = state.showOriginal[m.id] ? 'AI view' : (hasHtml ? 'HTML' : 'Original');
  if (state.expandedMsgs.has(idx) && bodyEl) {
    bodyEl.innerHTML = _bodyContent(idx);
    if (!state.showOriginal[m.id] && !state.formatCache[m.id]) loadFormatted(idx);
  }
}

function _renderParas(paras) {
  if (!paras||!paras.length) return '<div style="color:#5ba4cf;font-size:12px">(no content)</div>';
  return paras.map(p=>{
    const cls='i-'+(INTENT_CLS[p.intent]||'context');
    const factHtml=p.fact_concern
      ?`<div class="fact-warn"><span>⚠️</span><span>${esc(p.fact_concern)}</span></div>`:'';
    return `<div class="para-blk">
      <div class="intent-pill ${cls}">${esc(p.emoji||'')} ${esc(p.intent||'FYI')}</div>
      <div class="para-txt">${linkify(p.text||'')}</div>
      ${factHtml}
    </div>`;
  }).join('');
}
