// ── Reply (2-step) ─────────────────────────────────────────────────────────────
let _replyState = {thread:null, to:[], cc:[]};

async function openReply(enc) {
  _replyState.thread = decodeThread(enc);
  _activeThread = _replyState.thread;
  const t = _replyState.thread;

  // Build reply-all from ALL messages in thread: collect every sender + To across all msgs into To,
  // every CC across all msgs into CC; deduplicate; exclude self; anyone in To is removed from CC.
  _replyState.to = [];
  _replyState.cc = [];
  const myAddr = MY_EMAIL.toLowerCase();
  const addUniq = (list, r) => {
    if (!r || !r.address) return;
    if (r.address.toLowerCase() === myAddr) return; // exclude self
    if (!list.find(x=>x.address.toLowerCase()===r.address.toLowerCase())) list.push(r);
  };
  for (const msg of state.currentMsgs) {
    if (msg.from_address) addUniq(_replyState.to, {name:msg.from_name||msg.from_address, address:msg.from_address});
    for (const r of (msg.to_recipients||[])) addUniq(_replyState.to, r);
    for (const r of (msg.cc_recipients||[])) addUniq(_replyState.cc, r);
  }
  // Remove anyone already in To from CC
  _replyState.cc = _replyState.cc.filter(
    r => !_replyState.to.find(t => t.address.toLowerCase() === r.address.toLowerCase())
  );

  document.getElementById('reply-sub').textContent = `Re: ${t.subject||''}`;
  const bodyEl = document.getElementById('reply-body');
  bodyEl.value = '';
  bodyEl.placeholder = 'Generating reply…';
  _renderRecipFields();
  document.getElementById('reply-modal').classList.add('open');
  document.getElementById('reply-generating').style.display = 'flex';

  // Auto-populate with suggested reply
  try {
    // Use cached suggestedReply if available
    const thread = state.threadMap[t.conversationKey];
    if (thread && thread.suggestedReply) {
      bodyEl.value = thread.suggestedReply;
      bodyEl.placeholder = '';
    } else {
      const d = await fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({conversationKey: t.conversationKey})}).then(r=>r.json()).catch(()=>null);
      if (d && d.reply) {
        bodyEl.value = d.reply;
        if (thread) thread.suggestedReply = d.reply;
      }
      bodyEl.placeholder = '';
    }
  } catch(e) { bodyEl.placeholder = 'Write your reply…'; }
  document.getElementById('reply-generating').style.display = 'none';
  setTimeout(()=>{ bodyEl.focus(); bodyEl.setSelectionRange(0,0); }, 50);
}

async function regenerateReply() {
  const t = _replyState.thread;
  if (!t) return;
  const bodyEl = document.getElementById('reply-body');
  const context = bodyEl.value.trim();
  const btn = document.getElementById('reply-regen-btn');
  btn.disabled = true; btn.textContent = '↺ Generating…';
  document.getElementById('reply-generating').style.display = 'flex';
  try {
    const d = await fetch('/api/suggested_reply', {method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: t.conversationKey, context})}).then(r=>r.json()).catch(()=>null);
    if (d && d.reply) {
      bodyEl.value = d.reply;
      const thread = state.threadMap[t.conversationKey];
      if (thread) thread.suggestedReply = d.reply;
    }
  } finally {
    document.getElementById('reply-generating').style.display = 'none';
    btn.disabled = false; btn.textContent = '↺ Regenerate';
  }
}

async function sendReply() {
  const body=document.getElementById('reply-body').value.trim();
  if (!body) return;
  const t=_replyState.thread||_activeThread;
  const to=_replyState.to.map(r=>r.address).filter(Boolean);
  const cc=_replyState.cc.map(r=>r.address).filter(Boolean);
  closeModals();
  await _act('/api/reply/'+t.latestId,{body,conversationKey:t.conversationKey,to,cc},t.conversationKey);
}

function _renderRecipFields() {
  _renderTags('reply-to-field', _replyState.to, 'to');
  _renderTags('reply-cc-field', _replyState.cc, 'cc');
}
function _renderTags(fieldId, list, field) {
  const el = document.getElementById(fieldId);
  if (!list.length) { el.innerHTML='<span class="recip-empty">none</span>'; return; }
  el.innerHTML = list.map(r=>
    `<span class="recip-tag" data-field="${field}" data-addr="${esc(r.address)}">${esc(r.name||r.address)}<span class="rm" data-rm-field="${field}" data-rm-addr="${esc(r.address)}">×</span></span>`
  ).join('');
  el.querySelectorAll('.rm').forEach(btn=>{
    btn.addEventListener('click', ()=>removeRecip(btn.dataset.rmField, btn.dataset.rmAddr));
  });
}
function removeRecip(field, address) {
  _replyState[field] = _replyState[field].filter(r=>r.address!==address);
  _renderRecipFields();
}

// ── Inline reply ───────────────────────────────────────────────────────────────
async function sendInlineReply(enc) {
  const t = decodeThread(enc);
  const ta = document.getElementById('inline-reply-'+enc);
  const body = ta ? ta.value.trim() : '';
  if (!body) return;
  const res = await fetch('/api/reply/'+t.latestId, {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({body, conversationKey: t.conversationKey, to: [], cc: []})
  }).then(r=>r.json()).catch(()=>null);
  if (!res || !res.ok) { alert('Error: '+(res&&res.error||'Unknown error')); return; }
  delete state.threadMap[t.conversationKey];
  for (const g of state.groups) g.threads = g.threads.filter(th=>th.conversationKey!==t.conversationKey);
  state.groups = state.groups.filter(g=>g.threads.length>0);
  if (state.selectedKey===t.conversationKey) {
    state.selectedKey = null;
    document.getElementById('thread-detail').style.display='none';
    document.getElementById('empty-pane').style.display='flex';
    const rb=document.getElementById('resync-thread-btn');
    if(rb) rb.disabled=true;
  }
  renderSidebar();
  updateCounts(null, Object.keys(state.threadMap).length);
}

async function regenerateInlineReply(enc) {
  const t = decodeThread(enc);
  const ta = document.getElementById('inline-reply-'+enc);
  if (ta) { ta.value = '⏳ Generating...'; ta.disabled = true; }
  try {
    const res = await fetch('/api/suggested_reply', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({conversationKey: t.conversationKey})
    }).then(r=>r.json());
    if (res.reply) {
      const thread = state.threadMap[t.conversationKey];
      if (thread) thread.suggestedReply = res.reply;
      if (ta) { ta.value = res.reply; ta.disabled = false; }
      else {
        // textarea may not exist yet (was "Generate reply..." button) — re-render header
        const fullThread = state.threadMap[t.conversationKey];
        if (fullThread) _renderThreadHdr(fullThread);
      }
    } else {
      if (ta) { ta.value = ''; ta.disabled = false; }
    }
  } catch(e) {
    if (ta) { ta.value = ''; ta.disabled = false; }
    alert('Error regenerating reply: '+e);
  }
}
