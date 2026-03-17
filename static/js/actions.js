// ── Actions ────────────────────────────────────────────────────────────────────
function openFile(enc) {
  _activeThread=decodeThread(enc);
  const suggested=_activeThread.suggestedFolder||'';
  document.getElementById('file-sub').textContent=`"${_activeThread.subject||''}"`;
  const populate=(folders,effortsFolders)=>{
    const es=new Set(effortsFolders||[]);
    const eff=folders.filter(n=>es.has(n)),oth=folders.filter(n=>!es.has(n));
    let h='';
    if (eff.length) h+=`<optgroup label="Efforts">`+eff.map(n=>`<option value="${esc(n)}"${n===suggested?' selected':''}>${esc(n)}</option>`).join('')+`</optgroup>`;
    if (oth.length) h+=`<optgroup label="Other">`+oth.map(n=>`<option value="${esc(n)}"${n===suggested?' selected':''}>${esc(n)}</option>`).join('')+`</optgroup>`;
    document.getElementById('folder-select').innerHTML=h;
  };
  if (state.folders.length){populate(state.folders,state.effortsFolders);document.getElementById('file-modal').classList.add('open');}
  else fetch('/api/folders').then(r=>r.json()).then(d=>{state.folders=d.folders||[];state.effortsFolders=d.effortsFolders||[];populate(state.folders,state.effortsFolders);document.getElementById('file-modal').classList.add('open');});
}

async function openDelete(enc) {
  const t=decodeThread(enc);
  await _act('/api/delete',{ids:t.emailIds,conversationKey:t.conversationKey},t.conversationKey);
}

function closeModals() {
  document.querySelectorAll('.modal-overlay').forEach(m=>m.classList.remove('open'));
  _activeThread=null;
  // Close any open people dropdowns
  document.querySelectorAll('.people-dropdown').forEach(d=>d.classList.remove('open'));
}

async function fileThread() {
  const folder=document.getElementById('folder-select').value;
  const t=_activeThread; closeModals();
  await _act('/api/move',{ids:t.emailIds,folder,conversationKey:t.conversationKey},t.conversationKey);
}

async function quickFile(enc,folder) {
  const t=decodeThread(enc);
  await _act('/api/move',{ids:t.emailIds,folder,conversationKey:t.conversationKey},t.conversationKey);
}

async function confirmDelete() {
  const t=_activeThread; closeModals();
  await _act('/api/delete',{ids:t.emailIds,conversationKey:t.conversationKey},t.conversationKey);
}

async function toggleFlag(enc) {
  const t=decodeThread(enc);
  const thread=state.threadMap[t.conversationKey];
  if (!thread) return;
  const nowFlagged=!thread.isFlagged;
  const d=await fetch('/api/flag',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({conversationKey:t.conversationKey,flagged:nowFlagged})}).then(r=>r.json()).catch(()=>null);
  if (!d||!d.ok) return;
  thread.isFlagged=nowFlagged;
  // Update flag button immediately
  const btn=document.getElementById('flag-btn-'+t.conversationKey);
  if (btn){
    btn.textContent=nowFlagged?'🚩 Flagged':'🚩 Flag';
    btn.className='btn btn-flag btn-sm'+(nowFlagged?' flagged':'');
  }
  renderSidebar();
}

function _showActSpinner(msg) {
  let el = document.getElementById('act-spinner');
  if (!el) {
    el = document.createElement('div');
    el.id = 'act-spinner';
    el.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#0d2040;border:1px solid #2a5a8a;border-radius:10px;padding:10px 20px;display:flex;align-items:center;gap:10px;font-size:13px;color:#c9d1d9;z-index:99999;box-shadow:0 8px 24px rgba(0,0,0,.5);transition:opacity .2s;';
    document.body.appendChild(el);
  }
  el.innerHTML = `<div class="spinner spinner-sm" style="border-top-color:#58a6ff"></div>${msg}`;
  el.style.opacity = '1';
}

function _hideActSpinner() {
  const el = document.getElementById('act-spinner');
  if (el) { el.style.opacity = '0'; setTimeout(()=>el.remove(), 200); }
}

async function _act(url,body,convKey) {
  const label = url.includes('delete') ? 'Deleting…' : url.includes('move') ? 'Filing…' : 'Sending…';
  _showActSpinner(label);
  const res=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const d=await res.json().catch(()=>({}));
  _hideActSpinner();
  if (!d.ok) return alert('Error: '+(d.error||'Unknown error'));
  // Find next mailbox row before removing from DOM
  const mboxRow = document.querySelector(`.mbox-row[data-key="${CSS.escape(convKey)}"]`);
  const nextRow = mboxRow?.nextElementSibling;
  if (mboxRow) mboxRow.remove();
  delete state.threadMap[convKey];
  for (const g of state.groups) g.threads=g.threads.filter(t=>t.conversationKey!==convKey);
  state.groups=state.groups.filter(g=>g.threads.length>0);
  if (state.selectedKey===convKey) {
    state.selectedKey=null;
    const rb=document.getElementById('resync-thread-btn');
    if(rb) rb.disabled=true;
    if (state.mailboxContext && nextRow && nextRow.dataset.key) {
      // Advance to next thread in folder
      openMailboxThread(nextRow.dataset.key, nextRow.dataset.folder || mailboxCurrentFolder);
    } else if (state.mailboxContext) {
      backToMailboxList();
    } else {
      document.getElementById('thread-detail').style.display='none';
      document.getElementById('empty-pane').style.display='flex';
    }
  }
  renderSidebar();
  updateCounts(null,Object.keys(state.threadMap).length);
}
