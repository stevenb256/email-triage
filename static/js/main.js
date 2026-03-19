// ── State ──────────────────────────────────────────────────────────────────────
let state = {
  groups: [],
  threadMap: {},
  selectedKey: null,
  collapsedTopics: new Set(),
  expandedMsgs: new Set(),// indices of expanded messages in thread-detail view
  currentMsgs: [],        // messages for selected thread (newest-first)
  latestTs: '',
  folders: [],
  effortsFolders: [],
  pollTimer: null,
  triageActions: {},      // map of conversationKey → {type: 'send'|'delete'|'file', reply: string}
  triageView: false,
  mailboxContext: false,  // true when thread opened from mailbox view
  collapsedTriageTopics: new Set(),
  triageFocusIdx: -1,
  expandedTriageRows: new Set(),
  triageMsgCache: {},
  triageExpandedMsgs: {}, // convKey → Set of expanded msg indices
  triageMsgSummaries: {}, // msgId → summary string (or '' if done, null if pending)
};
let _activeThread = null;
let MY_EMAIL = '';
let _pdCache = null;

// ── Init ───────────────────────────────────────────────────────────────────────
async function init() {
  const d = await fetch('/api/threads').then(r=>r.json()).catch(()=>null);
  if (!d) { setTimeout(init,3000); return; }
  if (d.groups.length===0 && d.syncStatus.running) {
    document.getElementById('load-msg').textContent='Syncing inbox for the first time…';
    updateSyncStatus(d.syncStatus);
    setTimeout(init,3000); return;
  }
  state.groups = d.groups;
  state.threadMap = {};
  state.groups.forEach(g=>g.threads.forEach(t=>{state.threadMap[t.conversationKey]=t;}));
  state.collapsedTopics = new Set(state.groups.map(g=>g.topic));
  state.latestTs = d.latestTs;
  const splash = document.getElementById('first-load');
  setTimeout(() => {
    splash.classList.add('fade-out');
    setTimeout(()=>{ splash.style.display='none'; }, 650);
  }, 1000);
  updateCounts(d.emailCount, Object.keys(state.threadMap).length);
  updateSyncStatus(d.syncStatus);

  renderSidebar();
  renderTodayCal();
  openTriageSheet();
  schedulePoll();

  // Load my email for reply-all filtering + preload people cache
  fetch('/api/my_email').then(r=>r.json()).then(d=>{ if(d.email) MY_EMAIL=d.email; }).catch(()=>{});
  fetch('/api/people').then(r=>r.json()).then(d=>{ if(d.people) _pdCache=d.people; }).catch(()=>{});
}

// ── Poll ───────────────────────────────────────────────────────────────────────
function schedulePoll() {
  clearTimeout(state.pollTimer);
  state.pollTimer = setTimeout(pollUpdates,10000);
}
async function pollUpdates() {
  const d = await fetch(`/api/updates?since=${encodeURIComponent(state.latestTs)}`).then(r=>r.json()).catch(()=>null);
  if (!d) { schedulePoll(); return; }
  updateSyncStatus(d.syncStatus);
  // Update next meeting display if present
  try {
    const nm = d.nextMeeting || {};
    const nmEl = document.getElementById('next-meeting');
    if (nm && nm.subject && nm.start_time) {
      nmEl.textContent = `Next: ${esc(nm.subject)} · ${fmtUntil(nm.start_time)}`;
    } else if (nmEl) { nmEl.textContent = ''; }
  } catch(e) {}
  if (d.threads && d.threads.length>0) {
    let newCount=0;
    for (const t of d.threads) {
      if (t.updatedAt>state.latestTs) state.latestTs=t.updatedAt;
      const isNew=!state.threadMap[t.conversationKey];
      state.threadMap[t.conversationKey]=t;
      if (isNew) { newCount++; _insertGroup(t); } else { _updateGroup(t); }
    }
    renderSidebar();
    updateCounts(null, Object.keys(state.threadMap).length);
    if (newCount>0) _showNewBadge(newCount);
    if (state.selectedKey && state.threadMap[state.selectedKey]) {
      _renderThreadHdr(state.threadMap[state.selectedKey]);
    }
  }
  schedulePoll();
}
function _insertGroup(t) {
  let g=state.groups.find(g=>g.topic===t.topic);
  if (!g){g={topic:t.topic,threads:[]};state.groups.unshift(g);}
  g.threads.unshift(t);
}
function _updateGroup(t) {
  for (const g of state.groups) {
    const i=g.threads.findIndex(x=>x.conversationKey===t.conversationKey);
    if (i>=0){if(g.topic!==t.topic){g.threads.splice(i,1);_insertGroup(t);}else{g.threads[i]=t;}return;}
  }
  _insertGroup(t);
}

// ── Tab switching ──────────────────────────────────────────────────────────────
let activeTab = 'mailbox';
function switchTab(tab) {
  activeTab = tab;
  clearSearch();
  if (state.triageView) { state.triageView = false; document.removeEventListener('keydown', _triageKeydown); }
  if (tab !== 'mailbox') _mboxUnregisterKeys();
  // Sidebar nav highlight — clear all folder items then re-apply
  document.querySelectorAll('.folder-item.active').forEach(e=>e.classList.remove('active'));
  const calBtn = document.getElementById('nav-calendar');
  if (calBtn) calBtn.classList.toggle('active', tab==='calendar');
  // Title bar actions always visible
  document.getElementById('triage-actions').style.display = 'flex';
  // Sidebar always visible
  document.getElementById('sidebar').style.display = '';
  document.getElementById('resize-handle').style.display = '';
  // Hide all right-pane views
  ['empty-pane','thread-detail','triage-pane','mailbox-pane','calendar-pane','search-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  // Show tab-specific view
  if (tab === 'email') {
    const hasThread = !!document.getElementById('thread-detail').dataset.loaded;
    document.getElementById(hasThread ? 'thread-detail' : 'empty-pane').style.display = '';
    initMailbox();
  } else if (tab === 'mailbox') {
    document.getElementById('mailbox-pane').style.display = 'flex';
    initMailbox();
  } else if (tab === 'calendar') {
    document.getElementById('calendar-pane').style.display = 'flex';
    renderCalendar();
  }
}

// ── Toast ──────────────────────────────────────────────────────────────────────
function showToast(msg, isError=false) {
  let el = document.getElementById('app-toast');
  if (!el) { el = document.createElement('div'); el.id='app-toast'; el.style.cssText='position:fixed;bottom:24px;right:24px;background:#1a3a5c;color:#c9d1d9;border:1px solid #2a5a8a;border-radius:8px;padding:10px 18px;font-size:13px;z-index:99999;transition:opacity .3s;'; document.body.appendChild(el); }
  if (isError) el.style.borderColor='#f85149';
  else el.style.borderColor='#2a5a8a';
  el.textContent = msg;
  el.style.opacity='1';
  clearTimeout(el._t);
  el._t = setTimeout(()=>el.style.opacity='0', 3000);
}

// ── Resizable sidebar ──────────────────────────────────────────────────────────
(function(){
  const handle=document.getElementById('resize-handle');
  const sidebar=document.getElementById('sidebar');
  let dragging=false,startX=0,startW=0;
  handle.addEventListener('mousedown',e=>{
    dragging=true;startX=e.clientX;startW=sidebar.offsetWidth;
    handle.classList.add('dragging');
    document.body.style.cursor='ew-resize';
    document.body.style.userSelect='none';
    e.preventDefault();
  });
  document.addEventListener('mousemove',e=>{
    if(!dragging)return;
    const w=Math.max(160,Math.min(520,startW+(e.clientX-startX)));
    sidebar.style.width=w+'px';
  });
  document.addEventListener('mouseup',()=>{
    if(!dragging)return;
    dragging=false;handle.classList.remove('dragging');
    document.body.style.cursor='';document.body.style.userSelect='';
  });
})();

// ── Modal overlay close on background click ────────────────────────────────────
document.querySelectorAll('.modal-overlay').forEach(m=>
  m.addEventListener('click',e=>{if(e.target===m)closeModals();}));

// ── Global Escape key closes open modals ────────────────────────────────────────
document.addEventListener('keydown', e => {
  if (e.key !== 'Escape') return;
  // Don't interfere with people dropdown Escape (handled in compose.js)
  const openDropdown = document.querySelector('.people-dropdown.open');
  if (openDropdown) return;
  const openModal = document.querySelector('.modal-overlay.open');
  if (openModal) { e.preventDefault(); closeModals(); }
});

init();
