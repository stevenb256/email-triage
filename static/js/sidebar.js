// ── Sidebar ────────────────────────────────────────────────────────────────────
function renderSidebar() {
  // Update triage count badge
  const countEl = document.getElementById('triage-count');
  if (countEl) {
    const n = Object.keys(state.threadMap).length;
    countEl.textContent = n > 0 ? n : '';
  }
}
function _threadItemHTML(t) {
  const active = t.conversationKey===state.selectedKey;
  const urgCls = t.urgency==='high'?' urg-high':'';
  const unreadCls = t.hasUnread?' unread':'';
  const flaggedCls = t.isFlagged?' flagged':'';
  return `<div class="thread-item${active?' active':''}${urgCls}${unreadCls}${flaggedCls}" data-key="${esc(t.conversationKey)}" onclick="selectThread(this.getAttribute('data-key'))">
    ${t.hasUnread?'<div class="ti-dot"></div>':'<div class="ti-dot-empty"></div>'}
    <div class="ti-body">
      <div class="ti-subj">${esc(t.subject||'(No subject)')}</div>
      <div class="ti-meta"><span>${esc((t.participants||[])[0]||'')}</span><span>${esc(fmtDate(t.latestReceived||''))}</span></div>
    </div>
  </div>`;
}
function toggleGroup(topic) {
  state.collapsedTopics.has(topic)?state.collapsedTopics.delete(topic):state.collapsedTopics.add(topic);
  renderSidebar();
}

// ── Mailbox ────────────────────────────────────────────────────────────────────
let mailboxFolderLoaded = false;
async function initMailbox() {
  if (mailboxFolderLoaded) return;
  mailboxFolderLoaded = true;
  const r = await fetch('/api/mailbox/folders').then(r=>r.json()).catch(()=>null);
  if (!r) return;
  const tree = document.getElementById('folder-tree');
  const fItem = (f, path) => {
    const cnt = f.count ? `<span class="folder-item-count">${f.count.toLocaleString()}</span>` : '';
    return `<div class="folder-item" data-folder="${esc(path)}" onclick="selectMailboxFolder('${esc(path)}',this)"><span>${f.icon}</span><span class="folder-item-name">${esc(f.name)}</span>${cnt}</div>`;
  };
  const triageCount = Object.keys(state.threadMap).length;
  const triageItem = `<div class="folder-item" id="nav-triage" onclick="openTriageSheet()"><span>📋</span><span class="folder-item-name">Triage</span>${triageCount > 0 ? `<span class="folder-item-count" id="triage-count">${triageCount}</span>` : `<span class="folder-item-count" id="triage-count"></span>`}</div>`;
  const calItem = `<div class="folder-item" id="nav-calendar" onclick="switchTab('calendar')"><span>📅</span><span class="folder-item-name">Calendar</span></div>`;

  const folderHtml = r.folders.map(f => {
    if (f.children && f.children.length) {
      return `<div>
        <div class="folder-group-hdr" onclick="this.classList.toggle('open');this.nextElementSibling.classList.toggle('open')">
          <span>${f.icon}</span><span class="folder-item-name">${esc(f.name)}</span>
          <span class="folder-group-chevron">▾</span>
        </div>
        <div class="folder-group-children">
          ${f.children.map(c=>fItem(c, c.path||c.name)).join('')}
        </div>
      </div>`;
    }
    return fItem(f, f.name);
  });

  // Insert calendar after Archive, or after the last non-system folder
  const archiveIdx = r.folders.findIndex(f => f.name === 'Archive');
  const calIdx = archiveIdx >= 0 ? archiveIdx + 1 : r.folders.findIndex(f => f.name === 'Deleted Items');
  if (calIdx >= 0) {
    folderHtml.splice(calIdx, 0, calItem);
  } else {
    folderHtml.push(calItem);
  }

  tree.innerHTML = triageItem + folderHtml.join('');
  loadTopContacts();
}


// ── Today calendar widget ──────────────────────────────────────────────────────
let _todayCalOffset = 0; // 0 = today, -1 = yesterday, +1 = tomorrow, etc.

function todayCalMove(delta) {
  _todayCalOffset += delta;
  renderTodayCal();
}

async function renderTodayCal() {
  const base = new Date();
  base.setDate(base.getDate() + _todayCalOffset);
  const start = new Date(base); start.setHours(0,0,0,0);
  const end   = new Date(base); end.setHours(23,59,59,0);
  const startISO = start.toISOString().slice(0,19);
  const endISO   = end.toISOString().slice(0,19);

  // Update header label
  const dateEl = document.getElementById('today-cal-date');
  if (dateEl) {
    if (_todayCalOffset === 0) {
      const now = new Date();
      const dow  = now.toLocaleDateString([],{weekday:'short'});
      const mo   = now.toLocaleDateString([],{month:'short'});
      const day  = now.getDate();
      const time = now.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
      dateEl.textContent = `${dow} ${mo} ${day} · ${time}`;
    } else {
      dateEl.textContent = base.toLocaleDateString([],{weekday:'short',month:'short',day:'numeric'});
    }
  }

  let events = [];
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(startISO)}&end=${encodeURIComponent(endISO)}`);
    const d = await r.json();
    events = (d.events||[]).filter(ev => {
      const st = ev.start_time||'';
      return st.length > 10 && !/T00:00:00/.test(st);
    });
  } catch(e) {}
  const list = document.getElementById('today-cal-list');
  if (!list) return;
  if (!events.length) { list.innerHTML = '<div class="today-cal-empty">No meetings</div>'; return; }
  function evDotColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length][1];
  }
  const now = new Date();
  list.innerHTML = events.map(ev => {
    const st = new Date(ev.start_time);
    const timeStr = st.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const color = evDotColor(ev.subject||'');
    const past = _todayCalOffset < 0 || (_todayCalOffset === 0 && st < now);
    return `<div class="today-ev" style="${past?'opacity:.45':''}">
      <div class="today-ev-dot" style="background:${color}"></div>
      <span class="today-ev-time">${timeStr}</span>
      <span class="today-ev-title" title="${esc(ev.subject)}">${esc(ev.subject||'(No title)')}</span>
    </div>`;
  }).join('');
  updateWeekHours();
}

async function updateWeekHours() {
  const today = new Date();
  const dow = today.getDay();
  const mon = new Date(today); mon.setDate(today.getDate() - ((dow+6)%7)); mon.setHours(0,0,0,0);
  const sat = new Date(mon); sat.setDate(mon.getDate()+5);
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(mon.toISOString().slice(0,19))}&end=${encodeURIComponent(sat.toISOString().slice(0,19))}`);
    const d = await r.json();
    const evs = (d.events||[]).filter(ev => {
      const st=ev.start_time||'';
      if (st.length<=10 || /T00:00:00/.test(st)) return false;
      if (/NO\s*MTGS/i.test(ev.subject||'')) return false;
      return true;
    });
    let mins = 0;
    evs.forEach(ev => {
      const st=new Date(ev.start_time), et=new Date(ev.end_time||ev.start_time);
      mins += Math.max(0, (et-st)/60000);
    });
    const hrs = mins/60;
    const hrsStr = Number.isInteger(hrs) ? hrs.toString() : hrs.toFixed(1);
    const el = document.getElementById('week-hours-line');
    if (el) el.textContent = `${hrsStr}h in meetings this week`;
  } catch(e) {}
}

let mailboxCurrentFolder = null;
async function selectMailboxFolder(folder, el) {
  mailboxCurrentFolder = folder;
  // Always switch to mailbox tab — ensures triage/calendar/other panes are hidden
  if (activeTab !== 'mailbox' || state.triageView) switchTab('mailbox');
  document.querySelectorAll('.folder-item.active').forEach(e=>e.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('mailbox-folder-name').textContent = folder.split('/').pop();
  document.getElementById('mailbox-folder-count').textContent = '';
  document.getElementById('mailbox-list').innerHTML = '<div class="mailbox-empty">Loading…</div>';
  // Ensure mailbox pane is shown (user might have drilled into a thread)
  document.getElementById('mailbox-pane').style.display = 'flex';
  document.getElementById('thread-detail').style.display = 'none';
  const r = await fetch(`/api/mailbox/folder?folder=${encodeURIComponent(folder)}`).then(r=>r.json()).catch(()=>null);
  if (!r) { document.getElementById('mailbox-list').innerHTML = '<div class="mailbox-empty">Error loading folder</div>'; return; }
  document.getElementById('mailbox-folder-count').textContent = `${r.total} thread${r.total!==1?'s':''}`;
  const list = document.getElementById('mailbox-list');
  if (!r.threads.length) { list.innerHTML = '<div class="mailbox-empty">No messages</div>'; return; }
  list.innerHTML = r.threads.map(t => {
    const ck = esc(t.conversationKey);
    const fl = esc(folder);
    return `<div class="mbox-row${!t.isRead?' unread':''}" data-key="${ck}" data-folder="${fl}"
        onclick="openMailboxThread(this.dataset.key, this.dataset.folder)">
      ${!t.isRead ? '<div class="mbox-dot"></div>' : '<div class="mbox-dot-empty"></div>'}
      <div class="mbox-body">
        <div class="mbox-subj">${esc(t.subject)}</div>
        <div class="mbox-meta">
          <span class="mbox-from">${esc(t.fromName||t.fromAddress)}</span>
          <span class="mbox-date">${esc(fmtDate(t.date))}</span>
        </div>
        <div class="mbox-preview">${esc(decodeEntities(t.preview))}</div>
      </div>
      ${t.messageCount>1?`<div class="mbox-cnt">${t.messageCount}</div>`:''}
      <div class="mbox-actions" onclick="event.stopPropagation()">
        <button class="mbox-act-btn mbox-act-reply" onclick="mboxQuickReply(this.closest('.mbox-row').dataset.key, this.closest('.mbox-row').dataset.folder)">↩ Reply</button>
        <button class="mbox-act-btn mbox-act-del" onclick="mboxQuickDelete(this.closest('.mbox-row').dataset.key)">🗑</button>
      </div>
    </div>`;
  }).join('');
  _mboxFocusIdx = -1;
  _mboxRegisterKeys();
}

// ── Top Collaborators widget ───────────────────────────────────────────────────
async function loadTopContacts() {
  const list = document.getElementById('top-contacts-list');
  if (!list) return;
  const r = await fetch('/api/top_contacts?n=10').then(r=>r.json()).catch(()=>null);
  if (!r || !r.contacts || !r.contacts.length) {
    list.innerHTML = '<div class="top-contacts-loading">No data yet</div>';
    return;
  }
  list.innerHTML = r.contacts.map(c => {
    const name = c.name || c.email;
    return `<div class="top-contact-row" title="${esc(c.email)}">
      <span class="top-contact-av" style="background:${avColor(name)}" data-av-email="${esc(c.email.toLowerCase())}">${initials(name)}</span>
      <span class="top-contact-name">${esc(name)}</span>
      <span class="top-contact-freq">${c.frequency}</span>
    </div>`;
  }).join('');
  // Load profile images for top contacts
  const contactEmails = r.contacts.map(c=>c.email).filter(Boolean);
  if (contactEmails.length) loadProfileImages(contactEmails);
}
