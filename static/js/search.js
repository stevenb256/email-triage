// ── Search ─────────────────────────────────────────────────────────────────────
let searchTimeout = null;

async function doSearch(q) {
  q = (q||'').trim();
  if (q.length < 2) return;
  // Hide all right-pane content, show search pane
  ['empty-pane','thread-detail','triage-pane','mailbox-pane','calendar-pane']
    .forEach(id => document.getElementById(id).style.display = 'none');
  const pane = document.getElementById('search-pane');
  pane.style.display = 'flex';
  pane.style.flexDirection = 'column';
  document.getElementById('search-hdr').textContent = `Searching for "${q}"…`;
  document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">Searching…</div>';
  const r = await fetch(`/api/search?q=${encodeURIComponent(q)}`).then(r=>r.json()).catch(()=>null);
  if (!r) { document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">Error</div>'; return; }
  document.getElementById('search-hdr').textContent = `${r.count} result${r.count!==1?'s':''} for "${q}"`;
  if (!r.results.length) { document.getElementById('search-results').innerHTML = '<div class="mailbox-empty">No results</div>'; return; }
  const resultsEl = document.getElementById('search-results');
  resultsEl.innerHTML = r.results.map((e,i) => `
    <div class="search-row" data-search-idx="${i}">
      <div class="search-row-body">
        <div class="search-row-subj">${esc(e.subject||'(No subject)')}${e.score?` <span class="search-score">${Math.round(e.score*100)}%</span>`:''}</div>
        <div class="search-row-meta">${avatarHTML(e.from_name||e.from_address||'', e.from_address||'', 16, 'margin-right:4px;vertical-align:middle;display:inline-flex')} ${esc(e.from_name||e.from_address||'')} · ${esc(fmtDate(e.received_date_time||''))}</div>
        <div class="search-row-preview">${esc(decodeEntities(e.body_preview||''))}</div>
      </div>
      <div class="search-row-folder">${esc(e.folder||'Unknown')}</div>
    </div>`).join('');
  // Load profile images for search result senders
  const searchEmails = [...new Set(r.results.map(e=>e.from_address).filter(Boolean))];
  if (searchEmails.length) loadProfileImages(searchEmails);
  // Delegated click — avoids escaping issues with special chars in conversation keys
  window._searchResults = r.results;
  resultsEl.onclick = (ev) => {
    const row = ev.target.closest('[data-search-idx]');
    if (!row) return;
    const e = window._searchResults[+row.dataset.searchIdx];
    if (e) openSearchResult(e.conversation_key, e.folder||'', e.id);
  };
}

function clearSearch() {
  document.getElementById('search-pane').style.display = 'none';
}

async function openSearchResult(convKey, folder, emailId) {
  document.getElementById('search-pane').style.display = 'none';
  const inInbox = (folder||'').toLowerCase() === 'inbox';
  if (inInbox && state.threadMap[convKey]) {
    switchTab('email');
    selectThread(convKey);
  } else {
    switchTab('mailbox');
    // Small delay to let mailbox init
    await initMailbox();
    // Pre-highlight the folder and load the thread
    const folderEl = document.querySelector(`.folder-item[data-folder="${CSS.escape(folder)}"]`);
    await openMailboxThread(convKey, folder);
  }
}
