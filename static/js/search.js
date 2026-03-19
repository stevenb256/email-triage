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
  document.getElementById('search-results').innerHTML = r.results.map(e => `
    <div class="search-row" onclick="openSearchResult('${esc(e.conversation_key)}','${esc(e.folder||'')}','${esc(e.id)}')">
      <div class="search-row-body">
        <div class="search-row-subj">${esc(e.subject||'(No subject)')}</div>
        <div class="search-row-meta">${esc(e.from_name||e.from_address||'')} · ${esc(fmtDate((e.received_date_time||'').slice(0,19)))}</div>
        <div class="search-row-preview">${esc(decodeEntities(e.body_preview||''))}</div>
      </div>
      <div class="search-row-folder">${esc(e.folder||'Unknown')}</div>
    </div>`).join('');
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
