// ── Mailbox keyboard navigation ─────────────────────────────────────────────────
let _mboxFocusIdx = -1;

function _mboxGetRows() {
  return [...document.querySelectorAll('#mailbox-list .mbox-row')];
}

function _mboxSetFocus(idx, scroll=true) {
  document.querySelectorAll('.mbox-row.focused').forEach(r=>r.classList.remove('focused'));
  const rows = _mboxGetRows();
  if (idx < 0 || idx >= rows.length) { _mboxFocusIdx = -1; return; }
  _mboxFocusIdx = idx;
  rows[idx].classList.add('focused');
  if (scroll) rows[idx].scrollIntoView({block:'nearest', behavior:'smooth'});
}

function _mboxRegisterKeys() {
  document.removeEventListener('keydown', _mboxKeydown);
  document.addEventListener('keydown', _mboxKeydown);
}

function _mboxUnregisterKeys() {
  document.removeEventListener('keydown', _mboxKeydown);
}

function _mboxKeydown(e) {
  // Ignore when typing in inputs or a modal is open
  if (['INPUT','TEXTAREA','SELECT'].includes(e.target.tagName)) return;
  if (document.querySelector('.modal-overlay.open')) return;

  const listEl = document.getElementById('mailbox-pane');
  const threadEl = document.getElementById('thread-detail');
  const inList = listEl && listEl.style.display !== 'none';
  const inThread = threadEl && threadEl.style.display !== 'none' && threadEl.dataset.loaded;

  if (!inList && !inThread) return;

  const rows = _mboxGetRows();

  if (inList) {
    switch(e.key) {
      case 'ArrowDown': case 'j':
        e.preventDefault();
        _mboxSetFocus(_mboxFocusIdx < 0 ? 0 : Math.min(_mboxFocusIdx + 1, rows.length - 1));
        return;
      case 'ArrowUp': case 'k':
        e.preventDefault();
        _mboxSetFocus(_mboxFocusIdx <= 0 ? 0 : _mboxFocusIdx - 1);
        return;
      case 'Enter': case 'ArrowRight': {
        e.preventDefault();
        const row = rows[_mboxFocusIdx];
        if (row) openMailboxThread(row.dataset.key, row.dataset.folder || mailboxCurrentFolder);
        return;
      }
      case 'r': {
        const row = rows[_mboxFocusIdx];
        if (row) mboxQuickReply(row.dataset.key, row.dataset.folder || mailboxCurrentFolder);
        return;
      }
      case 'd': {
        e.preventDefault();
        const row = rows[_mboxFocusIdx];
        if (row) {
          const nextIdx = Math.min(_mboxFocusIdx, rows.length - 2);
          mboxQuickDelete(row.dataset.key).then(()=>setTimeout(()=>_mboxSetFocus(nextIdx), 50));
        }
        return;
      }
      case 'f': {
        const row = rows[_mboxFocusIdx];
        if (row) {
          const thread = state.threadMap[row.dataset.key];
          if (thread) openFile(encodeThread(thread));
        }
        return;
      }
      case 'u': {
        const row = rows[_mboxFocusIdx];
        if (row) _mboxToggleRead(row.dataset.key);
        return;
      }
    }
  }

  if (inThread) {
    switch(e.key) {
      case 'Escape': case 'ArrowLeft':
        e.preventDefault();
        backToMailboxList();
        return;
      case 'ArrowDown': case 'j': {
        e.preventDefault();
        const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
        if (idx < rows.length - 1) {
          _mboxFocusIdx = idx + 1;
          openMailboxThread(rows[idx+1].dataset.key, rows[idx+1].dataset.folder || mailboxCurrentFolder);
        }
        return;
      }
      case 'ArrowUp': case 'k': {
        e.preventDefault();
        const idx = rows.findIndex(r=>r.dataset.key===state.selectedKey);
        if (idx > 0) {
          _mboxFocusIdx = idx - 1;
          openMailboxThread(rows[idx-1].dataset.key, rows[idx-1].dataset.folder || mailboxCurrentFolder);
        }
        return;
      }
      case 'r': {
        const thread = state.threadMap[state.selectedKey];
        if (thread) openReply(encodeThread(thread));
        return;
      }
      case 'd': {
        e.preventDefault();
        if (state.selectedKey) mboxQuickDelete(state.selectedKey);
        return;
      }
      case 'f': {
        const thread = state.threadMap[state.selectedKey];
        if (thread) openFile(encodeThread(thread));
        return;
      }
    }
  }
}
