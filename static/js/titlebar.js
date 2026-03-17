// ── Titlebar sync controls ─────────────────────────────────────────────────────
async function triggerSync() {
  const d=await fetch('/api/sync_now',{method:'POST'}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
}
async function resyncThread() {
  const convKey=state.selectedKey;
  if (!convKey) return;
  const btn=document.getElementById('resync-thread-btn');
  btn.disabled=true; btn.textContent='↺ Resyncing…';
  const d=await fetch('/api/resync_thread',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({conversationKey:convKey})}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
  // Reload the thread view after resync completes
  setTimeout(async()=>{
    btn.disabled=false; btn.textContent='↺ Resync Thread';
    if (state.selectedKey===convKey) await selectThread(convKey);
  }, 1500);
}
async function reanalyzeAll() {
  const btn=document.getElementById('reanalyze-btn');
  btn.disabled=true; btn.textContent='⚙ Re-analyzing…';
  const d=await fetch('/api/reanalyze_all',{method:'POST'}).then(r=>r.json()).catch(()=>null);
  if (d) updateSyncStatus(d.syncStatus);
  setTimeout(()=>{btn.disabled=false;btn.textContent='⚙ Re-analyze';},3000);
}
