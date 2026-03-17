// ── Avatar helpers ─────────────────────────────────────────────────────────────
const AV_COLORS = ['#1f6feb','#1a7f37','#9a1c1c','#7d4e00','#6e3cc1','#b45309','#0284c7','#be185d'];
function avColor(name) {
  let h = 0; for (const c of String(name)) h = (h*31+c.charCodeAt(0))&0xffffffff;
  return AV_COLORS[Math.abs(h)%AV_COLORS.length];
}
function initials(name) {
  const p = String(name||'').trim().split(/\s+/);
  return p.length>=2?(p[0][0]+p[1][0]).toUpperCase():(p[0]||'?').slice(0,2).toUpperCase();
}

function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ── Encode/decode thread ───────────────────────────────────────────────────────
function encodeThread(t) {
  try {
    const j=JSON.stringify({conversationKey:t.conversationKey,latestId:t.latestId,emailIds:t.emailIds,subject:t.subject,messageCount:t.messageCount,suggestedReply:t.suggestedReply,suggestedFolder:t.suggestedFolder});
    return btoa(unescape(encodeURIComponent(j))).replace(/=/g,'');
  } catch(e){return '';}
}
function decodeThread(s) {
  try { return JSON.parse(decodeURIComponent(escape(atob(s.replace(/-/g,'+').replace(/_/g,'/'))))); }
  catch{return {};}
}

// ── Date formatting ────────────────────────────────────────────────────────────
function fmtDate(s) {
  if (!s) return '';
  const d=new Date(s),now=new Date(),diff=now-d;
  if (isNaN(d)) return '';
  if (diff<3600000){const m=Math.round(diff/60000);return m<1?'just now':`${m}m`;}
  if (diff<86400000) return d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
  if (diff<604800000) return d.toLocaleDateString([],{weekday:'short'});
  return d.toLocaleDateString([],{month:'short',day:'numeric'});
}
function fmtUntil(s) {
  if (!s) return '';
  const d=new Date(s),now=new Date();
  if (isNaN(d)) return '';
  const diff=d-now; // ms until event
  if (diff<=0) return 'now';
  const mins=Math.round(diff/60000);
  if (mins<60) return `in ${mins}m`;
  const hrs=Math.floor(diff/3600000);
  const rem=Math.round((diff%3600000)/60000);
  if (hrs<24) return rem>0?`in ${hrs}h ${rem}m`:`in ${hrs}h`;
  const days=Math.floor(diff/86400000);
  const time=d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
  if (days===0) return `today ${time}`;
  if (days===1) return `tomorrow ${time}`;
  const dow=d.toLocaleDateString([],{weekday:'short'});
  return `${dow} ${time}`;
}

// ── Sync status ────────────────────────────────────────────────────────────────
function updateSyncStatus(ss) {
  if (!ss) return;
  const dot=document.getElementById('sync-dot');
  const txt=document.getElementById('sync-txt');
  const wrap=document.getElementById('sync-bar-wrap');
  const bar=document.getElementById('sync-bar');
  if (ss.running) {
    dot.className='sync-dot syncing';
    txt.textContent=ss.progress||'Syncing…';
    if (ss.total>0){const pct=Math.round((ss.done/ss.total)*100);wrap.style.display='block';bar.style.width=Math.max(4,pct)+'%';}
    else wrap.style.display='none';
  } else {
    if(wrap) wrap.style.display='none';
    if (ss.lastError){dot.className='sync-dot error';txt.textContent='Sync error: '+ss.lastError;}
    else if (ss.lastSync){
      dot.className='sync-dot';
      const mins=Math.round((Date.now()-new Date(ss.lastSync))/60000);
      const ago=mins<1?'just now':mins===1?'1 min ago':`${mins} min ago`;
      txt.textContent=`Synced ${ago}`+(ss.threadsUpdated>0?` · ${ss.threadsUpdated} updated`:'');
    } else {dot.className='sync-dot syncing';txt.textContent='Waiting for first sync…';}
  }
}
function _showNewBadge(n) {
  const w=document.getElementById('new-badge-wrap');
  w.innerHTML=`<span class="new-badge">${n} new</span>`;
  setTimeout(()=>{w.innerHTML='';},5000);
}
function updateCounts(emailCount,threadCount) {
  const el=document.getElementById('sidebar-counts');
  if (!el) return;
  el.textContent=emailCount!==null?`${emailCount} emails · ${threadCount} threads`:`${threadCount} threads`;
}

// ── Text helpers ───────────────────────────────────────────────────────────────
function highlightEmails(html) {
  return html.replace(/([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})/g,
    '<span class="eml">$1</span>');
}
function linkify(text) {
  // Split on URLs, linkify those parts, highlight emails in non-URL parts
  const parts=text.split(/(https?:\/\/[^\s<>"')\]]+)/g);
  return parts.map((part,i)=>{
    if(i%2===1){const href=esc(part);return `<a href="${href}" target="_blank" rel="noopener noreferrer" class="link">${href}</a>`;}
    return highlightEmails(esc(part));
  }).join('');
}

// Render AI summary — handles both legacy plain text and new 3-part format
// (parts separated by blank lines: FACTS, OPEN QUESTIONS, NEXT ACTION)
function _renderSummary(summary) {
  if (!summary) return '';
  const parts = summary.split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);
  if (parts.length < 2) {
    // Legacy single-block summary — just render as-is
    return `<span>${esc(summary)}</span>`;
  }
  const labels = ['📋 Facts', '❓ Open Questions / Blockers', '⚡ Your Next Action'];
  return parts.map((p, i) => {
    const label = labels[i] || '';
    // Render bullet lines as a list if they start with •, -, or *
    const lines = p.split('\n').map(l => l.trim()).filter(Boolean);
    const isList = lines.length > 1 && lines.every(l => /^[•\-\*]/.test(l));
    const body = isList
      ? `<ul class="sum-list">${lines.map(l => `<li>${esc(l.replace(/^[•\-\*]\s*/,''))}</li>`).join('')}</ul>`
      : `<span>${esc(p)}</span>`;
    return `<div class="sum-part"><span class="sum-part-lbl">${label}</span>${body}</div>`;
  }).join('');
}

// intent → CSS class suffix
const INTENT_CLS = {
  'Status Update':'status-update','Request':'request','Decision':'decision',
  'Question':'question','Action Item':'action-item','Context':'context',
  'FYI':'fyi','Warning':'warning','Introduction':'introduction','Closing':'closing'
};

const ACTION_REC = {
  reply:  {cls:'action-rec-reply',  icon:'↩', label:'Reply Needed'},
  delete: {cls:'action-rec-delete', icon:'🗑', label:'Safe to Delete'},
  file:   {cls:'action-rec-file',   icon:'📁', label:'File for Reference'},
  read:   {cls:'action-rec-read',   icon:'👁', label:'FYI Only'},
  done:   {cls:'action-rec-done',   icon:'✓',  label:'Resolved'},
};
