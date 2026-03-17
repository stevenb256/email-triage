// ── Compose ────────────────────────────────────────────────────────────────────
let _composeState = {to:[], cc:[]};
let _pdTimers = {};

function openCompose() {
  _composeState = {to:[], cc:[]};
  document.getElementById('compose-to-field').innerHTML='';
  document.getElementById('compose-cc-field').innerHTML='';
  document.getElementById('compose-to-input').value='';
  document.getElementById('compose-cc-input').value='';
  document.getElementById('compose-subject').value='';
  document.getElementById('compose-body').value='';
  document.getElementById('compose-modal').classList.add('open');
  setTimeout(()=>document.getElementById('compose-to-input').focus(), 50);
}

function focusComposeInput(field) {
  document.getElementById(`compose-${field}-input`).focus();
}

function _renderComposeTags(field) {
  const el = document.getElementById(`compose-${field}-field`);
  el.innerHTML = _composeState[field].map(r=>
    `<span class="recip-tag" data-addr="${esc(r.address)}">${esc(r.name||r.address)}<span class="rm" data-rm-field="${field}" data-rm-addr="${esc(r.address)}">×</span></span>`
  ).join('');
  el.querySelectorAll('.rm').forEach(btn=>{
    btn.addEventListener('click', ()=>{ _composeState[btn.dataset.rmField]=_composeState[btn.dataset.rmField].filter(r=>r.address!==btn.dataset.rmAddr); _renderComposeTags(btn.dataset.rmField); });
  });
}

function _addComposeRecip(field, person) {
  if (!person.address) return;
  if (!_composeState[field].find(r=>r.address.toLowerCase()===person.address.toLowerCase()))
    _composeState[field].push(person);
  _renderComposeTags(field);
  const inp = document.getElementById(`compose-${field}-input`);
  inp.value = '';
  document.getElementById(`pd-${field}`).classList.remove('open');
}

async function peopleSuggest(inp, field) {
  const q = inp.value.trim();
  if (!q) { document.getElementById(`pd-${field}`).classList.remove('open'); return; }
  clearTimeout(_pdTimers[field]);
  _pdTimers[field] = setTimeout(async ()=>{
    let people;
    if (_pdCache) {
      people = _pdCache.filter(p=>(p.name||'').toLowerCase().includes(q.toLowerCase())||(p.address||'').toLowerCase().includes(q.toLowerCase())).slice(0,12);
    } else {
      const d = await fetch(`/api/people?q=${encodeURIComponent(q)}`).then(r=>r.json()).catch(()=>null);
      people = d ? d.people : [];
    }
    const dd = document.getElementById(`pd-${field}`);
    if (!people.length) { dd.classList.remove('open'); return; }
    dd.innerHTML = people.map((p,i)=>
      `<div class="pd-item" data-i="${i}"><span class="pd-item-name">${esc(p.name||p.address)}</span><span class="pd-item-addr">${esc(p.address)}</span></div>`
    ).join('');
    dd._people = people;
    dd.querySelectorAll('.pd-item').forEach((item,i)=>{
      item.addEventListener('mousedown', e=>{ e.preventDefault(); _addComposeRecip(field, dd._people[i]); });
    });
    dd.classList.add('open');
  }, 150);
}

function peopleSuggestKey(e, field) {
  const dd = document.getElementById(`pd-${field}`);
  const inp = document.getElementById(`compose-${field}-input`);
  if (e.key==='Enter' || e.key===',') {
    e.preventDefault();
    // If dropdown has active item, use it; otherwise treat as raw email
    const active = dd.querySelector('.pd-item.active');
    if (active && dd._people) {
      const i = parseInt(active.dataset.i);
      _addComposeRecip(field, dd._people[i]);
    } else if (inp.value.includes('@')) {
      _addComposeRecip(field, {name:'', address:inp.value.trim()});
    }
    return;
  }
  if (e.key==='ArrowDown'||e.key==='ArrowUp') {
    e.preventDefault();
    const items = [...dd.querySelectorAll('.pd-item')];
    if (!items.length) return;
    const cur = dd.querySelector('.pd-item.active');
    const idx = cur ? items.indexOf(cur) : -1;
    if (cur) cur.classList.remove('active');
    const next = items[(idx + (e.key==='ArrowDown'?1:-1) + items.length) % items.length];
    next.classList.add('active');
    return;
  }
  if (e.key==='Escape') { dd.classList.remove('open'); }
  if (e.key==='Backspace' && !inp.value && _composeState[field].length) {
    _composeState[field].pop();
    _renderComposeTags(field);
  }
}

async function sendNewMessage() {
  const to = _composeState.to.map(r=>r.address).filter(Boolean);
  const cc = _composeState.cc.map(r=>r.address).filter(Boolean);
  const subject = document.getElementById('compose-subject').value.trim();
  const body = document.getElementById('compose-body').value.trim();
  if (!to.length) { alert('Please add at least one recipient.'); return; }
  if (!subject) { alert('Please add a subject.'); return; }
  closeModals();
  const d = await fetch('/api/send_new', {method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({to, cc, subject, body})}).then(r=>r.json()).catch(()=>null);
  if (d && d.ok) {
    showToast('Message sent');
  } else {
    showToast('Send failed: '+(d&&d.error||'unknown error'), true);
  }
}
