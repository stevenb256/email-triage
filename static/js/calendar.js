// ── Calendar ───────────────────────────────────────────────────────────────────
const CAL_COLORS = [
  ['#1f6feb','#58a6ff'],['#388bfd22','#79c0ff'],['#1a7f37','#3fb950'],
  ['#6e40c9','#bc8cff'],['#b45309','#d18616'],['#0e7490','#06b6d4'],
  ['#9a1c1c','#f85149'],
];
let calWeekOffset = 0;
let calDayOffset  = 0;   // days from today (day view)
let calViewMode   = 'day';
let calEvents     = [];
let calPrepCache  = {};  // eventId -> {headsup, topics}

function calGetWeekStart(offset) {
  const d = new Date();
  const day = d.getDay();
  const mon = new Date(d);
  mon.setDate(d.getDate() - ((day+6)%7) + offset*7);
  mon.setHours(0,0,0,0);
  return mon;
}

function calSetView(mode) {
  calViewMode = mode;
  document.getElementById('cal-view-day')?.classList.toggle('active', mode==='day');
  document.getElementById('cal-view-week')?.classList.toggle('active', mode==='week');
  renderCalendar();
}

function calMove(dir) {
  if (calViewMode==='day') calDayOffset += dir; else calWeekOffset += dir;
  renderCalendar();
}

function calGoToday() { calWeekOffset = 0; calDayOffset = 0; renderCalendar(); }

async function renderCalendar() {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const DOW_LONG = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  let startISO, endISO, title;

  if (calViewMode === 'day') {
    const base = new Date(); base.setHours(0,0,0,0);
    const day = new Date(base); day.setDate(base.getDate() + calDayOffset);
    const next = new Date(day); next.setDate(day.getDate() + 1);
    startISO = day.toISOString().slice(0,19);
    endISO   = next.toISOString().slice(0,19);
    title = `${DOW_LONG[day.getDay()]}, ${months[day.getMonth()]} ${day.getDate()}, ${day.getFullYear()}`;
  } else {
    const weekStart = calGetWeekStart(calWeekOffset);
    const weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate()+7);
    startISO = weekStart.toISOString().slice(0,19);
    endISO   = weekEnd.toISOString().slice(0,19);
    const s = weekStart, e = new Date(weekEnd); e.setDate(e.getDate()-1);
    title = s.getMonth()===e.getMonth()
      ? `${months[s.getMonth()]} ${s.getDate()} – ${e.getDate()}, ${s.getFullYear()}`
      : `${months[s.getMonth()]} ${s.getDate()} – ${months[e.getMonth()]} ${e.getDate()}, ${s.getFullYear()}`;
  }
  document.getElementById('cal-title').textContent = title;

  document.getElementById('cal-loading').style.display = 'flex';
  document.getElementById('cal-scroll-wrap').style.display = 'none';
  try {
    const r = await fetch(`/api/calendar?start=${encodeURIComponent(startISO)}&end=${encodeURIComponent(endISO)}`);
    const d = await r.json();
    calEvents = d.events || [];
  } catch(e) { calEvents = []; }
  document.getElementById('cal-loading').style.display = 'none';
  document.getElementById('cal-scroll-wrap').style.display = '';

  const grid = document.getElementById('cal-grid');
  if (calViewMode === 'day') {
    grid.style.gridTemplateColumns = '52px 1fr';
    buildDayView();
  } else {
    grid.style.gridTemplateColumns = '52px repeat(7,minmax(0,1fr))';
    buildCalGrid(calGetWeekStart(calWeekOffset));
  }
}

function buildCalGrid(weekStart) {
  const grid = document.getElementById('cal-grid');
  const today = new Date(); today.setHours(0,0,0,0);
  const days = Array.from({length:7}, (_,i) => { const d=new Date(weekStart); d.setDate(weekStart.getDate()+i); return d; });
  const SLOT_H = 24; // px per 30-min slot
  const HOUR_START = 7, HOUR_END = 21; // visible hours
  const SLOTS = (HOUR_END - HOUR_START) * 2;

  // Separate all-day events
  const allDayEvs = calEvents.filter(ev => {
    const st = ev.start_time || '';
    return st.length <= 10 || /T00:00:00/.test(st) && /T00:00:00/.test(ev.end_time||'');
  });
  const timedEvs = calEvents.filter(ev => !allDayEvs.includes(ev));

  // Assign colors per event title hash
  function evColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length];
  }

  let html = '';

  // Corner + day headers
  html += `<div class="cal-hdr-corner"></div>`;
  days.forEach((d,i) => {
    const isToday = d.getTime()===today.getTime();
    const dow = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'][i];
    html += `<div class="cal-hdr-cell${isToday?' today':''}">
      <div class="cal-hdr-dow">${dow}</div>
      <div class="cal-hdr-day">${d.getDate()}</div>
    </div>`;
  });

  // All-day row
  html += `<div style="font-size:9px;color:#5ba4cf;text-align:right;padding:2px 4px 2px 0;border-right:1px solid #1a3252;border-bottom:2px solid #1a3252;">all day</div>`;
  days.forEach((d,i) => {
    const isToday = d.getTime()===today.getTime();
    const dayStr = d.toISOString().slice(0,10);
    const evs = allDayEvs.filter(ev => (ev.start_time||'').startsWith(dayStr));
    html += `<div class="cal-all-day-cell${isToday?' today-col':''}">`;
    evs.forEach(ev => { html += `<div class="cal-all-day-event" title="${esc(ev.subject)}">${esc(ev.subject)}</div>`; });
    html += `</div>`;
  });

  // Time slots
  for (let slot=0; slot<SLOTS; slot++) {
    const totalMins = (HOUR_START * 60) + slot * 30;
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const isHour = m===0;
    if (isHour) {
      const label = h===12?'12pm':h>12?`${h-12}pm`:`${h}am`;
      html += `<div class="cal-time-label" style="height:${SLOT_H}px;${isHour?'':'border-top:none'}">${label}</div>`;
    } else {
      html += `<div class="cal-time-label" style="height:${SLOT_H}px;border-top:none"></div>`;
    }
    days.forEach((d,di) => {
      const isToday = d.getTime()===today.getTime();
      html += `<div class="cal-cell${isToday?' today-col':''}${isHour?' hour-start':''}" style="height:${SLOT_H}px"></div>`;
    });
  }

  grid.innerHTML = html;

  // Position timed events as absolute overlays
  // We need to position them inside the correct cell after render
  // Use a post-render approach: collect cells by [day][slot]
  const cells = grid.querySelectorAll('.cal-cell');
  const cellMap = {}; // "dayIdx-slot" -> cell el
  let ci = 0;
  for (let slot=0; slot<SLOTS; slot++) {
    for (let di=0; di<7; di++) {
      cellMap[`${di}-${slot}`] = cells[ci++];
    }
  }

  // For each timed event, find its day col and slot range
  timedEvs.forEach((ev, idx) => {
    const st = new Date(ev.start_time);
    const et = new Date(ev.end_time || ev.start_time);
    const evDay = new Date(st); evDay.setHours(0,0,0,0);
    const di = days.findIndex(d => d.getTime()===evDay.getTime());
    if (di < 0) return;

    const startMins = st.getHours()*60 + st.getMinutes();
    const endMins = et.getHours()*60 + et.getMinutes() || startMins + 30;
    const startSlot = Math.max(0, Math.floor((startMins - HOUR_START*60)/30));
    const durationSlots = Math.max(1, Math.ceil((endMins - startMins)/30));

    const anchorCell = cellMap[`${di}-${startSlot}`];
    if (!anchorCell) return;

    const topOffset = ((startMins - HOUR_START*60) % 30) / 30 * SLOT_H;
    const height = Math.max(SLOT_H-2, durationSlots * SLOT_H - 2);
    const [bg, fg] = evColor(ev.subject||'');
    const timeStr = st.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});

    const el = document.createElement('div');
    el.className = 'cal-event';
    el.style.cssText = `top:${topOffset}px;height:${height}px;background:${bg}33;border:1px solid ${bg}88;color:${fg};`;
    el.title = `${ev.subject}\n${timeStr}${ev.location?' · '+ev.location:''}`;
    const timeEl = durationSlots > 1 ? `<div class="cal-event-time">${timeStr}</div>` : '';
    el.innerHTML = `<div class="cal-event-title">${esc(ev.subject||'(No title)')}</div>${timeEl}`;
    anchorCell.style.position = 'relative';
    anchorCell.appendChild(el);
  });

  // Scroll to 8am (1hr from HOUR_START=7 = 2 slots)
  document.getElementById('cal-scroll-wrap').scrollTop = 2 * SLOT_H;
}

function buildDayView() {
  const grid = document.getElementById('cal-grid');
  const base = new Date(); base.setHours(0,0,0,0);
  const day  = new Date(base); day.setDate(base.getDate() + calDayOffset);
  const today = new Date(); today.setHours(0,0,0,0);
  const isToday = day.getTime() === today.getTime();
  const SLOT_H = 32;
  const HOUR_START = 7, HOUR_END = 21;
  const SLOTS = (HOUR_END - HOUR_START) * 2;
  const dayStr = day.toISOString().slice(0,10);

  function evColor(subj) {
    let h=0; for(const c of subj) h=(h*31+c.charCodeAt(0))&0xffff;
    return CAL_COLORS[h % CAL_COLORS.length];
  }

  const allDayEvs = calEvents.filter(ev => {
    const st = ev.start_time||'';
    return st.length<=10 || (/T00:00:00/.test(st) && /T00:00:00/.test(ev.end_time||''));
  });
  const timedEvs = calEvents.filter(ev => !allDayEvs.includes(ev));

  const dow = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][day.getDay()];
  let html = `<div class="cal-hdr-corner"></div>
    <div class="cal-hdr-cell${isToday?' today':''}">
      <div class="cal-hdr-dow">${dow}</div>
      <div class="cal-hdr-day">${day.getDate()}</div>
    </div>
    <div style="font-size:9px;color:#5ba4cf;text-align:right;padding:2px 4px 2px 0;border-right:1px solid #1a3252;border-bottom:2px solid #1a3252;">all day</div>
    <div class="cal-all-day-cell${isToday?' today-col':''}">`;
  allDayEvs.filter(ev=>(ev.start_time||'').startsWith(dayStr)).forEach(ev=>{
    html+=`<div class="cal-all-day-event">${esc(ev.subject)}</div>`;
  });
  html += `</div>`;

  for (let slot=0; slot<SLOTS; slot++) {
    const totalMins = (HOUR_START*60) + slot*30;
    const h = Math.floor(totalMins/60), m = totalMins%60;
    const isHour = m===0;
    const label = isHour ? (h===12?'12pm':h>12?`${h-12}pm`:`${h}am`) : '';
    html += `<div class="cal-time-label" style="height:${SLOT_H}px;${isHour?'':'border-top:none'}">${label}</div>`;
    html += `<div class="cal-cell cal-day-cell${isToday?' today-col':''}${isHour?' hour-start':''}" style="height:${SLOT_H}px"></div>`;
  }
  grid.innerHTML = html;

  const cells = Array.from(grid.querySelectorAll('.cal-day-cell'));

  timedEvs.forEach(ev => {
    const st = new Date(ev.start_time);
    const et = new Date(ev.end_time || ev.start_time);
    const startMins = st.getHours()*60 + st.getMinutes();
    const endMins   = et.getHours()*60 + et.getMinutes() || startMins+30;
    const startSlot = Math.max(0, Math.floor((startMins - HOUR_START*60)/30));
    const durSlots  = Math.max(2, Math.ceil((endMins - startMins)/30));
    const anchor    = cells[startSlot];
    if (!anchor) return;

    const topOffset = ((startMins - HOUR_START*60) % 30) / 30 * SLOT_H;
    const height    = Math.max(SLOT_H*2-2, durSlots * SLOT_H - 2);
    const [bg, fg]  = evColor(ev.subject||'');
    const timeStr   = st.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const endStr    = et.toLocaleTimeString([],{hour:'numeric',minute:'2-digit'});
    const prepId    = 'prep-'+ev.id.replace(/[^a-zA-Z0-9]/g,'_');

    const el = document.createElement('div');
    el.className = 'cal-event cal-day-event';
    el.style.cssText = `top:${topOffset}px;height:${height}px;`+
      `background:${bg}44;border-left:3px solid ${bg};`+
      `border-top:1px solid ${bg}88;border-right:1px solid ${bg}44;`+
      `border-bottom:1px solid ${bg}44;color:${fg};padding:5px 8px;`;
    el.innerHTML = `<div class="cal-day-ev-hdr">
        <div class="cal-day-ev-title">${esc(ev.subject||'(No title)')}</div>
        <div class="cal-day-ev-time">${timeStr}–${endStr}${ev.location?' · '+esc(ev.location):''}</div>
      </div>`;
    anchor.style.position = 'relative';
    anchor.appendChild(el);
  });

  document.getElementById('cal-scroll-wrap').scrollTop = 2 * SLOT_H;
}

async function loadMeetingPrep(ev, prepId) {
  const cacheKey = ev.id;
  if (calPrepCache[cacheKey]) { _renderPrep(prepId, calPrepCache[cacheKey]); return; }
  try {
    const r = await fetch('/api/meeting_prep', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({subject:ev.subject, attendees:ev.attendees,
        start_time:ev.start_time, end_time:ev.end_time, location:ev.location})
    }).then(r=>r.json());
    if (r.ok) { calPrepCache[cacheKey]=r; _renderPrep(prepId, r); }
    else { const el=document.getElementById(prepId); if(el) el.innerHTML=''; }
  } catch(e) { const el=document.getElementById(prepId); if(el) el.innerHTML=''; }
}

function _renderPrep(prepId, prep) {
  const el = document.getElementById(prepId);
  if (!el) return;
  const topics = (prep.topics||[]).map(t=>`<div class="cal-prep-topic">• ${esc(t)}</div>`).join('');
  el.innerHTML = `<div class="cal-prep-headsup">${esc(prep.headsup||'')}</div>`+
    (topics ? `<div class="cal-prep-topics">${topics}</div>` : '');
}
