// ── Constants ──────────────────────────────────────────────────────────────
const PRESETS_CS = [
  "Tech CS Monitoring",
  "Trust Pilot Review monitoring",
  "Weekly 1on1 meetings with the CS and Routers",
  "Shift scheduling management to maximize coverage",
  "CS and Router Inbox Audits",
  "Monitoring Akute for conversations sent through the portal and rerouting via Intercom to the appropriate teams",
  "Monitoring Akute fax inbox and rerouting faxes to appropriate patient charts",
  "Auditing Knowledge base and Macros for duplicate or outdated information"
];

const PRESETS_OPERATIONS = [
  "Daily operations review",
  "Process improvements implemented",
  "Team coordination and scheduling",
  "Issue resolution and escalations",
  "Documentation updates",
  "Training and onboarding",
  "Cross-team collaboration"
];

const DEF_MEETINGS_CS = [
  "1:1 with Shift Supervisors, Front and Back Office Team Leads, CCOO",
  "CS Team Meeting, CS Operations Sync"
];

const DEF_MEETINGS_OPERATIONS = [
  "Operations standup",
  "Team sync",
  "Leadership update"
];

const DEF_SYNC_CS = [
  "Incarcerated Patients Protocol",
  "New HRT Patients seeing that labs are still included",
  "Addressing HRT sign-up timeframe and labs",
  "Async Form responses from patients",
  "Maryland Lab Coverage",
  "Sending Tickets and Conversations to Shift Supervisors",
  "Tickets versus Conversations",
  "System Outage protocol",
  "Medical Critical Escalations when Shift Supervisors are OOO (During after hour shifts)"
];

const DEF_SYNC_OPERATIONS = [
  "Upcoming initiatives",
  "Resource allocation",
  "Blockers and risks"
];

const LS_KEY = 'fountain_report_v3';
const SLACK_CHAR_LIMIT = 4000;

// ── State ──────────────────────────────────────────────────────────────────
let activeWeekId = '';
let meetings = [];
let bullets = [];       // { text, carryForward }
let syncItems = [];
let presetChecks = {};
let notesForNextWeek = '';
let activeTemplate = 'cs';
let saveTimer = null, statusTimer = null;

function getPresets() { return activeTemplate === 'operations' ? PRESETS_OPERATIONS : PRESETS_CS; }
function getDefMeetings() { return activeTemplate === 'operations' ? DEF_MEETINGS_OPERATIONS : DEF_MEETINGS_CS; }
function getDefSync() { return activeTemplate === 'operations' ? DEF_SYNC_OPERATIONS : DEF_SYNC_CS; }

// ── Week helpers ───────────────────────────────────────────────────────────
function weekId(d) {
  d = d ? new Date(d) : new Date();
  const day = d.getDay();
  const mon = new Date(d); mon.setDate(d.getDate() - ((day + 6) % 7)); mon.setHours(0,0,0,0);
  const fri = new Date(mon); fri.setDate(mon.getDate() + 4);
  const f = x => `${x.getMonth()+1}/${x.getDate()}/${String(x.getFullYear()).slice(-2)}`;
  return f(mon) + '~' + f(fri);
}

function weekDates(d) {
  d = d ? new Date(d) : new Date();
  const day = d.getDay();
  const mon = new Date(d); mon.setDate(d.getDate() - ((day + 6) % 7)); mon.setHours(0,0,0,0);
  const fri = new Date(mon); fri.setDate(mon.getDate() + 4);
  const f  = x => `${x.getMonth()+1}/${x.getDate()}/${String(x.getFullYear()).slice(-2)}`;
  const mf = x => `${String(x.getMonth()+1).padStart(2,'0')}/${String(x.getDate()).padStart(2,'0')}/${x.getFullYear()}`;
  return { short: f(mon)+'-'+f(fri), metric: mf(mon)+'-'+mf(fri), label: f(mon)+' – '+f(fri) };
}

function weekLabel(wid) {
  const parts = wid.split('~');
  return (parts[0] || '') + ' – ' + (parts[1] || '');
}

// ── Storage ────────────────────────────────────────────────────────────────
function getStore() {
  try { return JSON.parse(localStorage.getItem(LS_KEY)) || { weeks:{}, order:[] }; }
  catch { return { weeks:{}, order:[] }; }
}
function saveStore(data) { try { localStorage.setItem(LS_KEY, JSON.stringify(data)); } catch {} }

// ── Init ───────────────────────────────────────────────────────────────────
function init() {
  activeTemplate = localStorage.getItem('ss_report_template') || 'cs';
  const sel = document.getElementById('templateSelect');
  if (sel) sel.value = activeTemplate;

  const dark = localStorage.getItem('ss_report_dark') === '1';
  if (dark) document.body.classList.add('dark');
  updateDarkModeIcon();

  activeWeekId = weekId();
  const dates = weekDates();
  document.getElementById('weekBadge').textContent = 'Week of ' + dates.label;
  loadWeek(activeWeekId);
  showReminderBanner();
  setupKeyboardShortcuts();
}

// ── Load / Save ────────────────────────────────────────────────────────────
function loadWeek(wid) {
  activeWeekId = wid;
  const store = getStore();
  const saved = store.weeks[wid];
  const dates = weekDates();

  if (saved) {
    setF('teamName', saved.teamName || 'CS/Refills and Clarifications/RN/Shift Supervisor');
    setF('dateRange', saved.dateRange || '');
    setF('metricPeriod', saved.metricPeriod || '');
    setF('totalConversations', saved.totalConversations || '');
    setF('medianResponseTime', saved.medianResponseTime || '');
    setF('responseGoal', saved.responseGoal || '2 minutes 30 seconds');
    setF('syncMeeting', saved.syncMeeting || '');
    setF('privateNotes', saved.privateNotes || '');
    setF('notesForNextWeek', saved.notesForNextWeek || '');
    meetings = [...(saved.meetings || getDefMeetings())];
    bullets = (saved.bullets || ['']).map(b => typeof b === 'string' ? {text:b, carryForward:false} : b);
    syncItems = [...(saved.syncItems || getDefSync())];
    presetChecks = {...(saved.presetChecks || {})};
    notesForNextWeek = saved.notesForNextWeek || '';
  } else {
    const carryBullets = getCarryForwardBullets(wid, store);
    const carryNotes = getCarryForwardNotes(wid, store);

    setF('teamName', 'CS/Refills and Clarifications/RN/Shift Supervisor');
    setF('dateRange', wid === weekId() ? dates.short : weekLabel(wid));
    setF('metricPeriod', wid === weekId() ? dates.metric : '');
    setF('totalConversations', '');
    setF('medianResponseTime', '');
    setF('responseGoal', '2 minutes 30 seconds');
    setF('syncMeeting', activeTemplate === 'operations' ? 'Upcoming topics:' : 'Front Office Sync Monthly Meeting on [DATE] to review the following:');
    setF('privateNotes', '');
    setF('notesForNextWeek', carryNotes);
    meetings = [...getDefMeetings()];
    bullets = carryBullets.length > 0 ? carryBullets : [{text:'', carryForward:false}];
    syncItems = [...getDefSync()];
    presetChecks = {};
    notesForNextWeek = carryNotes;
  }

  updateCopyLastWeekVisibility();

  renderAll();
  renderHistory();
  renderSparkline();
  update();
}

function getCarryForwardBullets(currentWid, store) {
  const sortedOrder = [...store.order].sort((a,b) => b.localeCompare(a));
  const prevWid = sortedOrder.find(wid => wid < currentWid);
  if (!prevWid) return [];
  const prevData = store.weeks[prevWid];
  if (!prevData || !prevData.bullets) return [];
  return prevData.bullets
    .filter(b => typeof b === 'object' && b.carryForward && b.text.trim())
    .map(b => ({text: b.text, carryForward: false}));
}

function getCarryForwardNotes(currentWid, store) {
  const sortedOrder = [...store.order].sort((a,b) => b.localeCompare(a));
  const prevWid = sortedOrder.find(wid => wid < currentWid);
  if (!prevWid) return '';
  const prevData = store.weeks[prevWid];
  return (prevData && prevData.notesForNextWeek) ? prevData.notesForNextWeek : '';
}

function saveWeek() {
  const store = getStore();
  const wid = activeWeekId;
  notesForNextWeek = getF('notesForNextWeek');
  store.weeks[wid] = {
    teamName: getF('teamName'), dateRange: getF('dateRange'),
    metricPeriod: getF('metricPeriod'), totalConversations: getF('totalConversations'),
    medianResponseTime: getF('medianResponseTime'), responseGoal: getF('responseGoal'),
    syncMeeting: getF('syncMeeting'), privateNotes: getF('privateNotes'),
    notesForNextWeek, meetings: [...meetings], bullets: bullets.map(b => ({...b})),
    syncItems: [...syncItems], presetChecks: {...presetChecks},
    savedAt: new Date().toISOString()
  };
  if (!store.order.includes(wid)) {
    store.order = [wid, ...store.order].sort((a,b) => b.localeCompare(a));
  }
  saveStore(store);
  showSaved();
  renderHistory();
  renderSparkline();
}

function debounce() {
  clearTimeout(saveTimer);
  setSaveStatus('saving');
  saveTimer = setTimeout(saveWeek, 700);
}

// ── Sparkline ──────────────────────────────────────────────────────────────
function renderSparkline() {
  const store = getStore();
  const sorted = [...store.order].sort((a,b) => a.localeCompare(b));
  const dataPoints = sorted
    .map(wid => ({ wid, val: parseInt(store.weeks[wid]?.totalConversations || '0', 10) }))
    .filter(d => d.val > 0)
    .slice(-6);

  const wrap = document.getElementById('sparkWrap');
  if (dataPoints.length < 2) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';

  const W = 160, H = 48, PAD = 6;
  const vals = dataPoints.map(d => d.val);
  const min = Math.min(...vals), max = Math.max(...vals);
  const range = max - min || 1;

  const xs = dataPoints.map((_, i) => PAD + (i / (dataPoints.length - 1)) * (W - PAD * 2));
  const ys = vals.map(v => H - PAD - ((v - min) / range) * (H - PAD * 2));

  const polyline = xs.map((x, i) => `${x},${ys[i]}`).join(' ');
  const fillPts = `${xs[0]},${H} ` + xs.map((x, i) => `${x},${ys[i]}`).join(' ') + ` ${xs[xs.length-1]},${H}`;

  const svg = document.getElementById('sparkSvg');
  svg.innerHTML = `
    <defs>
      <linearGradient id="sg" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0%" stop-color="#2563eb" stop-opacity="0.3"/>
        <stop offset="100%" stop-color="#2563eb" stop-opacity="0"/>
      </linearGradient>
    </defs>
    <polygon points="${fillPts}" fill="url(#sg)"/>
    <polyline points="${polyline}" fill="none" stroke="#2563eb" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>
    ${dataPoints.map((d, i) => `
      <circle cx="${xs[i]}" cy="${ys[i]}" r="2.5" fill="${d.wid === activeWeekId ? '#0ea5e9' : '#2563eb'}"/>
    `).join('')}
  `;

  const labelsEl = document.getElementById('sparkLabels');
  const first = dataPoints[0].wid.split('~')[0];
  const last = dataPoints[dataPoints.length-1].wid.split('~')[0];
  labelsEl.innerHTML = `<span>${first}</span><span>${last}</span>`;
}

// ── History sidebar ────────────────────────────────────────────────────────
function renderHistory() {
  const store = getStore();
  const thisWeek = weekId();
  const allWeeks = [...new Set([thisWeek, ...store.order])].sort((a,b) => b.localeCompare(a));
  const list = document.getElementById('historyList');
  list.innerHTML = '';
  allWeeks.forEach(wid => {
    const btn = document.createElement('button');
    btn.className = 'history-week' + (wid === activeWeekId ? ' active' : '');
    const total = store.weeks[wid]?.totalConversations;
    btn.innerHTML = `<b>${wid === thisWeek ? 'This Week' : 'Week of'}</b>
      <small>${weekLabel(wid)}${total ? ' · '+Number(total).toLocaleString() : ''}</small>`;
    btn.onclick = () => { if (wid !== activeWeekId) loadWeek(wid); };
    list.appendChild(btn);
  });
}

function startNewWeek() {
  saveWeek();
  const now = new Date();
  const day = now.getDay();
  const nextMon = new Date(now);
  nextMon.setDate(now.getDate() + (7 - ((day + 6) % 7)));
  loadWeek(weekId(nextMon));
}

// ── Render ─────────────────────────────────────────────────────────────────
function renderAll() {
  renderPresets(); renderMeetings(); renderBullets(); renderSyncItems(); bindInputs();
}

function renderPresets() {
  const g = document.getElementById('presetsGrid'); g.innerHTML = '';
  const presets = getPresets();
  presets.forEach((p, i) => {
    if (presetChecks[i] === undefined) presetChecks[i] = false;
    const row = document.createElement('label');
    row.className = 'preset-row';
    row.innerHTML = `<input type="checkbox" ${presetChecks[i]?'checked':''} onchange="presetChecks[${i}]=this.checked;debounce();update()"><span>${esc(p)}</span>`;
    g.appendChild(row);
  });
}

function renderMeetings() {
  const l = document.getElementById('meetingsList'); l.innerHTML = '';
  meetings.forEach((m, i) => {
    const r = document.createElement('div'); r.className = 'bullet-row';
    r.innerHTML = `<input type="text" value="${esc(m)}" oninput="meetings[${i}]=this.value;debounce();update()" placeholder="Meeting name...">
      <button class="btn-remove" onclick="meetings.splice(${i},1);renderMeetings();debounce();update()">✕</button>`;
    l.appendChild(r);
  });
}
function addMeeting() { meetings.push(''); renderMeetings(); update(); }

function renderBullets() {
  const l = document.getElementById('bulletsList'); l.innerHTML = '';
  bullets.forEach((b, i) => {
    const r = document.createElement('div'); r.className = 'bullet-row';
    const carryOn = b.carryForward;
    r.innerHTML = `
      <input type="text" value="${esc(b.text)}" oninput="bullets[${i}].text=this.value;debounce();update()" placeholder="What was completed or updated this week?">
      <button class="btn-carry ${carryOn?'on':''}" title="${carryOn?'Will carry to next week':'Carry to next week'}" onclick="toggleCarry(${i})">→</button>
      <button class="btn-remove" onclick="bullets.splice(${i},1);renderBullets();debounce();update()">✕</button>`;
    l.appendChild(r);
  });
}
function addBullet() { bullets.push({text:'', carryForward:false}); renderBullets(); update(); }
function toggleCarry(i) {
  bullets[i].carryForward = !bullets[i].carryForward;
  renderBullets(); debounce(); update();
}

function renderSyncItems() {
  const l = document.getElementById('syncItems'); l.innerHTML = '';
  syncItems.forEach((s, i) => {
    const r = document.createElement('div'); r.className = 'bullet-row';
    r.innerHTML = `<input type="text" value="${esc(s)}" oninput="syncItems[${i}]=this.value;debounce();update()" placeholder="Agenda item...">
      <button class="btn-remove" onclick="syncItems.splice(${i},1);renderSyncItems();debounce();update()">✕</button>`;
    l.appendChild(r);
  });
}
function addSyncItem() { syncItems.push(''); renderSyncItems(); update(); }

function bindInputs() {
  ['teamName','dateRange','metricPeriod','totalConversations','medianResponseTime','responseGoal','syncMeeting','privateNotes','notesForNextWeek']
    .forEach(id => {
      const el = document.getElementById(id);
      if (el) el.oninput = () => { debounce(); update(); };
    });

  const templateSel = document.getElementById('templateSelect');
  if (templateSel) templateSel.onchange = () => {
    activeTemplate = templateSel.value;
    localStorage.setItem('ss_report_template', activeTemplate);
    presetChecks = {};
    renderPresets();
    update();
  };
}

// ── Goal indicator ─────────────────────────────────────────────────────────
function parseTimeToSeconds(str) {
  if (!str) return null;
  const m = str.match(/(\d+)\s*min(?:utes?)?\s*(\d+)?\s*sec?(?:onds?)?/i);
  if (m) return parseInt(m[1]) * 60 + parseInt(m[2] || '0');
  const s = str.match(/^(\d+):(\d+)$/);
  if (s) return parseInt(s[1]) * 60 + parseInt(s[2]);
  return null;
}

function updateGoalIndicator() {
  const medianStr = getF('medianResponseTime');
  const goalStr   = getF('responseGoal');
  const el = document.getElementById('goalIndicator');
  if (!medianStr || !goalStr) { el.innerHTML = ''; return; }

  const medianSec = parseTimeToSeconds(medianStr);
  const goalSec   = parseTimeToSeconds(goalStr);
  if (!medianSec || !goalSec) { el.innerHTML = ''; return; }

  const pct = medianSec / goalSec;
  let color, label;
  if (pct <= 0.85)      { color = '#059669'; label = '● Under goal'; }
  else if (pct <= 1.0)  { color = '#d97706'; label = '● Near goal'; }
  else                  { color = '#dc2626'; label = '● Over goal'; }

  el.innerHTML = `<span style="color:${color};font-size:10px">${label} (goal: ${goalStr})</span>`;
}

// ── Last week metric hint + comparison ─────────────────────────────────────
function updateMetricHints() {
  const store = getStore();
  const sorted = [...store.order].sort((a,b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid !== activeWeekId);
  const convHint = document.getElementById('convHint');

  if (prevWid && store.weeks[prevWid]) {
    const prev = store.weeks[prevWid];
    const parts = prevWid.split('~');
    const label = parts[0] || prevWid;
    const prevVal = parseInt(prev.totalConversations || '0', 10);
    const currVal = parseInt(getF('totalConversations') || '0', 10);

    let html = `Last week (${label}): ${prevVal > 0 ? Number(prevVal).toLocaleString() : '—'}`;
    if (prevVal > 0 && currVal > 0 && prevVal !== currVal) {
      const pct = Math.round(((currVal - prevVal) / prevVal) * 100);
      const cls = pct > 0 ? 'up' : (pct < 0 ? 'down' : 'same');
      const arrow = pct > 0 ? '↑' : '↓';
      html += ` <span class="metric-comparison ${cls}">(${arrow} ${Math.abs(pct)}%)</span>`;
    }
    convHint.innerHTML = html;
  } else {
    convHint.textContent = '';
  }
}

function updateCopyLastWeekVisibility() {
  const store = getStore();
  const sorted = [...store.order].sort((a,b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid < activeWeekId);
  const btn = document.getElementById('copyLastWeekBtn');
  if (btn) btn.style.display = prevWid ? 'block' : 'none';
}

function copyFromLastWeek() {
  const store = getStore();
  const sorted = [...store.order].sort((a,b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid < activeWeekId);
  if (!prevWid || !store.weeks[prevWid]) return;
  const prev = store.weeks[prevWid];

  meetings = [...(prev.meetings || [])];
  bullets = (prev.bullets || []).map(b => typeof b === 'string' ? {text:b, carryForward:false} : {...b});
  syncItems = [...(prev.syncItems || [])];
  presetChecks = {...(prev.presetChecks || {})};
  setF('notesForNextWeek', prev.notesForNextWeek || '');
  notesForNextWeek = prev.notesForNextWeek || '';

  renderAll();
  debounce();
  update();
  showSaved();
}

// ── Generate report ────────────────────────────────────────────────────────
function generatePlain() {
  let out = '';
  const team = getF('teamName').trim(), dr = getF('dateRange').trim();
  if (team || dr) out += `${team} Team Weekly Update ${dr}\n`;

  const period = getF('metricPeriod').trim(), total = getF('totalConversations').trim();
  const median = getF('medianResponseTime').trim(), goal = getF('responseGoal').trim();
  if (period||total||median||goal) {
    out += `CS Router Metrics: ${period}\n`;
    if (total) out += `Total Conversations: ${Number(total).toLocaleString()}\n`;
    if (median) out += `Median Response Time: ${median}\n`;
    if (goal) out += `Response Time Goal: ${goal}\n`;
  }

  meetings.filter(m=>m.trim()).forEach(m => out += `${m}\n`);

  [...getPresets().filter((_,i)=>presetChecks[i]), ...bullets.filter(b=>b.text.trim()).map(b=>b.text)]
    .forEach(b => out += `• ${b}\n`);

  const notesNext = getF('notesForNextWeek').trim();
  if (notesNext) {
    out += `\nNotes for Next Week:\n${notesNext}\n`;
  }

  const syncMtg = getF('syncMeeting').trim();
  const sItems = syncItems.filter(s=>s.trim());
  if (syncMtg||sItems.length) {
    if (syncMtg) out += `${syncMtg}\n`;
    sItems.forEach(s => out += `• ${s}\n`);
  }
  return out.trim();
}

function generateSlack() {
  let out = '';
  const team = getF('teamName').trim(), dr = getF('dateRange').trim();
  if (team || dr) out += `*${team} Team Weekly Update ${dr}*\n`;

  const period = getF('metricPeriod').trim(), total = getF('totalConversations').trim();
  const median = getF('medianResponseTime').trim(), goal = getF('responseGoal').trim();
  if (period||total||median||goal) {
    out += `\n*CS Router Metrics: ${period}*\n`;
    if (total) out += `Total Conversations: ${Number(total).toLocaleString()}\n`;
    if (median) out += `Median Response Time: ${median}\n`;
    if (goal) out += `Response Time Goal: ${goal}\n`;
  }

  const mtgs = meetings.filter(m=>m.trim());
  if (mtgs.length) {
    out += '\n';
    mtgs.forEach(m => out += `${m}\n`);
  }

  const allBullets = [...getPresets().filter((_,i)=>presetChecks[i]), ...bullets.filter(b=>b.text.trim()).map(b=>b.text)];
  if (allBullets.length) {
    out += '\n';
    allBullets.forEach(b => out += `• ${b}\n`);
  }

  const notesNext = getF('notesForNextWeek').trim();
  if (notesNext) {
    out += `\n*Notes for Next Week:*\n${notesNext}\n`;
  }

  const syncMtg = getF('syncMeeting').trim();
  const sItems = syncItems.filter(s=>s.trim());
  if (syncMtg||sItems.length) {
    out += '\n';
    if (syncMtg) out += `*${syncMtg}*\n`;
    sItems.forEach(s => out += `• ${s}\n`);
  }
  return out.trim();
}

// ── Update UI ──────────────────────────────────────────────────────────────
function update() {
  const text = generatePlain();
  document.getElementById('previewBox').textContent = text || 'Start filling in the form to see your report...';

  const len = text.length;
  const ccEl = document.getElementById('charCount');
  let cls = '';
  if (len > SLACK_CHAR_LIMIT) cls = 'char-danger';
  else if (len > SLACK_CHAR_LIMIT * 0.85) cls = 'char-warn';
  ccEl.innerHTML = `<span class="${cls}">${len.toLocaleString()}</span> chars &nbsp;|&nbsp; <span class="${cls}">${text.split(' ').filter(Boolean).length}</span> words`;

  updateGoalIndicator();
  updateMetricHints();
  showReminderBanner();
}

// ── Copy ───────────────────────────────────────────────────────────────────
async function copyPlain() {
  const text = generatePlain();
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    flash('copyBtn', '✓ Copied!', 'btn-copy copied');
  } catch { alert('Copy failed.'); }
}

async function copySlack() {
  const text = generateSlack();
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    flash('slackBtn', '✓ Copied!', 'btn-slack copied');
  } catch { alert('Copy failed.'); }
}

function flash(id, label, cls) {
  const btn = document.getElementById(id);
  const orig = btn.textContent;
  const origCls = btn.className;
  btn.textContent = label; btn.className = cls;
  setTimeout(() => { btn.textContent = orig; btn.className = origCls; }, 2000);
}

// ── Reset ──────────────────────────────────────────────────────────────────
function resetWeek() {
  if (!confirm('Reset this week to blank defaults?')) return;
  const store = getStore();
  delete store.weeks[activeWeekId];
  store.order = store.order.filter(id => id !== activeWeekId);
  saveStore(store);
  loadWeek(activeWeekId);
}

// ── Save status ────────────────────────────────────────────────────────────
function setSaveStatus() {
  const el = document.getElementById('saveStatus');
  el.style.color = '#d97706'; el.textContent = '● Saving…';
}
function showSaved() {
  const el = document.getElementById('saveStatus');
  el.style.color = '#059669'; el.textContent = '✓ Saved';
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => { el.textContent = ''; }, 3000);
}

// ── Helpers ────────────────────────────────────────────────────────────────
function getF(id) { return (document.getElementById(id)||{}).value || ''; }
function setF(id, val) { const el = document.getElementById(id); if(el) el.value = val; }
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ── How to Use toggle ─────────────────────────────────────────────────────
function toggleHowTo() {
  const btn = document.getElementById('howToToggle');
  const content = document.getElementById('howToContent');
  const isOpen = content.classList.toggle('is-open');
  btn.setAttribute('aria-expanded', isOpen);
}

// ── Print ─────────────────────────────────────────────────────────────────
function printReport() {
  const text = generatePlain();
  const container = document.getElementById('printContainer');
  if (!container) return;
  container.textContent = text || 'No content to print.';
  container.style.display = 'block';
  window.print();
  container.style.display = 'none';
}

// ── Export / Import ───────────────────────────────────────────────────────
function exportData() {
  const store = getStore();
  const blob = new Blob([JSON.stringify(store, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `ss-report-backup-${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
}

function importData(ev) {
  const file = ev.target.files?.[0];
  if (!file) return;
  const r = new FileReader();
  r.onload = () => {
    try {
      const data = JSON.parse(r.result);
      if (data.weeks && Array.isArray(data.order)) {
        saveStore(data);
        loadWeek(activeWeekId);
        showSaved();
        alert('Data imported successfully.');
      } else {
        alert('Invalid backup file format.');
      }
    } catch (e) {
      alert('Could not parse file. ' + e.message);
    }
  };
  r.readAsText(file);
  ev.target.value = '';
}

// ── Dark mode ─────────────────────────────────────────────────────────────
function toggleDarkMode() {
  document.body.classList.toggle('dark');
  const isDark = document.body.classList.contains('dark');
  localStorage.setItem('ss_report_dark', isDark ? '1' : '0');
  updateDarkModeIcon();
}

function updateDarkModeIcon() {
  const btn = document.getElementById('darkModeToggle');
  if (!btn) return;
  btn.textContent = document.body.classList.contains('dark') ? '☀️' : '🌙';
  btn.title = document.body.classList.contains('dark') ? 'Switch to light mode' : 'Switch to dark mode';
}

// ── Week-end reminder ─────────────────────────────────────────────────────
function showReminderBanner() {
  const banner = document.getElementById('reminderBanner');
  const textEl = document.getElementById('reminderText');
  if (!banner || !textEl) return;

  const now = new Date();
  const day = now.getDay();
  const fri = day === 5;
  const satSun = day === 0 || day === 6;
  const hasMetrics = parseInt(getF('totalConversations') || '0', 10) > 0;

  if (fri && !hasMetrics) {
    banner.style.display = 'block';
    textEl.textContent = '📋 Report due today — fill in your metrics and accomplishments.';
  } else if (fri) {
    banner.style.display = 'block';
    textEl.textContent = '📋 Report due today — don\'t forget to copy and submit.';
  } else if (day === 4) {
    banner.style.display = 'block';
    textEl.textContent = '📋 Report due tomorrow (Friday).';
  } else if (!satSun) {
    const daysToFri = (5 - day + 7) % 7;
    banner.style.display = 'block';
    textEl.textContent = `📋 Report due in ${daysToFri} day${daysToFri === 1 ? '' : 's'} (Friday).`;
  } else {
    banner.style.display = 'none';
  }
}

// ── Keyboard shortcuts ────────────────────────────────────────────────────
function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'Enter') {
      e.preventDefault();
      copySlack();
    } else if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      copyPlain();
    }
  });
}

init();
