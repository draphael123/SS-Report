// ══════════════════════════════════════════════════════════════════════════════
// Weekly Update Builder - Refactored with Security, Accessibility & Modern Patterns
// ══════════════════════════════════════════════════════════════════════════════

// ── Type Definitions (JSDoc) ─────────────────────────────────────────────────

/**
 * @typedef {Object} Bullet
 * @property {string} text - The bullet text content
 * @property {boolean} carryForward - Whether to carry to next week
 */

/**
 * @typedef {Object} WeekData
 * @property {string} teamName
 * @property {string} dateRange
 * @property {string} metricPeriod
 * @property {string} totalConversations
 * @property {string} medianResponseTime
 * @property {string} responseGoal
 * @property {string} syncMeeting
 * @property {string} privateNotes
 * @property {string} notesForNextWeek
 * @property {string[]} meetings
 * @property {Bullet[]} bullets
 * @property {string[]} syncItems
 * @property {Object<number, boolean>} presetChecks
 * @property {string} savedAt
 */

/**
 * @typedef {Object} StoreData
 * @property {Object<string, WeekData>} weeks
 * @property {string[]} order
 * @property {number} version
 */

/**
 * @typedef {Object} Template
 * @property {string} id
 * @property {string} name
 * @property {string[]} presets
 * @property {string[]} defaultMeetings
 * @property {string[]} defaultSync
 */

/**
 * @typedef {Object} HistoryEntry
 * @property {string[]} meetings
 * @property {Bullet[]} bullets
 * @property {string[]} syncItems
 * @property {Object<number, boolean>} presetChecks
 * @property {string} notesForNextWeek
 */

// ── Constants ────────────────────────────────────────────────────────────────

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
const LS_TEMPLATES_KEY = 'ss_report_custom_templates';
const SLACK_CHAR_LIMIT = 4000;
const CURRENT_DATA_VERSION = 2;
const MAX_HISTORY_SIZE = 50;
const DEBOUNCE_DELAY = 700;
const TOAST_DURATION = 3000;
const STATUS_CLEAR_DELAY = 3000;

// ── Application State ────────────────────────────────────────────────────────

const AppState = {
  activeWeekId: '',
  meetings: [],
  bullets: [],
  syncItems: [],
  presetChecks: {},
  notesForNextWeek: '',
  activeTemplate: 'cs',
  customTemplates: [],
  // Undo/Redo
  history: [],
  historyIndex: -1,
  isUndoRedo: false,
  // Timers
  saveTimer: null,
  statusTimer: null,
  // Dirty tracking for selective rendering
  dirty: {
    meetings: true,
    bullets: true,
    syncItems: true,
    presets: true,
    preview: true
  }
};

// ── Built-in Templates ───────────────────────────────────────────────────────

/** @type {Template[]} */
const BUILT_IN_TEMPLATES = [
  {
    id: 'cs',
    name: 'CS Weekly',
    presets: PRESETS_CS,
    defaultMeetings: DEF_MEETINGS_CS,
    defaultSync: DEF_SYNC_CS
  },
  {
    id: 'operations',
    name: 'Operations',
    presets: PRESETS_OPERATIONS,
    defaultMeetings: DEF_MEETINGS_OPERATIONS,
    defaultSync: DEF_SYNC_OPERATIONS
  }
];

// ── Template Helpers ─────────────────────────────────────────────────────────

/**
 * Get all available templates (built-in + custom)
 * @returns {Template[]}
 */
function getAllTemplates() {
  return [...BUILT_IN_TEMPLATES, ...AppState.customTemplates];
}

/**
 * Get current template
 * @returns {Template}
 */
function getCurrentTemplate() {
  const all = getAllTemplates();
  return all.find(t => t.id === AppState.activeTemplate) || BUILT_IN_TEMPLATES[0];
}

/**
 * Get presets for current template
 * @returns {string[]}
 */
function getPresets() {
  return getCurrentTemplate().presets;
}

/**
 * Get default meetings for current template
 * @returns {string[]}
 */
function getDefMeetings() {
  return getCurrentTemplate().defaultMeetings;
}

/**
 * Get default sync items for current template
 * @returns {string[]}
 */
function getDefSync() {
  return getCurrentTemplate().defaultSync;
}

// ── Week Helpers ─────────────────────────────────────────────────────────────

/**
 * Calculate Monday of the week for a given date
 * @param {Date} d - Date to calculate from
 * @returns {Date} Monday of that week
 */
function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  date.setDate(date.getDate() - ((day + 6) % 7));
  date.setHours(0, 0, 0, 0);
  return date;
}

/**
 * Format date as M/D/YY
 * @param {Date} d
 * @returns {string}
 */
function formatShort(d) {
  return `${d.getMonth() + 1}/${d.getDate()}/${String(d.getFullYear()).slice(-2)}`;
}

/**
 * Format date as MM/DD/YYYY
 * @param {Date} d
 * @returns {string}
 */
function formatMetric(d) {
  return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
}

/**
 * Generate week ID from date (Mon~Fri format)
 * @param {Date|string} [d] - Date to use, defaults to today
 * @returns {string}
 */
function weekId(d) {
  const date = d ? new Date(d) : new Date();
  const mon = getMonday(date);
  const fri = new Date(mon);
  fri.setDate(mon.getDate() + 4);
  return formatShort(mon) + '~' + formatShort(fri);
}

/**
 * Get various date formats for a week
 * @param {Date|string} [d] - Date to use
 * @returns {{short: string, metric: string, label: string}}
 */
function weekDates(d) {
  const date = d ? new Date(d) : new Date();
  const mon = getMonday(date);
  const fri = new Date(mon);
  fri.setDate(mon.getDate() + 4);
  return {
    short: formatShort(mon) + '-' + formatShort(fri),
    metric: formatMetric(mon) + '-' + formatMetric(fri),
    label: formatShort(mon) + ' – ' + formatShort(fri)
  };
}

/**
 * Convert week ID to display label
 * @param {string} wid - Week ID
 * @returns {string}
 */
function weekLabel(wid) {
  const parts = wid.split('~');
  return (parts[0] || '') + ' – ' + (parts[1] || '');
}

/**
 * Validate week ID format
 * @param {string} wid
 * @returns {boolean}
 */
function isValidWeekId(wid) {
  if (typeof wid !== 'string') return false;
  return /^\d{1,2}\/\d{1,2}\/\d{2}~\d{1,2}\/\d{1,2}\/\d{2}$/.test(wid);
}

// ── Input Validation ─────────────────────────────────────────────────────────

/**
 * Parse time string to seconds
 * @param {string} str - Time string like "2 minutes 30 seconds" or "2:30"
 * @returns {number|null}
 */
function parseTimeToSeconds(str) {
  if (!str || typeof str !== 'string') return null;

  // Format: "X minutes Y seconds" or "X min Y sec"
  const minSec = str.match(/(\d+)\s*min(?:utes?)?\s*(\d+)?\s*sec?(?:onds?)?/i);
  if (minSec) {
    const mins = parseInt(minSec[1], 10);
    const secs = parseInt(minSec[2] || '0', 10);
    if (mins >= 0 && mins < 1000 && secs >= 0 && secs < 60) {
      return mins * 60 + secs;
    }
  }

  // Format: "M:SS" or "MM:SS"
  const colonFormat = str.match(/^(\d{1,3}):(\d{2})$/);
  if (colonFormat) {
    const mins = parseInt(colonFormat[1], 10);
    const secs = parseInt(colonFormat[2], 10);
    if (mins >= 0 && mins < 1000 && secs >= 0 && secs < 60) {
      return mins * 60 + secs;
    }
  }

  return null;
}

/**
 * Validate imported data structure
 * @param {any} data
 * @returns {{valid: boolean, error?: string}}
 */
function validateImportData(data) {
  if (!data || typeof data !== 'object') {
    return { valid: false, error: 'Invalid data format' };
  }

  if (!data.weeks || typeof data.weeks !== 'object') {
    return { valid: false, error: 'Missing or invalid weeks data' };
  }

  if (!Array.isArray(data.order)) {
    return { valid: false, error: 'Missing or invalid order array' };
  }

  // Validate each week ID in order
  for (const wid of data.order) {
    if (!isValidWeekId(wid)) {
      return { valid: false, error: `Invalid week ID format: ${wid}` };
    }
    if (!data.weeks[wid]) {
      return { valid: false, error: `Missing week data for: ${wid}` };
    }
  }

  // Validate week data structure
  for (const [wid, weekData] of Object.entries(data.weeks)) {
    if (!isValidWeekId(wid)) {
      return { valid: false, error: `Invalid week ID: ${wid}` };
    }
    if (typeof weekData !== 'object') {
      return { valid: false, error: `Invalid week data for: ${wid}` };
    }
    // Check for required string fields
    const stringFields = ['teamName', 'dateRange'];
    for (const field of stringFields) {
      if (weekData[field] !== undefined && typeof weekData[field] !== 'string') {
        return { valid: false, error: `Invalid ${field} in week ${wid}` };
      }
    }
    // Check arrays
    if (weekData.meetings && !Array.isArray(weekData.meetings)) {
      return { valid: false, error: `Invalid meetings array in week ${wid}` };
    }
    if (weekData.bullets && !Array.isArray(weekData.bullets)) {
      return { valid: false, error: `Invalid bullets array in week ${wid}` };
    }
    if (weekData.syncItems && !Array.isArray(weekData.syncItems)) {
      return { valid: false, error: `Invalid syncItems array in week ${wid}` };
    }
  }

  return { valid: true };
}

/**
 * Sanitize string for safe display
 * @param {any} s
 * @returns {string}
 */
function sanitize(s) {
  if (s === null || s === undefined) return '';
  return String(s);
}

// ── Data Migration ───────────────────────────────────────────────────────────

/**
 * Migrate bullets from string format to object format
 * @param {any} bullet
 * @returns {Bullet}
 */
function migrateBullet(bullet) {
  if (typeof bullet === 'string') {
    return { text: bullet, carryForward: false };
  }
  if (typeof bullet === 'object' && bullet !== null) {
    return {
      text: sanitize(bullet.text),
      carryForward: Boolean(bullet.carryForward)
    };
  }
  return { text: '', carryForward: false };
}

/**
 * Migrate store data to current version
 * @param {any} data
 * @returns {StoreData}
 */
function migrateData(data) {
  const version = data.version || 1;

  if (version < 2) {
    // Migrate bullets from string to object format
    if (data.weeks) {
      for (const wid of Object.keys(data.weeks)) {
        const week = data.weeks[wid];
        if (week.bullets && Array.isArray(week.bullets)) {
          week.bullets = week.bullets.map(migrateBullet);
        }
      }
    }
    data.version = 2;
  }

  return data;
}

// ── Storage ──────────────────────────────────────────────────────────────────

/**
 * Get store from localStorage with migration
 * @returns {StoreData}
 */
function getStore() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { weeks: {}, order: [], version: CURRENT_DATA_VERSION };

    let data = JSON.parse(raw);
    data = migrateData(data);
    return data;
  } catch (e) {
    console.error('Failed to load store:', e);
    return { weeks: {}, order: [], version: CURRENT_DATA_VERSION };
  }
}

/**
 * Save store to localStorage
 * @param {StoreData} data
 * @returns {boolean} Success
 */
function saveStore(data) {
  try {
    data.version = CURRENT_DATA_VERSION;
    localStorage.setItem(LS_KEY, JSON.stringify(data));
    return true;
  } catch (e) {
    console.error('Failed to save store:', e);
    if (e.name === 'QuotaExceededError') {
      showToast('Storage quota exceeded. Please export and clear old data.', 'error');
    } else {
      showToast('Failed to save data.', 'error');
    }
    return false;
  }
}

/**
 * Load custom templates from localStorage
 */
function loadCustomTemplates() {
  try {
    const raw = localStorage.getItem(LS_TEMPLATES_KEY);
    if (raw) {
      AppState.customTemplates = JSON.parse(raw);
    }
  } catch (e) {
    console.error('Failed to load custom templates:', e);
  }
}

/**
 * Save custom templates to localStorage
 */
function saveCustomTemplates() {
  try {
    localStorage.setItem(LS_TEMPLATES_KEY, JSON.stringify(AppState.customTemplates));
  } catch (e) {
    console.error('Failed to save custom templates:', e);
  }
}

// ── Toast Notification System ────────────────────────────────────────────────

/**
 * Show a toast notification
 * @param {string} message - Message to display
 * @param {'success'|'error'|'warning'|'info'} [type='info'] - Toast type
 * @param {number} [duration=3000] - Duration in ms
 */
function showToast(message, type = 'info', duration = TOAST_DURATION) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.setAttribute('role', 'alert');
  toast.setAttribute('aria-live', 'polite');

  const icon = document.createElement('span');
  icon.className = 'toast-icon';
  icon.setAttribute('aria-hidden', 'true');
  switch (type) {
    case 'success': icon.textContent = '\u2713'; break;
    case 'error': icon.textContent = '\u2717'; break;
    case 'warning': icon.textContent = '\u26A0'; break;
    default: icon.textContent = '\u2139'; break;
  }

  const text = document.createElement('span');
  text.className = 'toast-message';
  text.textContent = message;

  const closeBtn = document.createElement('button');
  closeBtn.className = 'toast-close';
  closeBtn.setAttribute('aria-label', 'Close notification');
  closeBtn.textContent = '\u00D7';
  closeBtn.onclick = () => removeToast(toast);

  toast.appendChild(icon);
  toast.appendChild(text);
  toast.appendChild(closeBtn);
  container.appendChild(toast);

  // Trigger animation
  requestAnimationFrame(() => {
    toast.classList.add('toast-visible');
  });

  // Auto-remove
  setTimeout(() => removeToast(toast), duration);
}

/**
 * Remove a toast element
 * @param {HTMLElement} toast
 */
function removeToast(toast) {
  toast.classList.remove('toast-visible');
  toast.classList.add('toast-hiding');
  setTimeout(() => {
    if (toast.parentNode) {
      toast.parentNode.removeChild(toast);
    }
  }, 300);
}

// ── Undo/Redo System ─────────────────────────────────────────────────────────

/**
 * Get current state snapshot for history
 * @returns {HistoryEntry}
 */
function getStateSnapshot() {
  return {
    meetings: [...AppState.meetings],
    bullets: AppState.bullets.map(b => ({ ...b })),
    syncItems: [...AppState.syncItems],
    presetChecks: { ...AppState.presetChecks },
    notesForNextWeek: getF('notesForNextWeek')
  };
}

/**
 * Push current state to history
 */
function pushHistory() {
  if (AppState.isUndoRedo) return;

  // Remove any redo states
  if (AppState.historyIndex < AppState.history.length - 1) {
    AppState.history = AppState.history.slice(0, AppState.historyIndex + 1);
  }

  AppState.history.push(getStateSnapshot());

  // Limit history size
  if (AppState.history.length > MAX_HISTORY_SIZE) {
    AppState.history.shift();
  } else {
    AppState.historyIndex++;
  }

  updateUndoRedoButtons();
}

/**
 * Restore state from history entry
 * @param {HistoryEntry} entry
 */
function restoreState(entry) {
  AppState.meetings = [...entry.meetings];
  AppState.bullets = entry.bullets.map(b => ({ ...b }));
  AppState.syncItems = [...entry.syncItems];
  AppState.presetChecks = { ...entry.presetChecks };
  setF('notesForNextWeek', entry.notesForNextWeek);

  markAllDirty();
  renderAll();
  update();
  debounce();
}

/**
 * Undo last change
 */
function undo() {
  if (AppState.historyIndex <= 0) {
    showToast('Nothing to undo', 'info');
    return;
  }

  AppState.isUndoRedo = true;
  AppState.historyIndex--;
  restoreState(AppState.history[AppState.historyIndex]);
  AppState.isUndoRedo = false;

  updateUndoRedoButtons();
  showToast('Undone', 'info', 1500);
}

/**
 * Redo last undone change
 */
function redo() {
  if (AppState.historyIndex >= AppState.history.length - 1) {
    showToast('Nothing to redo', 'info');
    return;
  }

  AppState.isUndoRedo = true;
  AppState.historyIndex++;
  restoreState(AppState.history[AppState.historyIndex]);
  AppState.isUndoRedo = false;

  updateUndoRedoButtons();
  showToast('Redone', 'info', 1500);
}

/**
 * Update undo/redo button states
 */
function updateUndoRedoButtons() {
  const undoBtn = document.getElementById('undoBtn');
  const redoBtn = document.getElementById('redoBtn');

  if (undoBtn) {
    undoBtn.disabled = AppState.historyIndex <= 0;
  }
  if (redoBtn) {
    redoBtn.disabled = AppState.historyIndex >= AppState.history.length - 1;
  }
}

/**
 * Clear history (e.g., when switching weeks)
 */
function clearHistory() {
  AppState.history = [getStateSnapshot()];
  AppState.historyIndex = 0;
  updateUndoRedoButtons();
}

// ── Multi-Tab Conflict Handling ──────────────────────────────────────────────

/**
 * Handle storage events from other tabs
 * @param {StorageEvent} e
 */
function handleStorageChange(e) {
  if (e.key !== LS_KEY) return;

  // Another tab changed the data
  showConflictModal();
}

/**
 * Show conflict resolution modal
 */
function showConflictModal() {
  const modal = document.getElementById('conflictModal');
  if (modal) {
    modal.classList.add('is-open');
    modal.setAttribute('aria-hidden', 'false');
    // Focus the modal for accessibility
    const firstBtn = modal.querySelector('button');
    if (firstBtn) firstBtn.focus();
  }
}

/**
 * Hide conflict modal
 */
function hideConflictModal() {
  const modal = document.getElementById('conflictModal');
  if (modal) {
    modal.classList.remove('is-open');
    modal.setAttribute('aria-hidden', 'true');
  }
}

/**
 * Keep current data, ignore external changes
 */
function conflictKeepMine() {
  hideConflictModal();
  saveWeek();
  showToast('Kept your changes', 'success');
}

/**
 * Load external data, discard current
 */
function conflictLoadExternal() {
  hideConflictModal();
  loadWeek(AppState.activeWeekId);
  showToast('Loaded external changes', 'success');
}

// ── Initialize ───────────────────────────────────────────────────────────────

/**
 * Initialize the application
 */
function init() {
  // Load template preference
  loadCustomTemplates();
  AppState.activeTemplate = localStorage.getItem('ss_report_template') || 'cs';

  const sel = document.getElementById('templateSelect');
  if (sel) {
    populateTemplateSelect();
    sel.value = AppState.activeTemplate;
  }

  // Load dark mode preference
  const dark = localStorage.getItem('ss_report_dark') === '1';
  if (dark) document.body.classList.add('dark');
  updateDarkModeIcon();

  // Initialize week
  AppState.activeWeekId = weekId();
  const dates = weekDates();

  const weekBadge = document.getElementById('weekBadge');
  if (weekBadge) {
    weekBadge.textContent = 'Week of ' + dates.label;
  }

  loadWeek(AppState.activeWeekId);
  showReminderBanner();
  setupKeyboardShortcuts();

  // Listen for storage changes from other tabs
  window.addEventListener('storage', handleStorageChange);

  // Initialize undo/redo
  clearHistory();
}

// ── Load / Save ──────────────────────────────────────────────────────────────

/**
 * Load a specific week's data
 * @param {string} wid - Week ID to load
 */
function loadWeek(wid) {
  AppState.activeWeekId = wid;
  const store = getStore();
  const saved = store.weeks[wid];
  const dates = weekDates();

  if (saved) {
    setF('teamName', sanitize(saved.teamName) || 'CS/Refills and Clarifications/RN/Shift Supervisor');
    setF('dateRange', sanitize(saved.dateRange) || '');
    setF('metricPeriod', sanitize(saved.metricPeriod) || '');
    setF('totalConversations', sanitize(saved.totalConversations) || '');
    setF('medianResponseTime', sanitize(saved.medianResponseTime) || '');
    setF('responseGoal', sanitize(saved.responseGoal) || '2 minutes 30 seconds');
    setF('syncMeeting', sanitize(saved.syncMeeting) || '');
    setF('privateNotes', sanitize(saved.privateNotes) || '');
    setF('notesForNextWeek', sanitize(saved.notesForNextWeek) || '');
    AppState.meetings = (saved.meetings || getDefMeetings()).map(sanitize);
    AppState.bullets = (saved.bullets || [{ text: '', carryForward: false }]).map(migrateBullet);
    AppState.syncItems = (saved.syncItems || getDefSync()).map(sanitize);
    AppState.presetChecks = { ...(saved.presetChecks || {}) };
    AppState.notesForNextWeek = sanitize(saved.notesForNextWeek) || '';
  } else {
    const carryBullets = getCarryForwardBullets(wid, store);
    const carryNotes = getCarryForwardNotes(wid, store);

    setF('teamName', 'CS/Refills and Clarifications/RN/Shift Supervisor');
    setF('dateRange', wid === weekId() ? dates.short : weekLabel(wid));
    setF('metricPeriod', wid === weekId() ? dates.metric : '');
    setF('totalConversations', '');
    setF('medianResponseTime', '');
    setF('responseGoal', '2 minutes 30 seconds');
    setF('syncMeeting', AppState.activeTemplate === 'operations'
      ? 'Upcoming topics:'
      : 'Front Office Sync Monthly Meeting on [DATE] to review the following:');
    setF('privateNotes', '');
    setF('notesForNextWeek', carryNotes);
    AppState.meetings = [...getDefMeetings()];
    AppState.bullets = carryBullets.length > 0 ? carryBullets : [{ text: '', carryForward: false }];
    AppState.syncItems = [...getDefSync()];
    AppState.presetChecks = {};
    AppState.notesForNextWeek = carryNotes;
  }

  updateCopyLastWeekVisibility();
  markAllDirty();
  renderAll();
  renderHistory();
  renderSparkline();
  update();
  clearHistory();
}

/**
 * Get carry-forward bullets from previous week
 * @param {string} currentWid
 * @param {StoreData} store
 * @returns {Bullet[]}
 */
function getCarryForwardBullets(currentWid, store) {
  const sortedOrder = [...store.order].sort((a, b) => b.localeCompare(a));
  const prevWid = sortedOrder.find(wid => wid < currentWid);
  if (!prevWid) return [];

  const prevData = store.weeks[prevWid];
  if (!prevData || !prevData.bullets) return [];

  return prevData.bullets
    .filter(b => typeof b === 'object' && b.carryForward && b.text && b.text.trim())
    .map(b => ({ text: b.text, carryForward: false }));
}

/**
 * Get notes from previous week
 * @param {string} currentWid
 * @param {StoreData} store
 * @returns {string}
 */
function getCarryForwardNotes(currentWid, store) {
  const sortedOrder = [...store.order].sort((a, b) => b.localeCompare(a));
  const prevWid = sortedOrder.find(wid => wid < currentWid);
  if (!prevWid) return '';

  const prevData = store.weeks[prevWid];
  return (prevData && prevData.notesForNextWeek) ? prevData.notesForNextWeek : '';
}

/**
 * Save current week's data
 */
function saveWeek() {
  const store = getStore();
  const wid = AppState.activeWeekId;

  AppState.notesForNextWeek = getF('notesForNextWeek');

  store.weeks[wid] = {
    teamName: getF('teamName'),
    dateRange: getF('dateRange'),
    metricPeriod: getF('metricPeriod'),
    totalConversations: getF('totalConversations'),
    medianResponseTime: getF('medianResponseTime'),
    responseGoal: getF('responseGoal'),
    syncMeeting: getF('syncMeeting'),
    privateNotes: getF('privateNotes'),
    notesForNextWeek: AppState.notesForNextWeek,
    meetings: [...AppState.meetings],
    bullets: AppState.bullets.map(b => ({ ...b })),
    syncItems: [...AppState.syncItems],
    presetChecks: { ...AppState.presetChecks },
    savedAt: new Date().toISOString()
  };

  if (!store.order.includes(wid)) {
    store.order = [wid, ...store.order].sort((a, b) => b.localeCompare(a));
  }

  if (saveStore(store)) {
    showSaved();
    renderHistory();
    renderSparkline();
  }
}

/**
 * Debounced save with history push
 */
function debounce() {
  clearTimeout(AppState.saveTimer);
  setSaveStatus('saving');
  pushHistory();
  AppState.saveTimer = setTimeout(saveWeek, DEBOUNCE_DELAY);
}

// ── Mark Dirty for Selective Rendering ───────────────────────────────────────

function markAllDirty() {
  AppState.dirty.meetings = true;
  AppState.dirty.bullets = true;
  AppState.dirty.syncItems = true;
  AppState.dirty.presets = true;
  AppState.dirty.preview = true;
}

// ── Sparkline ────────────────────────────────────────────────────────────────

/**
 * Render the sparkline chart (using DOM methods for security)
 */
function renderSparkline() {
  const store = getStore();
  const sorted = [...store.order].sort((a, b) => a.localeCompare(b));
  const dataPoints = sorted
    .map(wid => ({ wid, val: parseInt(store.weeks[wid]?.totalConversations || '0', 10) }))
    .filter(d => d.val > 0)
    .slice(-6);

  const wrap = document.getElementById('sparkWrap');
  const svg = document.getElementById('sparkSvg');
  const labelsEl = document.getElementById('sparkLabels');

  if (!wrap || !svg || !labelsEl) return;

  if (dataPoints.length < 2) {
    wrap.style.display = 'none';
    return;
  }
  wrap.style.display = 'block';

  const W = 160, H = 48, PAD = 6;
  const vals = dataPoints.map(d => d.val);
  const min = Math.min(...vals), max = Math.max(...vals);
  const range = max - min || 1;

  const xs = dataPoints.map((_, i) => PAD + (i / (dataPoints.length - 1)) * (W - PAD * 2));
  const ys = vals.map(v => H - PAD - ((v - min) / range) * (H - PAD * 2));

  // Clear existing content
  while (svg.firstChild) {
    svg.removeChild(svg.firstChild);
  }

  const ns = 'http://www.w3.org/2000/svg';

  // Create gradient
  const defs = document.createElementNS(ns, 'defs');
  const gradient = document.createElementNS(ns, 'linearGradient');
  gradient.setAttribute('id', 'sg');
  gradient.setAttribute('x1', '0');
  gradient.setAttribute('y1', '0');
  gradient.setAttribute('x2', '0');
  gradient.setAttribute('y2', '1');

  const stop1 = document.createElementNS(ns, 'stop');
  stop1.setAttribute('offset', '0%');
  stop1.setAttribute('stop-color', '#2563eb');
  stop1.setAttribute('stop-opacity', '0.3');

  const stop2 = document.createElementNS(ns, 'stop');
  stop2.setAttribute('offset', '100%');
  stop2.setAttribute('stop-color', '#2563eb');
  stop2.setAttribute('stop-opacity', '0');

  gradient.appendChild(stop1);
  gradient.appendChild(stop2);
  defs.appendChild(gradient);
  svg.appendChild(defs);

  // Create title for accessibility
  const title = document.createElementNS(ns, 'title');
  title.textContent = `Conversation trend: ${vals.join(', ')} over ${dataPoints.length} weeks`;
  svg.appendChild(title);

  // Create fill polygon
  const fillPts = `${xs[0]},${H} ` + xs.map((x, i) => `${x},${ys[i]}`).join(' ') + ` ${xs[xs.length - 1]},${H}`;
  const polygon = document.createElementNS(ns, 'polygon');
  polygon.setAttribute('points', fillPts);
  polygon.setAttribute('fill', 'url(#sg)');
  svg.appendChild(polygon);

  // Create line
  const polyline = document.createElementNS(ns, 'polyline');
  polyline.setAttribute('points', xs.map((x, i) => `${x},${ys[i]}`).join(' '));
  polyline.setAttribute('fill', 'none');
  polyline.setAttribute('stroke', '#2563eb');
  polyline.setAttribute('stroke-width', '1.5');
  polyline.setAttribute('stroke-linejoin', 'round');
  polyline.setAttribute('stroke-linecap', 'round');
  svg.appendChild(polyline);

  // Create circles
  dataPoints.forEach((d, i) => {
    const circle = document.createElementNS(ns, 'circle');
    circle.setAttribute('cx', String(xs[i]));
    circle.setAttribute('cy', String(ys[i]));
    circle.setAttribute('r', '2.5');
    circle.setAttribute('fill', d.wid === AppState.activeWeekId ? '#0ea5e9' : '#2563eb');

    // Add tooltip
    const tooltipTitle = document.createElementNS(ns, 'title');
    tooltipTitle.textContent = `${d.wid}: ${d.val.toLocaleString()} conversations`;
    circle.appendChild(tooltipTitle);

    svg.appendChild(circle);
  });

  // Update labels
  while (labelsEl.firstChild) {
    labelsEl.removeChild(labelsEl.firstChild);
  }

  const first = dataPoints[0].wid.split('~')[0];
  const last = dataPoints[dataPoints.length - 1].wid.split('~')[0];

  const firstSpan = document.createElement('span');
  firstSpan.textContent = first;
  const lastSpan = document.createElement('span');
  lastSpan.textContent = last;

  labelsEl.appendChild(firstSpan);
  labelsEl.appendChild(lastSpan);
}

// ── History Sidebar ──────────────────────────────────────────────────────────

/**
 * Render the week history sidebar
 */
function renderHistory() {
  const store = getStore();
  const thisWeek = weekId();
  const allWeeks = [...new Set([thisWeek, ...store.order])].sort((a, b) => b.localeCompare(a));
  const list = document.getElementById('historyList');

  if (!list) return;

  // Clear existing
  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }

  allWeeks.forEach(wid => {
    const btn = document.createElement('button');
    btn.className = 'history-week' + (wid === AppState.activeWeekId ? ' active' : '');
    btn.setAttribute('aria-pressed', wid === AppState.activeWeekId ? 'true' : 'false');

    const isThisWeek = wid === thisWeek;
    const total = store.weeks[wid]?.totalConversations;

    const boldText = document.createElement('b');
    boldText.textContent = isThisWeek ? 'This Week' : 'Week of';

    const smallText = document.createElement('small');
    let labelText = weekLabel(wid);
    if (total) {
      labelText += ' \u00B7 ' + Number(total).toLocaleString();
    }
    smallText.textContent = labelText;

    btn.appendChild(boldText);
    btn.appendChild(smallText);

    btn.onclick = () => {
      if (wid !== AppState.activeWeekId) {
        loadWeek(wid);
      }
    };

    list.appendChild(btn);
  });
}

/**
 * Start a new week
 */
function startNewWeek() {
  saveWeek();
  const now = new Date();
  const day = now.getDay();
  const nextMon = new Date(now);
  nextMon.setDate(now.getDate() + (7 - ((day + 6) % 7)));
  loadWeek(weekId(nextMon));
}

// ── Generic List Renderer ────────────────────────────────────────────────────

/**
 * @typedef {Object} ListOptions
 * @property {string} placeholder - Input placeholder text
 * @property {function(number, string): void} onUpdate - Called when item updated
 * @property {function(number): void} onRemove - Called when item removed
 * @property {boolean} [showCarry] - Show carry-forward button
 * @property {function(number): boolean} [getCarryState] - Get carry state for index
 * @property {function(number): void} [onToggleCarry] - Toggle carry state
 */

/**
 * Render an editable list (meetings, bullets, sync items)
 * @param {string} containerId - Container element ID
 * @param {string[]|Bullet[]} items - Items to render
 * @param {ListOptions} options - Rendering options
 */
function renderEditableList(containerId, items, options) {
  const container = document.getElementById(containerId);
  if (!container) return;

  // Clear existing
  while (container.firstChild) {
    container.removeChild(container.firstChild);
  }

  items.forEach((item, i) => {
    const row = document.createElement('div');
    row.className = 'bullet-row';

    const input = document.createElement('input');
    input.type = 'text';
    input.value = typeof item === 'object' ? item.text : item;
    input.placeholder = options.placeholder;
    input.setAttribute('aria-label', `${options.placeholder} ${i + 1}`);
    input.oninput = () => {
      options.onUpdate(i, input.value);
      debounce();
      update();
    };

    row.appendChild(input);

    // Carry-forward button for bullets
    if (options.showCarry && options.getCarryState && options.onToggleCarry) {
      const carryBtn = document.createElement('button');
      const isCarry = options.getCarryState(i);
      carryBtn.className = 'btn-carry' + (isCarry ? ' on' : '');
      carryBtn.title = isCarry ? 'Will carry to next week' : 'Carry to next week';
      carryBtn.setAttribute('aria-label', isCarry ? 'Will carry to next week' : 'Carry to next week');
      carryBtn.setAttribute('aria-pressed', isCarry ? 'true' : 'false');
      carryBtn.textContent = '\u2192';
      carryBtn.onclick = () => options.onToggleCarry(i);
      row.appendChild(carryBtn);
    }

    // Remove button
    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.title = 'Remove';
    removeBtn.setAttribute('aria-label', 'Remove item');
    removeBtn.textContent = '\u2715';
    removeBtn.onclick = () => {
      options.onRemove(i);
      debounce();
      update();
    };

    row.appendChild(removeBtn);
    container.appendChild(row);
  });
}

// ── Render Functions ─────────────────────────────────────────────────────────

/**
 * Render all editable sections
 */
function renderAll() {
  renderPresets();
  renderMeetings();
  renderBullets();
  renderSyncItems();
  bindInputs();
}

/**
 * Render preset checkboxes
 */
function renderPresets() {
  const g = document.getElementById('presetsGrid');
  if (!g) return;

  // Clear existing
  while (g.firstChild) {
    g.removeChild(g.firstChild);
  }

  const presets = getPresets();
  presets.forEach((p, i) => {
    if (AppState.presetChecks[i] === undefined) {
      AppState.presetChecks[i] = false;
    }

    const row = document.createElement('label');
    row.className = 'preset-row';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = AppState.presetChecks[i];
    checkbox.setAttribute('aria-label', p);
    checkbox.onchange = () => {
      AppState.presetChecks[i] = checkbox.checked;
      debounce();
      update();
    };

    const span = document.createElement('span');
    span.textContent = p;

    row.appendChild(checkbox);
    row.appendChild(span);
    g.appendChild(row);
  });
}

/**
 * Render meetings list
 */
function renderMeetings() {
  renderEditableList('meetingsList', AppState.meetings, {
    placeholder: 'Meeting name...',
    onUpdate: (i, val) => { AppState.meetings[i] = val; },
    onRemove: (i) => {
      AppState.meetings.splice(i, 1);
      renderMeetings();
    }
  });
}

/**
 * Render bullets list
 */
function renderBullets() {
  renderEditableList('bulletsList', AppState.bullets, {
    placeholder: 'What was completed or updated this week?',
    onUpdate: (i, val) => { AppState.bullets[i].text = val; },
    onRemove: (i) => {
      AppState.bullets.splice(i, 1);
      renderBullets();
    },
    showCarry: true,
    getCarryState: (i) => AppState.bullets[i].carryForward,
    onToggleCarry: (i) => {
      AppState.bullets[i].carryForward = !AppState.bullets[i].carryForward;
      renderBullets();
      debounce();
      update();
    }
  });
}

/**
 * Render sync items list
 */
function renderSyncItems() {
  renderEditableList('syncItems', AppState.syncItems, {
    placeholder: 'Agenda item...',
    onUpdate: (i, val) => { AppState.syncItems[i] = val; },
    onRemove: (i) => {
      AppState.syncItems.splice(i, 1);
      renderSyncItems();
    }
  });
}

/**
 * Add a new meeting
 */
function addMeeting() {
  AppState.meetings.push('');
  renderMeetings();
  update();
  // Focus the new input
  const inputs = document.querySelectorAll('#meetingsList input');
  if (inputs.length > 0) {
    inputs[inputs.length - 1].focus();
  }
}

/**
 * Add a new bullet
 */
function addBullet() {
  AppState.bullets.push({ text: '', carryForward: false });
  renderBullets();
  update();
  const inputs = document.querySelectorAll('#bulletsList input');
  if (inputs.length > 0) {
    inputs[inputs.length - 1].focus();
  }
}

/**
 * Add a new sync item
 */
function addSyncItem() {
  AppState.syncItems.push('');
  renderSyncItems();
  update();
  const inputs = document.querySelectorAll('#syncItems input');
  if (inputs.length > 0) {
    inputs[inputs.length - 1].focus();
  }
}

/**
 * Bind input change handlers
 */
function bindInputs() {
  const fields = [
    'teamName', 'dateRange', 'metricPeriod', 'totalConversations',
    'medianResponseTime', 'responseGoal', 'syncMeeting', 'privateNotes', 'notesForNextWeek'
  ];

  fields.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.oninput = () => {
        debounce();
        update();
      };
    }
  });

  const templateSel = document.getElementById('templateSelect');
  if (templateSel) {
    templateSel.onchange = () => {
      AppState.activeTemplate = templateSel.value;
      localStorage.setItem('ss_report_template', AppState.activeTemplate);
      AppState.presetChecks = {};
      renderPresets();
      update();
    };
  }
}

// ── Template Select Population ───────────────────────────────────────────────

/**
 * Populate the template select dropdown
 */
function populateTemplateSelect() {
  const sel = document.getElementById('templateSelect');
  if (!sel) return;

  // Clear existing options
  while (sel.firstChild) {
    sel.removeChild(sel.firstChild);
  }

  // Add built-in templates
  const builtInGroup = document.createElement('optgroup');
  builtInGroup.label = 'Built-in';
  BUILT_IN_TEMPLATES.forEach(t => {
    const opt = document.createElement('option');
    opt.value = t.id;
    opt.textContent = t.name;
    builtInGroup.appendChild(opt);
  });
  sel.appendChild(builtInGroup);

  // Add custom templates if any
  if (AppState.customTemplates.length > 0) {
    const customGroup = document.createElement('optgroup');
    customGroup.label = 'Custom';
    AppState.customTemplates.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.id;
      opt.textContent = t.name;
      customGroup.appendChild(opt);
    });
    sel.appendChild(customGroup);
  }
}

// ── Goal Indicator ───────────────────────────────────────────────────────────

/**
 * Update the goal indicator display
 */
function updateGoalIndicator() {
  const medianStr = getF('medianResponseTime');
  const goalStr = getF('responseGoal');
  const el = document.getElementById('goalIndicator');

  if (!el) return;

  if (!medianStr || !goalStr) {
    while (el.firstChild) el.removeChild(el.firstChild);
    return;
  }

  const medianSec = parseTimeToSeconds(medianStr);
  const goalSec = parseTimeToSeconds(goalStr);

  if (!medianSec || !goalSec) {
    while (el.firstChild) el.removeChild(el.firstChild);
    return;
  }

  const pct = medianSec / goalSec;
  let color, label, icon;

  if (pct <= 0.85) {
    color = '#059669';
    label = 'Under goal';
    icon = '\u25CF'; // filled circle
  } else if (pct <= 1.0) {
    color = '#d97706';
    label = 'Near goal';
    icon = '\u25CB'; // empty circle
  } else {
    color = '#dc2626';
    label = 'Over goal';
    icon = '\u25B2'; // triangle
  }

  // Clear and rebuild
  while (el.firstChild) el.removeChild(el.firstChild);

  const span = document.createElement('span');
  span.style.color = color;
  span.style.fontSize = '10px';
  span.setAttribute('role', 'status');
  span.setAttribute('aria-label', `${label}: goal is ${goalStr}`);
  span.textContent = `${icon} ${label} (goal: ${goalStr})`;

  el.appendChild(span);
}

// ── Metric Hints ─────────────────────────────────────────────────────────────

/**
 * Update metric comparison hints
 */
function updateMetricHints() {
  const store = getStore();
  const sorted = [...store.order].sort((a, b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid !== AppState.activeWeekId);
  const convHint = document.getElementById('convHint');

  if (!convHint) return;

  // Clear existing
  while (convHint.firstChild) convHint.removeChild(convHint.firstChild);

  if (prevWid && store.weeks[prevWid]) {
    const prev = store.weeks[prevWid];
    const parts = prevWid.split('~');
    const label = parts[0] || prevWid;
    const prevVal = parseInt(prev.totalConversations || '0', 10);
    const currVal = parseInt(getF('totalConversations') || '0', 10);

    const textNode = document.createTextNode(
      `Last week (${label}): ${prevVal > 0 ? Number(prevVal).toLocaleString() : '\u2014'}`
    );
    convHint.appendChild(textNode);

    if (prevVal > 0 && currVal > 0 && prevVal !== currVal) {
      const pct = Math.round(((currVal - prevVal) / prevVal) * 100);
      const cls = pct > 0 ? 'up' : (pct < 0 ? 'down' : 'same');
      const arrow = pct > 0 ? '\u2191' : '\u2193';

      const compSpan = document.createElement('span');
      compSpan.className = 'metric-comparison ' + cls;
      compSpan.setAttribute('aria-label', `${Math.abs(pct)}% ${pct > 0 ? 'increase' : 'decrease'} from last week`);
      compSpan.textContent = ` (${arrow} ${Math.abs(pct)}%)`;
      convHint.appendChild(compSpan);
    }
  }
}

/**
 * Update copy from last week button visibility
 */
function updateCopyLastWeekVisibility() {
  const store = getStore();
  const sorted = [...store.order].sort((a, b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid < AppState.activeWeekId);
  const btn = document.getElementById('copyLastWeekBtn');
  if (btn) {
    btn.style.display = prevWid ? 'block' : 'none';
  }
}

/**
 * Copy data from last week
 */
function copyFromLastWeek() {
  const store = getStore();
  const sorted = [...store.order].sort((a, b) => b.localeCompare(a));
  const prevWid = sorted.find(wid => wid < AppState.activeWeekId);

  if (!prevWid || !store.weeks[prevWid]) {
    showToast('No previous week data found', 'warning');
    return;
  }

  const prev = store.weeks[prevWid];

  AppState.meetings = (prev.meetings || []).map(sanitize);
  AppState.bullets = (prev.bullets || []).map(b =>
    typeof b === 'string'
      ? { text: b, carryForward: false }
      : { text: b.text, carryForward: false }
  );
  AppState.syncItems = (prev.syncItems || []).map(sanitize);
  AppState.presetChecks = { ...(prev.presetChecks || {}) };
  setF('notesForNextWeek', prev.notesForNextWeek || '');
  AppState.notesForNextWeek = prev.notesForNextWeek || '';

  renderAll();
  debounce();
  update();
  showToast('Copied from last week', 'success');
}

// ── Report Generation ────────────────────────────────────────────────────────

/**
 * Generate report content
 * @param {'plain'|'slack'} format - Output format
 * @returns {string}
 */
function generateReport(format) {
  const isSlack = format === 'slack';
  const bold = (text) => isSlack ? `*${text}*` : text;

  let out = '';

  // Header
  const team = getF('teamName').trim();
  const dr = getF('dateRange').trim();
  if (team || dr) {
    out += bold(`${team} Team Weekly Update ${dr}`) + '\n';
  }

  // Metrics section
  const period = getF('metricPeriod').trim();
  const total = getF('totalConversations').trim();
  const median = getF('medianResponseTime').trim();
  const goal = getF('responseGoal').trim();

  if (period || total || median || goal) {
    out += isSlack ? '\n' : '';
    out += bold(`CS Router Metrics: ${period}`) + '\n';
    if (total) out += `Total Conversations: ${Number(total).toLocaleString()}\n`;
    if (median) out += `Median Response Time: ${median}\n`;
    if (goal) out += `Response Time Goal: ${goal}\n`;
  }

  // Meetings
  const mtgs = AppState.meetings.filter(m => m.trim());
  if (mtgs.length) {
    out += isSlack ? '\n' : '';
    mtgs.forEach(m => out += `${m}\n`);
  }

  // Accomplishments (presets + bullets)
  const checkedPresets = getPresets().filter((_, i) => AppState.presetChecks[i]);
  const filledBullets = AppState.bullets.filter(b => b.text.trim()).map(b => b.text);
  const allBullets = [...checkedPresets, ...filledBullets];

  if (allBullets.length) {
    out += isSlack ? '\n' : '';
    allBullets.forEach(b => out += `\u2022 ${b}\n`);
  }

  // Notes for next week
  const notesNext = getF('notesForNextWeek').trim();
  if (notesNext) {
    out += '\n' + bold('Notes for Next Week:') + '\n' + notesNext + '\n';
  }

  // Sync items
  const syncMtg = getF('syncMeeting').trim();
  const sItems = AppState.syncItems.filter(s => s.trim());

  if (syncMtg || sItems.length) {
    out += '\n';
    if (syncMtg) out += bold(syncMtg) + '\n';
    sItems.forEach(s => out += `\u2022 ${s}\n`);
  }

  return out.trim();
}

/**
 * Generate plain text report (legacy wrapper)
 * @returns {string}
 */
function generatePlain() {
  return generateReport('plain');
}

/**
 * Generate Slack-formatted report (legacy wrapper)
 * @returns {string}
 */
function generateSlack() {
  return generateReport('slack');
}

// ── Update UI ────────────────────────────────────────────────────────────────

/**
 * Update the preview and related UI elements
 */
function update() {
  const text = generatePlain();
  const previewBox = document.getElementById('previewBox');

  if (previewBox) {
    previewBox.textContent = text || 'Start filling in the form to see your report...';
  }

  // Character/word count
  const len = text.length;
  const ccEl = document.getElementById('charCount');

  if (ccEl) {
    while (ccEl.firstChild) ccEl.removeChild(ccEl.firstChild);

    const charSpan = document.createElement('span');
    if (len > SLACK_CHAR_LIMIT) {
      charSpan.className = 'char-danger';
    } else if (len > SLACK_CHAR_LIMIT * 0.85) {
      charSpan.className = 'char-warn';
    }
    charSpan.textContent = len.toLocaleString();

    const wordCount = text.split(/\s+/).filter(Boolean).length;
    const wordSpan = document.createElement('span');
    wordSpan.className = charSpan.className;
    wordSpan.textContent = String(wordCount);

    ccEl.appendChild(charSpan);
    ccEl.appendChild(document.createTextNode(' chars \u00A0|\u00A0 '));
    ccEl.appendChild(wordSpan);
    ccEl.appendChild(document.createTextNode(' words'));
  }

  updateGoalIndicator();
  updateMetricHints();
  showReminderBanner();
}

// ── Copy Functions ───────────────────────────────────────────────────────────

/**
 * Copy plain text to clipboard
 */
async function copyPlain() {
  const text = generatePlain();
  if (!text) {
    showToast('Nothing to copy', 'warning');
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    flash('copyBtn', '\u2713 Copied!', 'btn-copy copied');
    showToast('Copied plain text', 'success', 1500);
  } catch (e) {
    console.error('Copy failed:', e);
    showToast('Copy failed. Please try again.', 'error');
  }
}

/**
 * Copy Slack-formatted text to clipboard
 */
async function copySlack() {
  const text = generateSlack();
  if (!text) {
    showToast('Nothing to copy', 'warning');
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    flash('slackBtn', '\u2713 Copied!', 'btn-slack copied');
    showToast('Copied for Slack', 'success', 1500);
  } catch (e) {
    console.error('Copy failed:', e);
    showToast('Copy failed. Please try again.', 'error');
  }
}

/**
 * Flash button with temporary state
 * @param {string} id - Button ID
 * @param {string} label - Temporary label
 * @param {string} cls - Temporary class
 */
function flash(id, label, cls) {
  const btn = document.getElementById(id);
  if (!btn) return;

  const orig = btn.textContent;
  const origCls = btn.className;
  btn.textContent = label;
  btn.className = cls;

  setTimeout(() => {
    btn.textContent = orig;
    btn.className = origCls;
  }, 2000);
}

// ── Reset ────────────────────────────────────────────────────────────────────

/**
 * Reset current week to defaults
 */
function resetWeek() {
  if (!confirm('Reset this week to blank defaults?')) return;

  const store = getStore();
  delete store.weeks[AppState.activeWeekId];
  store.order = store.order.filter(id => id !== AppState.activeWeekId);
  saveStore(store);
  loadWeek(AppState.activeWeekId);
  showToast('Week reset to defaults', 'info');
}

// ── Save Status ──────────────────────────────────────────────────────────────

/**
 * Show saving status
 */
function setSaveStatus() {
  const el = document.getElementById('saveStatus');
  if (el) {
    el.style.color = '#d97706';
    el.textContent = '\u25CF Saving\u2026';
    el.setAttribute('aria-label', 'Saving');
  }
}

/**
 * Show saved status
 */
function showSaved() {
  const el = document.getElementById('saveStatus');
  if (el) {
    el.style.color = '#059669';
    el.textContent = '\u2713 Saved';
    el.setAttribute('aria-label', 'Saved');

    clearTimeout(AppState.statusTimer);
    AppState.statusTimer = setTimeout(() => {
      el.textContent = '';
    }, STATUS_CLEAR_DELAY);
  }
}

// ── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Get form field value
 * @param {string} id - Element ID
 * @returns {string}
 */
function getF(id) {
  const el = document.getElementById(id);
  return el ? el.value || '' : '';
}

/**
 * Set form field value
 * @param {string} id - Element ID
 * @param {string} val - Value to set
 */
function setF(id, val) {
  const el = document.getElementById(id);
  if (el) el.value = val;
}

// ── How to Use Toggle ────────────────────────────────────────────────────────

/**
 * Toggle how-to section visibility
 */
function toggleHowTo() {
  const btn = document.getElementById('howToToggle');
  const content = document.getElementById('howToContent');

  if (!btn || !content) return;

  const isOpen = content.classList.toggle('is-open');
  btn.setAttribute('aria-expanded', String(isOpen));
}

// ── Print ────────────────────────────────────────────────────────────────────

/**
 * Print the report
 */
function printReport() {
  const text = generatePlain();
  const container = document.getElementById('printContainer');

  if (!container) return;

  container.textContent = text || 'No content to print.';
  container.style.display = 'block';
  window.print();
  container.style.display = 'none';
}

// ── Export / Import ──────────────────────────────────────────────────────────

/**
 * Export all data as JSON
 */
function exportData() {
  const store = getStore();
  const blob = new Blob([JSON.stringify(store, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `ss-report-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('Data exported', 'success');
}

/**
 * Import data from JSON file
 * @param {Event} ev - File input change event
 */
function importData(ev) {
  const file = ev.target.files?.[0];
  if (!file) return;

  const r = new FileReader();
  r.onload = () => {
    try {
      const data = JSON.parse(r.result);
      const validation = validateImportData(data);

      if (!validation.valid) {
        showToast(validation.error || 'Invalid backup file', 'error');
        return;
      }

      // Migrate if needed
      const migrated = migrateData(data);

      if (saveStore(migrated)) {
        loadWeek(AppState.activeWeekId);
        showToast('Data imported successfully', 'success');
      }
    } catch (e) {
      console.error('Import failed:', e);
      showToast('Could not parse file: ' + e.message, 'error');
    }
  };

  r.onerror = () => {
    showToast('Failed to read file', 'error');
  };

  r.readAsText(file);
  ev.target.value = '';
}

// ── Dark Mode ────────────────────────────────────────────────────────────────

/**
 * Toggle dark mode
 */
function toggleDarkMode() {
  document.body.classList.toggle('dark');
  const isDark = document.body.classList.contains('dark');
  localStorage.setItem('ss_report_dark', isDark ? '1' : '0');
  updateDarkModeIcon();
}

/**
 * Update dark mode toggle icon
 */
function updateDarkModeIcon() {
  const btn = document.getElementById('darkModeToggle');
  if (!btn) return;

  const isDark = document.body.classList.contains('dark');
  btn.textContent = isDark ? '\u2600' : '\uD83C\uDF19';
  btn.title = isDark ? 'Switch to light mode' : 'Switch to dark mode';
  btn.setAttribute('aria-label', isDark ? 'Switch to light mode' : 'Switch to dark mode');
}

// ── Reminder Banner ──────────────────────────────────────────────────────────

/**
 * Show/hide the reminder banner based on day of week
 */
function showReminderBanner() {
  const banner = document.getElementById('reminderBanner');
  const textEl = document.getElementById('reminderText');

  if (!banner || !textEl) return;

  const now = new Date();
  const day = now.getDay();
  const fri = day === 5;
  const satSun = day === 0 || day === 6;
  const hasMetrics = parseInt(getF('totalConversations') || '0', 10) > 0;

  let message = '';
  let show = false;

  if (fri && !hasMetrics) {
    show = true;
    message = 'Report due today \u2014 fill in your metrics and accomplishments.';
  } else if (fri) {
    show = true;
    message = "Report due today \u2014 don't forget to copy and submit.";
  } else if (day === 4) {
    show = true;
    message = 'Report due tomorrow (Friday).';
  } else if (!satSun) {
    const daysToFri = (5 - day + 7) % 7;
    show = true;
    message = `Report due in ${daysToFri} day${daysToFri === 1 ? '' : 's'} (Friday).`;
  }

  banner.style.display = show ? 'block' : 'none';
  banner.setAttribute('aria-hidden', show ? 'false' : 'true');
  textEl.textContent = message;
}

// ── Keyboard Shortcuts ───────────────────────────────────────────────────────

/**
 * Setup keyboard shortcuts
 */
function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    // Ctrl+Enter: Copy for Slack
    if (e.ctrlKey && e.key === 'Enter') {
      e.preventDefault();
      copySlack();
    }
    // Ctrl+Shift+C: Copy plain
    else if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'c') {
      e.preventDefault();
      copyPlain();
    }
    // Ctrl+Z: Undo
    else if (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === 'z') {
      // Only if not in an input/textarea
      if (document.activeElement?.tagName !== 'INPUT' && document.activeElement?.tagName !== 'TEXTAREA') {
        e.preventDefault();
        undo();
      }
    }
    // Ctrl+Y or Ctrl+Shift+Z: Redo
    else if ((e.ctrlKey && e.key.toLowerCase() === 'y') || (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'z')) {
      if (document.activeElement?.tagName !== 'INPUT' && document.activeElement?.tagName !== 'TEXTAREA') {
        e.preventDefault();
        redo();
      }
    }
    // Escape: Close modals
    else if (e.key === 'Escape') {
      hideConflictModal();
      closeMobileMenu();
    }
  });
}

// ── Mobile Menu ──────────────────────────────────────────────────────────────

/**
 * Toggle mobile sidebar menu
 */
function toggleMobileMenu() {
  const sidebar = document.querySelector('.history-pane');
  const overlay = document.getElementById('mobileOverlay');

  if (sidebar) {
    sidebar.classList.toggle('is-open');
  }
  if (overlay) {
    overlay.classList.toggle('is-visible');
  }
}

/**
 * Close mobile menu
 */
function closeMobileMenu() {
  const sidebar = document.querySelector('.history-pane');
  const overlay = document.getElementById('mobileOverlay');

  if (sidebar) {
    sidebar.classList.remove('is-open');
  }
  if (overlay) {
    overlay.classList.remove('is-visible');
  }
}

// ── Initialize on Load ───────────────────────────────────────────────────────

init();
