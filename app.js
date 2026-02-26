// ══════════════════════════════════════════════════════════════════════════════
// Weekly Update Builder - Refactored with Security, Accessibility & Modern Patterns
// ══════════════════════════════════════════════════════════════════════════════

// ── Type Definitions (JSDoc) ─────────────────────────────────────────────────

/**
 * @typedef {Object} Bullet
 * @property {string} id - Unique identifier for this item
 * @property {string} text - The bullet text content
 * @property {boolean} carryForward - Whether to carry to next week
 * @property {string|null} addedBy - Display name of user who added this item
 * @property {string|null} addedAt - ISO timestamp when item was added
 */

/**
 * @typedef {Object} Meeting
 * @property {string} id - Unique identifier for this item
 * @property {string} text - The meeting name
 * @property {string|null} addedBy - Display name of user who added this item
 * @property {string|null} addedAt - ISO timestamp when item was added
 */

/**
 * @typedef {Object} SyncItem
 * @property {string} id - Unique identifier for this item
 * @property {string} text - The agenda item text
 * @property {string|null} addedBy - Display name of user who added this item
 * @property {string|null} addedAt - ISO timestamp when item was added
 */

/**
 * @typedef {Object} FieldEdit
 * @property {string} editedBy - Display name of user who last edited this field
 * @property {string} editedAt - ISO timestamp when field was last edited
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
 * @property {Meeting[]} meetings
 * @property {Bullet[]} bullets
 * @property {SyncItem[]} syncItems
 * @property {Object<number, boolean>} presetChecks
 * @property {Object<string, FieldEdit>} fieldEdits - Track who edited each field
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
 * @property {Meeting[]} meetings
 * @property {Bullet[]} bullets
 * @property {SyncItem[]} syncItems
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
const CURRENT_DATA_VERSION = 4;
const MAX_HISTORY_SIZE = 50;
const DEBOUNCE_DELAY = 700;
const TOAST_DURATION = 3000;
const STATUS_CLEAR_DELAY = 3000;

// Firestore in-memory store when signed in (real-time synced)
let firestoreStoreCache = null;

// ── Application State ────────────────────────────────────────────────────────

const AppState = {
  activeWeekId: '',
  meetings: [],
  bullets: [],
  syncItems: [],
  presetChecks: {},
  notesForNextWeek: '',
  fieldEdits: {}, // Track who edited each field
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

/**
 * Generate a unique ID for items
 * @returns {string}
 */
function generateId() {
  return 'id-' + Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 9);
}

// ── Data Migration ───────────────────────────────────────────────────────────

/**
 * Migrate bullets from string format to object format
 * @param {any} bullet
 * @returns {Bullet}
 */
function migrateBullet(bullet) {
  if (typeof bullet === 'string') {
    return { id: generateId(), text: bullet, carryForward: false, addedBy: null, addedAt: null };
  }
  if (typeof bullet === 'object' && bullet !== null) {
    return {
      id: bullet.id || generateId(),
      text: sanitize(bullet.text),
      carryForward: Boolean(bullet.carryForward),
      addedBy: bullet.addedBy || null,
      addedAt: bullet.addedAt || null
    };
  }
  return { id: generateId(), text: '', carryForward: false, addedBy: null, addedAt: null };
}

/**
 * Migrate meetings from string format to object format
 * @param {any} meeting
 * @returns {Meeting}
 */
function migrateMeeting(meeting) {
  if (typeof meeting === 'string') {
    return { id: generateId(), text: meeting, addedBy: null, addedAt: null };
  }
  if (typeof meeting === 'object' && meeting !== null) {
    return {
      id: meeting.id || generateId(),
      text: sanitize(meeting.text),
      addedBy: meeting.addedBy || null,
      addedAt: meeting.addedAt || null
    };
  }
  return { id: generateId(), text: '', addedBy: null, addedAt: null };
}

/**
 * Migrate sync items from string format to object format
 * @param {any} item
 * @returns {SyncItem}
 */
function migrateSyncItem(item) {
  if (typeof item === 'string') {
    return { id: generateId(), text: item, addedBy: null, addedAt: null };
  }
  if (typeof item === 'object' && item !== null) {
    return {
      id: item.id || generateId(),
      text: sanitize(item.text),
      addedBy: item.addedBy || null,
      addedAt: item.addedAt || null
    };
  }
  return { id: generateId(), text: '', addedBy: null, addedAt: null };
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

  if (version < 3) {
    // Migrate meetings and syncItems from string to object format
    // Also ensure bullets have attribution fields
    if (data.weeks) {
      for (const wid of Object.keys(data.weeks)) {
        const week = data.weeks[wid];
        if (week.meetings && Array.isArray(week.meetings)) {
          week.meetings = week.meetings.map(migrateMeeting);
        }
        if (week.syncItems && Array.isArray(week.syncItems)) {
          week.syncItems = week.syncItems.map(migrateSyncItem);
        }
        if (week.bullets && Array.isArray(week.bullets)) {
          week.bullets = week.bullets.map(migrateBullet);
        }
      }
    }
    data.version = 3;
  }

  if (version < 4) {
    // Ensure all items have unique IDs
    if (data.weeks) {
      for (const wid of Object.keys(data.weeks)) {
        const week = data.weeks[wid];
        if (week.meetings && Array.isArray(week.meetings)) {
          week.meetings = week.meetings.map(m => ({ ...m, id: m.id || generateId() }));
        }
        if (week.syncItems && Array.isArray(week.syncItems)) {
          week.syncItems = week.syncItems.map(s => ({ ...s, id: s.id || generateId() }));
        }
        if (week.bullets && Array.isArray(week.bullets)) {
          week.bullets = week.bullets.map(b => ({ ...b, id: b.id || generateId() }));
        }
      }
    }
    data.version = 4;
  }

  return data;
}

// ── Storage ──────────────────────────────────────────────────────────────────

/**
 * Get store: from Firestore cache when signed in, else localStorage (with migration)
 * @returns {StoreData}
 */
function getStore() {
  if (window.firebaseAuth && window.firebaseAuth.isSignedIn() && firestoreStoreCache) {
    return migrateData({ ...firestoreStoreCache, weeks: { ...firestoreStoreCache.weeks } });
  }
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
 * Save store to localStorage and, when signed in, to Firestore
 * @param {StoreData} data
 * @returns {boolean} Success
 */
function saveStore(data) {
  try {
    data.version = CURRENT_DATA_VERSION;
    localStorage.setItem(LS_KEY, JSON.stringify(data));
    if (window.firebaseAuth && window.firebaseAuth.isSignedIn()) {
      window.firebaseAuth.saveStoreToFirestore(data);
    }
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
 * Called by firebase.js when Firestore store doc is updated (real-time sync)
 * @param {StoreData} store
 */
function onFirestoreStoreUpdated(store) {
  if (!store || !AppState.activeWeekId) return;
  loadWeek(AppState.activeWeekId);
}

/**
 * Set in-memory Firestore store (called by firebase.js on snapshot)
 * @param {StoreData} store
 */
function setFirestoreStore(store) {
  firestoreStoreCache = store ? { weeks: store.weeks || {}, order: store.order || [], version: store.version || CURRENT_DATA_VERSION } : null;
}

/**
 * Clear Firestore cache (called by firebase.js on sign-out)
 */
function clearFirestoreStore() {
  firestoreStoreCache = null;
}

/**
 * Get current user's display name for attribution
 * @returns {string|null}
 */
function getCurrentUserDisplayName() {
  if (window.firebaseAuth && window.firebaseAuth.isSignedIn()) {
    const user = window.firebaseAuth.getCurrentUser();
    if (user) {
      return user.displayName || user.email || 'Unknown';
    }
  }
  return null;
}

/**
 * Format a timestamp for display
 * @param {string|null} isoString
 * @returns {string}
 */
function formatAttributionDate(isoString) {
  if (!isoString) return '';
  try {
    const date = new Date(isoString);
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) +
      ', ' + date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });
  } catch {
    return '';
  }
}

/**
 * Generate a consistent color from a string (user name)
 * @param {string} str
 * @returns {string} HSL color string
 */
function stringToColor(str) {
  if (!str) return 'transparent';
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const hue = Math.abs(hash) % 360;
  return `hsl(${hue}, 70%, 50%)`;
}

// Expose for firebase.js
window.setFirestoreStore = setFirestoreStore;
window.clearFirestoreStore = clearFirestoreStore;

// ── Collaboration State ──────────────────────────────────────────────────────

const CollabState = {
  presenceUsers: [],
  comments: [],
  activities: [],
  expandedComments: {} // { itemId: true }
};

/**
 * Clear collaboration state (called on sign-out)
 */
function clearCollaborationState() {
  CollabState.presenceUsers = [];
  CollabState.comments = [];
  CollabState.activities = [];
  CollabState.expandedComments = {};
  renderPresenceAvatars();
  renderActivityFeed();
}
window.clearCollaborationState = clearCollaborationState;

/**
 * Render presence avatars in header
 */
function renderPresenceAvatars() {
  const container = document.getElementById('presenceAvatars');
  if (!container) return;

  container.innerHTML = '';
  const users = CollabState.presenceUsers;
  const currentUser = window.firebaseAuth?.getCurrentUser?.();
  const currentUid = currentUser?.uid;

  // Filter out current user and show max 5 avatars
  const otherUsers = users.filter(u => u.odingUserId !== currentUid);
  const displayUsers = otherUsers.slice(0, 5);
  const remaining = otherUsers.length - 5;

  displayUsers.forEach(user => {
    const avatar = document.createElement('div');
    avatar.className = 'presence-avatar';
    avatar.style.background = user.color || stringToColor(user.displayName);
    avatar.setAttribute('data-tooltip', user.displayName);

    if (user.photoURL) {
      const img = document.createElement('img');
      img.src = user.photoURL;
      img.alt = user.displayName;
      img.onerror = () => {
        img.remove();
        avatar.textContent = getInitials(user.displayName);
      };
      avatar.appendChild(img);
    } else {
      avatar.textContent = getInitials(user.displayName);
    }

    container.appendChild(avatar);
  });

  if (remaining > 0) {
    const countBadge = document.createElement('div');
    countBadge.className = 'presence-count';
    countBadge.textContent = '+' + remaining;
    countBadge.setAttribute('data-tooltip', remaining + ' more online');
    container.appendChild(countBadge);
  }
}

/**
 * Get initials from display name
 * @param {string} name
 * @returns {string}
 */
function getInitials(name) {
  if (!name) return '?';
  const parts = name.trim().split(' ');
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.substring(0, 2).toUpperCase();
}

/**
 * Format relative time for activity feed
 * @param {number} timestamp
 * @returns {string}
 */
function formatRelativeTime(timestamp) {
  const now = Date.now();
  const diff = now - timestamp;
  const minutes = Math.floor(diff / 60000);
  const hours = Math.floor(diff / 3600000);
  const days = Math.floor(diff / 86400000);

  if (minutes < 1) return 'just now';
  if (minutes < 60) return minutes + 'm ago';
  if (hours < 24) return hours + 'h ago';
  if (days < 7) return days + 'd ago';
  return new Date(timestamp).toLocaleDateString();
}

/**
 * Render activity feed in sidebar
 */
function renderActivityFeed() {
  const container = document.getElementById('activityFeed');
  const list = document.getElementById('activityList');
  if (!container || !list) return;

  const activities = CollabState.activities;
  const currentUser = window.firebaseAuth?.getCurrentUser?.();
  const currentUid = currentUser?.uid;

  // Show/hide feed based on sign-in status
  if (!window.firebaseAuth?.isSignedIn?.()) {
    container.style.display = 'none';
    return;
  }
  container.style.display = '';

  list.innerHTML = '';

  if (activities.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'activity-empty';
    empty.textContent = 'No recent activity';
    list.appendChild(empty);
    return;
  }

  activities.slice(0, 20).forEach(activity => {
    const item = document.createElement('div');
    item.className = 'activity-item';

    // Highlight if this activity mentions current user
    if (activity.mentions?.includes(currentUid)) {
      item.classList.add('is-mention');
    }

    const avatar = document.createElement('div');
    avatar.className = 'activity-avatar';
    avatar.style.background = activity.actorColor || stringToColor(activity.actorName);
    avatar.textContent = getInitials(activity.actorName);

    const content = document.createElement('div');
    content.className = 'activity-content';

    const text = document.createElement('div');
    text.className = 'activity-text';
    text.innerHTML = '<strong>' + escapeHtml(activity.actorName) + '</strong> ' + escapeHtml(activity.action);

    content.appendChild(text);

    if (activity.targetText) {
      const target = document.createElement('div');
      target.className = 'activity-target';
      target.textContent = '"' + activity.targetText + '"';
      content.appendChild(target);
    }

    const time = document.createElement('div');
    time.className = 'activity-time';
    time.textContent = formatRelativeTime(activity.timestamp);
    content.appendChild(time);

    item.appendChild(avatar);
    item.appendChild(content);
    list.appendChild(item);
  });
}

/**
 * Escape HTML for safe display
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── @Mentions System ─────────────────────────────────────────────────────────

const MentionState = {
  activeInput: null,
  dropdown: null,
  selectedIndex: 0,
  searchText: '',
  startPosition: 0
};

/**
 * Initialize mention detection on an input element
 * @param {HTMLInputElement} input
 */
function initMentionDetection(input) {
  input.addEventListener('input', handleMentionInput);
  input.addEventListener('keydown', handleMentionKeydown);
  input.addEventListener('blur', () => {
    // Delay to allow click on dropdown
    setTimeout(hideMentionDropdown, 150);
  });
}

/**
 * Handle input for mention detection
 * @param {Event} e
 */
function handleMentionInput(e) {
  const input = e.target;
  const value = input.value;
  const cursorPos = input.selectionStart;

  // Find @ before cursor
  const textBeforeCursor = value.substring(0, cursorPos);
  const atIndex = textBeforeCursor.lastIndexOf('@');

  if (atIndex === -1) {
    hideMentionDropdown();
    return;
  }

  // Check if @ is at start or after a space
  if (atIndex > 0 && textBeforeCursor[atIndex - 1] !== ' ') {
    hideMentionDropdown();
    return;
  }

  // Get search text after @
  const searchText = textBeforeCursor.substring(atIndex + 1);

  // Don't show if there's a space in the search (mention already completed)
  if (searchText.includes(' ') && searchText.length > 20) {
    hideMentionDropdown();
    return;
  }

  MentionState.activeInput = input;
  MentionState.searchText = searchText.toLowerCase();
  MentionState.startPosition = atIndex;
  MentionState.selectedIndex = 0;

  showMentionDropdown(input);
}

/**
 * Handle keydown for mention navigation
 * @param {KeyboardEvent} e
 */
function handleMentionKeydown(e) {
  if (!MentionState.dropdown || MentionState.dropdown.style.display === 'none') {
    return;
  }

  const items = MentionState.dropdown.querySelectorAll('.mention-item');
  if (items.length === 0) return;

  switch (e.key) {
    case 'ArrowDown':
      e.preventDefault();
      MentionState.selectedIndex = (MentionState.selectedIndex + 1) % items.length;
      updateMentionSelection(items);
      break;
    case 'ArrowUp':
      e.preventDefault();
      MentionState.selectedIndex = (MentionState.selectedIndex - 1 + items.length) % items.length;
      updateMentionSelection(items);
      break;
    case 'Enter':
    case 'Tab':
      if (items[MentionState.selectedIndex]) {
        e.preventDefault();
        const user = JSON.parse(items[MentionState.selectedIndex].dataset.user);
        insertMention(user);
      }
      break;
    case 'Escape':
      e.preventDefault();
      hideMentionDropdown();
      break;
  }
}

/**
 * Update visual selection in dropdown
 * @param {NodeList} items
 */
function updateMentionSelection(items) {
  items.forEach((item, i) => {
    item.classList.toggle('selected', i === MentionState.selectedIndex);
  });
}

/**
 * Show mention autocomplete dropdown
 * @param {HTMLInputElement} input
 */
function showMentionDropdown(input) {
  // Get users from presence
  const allUsers = CollabState.presenceUsers || [];
  const currentUser = window.firebaseAuth?.getCurrentUser?.();
  const currentUid = currentUser?.uid;

  // Filter users by search text, exclude current user
  const filteredUsers = allUsers.filter(user => {
    if (user.odingUserId === currentUid) return false;
    if (!MentionState.searchText) return true;
    return user.displayName.toLowerCase().includes(MentionState.searchText) ||
           (user.email && user.email.toLowerCase().includes(MentionState.searchText));
  }).slice(0, 5);

  if (filteredUsers.length === 0) {
    hideMentionDropdown();
    return;
  }

  // Create or update dropdown
  if (!MentionState.dropdown) {
    MentionState.dropdown = document.createElement('div');
    MentionState.dropdown.className = 'mention-dropdown';
    document.body.appendChild(MentionState.dropdown);
  }

  // Position dropdown below input
  const rect = input.getBoundingClientRect();
  MentionState.dropdown.style.left = rect.left + 'px';
  MentionState.dropdown.style.top = (rect.bottom + 4) + 'px';
  MentionState.dropdown.style.width = Math.min(rect.width, 280) + 'px';
  MentionState.dropdown.style.display = 'block';

  // Render users
  MentionState.dropdown.innerHTML = '';
  filteredUsers.forEach((user, i) => {
    const item = document.createElement('div');
    item.className = 'mention-item' + (i === MentionState.selectedIndex ? ' selected' : '');
    item.dataset.user = JSON.stringify(user);

    const avatar = document.createElement('div');
    avatar.className = 'mention-avatar';
    avatar.style.background = user.color || stringToColor(user.displayName);
    avatar.textContent = getInitials(user.displayName);

    const info = document.createElement('div');
    info.className = 'mention-info';

    const name = document.createElement('div');
    name.className = 'mention-name';
    name.textContent = user.displayName;

    const email = document.createElement('div');
    email.className = 'mention-email';
    email.textContent = user.email || '';

    info.appendChild(name);
    if (user.email) info.appendChild(email);

    item.appendChild(avatar);
    item.appendChild(info);

    item.onclick = () => insertMention(user);

    MentionState.dropdown.appendChild(item);
  });
}

/**
 * Hide mention dropdown
 */
function hideMentionDropdown() {
  if (MentionState.dropdown) {
    MentionState.dropdown.style.display = 'none';
  }
  MentionState.activeInput = null;
}

/**
 * Insert mention into input
 * @param {Object} user
 */
function insertMention(user) {
  const input = MentionState.activeInput;
  if (!input) return;

  const value = input.value;
  const before = value.substring(0, MentionState.startPosition);
  const after = value.substring(input.selectionStart);

  // Insert mention with special format: @[Name](uid:xxx)
  const mentionText = `@${user.displayName} `;
  const mentionData = `@[${user.displayName}](uid:${user.odingUserId})`;

  // For display, just show @Name, but store the full format
  input.value = before + mentionText + after;
  input.dataset.mentions = (input.dataset.mentions || '') + mentionData + '|';

  // Move cursor after mention
  const newPos = before.length + mentionText.length;
  input.setSelectionRange(newPos, newPos);
  input.focus();

  hideMentionDropdown();

  // Trigger input event to save
  input.dispatchEvent(new Event('input', { bubbles: true }));
}

/**
 * Parse mentions from text
 * @param {string} text
 * @returns {string[]} Array of user UIDs mentioned
 */
function parseMentions(text) {
  if (!text) return [];
  const mentionRegex = /@\[([^\]]+)\]\(uid:([^)]+)\)/g;
  const mentions = [];
  let match;
  while ((match = mentionRegex.exec(text)) !== null) {
    mentions.push(match[2]); // Push the UID
  }
  return mentions;
}

/**
 * Extract mentions from an input's data attribute
 * @param {HTMLInputElement} input
 * @returns {string[]} Array of user UIDs
 */
function getMentionsFromInput(input) {
  const mentionsData = input.dataset.mentions || '';
  return parseMentions(mentionsData);
}

/**
 * Handle presence update from Firebase
 * @param {Array} users
 */
function onPresenceUpdated(users) {
  CollabState.presenceUsers = users;
  renderPresenceAvatars();
}

/**
 * Handle activities update from Firebase
 * @param {Array} activities
 */
function onActivitiesUpdated(activities) {
  const currentUser = window.firebaseAuth?.getCurrentUser?.();
  const currentUid = currentUser?.uid;

  // Check for new mentions
  const oldActivities = CollabState.activities;
  if (oldActivities.length > 0 && activities.length > 0) {
    const latestOld = oldActivities[0]?.timestamp || 0;
    activities.forEach(activity => {
      if (activity.timestamp > latestOld &&
          activity.mentions?.includes(currentUid) &&
          activity.actorUid !== currentUid) {
        showToast(activity.actorName + ' mentioned you', 'info');
      }
    });
  }

  CollabState.activities = activities;
  renderActivityFeed();
}

/**
 * Handle comments update from Firebase
 * @param {Array} comments
 */
function onCommentsUpdated(comments) {
  CollabState.comments = comments;
  // Re-render item lists to show updated comment counts
  if (AppState.dirty.meetings || AppState.dirty.bullets || AppState.dirty.syncItems) {
    return; // Will be rendered soon anyway
  }
  renderMeetings();
  renderBullets();
  renderSyncItems();
}

/**
 * Log an activity
 * @param {string} type
 * @param {string} action
 * @param {string} targetText
 * @param {string[]} mentions
 */
function logActivity(type, action, targetText, mentions) {
  if (!window.firebaseAuth?.isSignedIn?.()) return;
  window.firebaseAuth.logActivity({
    weekId: AppState.activeWeekId,
    type: type,
    action: action,
    targetText: targetText || '',
    mentions: mentions || []
  });
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
    meetings: AppState.meetings.map(m => ({ ...m })),
    bullets: AppState.bullets.map(b => ({ ...b })),
    syncItems: AppState.syncItems.map(s => ({ ...s })),
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
  AppState.meetings = entry.meetings.map(m => ({ ...m }));
  AppState.bullets = entry.bullets.map(b => ({ ...b }));
  AppState.syncItems = entry.syncItems.map(s => ({ ...s }));
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

  // Firebase: init auth + real-time Firestore sync callback
  if (window.firebaseAuth) {
    window.firebaseAuth.init();
    window.firebaseAuth.setOnStoreUpdatedCallback(onFirestoreStoreUpdated);

    // Collaboration callbacks
    window.firebaseAuth.setOnPresenceUpdatedCallback(onPresenceUpdated);
    window.firebaseAuth.setOnActivitiesUpdatedCallback(onActivitiesUpdated);
    window.firebaseAuth.setOnCommentsUpdatedCallback(onCommentsUpdated);

    // Start presence heartbeat with current week getter
    window.firebaseAuth.startPresenceHeartbeat(() => AppState.activeWeekId);
  }

  // Handle page unload - remove presence
  window.addEventListener('beforeunload', () => {
    if (window.firebaseAuth?.removePresence) {
      window.firebaseAuth.removePresence();
    }
  });

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
    AppState.meetings = (saved.meetings || getDefMeetings()).map(migrateMeeting);
    AppState.bullets = (saved.bullets || [{ text: '', carryForward: false, addedBy: null, addedAt: null }]).map(migrateBullet);
    AppState.syncItems = (saved.syncItems || getDefSync()).map(migrateSyncItem);
    AppState.presetChecks = { ...(saved.presetChecks || {}) };
    AppState.fieldEdits = { ...(saved.fieldEdits || {}) };
    AppState.notesForNextWeek = sanitize(saved.notesForNextWeek) || '';
    renderFieldEdits();
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
    AppState.meetings = getDefMeetings().map(migrateMeeting);
    AppState.bullets = carryBullets.length > 0 ? carryBullets : [{ id: generateId(), text: '', carryForward: false, addedBy: null, addedAt: null }];
    AppState.syncItems = getDefSync().map(migrateSyncItem);
    AppState.presetChecks = {};
    AppState.fieldEdits = {};
    AppState.notesForNextWeek = carryNotes;
  }

  renderFieldEdits();
  updateCopyLastWeekVisibility();
  markAllDirty();
  renderAll();
  renderHistory();
  renderSparkline();
  update();
  clearHistory();

  // Subscribe to comments for this week
  if (window.firebaseAuth?.subscribeToComments) {
    window.firebaseAuth.subscribeToComments(wid);
  }

  // Update presence with current week
  if (window.firebaseAuth?.updatePresence) {
    window.firebaseAuth.updatePresence(wid);
  }
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
    meetings: AppState.meetings.map(m => ({ ...m })),
    bullets: AppState.bullets.map(b => ({ ...b })),
    syncItems: AppState.syncItems.map(s => ({ ...s })),
    presetChecks: { ...AppState.presetChecks },
    fieldEdits: { ...AppState.fieldEdits },
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
 * @param {Meeting[]|Bullet[]|SyncItem[]} items - Items to render
 * @param {ListOptions} options - Rendering options
 */
function renderEditableList(containerId, items, options) {
  const container = document.getElementById(containerId);
  if (!container) return;

  // Determine item type from container ID
  const itemType = containerId === 'meetingsList' ? 'meeting'
    : containerId === 'bulletsList' ? 'bullet'
    : containerId === 'syncItems' ? 'syncItem' : 'item';

  // Clear existing
  while (container.firstChild) {
    container.removeChild(container.firstChild);
  }

  items.forEach((item, i) => {
    const itemId = item.id || '';
    const wrapper = document.createElement('div');
    wrapper.className = 'bullet-item-wrapper';
    wrapper.setAttribute('data-item-id', itemId);

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

    // Initialize @mention detection if signed in
    if (window.firebaseAuth?.isSignedIn?.()) {
      initMentionDetection(input);
    }

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

    // Comment button (only if signed in)
    if (window.firebaseAuth?.isSignedIn?.() && itemId) {
      const commentCount = getCommentCountForItem(itemId);
      const commentBtn = document.createElement('button');
      commentBtn.className = 'btn-comment' + (commentCount > 0 ? ' has-comments' : '');
      commentBtn.title = commentCount > 0 ? `${commentCount} comment${commentCount > 1 ? 's' : ''}` : 'Add comment';
      commentBtn.setAttribute('aria-label', 'Comments');
      commentBtn.innerHTML = '\u{1F4AC}'; // Speech bubble emoji
      if (commentCount > 0) {
        const badge = document.createElement('span');
        badge.className = 'comment-badge';
        badge.textContent = commentCount;
        commentBtn.appendChild(badge);
      }
      commentBtn.onclick = () => toggleCommentThread(itemId, itemType, wrapper);
      row.appendChild(commentBtn);
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
    wrapper.appendChild(row);

    // Attribution line and colored border (if item has addedBy)
    if (typeof item === 'object' && item.addedBy) {
      const userColor = stringToColor(item.addedBy);
      wrapper.style.borderLeft = `3px solid ${userColor}`;
      wrapper.classList.add('has-attribution');

      const attribution = document.createElement('div');
      attribution.className = 'item-attribution';
      attribution.style.color = userColor;
      const dateStr = formatAttributionDate(item.addedAt);
      attribution.textContent = dateStr
        ? `Added by ${item.addedBy} on ${dateStr}`
        : `Added by ${item.addedBy}`;
      wrapper.appendChild(attribution);
    }

    // Render comment thread if expanded
    if (CollabState.expandedComments[itemId]) {
      renderCommentThread(itemId, itemType, wrapper);
    }

    container.appendChild(wrapper);
  });
}

/**
 * Get comment count for an item
 * @param {string} itemId
 * @returns {number}
 */
function getCommentCountForItem(itemId) {
  return CollabState.comments.filter(c => c.itemId === itemId && !c.resolved).length;
}

/**
 * Toggle comment thread visibility
 * @param {string} itemId
 * @param {string} itemType
 * @param {HTMLElement} wrapper
 */
function toggleCommentThread(itemId, itemType, wrapper) {
  if (CollabState.expandedComments[itemId]) {
    delete CollabState.expandedComments[itemId];
    const thread = wrapper.querySelector('.comment-thread');
    if (thread) thread.remove();
  } else {
    CollabState.expandedComments[itemId] = true;
    renderCommentThread(itemId, itemType, wrapper);
  }
}

/**
 * Render comment thread for an item
 * @param {string} itemId
 * @param {string} itemType
 * @param {HTMLElement} wrapper
 */
function renderCommentThread(itemId, itemType, wrapper) {
  // Remove existing thread if any
  const existing = wrapper.querySelector('.comment-thread');
  if (existing) existing.remove();

  const thread = document.createElement('div');
  thread.className = 'comment-thread';

  const comments = CollabState.comments.filter(c => c.itemId === itemId && !c.resolved);

  // Render existing comments
  comments.forEach(comment => {
    const commentEl = document.createElement('div');
    commentEl.className = 'comment-item';

    const avatar = document.createElement('div');
    avatar.className = 'comment-avatar';
    avatar.style.background = comment.authorColor || stringToColor(comment.authorName);
    avatar.textContent = getInitials(comment.authorName);

    const content = document.createElement('div');
    content.className = 'comment-content';

    const header = document.createElement('div');
    header.className = 'comment-header';
    header.innerHTML = '<strong>' + escapeHtml(comment.authorName) + '</strong> <span class="comment-time">' +
      formatRelativeTime(comment.createdAt) + '</span>';

    const text = document.createElement('div');
    text.className = 'comment-text';
    text.textContent = comment.text;

    content.appendChild(header);
    content.appendChild(text);

    commentEl.appendChild(avatar);
    commentEl.appendChild(content);

    // Delete button for own comments
    const currentUser = window.firebaseAuth?.getCurrentUser?.();
    if (currentUser && comment.authorUid === currentUser.uid) {
      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'comment-delete';
      deleteBtn.textContent = '\u2715';
      deleteBtn.title = 'Delete comment';
      deleteBtn.onclick = () => {
        window.firebaseAuth.deleteComment(comment.id).then(() => {
          showToast('Comment deleted', 'success');
        });
      };
      commentEl.appendChild(deleteBtn);
    }

    thread.appendChild(commentEl);
  });

  // Add comment input
  const inputRow = document.createElement('div');
  inputRow.className = 'comment-input-row';

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'comment-input';
  input.placeholder = 'Write a comment... (use @ to mention)';

  // Initialize mention detection
  initMentionDetection(input);

  input.onkeydown = (e) => {
    // Don't submit if mention dropdown is open
    if (MentionState.dropdown && MentionState.dropdown.style.display !== 'none') {
      return;
    }
    if (e.key === 'Enter' && input.value.trim()) {
      const mentions = getMentionsFromInput(input);
      submitComment(itemId, itemType, input.value.trim(), wrapper, mentions);
      input.value = '';
      input.dataset.mentions = '';
    }
  };

  const sendBtn = document.createElement('button');
  sendBtn.className = 'comment-send';
  sendBtn.textContent = 'Send';
  sendBtn.onclick = () => {
    if (input.value.trim()) {
      const mentions = getMentionsFromInput(input);
      submitComment(itemId, itemType, input.value.trim(), wrapper, mentions);
      input.value = '';
      input.dataset.mentions = '';
    }
  };

  inputRow.appendChild(input);
  inputRow.appendChild(sendBtn);
  thread.appendChild(inputRow);

  wrapper.appendChild(thread);

  // Focus the input
  input.focus();
}

/**
 * Submit a comment
 * @param {string} itemId
 * @param {string} itemType
 * @param {string} text
 * @param {HTMLElement} wrapper
 * @param {Array} mentions - Array of user UIDs being mentioned
 */
function submitComment(itemId, itemType, text, wrapper, mentions) {
  if (!window.firebaseAuth?.addComment) return;

  window.firebaseAuth.addComment({
    weekId: AppState.activeWeekId,
    itemId: itemId,
    itemType: itemType,
    text: text
  }).then(() => {
    showToast('Comment added', 'success');
    // Log activity with mentions for notifications
    if (mentions && mentions.length > 0) {
      logActivity('mention', 'mentioned you in a comment', text, mentions);
    }
    logActivity('comment', 'commented on a ' + itemType, text);
  }).catch(err => {
    console.error('Failed to add comment:', err);
    showToast('Failed to add comment', 'error');
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
    onUpdate: (i, val) => { AppState.meetings[i].text = val; },
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
    onUpdate: (i, val) => { AppState.syncItems[i].text = val; },
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
  const userName = getCurrentUserDisplayName();
  AppState.meetings.push({
    id: generateId(),
    text: '',
    addedBy: userName,
    addedAt: userName ? new Date().toISOString() : null
  });
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
  const userName = getCurrentUserDisplayName();
  AppState.bullets.push({
    id: generateId(),
    text: '',
    carryForward: false,
    addedBy: userName,
    addedAt: userName ? new Date().toISOString() : null
  });
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
  const userName = getCurrentUserDisplayName();
  AppState.syncItems.push({
    id: generateId(),
    text: '',
    addedBy: userName,
    addedAt: userName ? new Date().toISOString() : null
  });
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
        // Track who edited this field
        const userName = getCurrentUserDisplayName();
        if (userName) {
          AppState.fieldEdits[id] = {
            editedBy: userName,
            editedAt: new Date().toISOString()
          };
          updateFieldEditLabel(id);
        }
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

/**
 * Update field edit label for a specific field
 * @param {string} fieldId
 */
function updateFieldEditLabel(fieldId) {
  const el = document.getElementById(fieldId);
  if (!el) return;

  const edit = AppState.fieldEdits[fieldId];
  if (!edit) return;

  // Find wrapper - either .field or .card for textareas
  let wrapper = el.closest('.field');
  if (!wrapper) {
    wrapper = el.closest('.card');
  }
  if (!wrapper) return;

  let label = wrapper.querySelector(`.field-edit-label[data-field="${fieldId}"]`);
  if (!label) {
    label = document.createElement('div');
    label.className = 'field-edit-label';
    label.dataset.field = fieldId;
    // Insert after the element
    el.parentNode.insertBefore(label, el.nextSibling);
  }

  const userColor = stringToColor(edit.editedBy);
  label.style.color = userColor;
  label.textContent = `Edited by ${edit.editedBy}`;
  label.title = formatAttributionDate(edit.editedAt) || '';
}

/**
 * Render field edit labels for all fields
 */
function renderFieldEdits() {
  const fields = [
    'teamName', 'dateRange', 'metricPeriod', 'totalConversations',
    'medianResponseTime', 'responseGoal', 'syncMeeting', 'privateNotes', 'notesForNextWeek'
  ];

  fields.forEach(id => {
    const edit = AppState.fieldEdits[id];
    if (edit) {
      updateFieldEditLabel(id);
    } else {
      // Remove label if no edit info
      const label = document.querySelector(`.field-edit-label[data-field="${id}"]`);
      if (label) label.remove();
    }
  });
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

  AppState.meetings = (prev.meetings || []).map(migrateMeeting);
  AppState.bullets = (prev.bullets || []).map(b => {
    const migrated = migrateBullet(b);
    return { ...migrated, carryForward: false };
  });
  AppState.syncItems = (prev.syncItems || []).map(migrateSyncItem);
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
  const mtgs = AppState.meetings.filter(m => m.text && m.text.trim());
  if (mtgs.length) {
    out += isSlack ? '\n' : '';
    mtgs.forEach(m => out += `${m.text}\n`);
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
  const sItems = AppState.syncItems.filter(s => s.text && s.text.trim());

  if (syncMtg || sItems.length) {
    out += '\n';
    if (syncMtg) out += bold(syncMtg) + '\n';
    sItems.forEach(s => out += `\u2022 ${s.text}\n`);
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
