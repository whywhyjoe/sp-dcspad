// Central app state + debounced localStorage persistence.
// Kept as one serializable blob so a future SharePoint-backed storage
// layer (DevPadData/{user}.json) can replace localStorage wholesale.

const STORAGE_KEY = 'dcspad.v2.workspace';
const SAVE_DEBOUNCE_MS = 600;

const DEFAULTS = {
  html: '<div id="app">\n  <h2>Hello from DCSPad</h2>\n  <p>Edit HTML, CSS and JS, then press Run.</p>\n</div>\n',
  css: 'body {\n  font-family: "Segoe UI", sans-serif;\n  padding: 1rem;\n}\n',
  js: 'console.log("DCSPad ready", { when: new Date().toISOString() });\n',
  libraries: { enabled: [], pinned: ['pnpjs2'], custom: [] },
  settings: { autorun: false, jsAsModule: false, autoClearConsole: true, seenSplash: false, previewDark: true },
  layout: {
    sidebarW: 230, sidebarCollapsed: false,
    editorsFr: 1, runtimeFr: 1,
    previewFr: 1, diagH: 260,
    editorTab: 'js', diagTab: 'console',
  },
};

let state = load();
let saveTimer = null;
const listeners = new Set();

function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return structuredClone(DEFAULTS);
    const parsed = JSON.parse(raw);
    // Deep-merge over defaults so new fields added in later versions appear.
    return {
      ...structuredClone(DEFAULTS),
      ...parsed,
      libraries: { ...structuredClone(DEFAULTS.libraries), ...(parsed.libraries || {}) },
      settings: { ...structuredClone(DEFAULTS.settings), ...(parsed.settings || {}) },
      layout: { ...structuredClone(DEFAULTS.layout), ...(parsed.layout || {}) },
    };
  } catch {
    return structuredClone(DEFAULTS);
  }
}

function persist() {
  saveTimer = null;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    for (const fn of listeners) fn('saved');
  } catch (e) {
    console.warn('DCSPad: autosave failed', e);
  }
}

export function getState() { return state; }

export function update(patch) {
  Object.assign(state, patch);
  scheduleSave();
}

export function updateNested(section, patch) {
  Object.assign(state[section], patch);
  scheduleSave();
}

function scheduleSave() {
  for (const fn of listeners) fn('dirty');
  clearTimeout(saveTimer);
  saveTimer = setTimeout(persist, SAVE_DEBOUNCE_MS);
}

export function saveNow() {
  clearTimeout(saveTimer);
  persist();
}

export function onSaveStatus(fn) { listeners.add(fn); }

// Flush a save still sitting in the debounce window when the tab is
// closed, reloaded or backgrounded — otherwise the last ≤600 ms of edits
// are silently lost.
window.addEventListener('pagehide', () => {
  if (saveTimer) saveNow();
});
