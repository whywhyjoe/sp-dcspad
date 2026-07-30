// Central app state + debounced localStorage persistence.
// This module is the ONLY code that touches localStorage. It owns three
// documents — workspace (live, autosaved), snippets, catalog — so a
// future SharePoint-backed storage layer (DevPadData/{user}.json) can
// replace persistence wholesale by swapping just this seam.

const STORAGE_KEY = 'dcspad.v2.workspace';
export const CATALOG_KEY = 'dcspad.v2.catalog';
export const SNIPPETS_KEY = 'dcspad.v2.snippets';
const SAVE_DEBOUNCE_MS = 600;

const DEFAULTS = {
  projectName: '',
  html: '<div id="app">\n  <h2>Hello from DCSPad</h2>\n  <p>Edit HTML, CSS and JS, then press Run.</p>\n</div>\n',
  css: 'body {\n  font-family: "Segoe UI", sans-serif;\n  padding: 1rem;\n}\n',
  js: 'console.log("DCSPad ready", { when: new Date().toISOString() });\n',
  libraries: { enabled: [], pinned: ['pnpjs2'], custom: [] },
  settings: {
    autorun: false, jsAsModule: false, autoClearConsole: true, seenSplash: false,
    previewDark: true, diagFontSize: 12, editorFontSize: 13, wordWrap: false,
    spFilesWebUrl: '', spFilesFolder: '', browserHistory: [], browserFavorites: [],
    projectFileFingerprint: '',
  },
  layout: {
    sidebarW: 230, sidebarCollapsed: false,
    editorsFr: 1, runtimeFr: 1,
    previewFr: 1, diagH: 260, diagCollapsed: false,
    editorTab: 'js', diagTab: 'console',
    // Pane visibility (the topbar segmented toggles). sidebarCollapsed /
    // diagCollapsed above are legacy flags kept for shape stability: layout.js
    // reads them once to seed `panes` for pre-existing workspaces, then only
    // writes `panes`.
    panes: { resources: true, preview: true, console: true },
    snippetsPanelH: 210,
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
    // Surface it: without the 'error' event the status bar sticks at
    // "saving…" and the user closes the tab over unsaved work.
    console.warn('DCSPad: autosave failed', e);
    for (const fn of listeners) fn('error');
  }
}

export function getState() { return state; }

// The project-sized portion of a fresh workspace. New Project deliberately
// leaves app preferences, layout, snippets, and the framework catalog alone.
export function getNewProjectDefaults() {
  return {
    projectName: DEFAULTS.projectName,
    docs: {
      html: DEFAULTS.html,
      css: DEFAULTS.css,
      js: DEFAULTS.js,
    },
    libraries: {
      enabled: [...DEFAULTS.libraries.enabled],
      dcsUrl: '',
    },
    jsAsModule: DEFAULTS.settings.jsAsModule,
  };
}

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

// ---------------------------------------------------------------
// Named documents (catalog, snippets): small JSON collections beside
// the workspace. Loaded eagerly, written synchronously — they change
// on explicit user actions, not keystrokes, so no debounce.
// ---------------------------------------------------------------

// Returns the stored doc, or null when absent/corrupt (caller seeds).
export function loadDoc(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && Array.isArray(parsed.items) ? parsed : null;
  } catch {
    return null;
  }
}

// Returns false when the write failed (quota) so callers can surface it.
export function saveDoc(key, doc) {
  try {
    localStorage.setItem(key, JSON.stringify(doc));
    return true;
  } catch (e) {
    console.warn(`DCSPad: saving ${key} failed`, e);
    return false;
  }
}

// Session-random seed closes the (theoretical) cross-session collision
// window when two page loads mint their first id in the same millisecond.
const idSeed = Math.random().toString(36).slice(2, 6);
let idCounter = 0;
export function newId(prefix) {
  return `${prefix}_${Date.now().toString(36)}${idSeed}${(idCounter++).toString(36)}`;
}

// Flush a save still sitting in the debounce window when the tab is
// closed, reloaded or backgrounded — otherwise the last ≤600 ms of edits
// are silently lost.
window.addEventListener('pagehide', () => {
  if (saveTimer) saveNow();
});
