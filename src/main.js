// DCSPad bootstrap — wires modules together.

import { getState, updateNested, onSaveStatus } from './state.js';
import { initLayout } from './layout.js';
import { initEditors } from './editors.js';
import { initRunner, run as runnerRun, evalInFrame, mapSrcdocLineToUserJs, hasRun } from './runner.js';
import { initConsolePanel } from './console-panel.js';
import { initNetworkPanel, markRun as networkMarkRun } from './network-panel.js';
import {
  initLibraries, getEnabledLibraries, getCatalogDoc, replaceCatalog,
  unknownLibraryIds, refreshLibraryUI, getEnabledIntelligence,
} from './libraries.js';
import { initSnippets } from './snippets.js?v=2';
import { downloadText, wireJsonImport } from './io.js';
import { applyContextIndicators, getSpContext } from './bridge/sp-context.js';
import { showSplash } from './splash.js';
import { loadAppConfig } from './config.js';

const splashApi = showSplash();
splashApi.status('Restoring workspace…');
const configReady = loadAppConfig();
const state = getState();
applyContextIndicators();

// ---------- layout ----------
let editorsApi = null;
const layoutApi = initLayout({
  onEditorTabChange: (name) => editorsApi?.activate(name),
});

const isDiagVisible = (name) =>
  document.querySelector(`#diag-tabs .tab[data-diag="${name}"]`).classList.contains('active');

// ---------- editors ----------
splashApi.status('Starting Monaco editor…');
try {
  editorsApi = await initEditors({
    onChange: () => scheduleAutorun(),
    onRunShortcut: () => run(),
  });
} catch (error) {
  splashApi.fail(`Monaco failed to start — ${error.message || error}`);
  throw error;
}

// ---------- console + network ----------
const consoleApi = initConsolePanel({
  evalInFrame,
  mapSrcdocLine: mapSrcdocLineToUserJs,
  gotoJsLine: (line) => {
    layoutApi.selectEditorTab('js');
    editorsApi.gotoJsLine(line);
  },
  isConsoleVisible: () => isDiagVisible('console'),
});
const networkApi = initNetworkPanel({
  isNetworkVisible: () => isDiagVisible('network'),
});

// ---------- libraries ----------
const configResult = await configReady;
initLibraries({
  config: configResult.config,
  onChange: () => {
    scheduleAutorun();
    editorsApi.setIntelligencePacks(getEnabledIntelligence());
  },
  onStorageError: (msg) => reportStorageError(msg),
});
editorsApi.setIntelligencePacks(getEnabledIntelligence());

// ---------- snippets ----------
initSnippets({
  getSelection: (name) => editorsApi.getSelection(name),
  getDocs: () => editorsApi.getDocs(),
  insertAtCursor: (name, text) => editorsApi.insertAtCursor(name, text),
  selectEditorTab: (name) => layoutApi.selectEditorTab(name),
  onStorageError: (msg) => reportStorageError(msg),
});

// ---------- runner ----------
const statusRun = document.getElementById('status-run');
const SPINNER = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏'];
let spinnerTimer = null;

const runnerReady = initRunner({
  ...consoleApi.handlers,
  ...networkApi.handlers,
  loaded: (d) => {
    stopSpinner();
    statusRun.textContent = `ran in ${d.ms} ms`;
    statusRun.className = 'status-item ok';
  },
});

function startSpinner() {
  let i = 0;
  statusRun.className = 'status-item running';
  clearInterval(spinnerTimer);
  spinnerTimer = setInterval(() => {
    statusRun.textContent = `${SPINNER[i++ % SPINNER.length]} running`;
  }, 80);
}
function stopSpinner() {
  clearInterval(spinnerTimer);
  spinnerTimer = null;
}

async function run() {
  try {
    await runnerReady;
  } catch (e) {
    statusRun.textContent = e.message;
    statusRun.className = 'status-item error';
    return;
  }
  const settings = getState().settings;
  if (settings.autoClearConsole) consoleApi.clear();
  networkMarkRun();
  startSpinner();

  document.getElementById('btn-run').classList.remove('running');
  void document.getElementById('btn-run').offsetWidth;   // restart animation
  document.getElementById('btn-run').classList.add('running');
  const panel = document.getElementById('preview-panel');
  panel.classList.remove('sweeping');
  void panel.offsetWidth;
  panel.classList.add('sweeping');

  const { runNumber } = runnerRun({
    docs: editorsApi.getDocs(),
    libraries: getEnabledLibraries(),
    // Re-capture per run: on classic pages the host rewrites the
    // #__REQUESTDIGEST form field, and a bootstrap-time digest expires.
    spContext: getSpContext({ refresh: true }),
    settings,
  });
  if (!settings.autoClearConsole) consoleApi.runDivider(runNumber);

  // Safety: if load never fires (e.g. a library hangs), settle the status.
  setTimeout(() => {
    if (spinnerTimer) {
      stopSpinner();
      statusRun.textContent = 'still loading…';
      statusRun.className = 'status-item';
    }
  }, 15000);
}

document.getElementById('btn-run').addEventListener('click', run);
document.getElementById('btn-rerun').addEventListener('click', run);

// ---------- preview theme toggle ----------
const btnPreviewTheme = document.getElementById('btn-preview-theme');

function applyPreviewTheme() {
  const dark = getState().settings.previewDark;
  btnPreviewTheme.textContent = dark ? '☀' : '🌙';
  btnPreviewTheme.title = dark
    ? 'Switch preview to light — pad-only canvas color; your CSS still wins, and SharePoint pages are typically light'
    : 'Switch preview to dark — pad-only canvas color; your CSS still wins';
  document.getElementById('preview-host').classList.toggle('dark', dark);
}
applyPreviewTheme();

btnPreviewTheme.addEventListener('click', () => {
  updateNested('settings', { previewDark: !getState().settings.previewDark });
  applyPreviewTheme();
  if (hasRun()) run();
});

// ---------- auto-run ----------
const AUTORUN_DEBOUNCE_MS = 800;
let autorunTimer = null;
const chkAutorun = document.getElementById('chk-autorun');
chkAutorun.checked = state.settings.autorun;
document.getElementById('live-dot').classList.toggle('on', state.settings.autorun);

chkAutorun.addEventListener('change', () => {
  updateNested('settings', { autorun: chkAutorun.checked });
  document.getElementById('live-dot').classList.toggle('on', chkAutorun.checked);
  if (chkAutorun.checked) scheduleAutorun();
});

function scheduleAutorun() {
  if (!getState().settings.autorun) return;
  clearTimeout(autorunTimer);
  autorunTimer = setTimeout(run, AUTORUN_DEBOUNCE_MS);
}

// ---------- dropdown menus (settings, file) ----------
const menus = [
  { btn: document.getElementById('btn-settings'), menu: document.getElementById('settings-menu') },
  { btn: document.getElementById('btn-file'), menu: document.getElementById('file-menu') },
];
for (const { btn, menu } of menus) {
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    for (const m of menus) m.menu.hidden = true;
    menu.hidden = !open;
  });
}
document.addEventListener('click', (e) => {
  for (const { menu } of menus) {
    if (!menu.hidden && !menu.contains(e.target)) menu.hidden = true;
  }
});
const closeFileMenu = () => { document.getElementById('file-menu').hidden = true; };

// ---------- project save/load + pane exports ----------

// A warning from the pad itself, rendered into the console panel.
function padWarn(msg) {
  consoleApi.handlers.console({ level: 'warn', args: [{ t: 'str', v: `DCSPad: ${msg}` }] });
}

for (const warning of configResult.warnings) padWarn(warning);

document.getElementById('mi-save-project').addEventListener('click', () => {
  closeFileMenu();
  const s = getState();
  const file = {
    app: 'dcspad', kind: 'project', v: 1,
    savedAt: new Date().toISOString(),
    docs: editorsApi.getDocs(),
    libraries: { enabled: s.libraries.enabled, dcsUrl: s.libraries.dcsUrl },
    jsAsModule: s.settings.jsAsModule,
  };
  downloadText('dcspad-project.json', JSON.stringify(file, null, 2));
});

document.getElementById('mi-load-project').addEventListener('click', () => {
  closeFileMenu();
  document.getElementById('import-project-file').click();
});
wireJsonImport('import-project-file', (doc) => {
  if (!doc || doc.kind !== 'project' || typeof doc.docs !== 'object' || doc.docs === null) {
    alert('Not a DCSPad project file.');
    return;
  }
  const str = (v) => (typeof v === 'string' ? v : '');
  editorsApi.setDocs({ html: str(doc.docs.html), css: str(doc.docs.css), js: str(doc.docs.js) });

  const libs = doc.libraries || {};
  const enabled = Array.isArray(libs.enabled) ? libs.enabled.filter((id) => typeof id === 'string') : [];
  updateNested('libraries', {
    enabled,
    ...(typeof libs.dcsUrl === 'string' ? { dcsUrl: libs.dcsUrl } : {}),
  });
  if (typeof doc.jsAsModule === 'boolean') {
    updateNested('settings', { jsAsModule: doc.jsAsModule });
    // Resolved here, not via the chkModule const declared further down —
    // referencing it would work (the callback fires post-init) but is a
    // TDZ trap for anyone who reorders this file.
    document.getElementById('chk-module').checked = doc.jsAsModule;
    editorsApi.setJsAsModule(doc.jsAsModule);
  }
  refreshLibraryUI();
  editorsApi.setIntelligencePacks(getEnabledIntelligence());

  // Deliberately tolerant: a project may reference catalog entries that
  // were removed since it was saved. The run will fail visibly with
  // "X is not defined" — this warning just names the gap up front.
  const missing = unknownLibraryIds(enabled);
  if (missing.length) {
    padWarn(`this project references framework(s) not in your catalog: ${missing.join(', ')} — re-add them under Frameworks, or the run will fail where they're used`);
  }
  statusRun.textContent = 'project loaded — press Run';
  statusRun.className = 'status-item';
});

const PANE_EXPORTS = [
  ['mi-export-html', 'html', 'dcspad.html', 'text/html'],
  ['mi-export-css', 'css', 'dcspad.css', 'text/css'],
  ['mi-export-js', 'js', 'dcspad.js', 'text/javascript'],
];
for (const [id, pane, filename, type] of PANE_EXPORTS) {
  document.getElementById(id).addEventListener('click', () => {
    closeFileMenu();
    downloadText(filename, editorsApi.getDocs()[pane], type);
  });
}

// ---------- catalog file save/load ----------
document.getElementById('btn-catalog-export').addEventListener('click', () => {
  downloadText('dcspad-catalog.json', JSON.stringify(getCatalogDoc(), null, 2));
});
document.getElementById('btn-catalog-import').addEventListener('click', () => {
  document.getElementById('import-catalog-file').click();
});
wireJsonImport('import-catalog-file', (doc) => {
  if (!doc || !Array.isArray(doc.items)) { alert('Not a DCSPad catalog file.'); return; }
  const cur = getCatalogDoc().items.length;
  if (!confirm(`Replace your framework catalog (${cur} entries) with this file (${doc.items.length} entries)?`)) return;
  replaceCatalog(doc);
});

const chkModule = document.getElementById('chk-module');
chkModule.checked = state.settings.jsAsModule;
chkModule.addEventListener('change', () => {
  updateNested('settings', { jsAsModule: chkModule.checked });
  editorsApi.setJsAsModule(chkModule.checked);
});

const chkAutoclear = document.getElementById('chk-autoclear');
chkAutoclear.checked = state.settings.autoClearConsole;
chkAutoclear.addEventListener('change', () =>
  updateNested('settings', { autoClearConsole: chkAutoclear.checked }));

// Console/network text size stepper (10–18px, drives --diag-fs).
const DIAG_FS_MIN = 10, DIAG_FS_MAX = 18;
function applyDiagFontSize(px) {
  document.documentElement.style.setProperty('--diag-fs', `${px}px`);
  document.getElementById('diag-font-val').textContent = String(px);
}
applyDiagFontSize(state.settings.diagFontSize);
function stepDiagFontSize(delta) {
  const cur = getState().settings.diagFontSize;
  const next = Math.min(DIAG_FS_MAX, Math.max(DIAG_FS_MIN, cur + delta));
  if (next === cur) return;
  updateNested('settings', { diagFontSize: next });
  applyDiagFontSize(next);
}
document.getElementById('btn-diag-font-dec').addEventListener('click', () => stepDiagFontSize(-1));
document.getElementById('btn-diag-font-inc').addEventListener('click', () => stepDiagFontSize(+1));

// ---------- autosave tick + storage errors ----------
const saveEl = document.getElementById('status-save');
onSaveStatus((status) => {
  saveEl.classList.remove('saved', 'error');
  if (status === 'dirty') {
    saveEl.textContent = 'saving…';
  } else if (status === 'error') {
    saveEl.textContent = 'save failed — use File ▸ Save project';
    saveEl.classList.add('error');
  } else {
    saveEl.textContent = '✓ saved';
    saveEl.classList.add('saved');
  }
});

// Catalog/snippet write failures: the console entry persists even when
// a later (smaller) workspace autosave succeeds and retakes the status
// text, so the failure can't be silently papered over.
function reportStorageError(msg) {
  padWarn(`${msg} — use the ⤓ export buttons or File ▸ Save project to keep your work`);
  saveEl.textContent = 'save failed';
  saveEl.classList.add('error');
  saveEl.classList.remove('saved');
}

splashApi.status('Editor ready');
splashApi.finish();
