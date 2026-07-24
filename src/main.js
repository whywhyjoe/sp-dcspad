// DCSPad bootstrap — wires modules together.

import { getState, updateNested, onSaveStatus } from './state.js';
import { initLayout } from './layout.js';
import { initEditors } from './editors.js';
import { initRunner, run as runnerRun, evalInFrame, mapSrcdocLineToUserJs } from './runner.js';
import { initConsolePanel } from './console-panel.js';
import { initNetworkPanel, markRun as networkMarkRun } from './network-panel.js';
import { initLibraries, getEnabledLibraries } from './libraries.js';
import { applyContextIndicators } from './bridge/sp-context.js';
import { showSplash } from './splash.js';

const state = getState();

// ---------- layout ----------
const layoutApi = initLayout({
  onEditorTabChange: (name) => editorsApi.focus(name),
});

const isDiagVisible = (name) =>
  document.querySelector(`#diag-tabs .tab[data-diag="${name}"]`).classList.contains('active');

// ---------- editors ----------
const editorsApi = initEditors({
  onChange: () => scheduleAutorun(),
  onRunShortcut: () => run(),
});

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
initLibraries({ onChange: () => scheduleAutorun() });

// ---------- SP context ----------
const spContext = applyContextIndicators();

// ---------- boot splash ----------
showSplash({ spContext });

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
  await runnerReady;
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
    spContext,
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

// ---------- settings menu ----------
const settingsMenu = document.getElementById('settings-menu');
document.getElementById('btn-settings').addEventListener('click', (e) => {
  e.stopPropagation();
  settingsMenu.hidden = !settingsMenu.hidden;
});
document.addEventListener('click', (e) => {
  if (!settingsMenu.hidden && !settingsMenu.contains(e.target)) settingsMenu.hidden = true;
});

const chkModule = document.getElementById('chk-module');
chkModule.checked = state.settings.jsAsModule;
chkModule.addEventListener('change', () =>
  updateNested('settings', { jsAsModule: chkModule.checked }));

const chkAutoclear = document.getElementById('chk-autoclear');
chkAutoclear.checked = state.settings.autoClearConsole;
chkAutoclear.addEventListener('change', () =>
  updateNested('settings', { autoClearConsole: chkAutoclear.checked }));

// ---------- autosave tick ----------
const saveEl = document.getElementById('status-save');
onSaveStatus((status) => {
  if (status === 'dirty') {
    saveEl.textContent = 'saving…';
    saveEl.classList.remove('saved');
  } else {
    saveEl.textContent = '✓ saved';
    saveEl.classList.add('saved');
  }
});
