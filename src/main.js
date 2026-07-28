// DCSPad bootstrap — wires modules together.

import { getState, update, updateNested, onSaveStatus } from './state.js';
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
import { downloadText, wireJsonImport, wirePaneImport } from './io.js?v=2';
import { applyContextIndicators, getSpContext } from './bridge/sp-context.js';
import {
  connectSpWeb, getSpWebInfo, listFolder, readTextFile, writeTextFile,
} from './sp-files.js?v=3';
import { showSplash } from './splash.js';
import { loadAppConfig } from './config.js';
import { initDocs } from './docs.js';
import { initSpChromeToggle } from './sp-chrome.js';

const splashApi = showSplash();
splashApi.status('Restoring workspace…');
const configReady = loadAppConfig();
const state = getState();
const initialSpContext = applyContextIndicators();
const spChromeApi = initSpChromeToggle(initialSpContext);

// ---------- layout ----------
let editorsApi = null;
const layoutApi = initLayout({
  onEditorTabChange: (name) => {
    editorsApi?.activate(name);
    updateStatusLang(name);
  },
});

// Active-pane type badge in the status bar (reuses the snippet badge).
function updateStatusLang(name) {
  const badge = document.getElementById('status-lang');
  badge.textContent = name;
  badge.dataset.lang = name;
}
updateStatusLang(state.layout.editorTab);

// Unsaved marks: a 5px dot on a tab whose buffer changed since the last
// run — the preview is stale for that pane. Cleared by Run.
function markUnsaved(name) {
  const dot = document.getElementById(`unsaved-${name}`);
  if (dot) dot.hidden = false;
}
function clearUnsaved() {
  for (const name of ['html', 'css', 'js']) {
    const dot = document.getElementById(`unsaved-${name}`);
    if (dot) dot.hidden = true;
  }
}

const isDiagVisible = (name) =>
  document.querySelector(`#diag-tabs .tab[data-diag="${name}"]`).classList.contains('active');

// ---------- editors ----------
splashApi.status('Starting Monaco editor…');
try {
  editorsApi = await initEditors({
    onChange: (name) => { markUnsaved(name); scheduleAutorun(); },
    onRunShortcut: () => run(),
    onTogglePane: (name) => layoutApi.togglePane?.(name),
    onFontStep: (delta) => stepEditorFontSize(delta),
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
const docsApi = initDocs({
  config: configResult.config,
  layoutApi,
  onBrowse: () => openSpFiles('browser'),
  onError: (msg) => padWarn(msg),
});

const btnSp = document.getElementById('btn-sp');
if (btnSp) {
  if (!initialSpContext.live) {
    btnSp.hidden = true;
  } else {
    btnSp.hidden = false;
    btnSp.addEventListener('click', () => {
      const url = configResult.config?.workbench?.url || initialSpContext.webAbsoluteUrl || '/';
      const opened = window.open(url, 'dcspad-sp');
      opened?.focus?.();
    });
  }
}

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
    settleRunFeedback();
  },
});

// Run-feedback beats 2–4: the header sweep is CSS (.sweeping); a run that
// outlives ~1.2s loops the sweep until the frame reports; the resolved
// beat pops the ✓ timestamp chip (the ms figure lives in the status bar).
let longRunTimer = null;
function settleRunFeedback() {
  clearTimeout(longRunTimer);
  longRunTimer = null;
  const panel = document.getElementById('preview-panel');
  panel.classList.remove('running-long');
  const chip = document.getElementById('preview-run-chip');
  document.getElementById('preview-run-time').textContent = new Date().toTimeString().slice(0, 8);
  chip.hidden = false;
  chip.classList.remove('pop');
  void chip.offsetWidth;
  chip.classList.add('pop');
}

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
  panel.classList.remove('sweeping', 'running-long');
  void panel.offsetWidth;
  panel.classList.add('sweeping');
  clearUnsaved();
  clearTimeout(longRunTimer);
  longRunTimer = setTimeout(() => panel.classList.add('running-long'), 1200);

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
      clearTimeout(longRunTimer);
      document.getElementById('preview-panel').classList.remove('running-long');
    }
  }, 15000);
}

document.getElementById('btn-run').addEventListener('click', run);
document.getElementById('btn-rerun').addEventListener('click', run);

// ---------- preview theme toggle ----------
const btnPreviewTheme = document.getElementById('btn-preview-theme');

function applyPreviewTheme() {
  const dark = getState().settings.previewDark;
  // The button carries both sun/moon SVGs; CSS shows one per data-mode.
  btnPreviewTheme.dataset.mode = dark ? 'dark' : 'light';
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

// ---------- dropdown menus (settings, file, docs) ----------
const menus = [
  { btn: document.getElementById('btn-settings'), menu: document.getElementById('settings-menu') },
  { btn: document.getElementById('btn-file'), menu: document.getElementById('file-menu') },
  { btn: document.getElementById('btn-docs'), menu: document.getElementById('docs-menu') },
];
for (const { btn, menu } of menus) {
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    for (const item of menus) {
      item.menu.hidden = true;
      item.btn.setAttribute('aria-expanded', 'false');
    }
    menu.hidden = !open;
    btn.setAttribute('aria-expanded', String(open));
  });
}
document.addEventListener('click', (e) => {
  for (const { btn, menu } of menus) {
    if (!menu.hidden && !menu.contains(e.target)) {
      menu.hidden = true;
      btn.setAttribute('aria-expanded', 'false');
    }
  }
});
const closeFileMenu = () => {
  document.getElementById('file-menu').hidden = true;
  document.getElementById('btn-file').setAttribute('aria-expanded', 'false');
};

// ---------- project save/load + pane exports ----------

// A warning from the pad itself, rendered into the console panel.
function padWarn(msg) {
  consoleApi.handlers.console({ level: 'warn', args: [{ t: 'str', v: `DCSPad: ${msg}` }] });
}

for (const warning of configResult.warnings) padWarn(warning);

const appToast = document.getElementById('app-toast');
let toastTimer = null;
function showToast(message, tone = '') {
  clearTimeout(toastTimer);
  appToast.textContent = message;
  appToast.className = `app-toast${tone ? ` ${tone}` : ''}`;
  appToast.hidden = false;
  toastTimer = setTimeout(() => { appToast.hidden = true; }, 4200);
}

// Shared confirmation for local and SharePoint single-file imports.
const paneReplaceDialog = document.getElementById('pane-replace-dialog');
const paneReplaceTitle = document.getElementById('pane-replace-title');
const paneReplaceContext = document.getElementById('pane-replace-context');
const paneReplaceFile = document.getElementById('pane-replace-file');
const paneReplaceBadge = document.getElementById('pane-replace-badge');
let pendingPaneReplacement = null;

function cancelPaneReplacement() {
  pendingPaneReplacement = null;
  if (paneReplaceDialog.open) paneReplaceDialog.close();
}

function confirmPaneReplacement(candidate, onReplaced = () => {}) {
  pendingPaneReplacement = { ...candidate, onReplaced };
  const label = candidate.pane.toUpperCase();
  paneReplaceTitle.textContent = `Replace ${label} code?`;
  paneReplaceContext.textContent =
    `${candidate.fileName} will replace all code in the ${label} editor.`;
  paneReplaceFile.textContent = candidate.fileName;
  paneReplaceBadge.textContent = candidate.pane;
  paneReplaceBadge.dataset.lang = candidate.pane;
  if (!paneReplaceDialog.open) paneReplaceDialog.showModal();
  document.getElementById('pane-replace-confirm').focus();
}

document.getElementById('pane-replace-form').addEventListener('submit', (event) => {
  event.preventDefault();
  const pending = pendingPaneReplacement;
  if (!pending) return;
  editorsApi.setDocs({ [pending.pane]: pending.text });
  layoutApi.selectEditorTab(pending.pane);
  markUnsaved(pending.pane);
  pendingPaneReplacement = null;
  paneReplaceDialog.close();
  statusRun.textContent = `${pending.fileName} imported into ${pending.pane.toUpperCase()}`;
  statusRun.className = 'status-item';
  showToast(`${pending.fileName} replaced the ${pending.pane.toUpperCase()} editor.`, 'success');
  pending.onReplaced();
});
document.getElementById('pane-replace-cancel').addEventListener('click', cancelPaneReplacement);
document.getElementById('pane-replace-close').addEventListener('click', cancelPaneReplacement);
paneReplaceDialog.addEventListener('cancel', () => {
  pendingPaneReplacement = null;
});

// Project naming is an inline, persisted top-bar flow. A project-file save
// requested without a name resumes automatically after the user names it.
const projectNameDisplay = document.getElementById('project-name-display');
const projectNameText = document.getElementById('project-name-text');
const projectNameForm = document.getElementById('project-name-form');
const projectNameInput = document.getElementById('project-name-input');
const projectNameError = document.getElementById('project-name-error');
let saveProjectAfterNaming = false;

function projectName() {
  return typeof getState().projectName === 'string' ? getState().projectName.trim() : '';
}

function filenameBase() {
  const normalized = projectName()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80)
    .replace(/-+$/g, '');
  return normalized || 'dcspad';
}

function renderProjectName() {
  const name = projectName();
  projectNameText.textContent = name || '(untitled)';
  projectNameDisplay.classList.toggle('is-empty', !name);
  projectNameDisplay.title = name ? 'Edit project name' : 'Set project name';
  document.title = name
    ? `${name} — DCSPad`
    : 'DCSPad — SharePoint Developer Workbench';
}

function showProjectNameError(message = '') {
  projectNameError.textContent = message;
  projectNameError.hidden = !message;
  projectNameInput.classList.toggle('invalid', Boolean(message));
  projectNameInput.setAttribute('aria-invalid', String(Boolean(message)));
}

function startProjectNameEdit({ requiredForProjectSave = false } = {}) {
  saveProjectAfterNaming = requiredForProjectSave;
  projectNameInput.value = projectName();
  projectNameDisplay.hidden = true;
  projectNameForm.hidden = false;
  showProjectNameError(requiredForProjectSave && !projectName()
    ? 'Name this project to save a project file.'
    : '');
  requestAnimationFrame(() => {
    projectNameInput.focus();
    projectNameInput.select();
  });
}

function finishProjectNameEdit() {
  projectNameForm.hidden = true;
  projectNameDisplay.hidden = false;
  showProjectNameError('');
}

function cancelProjectNameEdit() {
  saveProjectAfterNaming = false;
  finishProjectNameEdit();
}

function downloadProject() {
  const s = getState();
  const file = {
    app: 'dcspad', kind: 'project', v: 1,
    name: projectName(),
    savedAt: new Date().toISOString(),
    docs: editorsApi.getDocs(),
    libraries: { enabled: s.libraries.enabled, dcsUrl: s.libraries.dcsUrl },
    jsAsModule: s.settings.jsAsModule,
  };
  downloadText(`${filenameBase()}.dcspad.json`, JSON.stringify(file, null, 2));
}

renderProjectName();
projectNameDisplay.addEventListener('click', () => startProjectNameEdit());
projectNameInput.addEventListener('input', () => showProjectNameError(''));
projectNameInput.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    event.preventDefault();
    cancelProjectNameEdit();
  }
});
document.getElementById('project-name-cancel').addEventListener('click', cancelProjectNameEdit);
projectNameForm.addEventListener('submit', (event) => {
  event.preventDefault();
  const name = projectNameInput.value.trim();
  if (!name) {
    showProjectNameError('Enter a project name.');
    projectNameInput.focus();
    return;
  }
  update({ projectName: name });
  renderProjectName();
  finishProjectNameEdit();
  if (saveProjectAfterNaming) {
    saveProjectAfterNaming = false;
    downloadProject();
  }
});

document.getElementById('mi-save-project').addEventListener('click', () => {
  closeFileMenu();
  if (!projectName()) {
    startProjectNameEdit({ requiredForProjectSave: true });
    return;
  }
  downloadProject();
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
  update({ projectName: str(doc.name || doc.projectName).trim() });
  renderProjectName();

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

document.getElementById('mi-import-pane').addEventListener('click', () => {
  closeFileMenu();
  document.getElementById('import-pane-file').click();
});
wirePaneImport(
  'import-pane-file',
  (candidate) => confirmPaneReplacement(candidate),
  (message) => {
    padWarn(message);
    showToast(message, 'error');
  },
);

const PANE_EXPORTS = [
  ['mi-export-html', 'html', 'html', 'text/html'],
  ['mi-export-css', 'css', 'css', 'text/css'],
  ['mi-export-js', 'js', 'js', 'text/javascript'],
];
for (const [id, pane, extension, type] of PANE_EXPORTS) {
  document.getElementById(id).addEventListener('click', () => {
    closeFileMenu();
    downloadText(`${filenameBase()}.${extension}`, editorsApi.getDocs()[pane], type);
  });
}

document.getElementById('mi-export-all').addEventListener('click', () => {
  closeFileMenu();
  const docs = editorsApi.getDocs();
  const exports = PANE_EXPORTS.filter(([, pane]) => docs[pane].trim());
  if (!exports.length) {
    alert('The HTML, CSS and JS panes are empty.');
    return;
  }
  for (const [, pane, extension, type] of exports) {
    downloadText(`${filenameBase()}.${extension}`, docs[pane], type);
  }
});

// ---------- SharePoint document-library transfer ----------
const spImportMenuItem = document.getElementById('mi-sp-import');
const spExportMenuItem = document.getElementById('mi-sp-export');
const browserBrowse = document.getElementById('browser-browse');
const spFilesDialog = document.getElementById('sp-files-dialog');
const spFilesTitle = document.getElementById('sp-files-title');
const spSiteForm = document.getElementById('sp-site-form');
const spSiteUrl = document.getElementById('sp-site-url');
const spSiteOpen = document.getElementById('sp-site-open');
const spExportControls = document.getElementById('sp-export-controls');
const spExportPane = document.getElementById('sp-export-pane');
const spExportName = document.getElementById('sp-export-name');
const spFolderPath = document.getElementById('sp-folder-path');
const spFolderUp = document.getElementById('sp-folder-up');
const spFilesList = document.getElementById('sp-files-list');
const spFilesEmpty = document.getElementById('sp-files-empty');
const spFilesError = document.getElementById('sp-files-error');
const spFilesNotice = document.getElementById('sp-files-notice');
const spFilesPrimary = document.getElementById('sp-files-primary');
let spFilesMode = 'import';
let spFolder = null;
let spSelectedFile = null;
let spFilesBusy = false;
let spOverwriteArmed = false;
let spTargetWebUrl = '';

const FOLDER_ICON = '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" aria-hidden="true"><path d="M1.8 4.2h4l1.3 1.4h7.1v7.2H1.8z"/><path d="M1.8 4.2V2.8h4.4l1.2 1.4"/></svg>';
const FILE_ICON = '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.35" stroke-linejoin="round" aria-hidden="true"><path d="M3 1.8h6.2L13 5.6v8.6H3z"/><path d="M9.2 1.8v3.8H13"/></svg>';
const browserFileLabel = (type) => ({
  html: 'HTML', markdown: 'MD', css: 'CSS', javascript: 'JS',
  json: 'JSON', csv: 'CSV', text: 'TXT',
}[type] || 'FILE');

function refreshSpMenuState(initial = null) {
  const ctx = initial || getSpContext({ refresh: true });
  const appliedContext = applyContextIndicators();
  spChromeApi.setContext(appliedContext);
  for (const item of [spImportMenuItem, spExportMenuItem]) {
    item.disabled = !ctx.live;
    item.title = ctx.live ? '' : 'Requires SP: Live';
  }
  browserBrowse.disabled = !ctx.live;
  browserBrowse.title = ctx.live ? 'Browse SharePoint' : 'Requires SP: Live';
  return ctx.live;
}
refreshSpMenuState(initialSpContext);
document.getElementById('btn-file').addEventListener('click', () => {
  refreshSpMenuState(getSpContext({ refresh: true }));
});

function setSpError(message = '') {
  spFilesError.textContent = message;
  spFilesError.hidden = !message;
}

function setSpNotice(message = '') {
  spFilesNotice.textContent = message;
  spFilesNotice.hidden = !message;
}

function resetOverwriteConfirmation() {
  spOverwriteArmed = false;
  setSpNotice('');
  if (spFilesMode === 'export') spFilesPrimary.textContent = 'Upload file';
}

function formatBytes(bytes) {
  if (!bytes) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(1)} MB`;
}

function parentSpPath(path, root) {
  if (!path || path === root) return root;
  const parent = path.slice(0, path.lastIndexOf('/')) || '/';
  return parent.length < root.length ? root : parent;
}

function renderSpFolder() {
  spFilesList.replaceChildren();
  const entries = [...spFolder.folders, ...spFolder.files];
  spFilesEmpty.hidden = entries.length > 0;
  spFolderPath.textContent = spFolder.path;
  spFolderPath.title = spFolder.path;
  spFolderUp.disabled = spFolder.path === spFolder.rootPath;

  for (const entry of entries) {
    const row = document.createElement('button');
    row.type = 'button';
    row.className = 'sp-file-row';
    row.dataset.kind = entry.kind;
    row.setAttribute('role', 'option');
    row.setAttribute('aria-selected', 'false');

    const icon = document.createElement('span');
    icon.className = 'sp-file-row__icon';
    icon.innerHTML = entry.kind === 'folder' ? FOLDER_ICON : FILE_ICON;
    const name = document.createElement('span');
    name.className = 'sp-file-row__name';
    name.textContent = entry.name;
    const meta = document.createElement('span');
    meta.className = 'sp-file-row__meta';
    meta.textContent = entry.kind === 'folder'
      ? 'folder'
      : `${spFilesMode === 'browser'
          ? browserFileLabel(entry.browserType)
          : entry.pane.toUpperCase()} · ${formatBytes(entry.length)}`;
    row.append(icon, name, meta);

    if (entry.kind === 'folder') {
      row.addEventListener('click', () => loadSpFolder(entry.serverRelativeUrl));
    } else {
      row.addEventListener('click', () => {
        spSelectedFile = entry;
        for (const other of spFilesList.querySelectorAll('.sp-file-row')) {
          const selected = other === row;
          other.classList.toggle('selected', selected);
          other.setAttribute('aria-selected', String(selected));
        }
        if (spFilesMode === 'import' || spFilesMode === 'browser') {
          spFilesPrimary.disabled = false;
        } else {
          spExportPane.value = entry.pane;
          spExportName.value = entry.name;
          resetOverwriteConfirmation();
        }
      });
    }
    spFilesList.append(row);
  }
}

async function loadSpFolder(path) {
  if (spFilesBusy) return;
  spFilesBusy = true;
  spSelectedFile = null;
  spFilesPrimary.disabled = true;
  setSpError('');
  resetOverwriteConfirmation();
  spFilesEmpty.hidden = true;
  spFilesList.innerHTML = '<div class="sp-files-empty">Loading SharePoint folder…</div>';
  try {
    spFolder = await listFolder(path, {
      webUrl: spTargetWebUrl,
      purpose: spFilesMode === 'browser' ? 'browser' : 'code',
    });
    updateNested('settings', { spFilesFolder: spFolder.path });
    renderSpFolder();
    if (spFilesMode === 'export') spFilesPrimary.disabled = false;
  } catch (error) {
    spFilesList.replaceChildren();
    setSpError(error.message || String(error));
  } finally {
    spFilesBusy = false;
  }
}

async function connectSpSite(candidateWebUrl, { restoreFolder = true } = {}) {
  if (spFilesBusy) return;
  spFilesBusy = true;
  spSelectedFile = null;
  spFolder = null;
  spFilesPrimary.disabled = true;
  spSiteOpen.disabled = true;
  setSpError('');
  setSpNotice('');
  spFilesEmpty.hidden = true;
  spFilesList.innerHTML = '<div class="sp-files-empty">Connecting to SharePoint site…</div>';

  let startPath = '';
  try {
    const previousWebUrl = getState().settings.spFilesWebUrl;
    const connected = await connectSpWeb(candidateWebUrl);
    spTargetWebUrl = connected.webUrl;
    spSiteUrl.value = connected.webUrl;
    const rememberedFolder = getState().settings.spFilesFolder;
    const sameRememberedWeb =
      previousWebUrl.replace(/\/+$/, '').toLowerCase()
      === connected.webUrl.replace(/\/+$/, '').toLowerCase();
    startPath = restoreFolder
      && sameRememberedWeb
      && rememberedFolder
      && (rememberedFolder === connected.rootPath
        || rememberedFolder.startsWith(`${connected.rootPath}/`))
      ? rememberedFolder
      : connected.rootPath;
    updateNested('settings', {
      spFilesWebUrl: connected.webUrl,
      spFilesFolder: startPath,
    });
  } catch (error) {
    spFilesList.replaceChildren();
    setSpError(error.message || String(error));
  } finally {
    spFilesBusy = false;
    spSiteOpen.disabled = false;
  }

  if (startPath) await loadSpFolder(startPath);
}

function exportExtension(pane) {
  return pane === 'js' ? 'js' : pane;
}

function defaultSpExportName() {
  return `${filenameBase()}.${exportExtension(spExportPane.value)}`;
}

async function openSpFiles(mode) {
  closeFileMenu();
  if (!refreshSpMenuState()) {
    showToast('SharePoint file transfer requires SP: Live.', 'error');
    return;
  }

  spFilesMode = mode;
  spSelectedFile = null;
  spFolder = null;
  setSpError('');
  resetOverwriteConfirmation();
  spExportControls.hidden = mode !== 'export';
  spFilesTitle.textContent = mode === 'import'
    ? 'Import from SharePoint'
    : mode === 'browser' ? 'Browse SharePoint' : 'Export to SharePoint';
  spFilesPrimary.textContent = mode === 'import'
    ? 'Continue'
    : mode === 'browser' ? 'Open file' : 'Upload file';
  spFilesList.setAttribute(
    'aria-label',
    mode === 'browser'
      ? 'SharePoint folders and Browser-supported files'
      : 'SharePoint folders and code files',
  );
  spFilesEmpty.textContent = mode === 'browser'
    ? 'No Browser-supported files in this folder.'
    : 'No HTML, CSS, or JavaScript files in this folder.';
  spFilesPrimary.disabled = true;

  if (mode === 'export') {
    const activePane = ['html', 'css', 'js'].includes(getState().layout.editorTab)
      ? getState().layout.editorTab
      : 'html';
    spExportPane.value = activePane;
    spExportName.value = defaultSpExportName();
  }

  if (!spFilesDialog.open) spFilesDialog.showModal();
  try {
    const defaultWebUrl =
      getState().settings.spFilesWebUrl || getSpWebInfo().webUrl;
    spSiteUrl.value = defaultWebUrl;
    await connectSpSite(defaultWebUrl);
  } catch (error) {
    setSpError(error.message || String(error));
  }
}

function closeSpFiles() {
  if (!spFilesBusy && spFilesDialog.open) spFilesDialog.close();
}

spImportMenuItem.addEventListener('click', () => openSpFiles('import'));
spExportMenuItem.addEventListener('click', () => openSpFiles('export'));
document.getElementById('sp-files-close').addEventListener('click', closeSpFiles);
document.getElementById('sp-files-cancel').addEventListener('click', closeSpFiles);
document.getElementById('sp-folder-refresh').addEventListener('click', () => {
  if (spFolder) loadSpFolder(spFolder.path);
});
spFolderUp.addEventListener('click', () => {
  if (spFolder) loadSpFolder(parentSpPath(spFolder.path, spFolder.rootPath));
});
spSiteForm.addEventListener('submit', (event) => {
  event.preventDefault();
  connectSpSite(spSiteUrl.value, { restoreFolder: false });
});
spSiteUrl.addEventListener('input', () => {
  const changed = spSiteUrl.value.trim().replace(/\/+$/, '').toLowerCase()
    !== spTargetWebUrl.replace(/\/+$/, '').toLowerCase();
  if (changed) {
    spFilesPrimary.disabled = true;
    setSpNotice('Choose Open site to browse this SharePoint site.');
  } else {
    setSpNotice('');
    if (spFilesMode === 'export' && spFolder) spFilesPrimary.disabled = false;
    if (spFilesMode === 'import' && spSelectedFile) spFilesPrimary.disabled = false;
    if (spFilesMode === 'browser' && spSelectedFile) spFilesPrimary.disabled = false;
  }
});
spExportPane.addEventListener('change', () => {
  spExportName.value = defaultSpExportName();
  resetOverwriteConfirmation();
});
spExportName.addEventListener('input', () => {
  setSpError('');
  resetOverwriteConfirmation();
});

spFilesPrimary.addEventListener('click', async () => {
  if (spFilesBusy || !spFolder) return;

  if (spFilesMode === 'browser') {
    if (!spSelectedFile) return;
    const url = new URL(spTargetWebUrl);
    // Assign the server-relative path as a pathname so SharePoint-valid
    // characters such as # stay part of the file path rather than becoming
    // a URL fragment.
    url.pathname = spSelectedFile.serverRelativeUrl;
    url.search = '';
    url.hash = '';
    spFilesDialog.close();
    await docsApi.loadAddress(url.href);
    return;
  }

  if (spFilesMode === 'import') {
    if (!spSelectedFile) return;
    spFilesBusy = true;
    spFilesPrimary.disabled = true;
    setSpError('');
    try {
      const candidate = await readTextFile(
        spSelectedFile.serverRelativeUrl,
        { webUrl: spTargetWebUrl },
      );
      confirmPaneReplacement(candidate, () => {
        if (spFilesDialog.open) spFilesDialog.close();
      });
    } catch (error) {
      setSpError(error.message || String(error));
    } finally {
      spFilesBusy = false;
      spFilesPrimary.disabled = !spSelectedFile;
    }
    return;
  }

  const pane = spExportPane.value;
  const name = spExportName.value.trim();
  const expected = pane === 'html' ? /\.(?:html|htm)$/i : new RegExp(`\\.${pane}$`, 'i');
  if (!/^[a-z0-9][a-z0-9._-]*\.(?:html?|css|js)$/i.test(name)) {
    setSpError(
      'Use a safe file name containing letters, numbers, dots, hyphens, or underscores.',
    );
    return;
  }
  if (!expected.test(name)) {
    setSpError(`The file extension must match the ${pane.toUpperCase()} editor.`);
    return;
  }
  const text = editorsApi.getDocs()[pane];
  if (!text.trim()) {
    setSpError(`The ${pane.toUpperCase()} editor is empty.`);
    return;
  }

  const existing = spFolder.files.find(
    (file) => file.name.localeCompare(name, undefined, { sensitivity: 'base' }) === 0,
  );
  if (existing && !spOverwriteArmed) {
    spOverwriteArmed = true;
    setSpNotice(`${existing.name} already exists. Choose Overwrite to replace it.`);
    spFilesPrimary.textContent = 'Overwrite';
    return;
  }

  spFilesBusy = true;
  spFilesPrimary.disabled = true;
  setSpError('');
  try {
    await writeTextFile(spFolder.path, name, text, {
      overwrite: Boolean(existing),
      webUrl: spTargetWebUrl,
    });
    showToast(`${name} uploaded to SharePoint.`, 'success');
    statusRun.textContent = `${name} uploaded to SharePoint`;
    statusRun.className = 'status-item';
    spFilesDialog.close();
  } catch (error) {
    setSpError(error.message || String(error));
    spFilesPrimary.disabled = false;
  } finally {
    spFilesBusy = false;
  }
});

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
// Lives at the right end of the Console/Network strip; the old settings-menu
// row is retired.
const DIAG_FS_MIN = 10, DIAG_FS_MAX = 18;
function applyDiagFontSize(px) {
  document.documentElement.style.setProperty('--diag-fs', `${px}px`);
  refreshStepperDisabled('btn-diag-font-dec', 'btn-diag-font-inc', px, DIAG_FS_MIN, DIAG_FS_MAX);
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

// Editor text size stepper (11–18px; line height locks to 1.7× inside the
// Monaco adapter). Ctrl/Cmd+= / − step whichever pane has focus: Monaco
// actions cover the editor, the listener below covers the console.
const EDITOR_FS_MIN = 11, EDITOR_FS_MAX = 18;
function refreshStepperDisabled(decId, incId, val, min, max) {
  document.getElementById(decId).disabled = val <= min;
  document.getElementById(incId).disabled = val >= max;
}
function applyEditorFontSize(px) {
  editorsApi.setFontSize(px);
  refreshStepperDisabled('btn-editor-font-dec', 'btn-editor-font-inc', px, EDITOR_FS_MIN, EDITOR_FS_MAX);
}
function stepEditorFontSize(delta) {
  const cur = getState().settings.editorFontSize;
  const next = Math.min(EDITOR_FS_MAX, Math.max(EDITOR_FS_MIN, cur + delta));
  if (next === cur) return;
  updateNested('settings', { editorFontSize: next });
  applyEditorFontSize(next);
}
refreshStepperDisabled('btn-editor-font-dec', 'btn-editor-font-inc',
  state.settings.editorFontSize, EDITOR_FS_MIN, EDITOR_FS_MAX);
document.getElementById('btn-editor-font-dec').addEventListener('click', () => stepEditorFontSize(-1));
document.getElementById('btn-editor-font-inc').addEventListener('click', () => stepEditorFontSize(+1));
document.getElementById('diag-panel').addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  if (e.key === '=' || e.key === '+') { e.preventDefault(); stepDiagFontSize(+1); }
  else if (e.key === '-') { e.preventDefault(); stepDiagFontSize(-1); }
});

// Word wrap (persisted; applied at editor creation and on toggle).
const btnWordWrap = document.getElementById('btn-word-wrap');
function reflectWordWrap() {
  const on = getState().settings.wordWrap;
  btnWordWrap.classList.toggle('active', on);
  btnWordWrap.setAttribute('aria-pressed', String(on));
}
reflectWordWrap();
btnWordWrap.addEventListener('click', () => {
  const on = !getState().settings.wordWrap;
  updateNested('settings', { wordWrap: on });
  editorsApi.setWordWrap(on);
  reflectWordWrap();
});

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
