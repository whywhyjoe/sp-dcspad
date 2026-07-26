// Library manager: the framework catalog — a single stored JSON document
// listing every library the pad can inject, rendered as the checkbox list
// in the sidebar. Seeded once from PRESETS, then the stored catalog is
// authoritative: entries (including seeded ones) can be added, removed
// and reordered. Enabled/pinned state stays in the workspace; the
// catalog is global across projects.
//
// Enabled libraries are injected into the assembled document as ordered,
// blocking tags in catalog order — same as a real page. Reordering
// matters: a plugin must sit below the library it extends.

import { getState, updateNested, loadDoc, saveDoc, newId, CATALOG_KEY } from './state.js';
import { el } from './inspect/tree-view.js';
import { applyFrameworkConfig } from './config.js';

// Seed only — after first boot the stored catalog is the truth.
export const PRESETS = [
  { id: 'dcs-standard', name: 'DCS Standard Include', needsConfig: true,
    hint: 'Set your org include URL once; stored with your workspace.' },
  { id: 'pnpjs2', name: 'PnPjs v2 (classic)', js: 'https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/2.15.0/pnpjs.es5.umd.bundle.min.js',
    intelligence: ['pnpjs-2.15.0'],
    hint: 'Exposes global pnp — use const { sp } = pnp;' },
  { id: 'alpine', name: 'Alpine.js', js: 'https://cdn.jsdelivr.net/npm/alpinejs@3/dist/cdn.min.js',
    intelligence: ['alpine-3'] },
  { id: 'chartjs', name: 'Chart.js', js: 'https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js' },
  { id: 'lodash', name: 'Lodash', js: 'https://cdn.jsdelivr.net/npm/lodash@4/lodash.min.js' },
  { id: 'exceljs', name: 'ExcelJS', js: 'https://cdn.jsdelivr.net/npm/exceljs@4/dist/exceljs.min.js' },
  { id: 'dayjs', name: 'Day.js', js: 'https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js' },
  { id: 'fusejs', name: 'Fuse.js', js: 'https://cdn.jsdelivr.net/npm/fuse.js@7/dist/fuse.min.js' },
  { id: 'marked', name: 'Marked', js: 'https://cdn.jsdelivr.net/npm/marked@12/marked.min.js' },
  { id: 'sortable', name: 'Sortable.js', js: 'https://cdn.jsdelivr.net/npm/sortablejs@1/Sortable.min.js' },
  { id: 'fabric', name: 'Fluent/Fabric Icons (CSS)', css: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css' },
];

let catalog = null;
let appConfig = null;
let onChangeCb = null;
let onStorageErrorCb = null;

const isCssUrl = (url) => /\.css(\?|$)/i.test(url);
const entryFromUrl = (url, name) => ({
  id: newId('lib'),
  name: name || url.split('/').pop() || url,
  js: isCssUrl(url) ? undefined : url,
  css: isCssUrl(url) ? url : undefined,
});

export function initLibraries({ config, onChange, onStorageError }) {
  appConfig = config;
  onChangeCb = onChange;
  onStorageErrorCb = onStorageError;
  catalog = loadDoc(CATALOG_KEY);

  if (!catalog) {
    // First boot on catalog-aware code: seed from PRESETS and migrate any
    // legacy workspace-local custom URLs into real catalog entries. They
    // were always-injected before, so they arrive enabled.
    catalog = { v: 1, items: structuredClone(PRESETS) };
    const libs = getState().libraries;
    if (libs.custom?.length) {
      const migratedIds = [];
      for (const url of libs.custom) {
        const entry = entryFromUrl(url);
        catalog.items.push(entry);
        migratedIds.push(entry.id);
      }
      updateNested('libraries', { custom: [], enabled: [...libs.enabled, ...migratedIds] });
    }
    persistCatalog();
  }

  render();
  document.getElementById('lib-custom-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const urlInput = document.getElementById('lib-custom-url');
    const nameInput = document.getElementById('lib-custom-name');
    const url = urlInput.value.trim();
    if (!url) return;
    const entry = entryFromUrl(url, nameInput.value.trim());
    catalog.items.push(entry);
    persistCatalog();
    // Adding a framework means "I want to use it now" — enable immediately.
    const enabled = new Set(getState().libraries.enabled);
    enabled.add(entry.id);
    updateNested('libraries', { enabled: [...enabled] });
    urlInput.value = '';
    nameInput.value = '';
    render();
    onChangeCb?.();
  });
  return { getEnabledLibraries };
}

function persistCatalog() {
  if (!saveDoc(CATALOG_KEY, catalog)) {
    onStorageErrorCb?.('framework catalog save failed (storage full?)');
  }
}

function render() {
  const libs = getState().libraries;
  const pinnedHost = document.getElementById('lib-pinned');
  const listHost = document.getElementById('lib-list');
  pinnedHost.textContent = '';
  listHost.textContent = '';

  for (const entry of catalog.items) {
    const pinned = libs.pinned.includes(entry.id);
    (pinned ? pinnedHost : listHost).append(catalogItem(entry, libs, pinned));
  }
}

function catalogItem(entry, libs, pinned) {
  const effective = applyFrameworkConfig(entry, appConfig);
  const item = el('label', 'lib-item');
  const chk = document.createElement('input');
  chk.type = 'checkbox';
  chk.checked = libs.enabled.includes(entry.id);

  const name = el('span', 'lib-name', effective.name);
  if (effective.hint) name.title = effective.hint;
  else if (effective.js || effective.css) {
    name.title = effective.js || effective.css;
    if (effective.fallbackJs) name.title += `\nFallback: ${effective.fallbackJs}`;
  }

  if (entry.needsConfig && !libs.dcsUrl) {
    item.classList.add('needs-config');
    name.title = entry.hint || 'Needs a URL';
  }

  chk.addEventListener('change', () => {
    if (entry.needsConfig && !getState().libraries.dcsUrl && chk.checked) {
      const url = prompt('URL for the DCS Standard Include (your org’s script/CSS bundle):');
      if (!url) { chk.checked = false; return; }
      updateNested('libraries', { dcsUrl: url.trim() });
      item.classList.remove('needs-config');
    }
    const enabled = new Set(getState().libraries.enabled);
    chk.checked ? enabled.add(entry.id) : enabled.delete(entry.id);
    updateNested('libraries', { enabled: [...enabled] });
    onChangeCb?.();
  });

  // Row tools. All live inside the <label>, so each must preventDefault
  // to stop the click from also toggling the checkbox.
  const tools = el('span', 'lib-tools');
  const tool = (cls, text, title, fn) => {
    const s = el('span', cls, text);
    s.title = title;
    s.addEventListener('click', (e) => { e.preventDefault(); fn(); });
    return s;
  };

  // Index is looked up at event time, not captured at render time: a
  // click that lands on a detached row (or any future re-entrancy)
  // must never splice by a stale position.
  const liveIdx = () => catalog.items.indexOf(entry);
  tools.append(
    tool('lib-move', '↑', 'Move up (injection order)', () => moveEntry(liveIdx(), -1)),
    tool('lib-move', '↓', 'Move down (injection order)', () => moveEntry(liveIdx(), +1)),
    tool('lib-pin' + (pinned ? ' pinned' : ''), pinned ? '★' : '☆', pinned ? 'Unpin' : 'Pin to top', () => {
      const pins = new Set(getState().libraries.pinned);
      pinned ? pins.delete(entry.id) : pins.add(entry.id);
      updateNested('libraries', { pinned: [...pins] });
      render();
    }),
    tool('lib-del', '✕', 'Remove from catalog', () => {
      const idx = liveIdx();
      if (idx === -1) return;
      if (!confirm(`Remove "${entry.name}" from the framework catalog?`)) return;
      catalog.items.splice(idx, 1);
      persistCatalog();
      const cur = getState().libraries;
      updateNested('libraries', {
        enabled: cur.enabled.filter((id) => id !== entry.id),
        pinned: cur.pinned.filter((id) => id !== entry.id),
      });
      render();
      onChangeCb?.();
    }),
  );

  item.append(chk, name, tools);
  return item;
}

function moveEntry(idx, delta) {
  const to = idx + delta;
  if (idx < 0 || to < 0 || to >= catalog.items.length) return;
  const [entry] = catalog.items.splice(idx, 1);
  catalog.items.splice(to, 0, entry);
  persistCatalog();
  render();
  onChangeCb?.();   // injection order changed — rerun matters
}

// Ordered list for the runner: enabled entries in catalog order.
export function getEnabledLibraries() {
  const libs = getState().libraries;
  const result = [];
  for (const entry of catalog.items) {
    if (!libs.enabled.includes(entry.id)) continue;
    if (entry.needsConfig) {
      if (libs.dcsUrl) {
        result.push({ name: entry.name, js: isCssUrl(libs.dcsUrl) ? undefined : libs.dcsUrl, css: isCssUrl(libs.dcsUrl) ? libs.dcsUrl : undefined });
      }
      continue;
    }
    const effective = applyFrameworkConfig(entry, appConfig);
    result.push({
      name: effective.name,
      js: effective.js,
      css: effective.css,
      fallbackJs: effective.fallbackJs,
      probeGlobal: effective.probeGlobal,
    });
  }
  return result;
}

// PnPjs types must follow the actual enabled runtime, not a catalog id: the
// catalog is user-editable and the same 2.15.0 UMD bundle is commonly added
// from jsDelivr under a custom name.
export function isPnpjs215Runtime(entry) {
  if (entry?.intelligence?.includes('pnpjs-2.15.0')) return true;
  const urls = [
    entry?.js,
    entry?.fallbackJs,
    entry?.configuredSources?.local,
    entry?.configuredSources?.cdn,
  ].map((url) => String(url || '').toLowerCase());
  return urls.some((url) =>
    url.includes('@pnp/pnpjs@2.15.0/')
    || url.includes('/pnp-pnpjs/2.15.0/')
    || url.includes('/pnpjs/2.15.0/'));
}

// Stored catalogs created before Alpine intelligence existed do not contain
// the pack metadata now present on PRESETS. Recognize the seeded entry and
// common Alpine v3 URLs so those existing workspaces gain intelligence
// without requiring a catalog reset or dcspad.config.json.
export function isAlpine3Runtime(entry) {
  if (entry?.intelligence?.includes('alpine-3')) return true;
  if (entry?.id === 'alpine') return true;
  const urls = [
    entry?.js,
    entry?.fallbackJs,
    entry?.configuredSources?.local,
    entry?.configuredSources?.cdn,
  ].map((url) => String(url || '').toLowerCase());
  return urls.some((url) =>
    url.includes('/alpinejs@3')
    || url.includes('/alpinejs/3.'));
}

export function hasEnabledPnpjs215Runtime() {
  const enabled = new Set(getState().libraries.enabled);
  return catalog.items.some((entry) =>
    enabled.has(entry.id)
    && isPnpjs215Runtime(applyFrameworkConfig(entry, appConfig)));
}

export function getEnabledIntelligence() {
  const enabled = new Set(getState().libraries.enabled);
  const packs = new Set();
  for (const entry of catalog.items) {
    if (!enabled.has(entry.id)) continue;
    const effective = applyFrameworkConfig(entry, appConfig);
    for (const pack of effective.intelligence || []) packs.add(pack);
    // Preserve compatibility for imported/custom catalogs created before
    // explicit intelligence metadata existed.
    if (isPnpjs215Runtime(effective)) packs.add('pnpjs-2.15.0');
    if (isAlpine3Runtime(effective)) packs.add('alpine-3');
  }
  return [...packs];
}

// ---------------------------------------------------------------
// For io.js (file export/import) and project-load warnings
// ---------------------------------------------------------------

export function getCatalogDoc() { return catalog; }

// Replace the catalog wholesale from an imported file. Minimal shape
// validation; returns false when the file isn't a catalog document.
export function replaceCatalog(doc) {
  if (!doc || !Array.isArray(doc.items)) return false;
  const items = doc.items.filter((it) => it && typeof it.id === 'string' && typeof it.name === 'string');
  catalog = { v: 1, items };
  persistCatalog();
  // Prune workspace ids that no longer resolve — otherwise dead
  // references accumulate in enabled/pinned across import cycles.
  const known = new Set(items.map((it) => it.id));
  const cur = getState().libraries;
  updateNested('libraries', {
    enabled: cur.enabled.filter((id) => known.has(id)),
    pinned: cur.pinned.filter((id) => known.has(id)),
  });
  render();
  onChangeCb?.();
  return true;
}

// Ids referenced by a loaded project but missing from the catalog.
export function unknownLibraryIds(ids) {
  const known = new Set(catalog.items.map((it) => it.id));
  return ids.filter((id) => !known.has(id));
}

// Re-render after workspace-level library state changed externally
// (e.g. a project file was loaded).
export function refreshLibraryUI() { render(); }
