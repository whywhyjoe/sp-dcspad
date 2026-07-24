// Library manager: preset catalog, enable/pin toggles, custom URLs.
// Enabled libraries are injected into the assembled document as ordered,
// blocking tags (catalog order, then customs) — same as a real page.

import { getState, updateNested } from './state.js';
import { el } from './inspect/tree-view.js';

export const PRESETS = [
  { id: 'dcs-standard', name: 'DCS Standard Include', needsConfig: true,
    hint: 'Set your org include URL once; stored with your workspace.' },
  { id: 'pnpjs2', name: 'PnPjs v2 (classic)', js: 'https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/2.15.0/pnpjs.es5.umd.bundle.min.js',
    hint: 'Exposes global pnp — use const { sp } = pnp;' },
  { id: 'alpine', name: 'Alpine.js', js: 'https://cdn.jsdelivr.net/npm/alpinejs@3/dist/cdn.min.js' },
  { id: 'chartjs', name: 'Chart.js', js: 'https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js' },
  { id: 'lodash', name: 'Lodash', js: 'https://cdn.jsdelivr.net/npm/lodash@4/lodash.min.js' },
  { id: 'exceljs', name: 'ExcelJS', js: 'https://cdn.jsdelivr.net/npm/exceljs@4/dist/exceljs.min.js' },
  { id: 'dayjs', name: 'Day.js', js: 'https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js' },
  { id: 'fusejs', name: 'Fuse.js', js: 'https://cdn.jsdelivr.net/npm/fuse.js@7/dist/fuse.min.js' },
  { id: 'marked', name: 'Marked', js: 'https://cdn.jsdelivr.net/npm/marked@12/marked.min.js' },
  { id: 'sortable', name: 'Sortable.js', js: 'https://cdn.jsdelivr.net/npm/sortablejs@1/Sortable.min.js' },
  { id: 'fabric', name: 'Fluent/Fabric Icons (CSS)', css: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css' },
];

let onChangeCb = null;

export function initLibraries({ onChange }) {
  onChangeCb = onChange;
  render();
  document.getElementById('lib-custom-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const input = document.getElementById('lib-custom-url');
    const url = input.value.trim();
    if (!url) return;
    const libs = getState().libraries;
    libs.custom.push(url);
    updateNested('libraries', { custom: libs.custom });
    input.value = '';
    render();
    onChangeCb?.();
  });
  return { getEnabledLibraries };
}

function render() {
  const libs = getState().libraries;
  const pinnedHost = document.getElementById('lib-pinned');
  const listHost = document.getElementById('lib-list');
  const customHost = document.getElementById('lib-custom-list');
  pinnedHost.textContent = '';
  listHost.textContent = '';
  customHost.textContent = '';

  for (const preset of PRESETS) {
    const pinned = libs.pinned.includes(preset.id);
    (pinned ? pinnedHost : listHost).append(presetItem(preset, libs, pinned));
  }

  libs.custom.forEach((url, i) => {
    const item = el('div', 'lib-item');
    const mark = el('span', '', '✓');
    mark.style.color = 'var(--accent)';
    const name = el('span', 'lib-name', url.split('/').pop() || url);
    name.title = url;
    const del = el('span', 'lib-del', '✕');
    del.title = 'Remove';
    del.addEventListener('click', () => {
      libs.custom.splice(i, 1);
      updateNested('libraries', { custom: libs.custom });
      render();
      onChangeCb?.();
    });
    item.append(mark, name, del);
    customHost.append(item);
  });
}

function presetItem(preset, libs, pinned) {
  const item = el('label', 'lib-item');
  const chk = document.createElement('input');
  chk.type = 'checkbox';
  chk.checked = libs.enabled.includes(preset.id);

  const name = el('span', 'lib-name', preset.name);
  if (preset.hint) name.title = preset.hint;

  if (preset.needsConfig && !libs.dcsUrl) {
    item.classList.add('needs-config');
    name.title = preset.hint || 'Needs a URL';
  }

  chk.addEventListener('change', () => {
    if (preset.needsConfig && !getState().libraries.dcsUrl && chk.checked) {
      const url = prompt('URL for the DCS Standard Include (your org’s script/CSS bundle):');
      if (!url) { chk.checked = false; return; }
      updateNested('libraries', { dcsUrl: url.trim() });
      item.classList.remove('needs-config');
    }
    const enabled = new Set(getState().libraries.enabled);
    chk.checked ? enabled.add(preset.id) : enabled.delete(preset.id);
    updateNested('libraries', { enabled: [...enabled] });
    onChangeCb?.();
  });

  const pin = el('span', 'lib-pin' + (pinned ? ' pinned' : ''), pinned ? '★' : '☆');
  pin.title = pinned ? 'Unpin' : 'Pin to top';
  pin.addEventListener('click', (e) => {
    e.preventDefault();
    const pins = new Set(getState().libraries.pinned);
    pinned ? pins.delete(preset.id) : pins.add(preset.id);
    updateNested('libraries', { pinned: [...pins] });
    render();
  });

  item.append(chk, name, pin);
  return item;
}

// Ordered list for the runner: presets in catalog order, then custom URLs.
export function getEnabledLibraries() {
  const libs = getState().libraries;
  const result = [];
  for (const preset of PRESETS) {
    if (!libs.enabled.includes(preset.id)) continue;
    if (preset.needsConfig) {
      if (libs.dcsUrl) {
        const isCss = /\.css(\?|$)/i.test(libs.dcsUrl);
        result.push({ name: preset.name, js: isCss ? undefined : libs.dcsUrl, css: isCss ? libs.dcsUrl : undefined });
      }
      continue;
    }
    result.push({ name: preset.name, js: preset.js, css: preset.css });
  }
  for (const url of libs.custom) {
    const isCss = /\.css(\?|$)/i.test(url);
    result.push({ name: url, js: isCss ? undefined : url, css: isCss ? url : undefined });
  }
  return result;
}
