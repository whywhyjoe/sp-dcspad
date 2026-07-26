// Network panel: rows for every fetch/XHR the preview makes, an
// "_api only" filter, and a detail pane that renders JSON responses
// through the SP-aware inspector.

import { renderValue, el } from './inspect/tree-view.js';
import { enhance } from './inspect/sp-shapes.js';

const requests = new Map();   // id -> { row, data }
let selectedId = null;
let deps = {};

export function initNetworkPanel({ isNetworkVisible }) {
  deps = { isNetworkVisible };
  document.getElementById('btn-clear-network').addEventListener('click', clear);
  document.getElementById('chk-api-only').addEventListener('change', applyApiFilter);
  return { handlers: { 'net-start': onStart, 'net-end': onEnd }, clear };
}

const isApiUrl = (url) => /\/_api\/|\/_vti_bin\//i.test(url);

const fmtSize = (bytes) =>
  bytes < 1024 ? `${bytes} B`
  : bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB`
  : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;

function onStart(d) {
  const tbody = document.getElementById('network-rows');
  const row = el('tr', 'network-row net-pending');
  row.dataset.id = d.id;
  if (isApiUrl(d.url)) row.classList.add('is-api');

  const tdMethod = el('td', 'net-method', d.method);
  const tdUrl = el('td', 'net-url', d.url);
  tdUrl.title = d.url;
  const tdStatus = el('td', 'net-status net-status-pending', '…');
  const tdTime = el('td', '', '');
  const tdSize = el('td', '', '');
  row.append(tdMethod, tdUrl, tdStatus, tdTime, tdSize);

  row.addEventListener('click', () => select(d.id));
  requests.set(d.id, { row, data: { ...d } });
  applyApiFilterTo(row);
  tbody.append(row);

  const wrap = document.getElementById('network-table-wrap');
  wrap.scrollTop = wrap.scrollHeight;
}

function onEnd(d) {
  const entry = requests.get(d.id);
  if (!entry) return;
  Object.assign(entry.data, d);
  const [, , tdStatus, tdTime, tdSize] = entry.row.children;
  tdStatus.textContent = d.failed ? '✕ failed' : String(d.status);
  tdStatus.className = 'net-status ' + (d.ok ? 'net-status-ok' : 'net-status-err');
  tdTime.textContent = `${d.ms} ms`;
  tdSize.textContent = d.size != null ? fmtSize(d.size) : '—';
  if (selectedId === d.id) renderDetail(entry.data);
  // Error dot on the Network tab: any failed request or 4xx/5xx status.
  // Cleared with the request log, not on tab focus.
  if (!d.ok) document.getElementById('network-badge').hidden = false;
}

function select(id) {
  selectedId = id;
  for (const { row } of requests.values()) row.classList.toggle('selected', row.dataset.id === id);
  renderDetail(requests.get(id).data);
}

function renderDetail(data) {
  const detail = document.getElementById('network-detail');
  detail.hidden = false;
  detail.textContent = '';

  const close = el('span', 'nd-close', '✕');
  close.addEventListener('click', () => {
    detail.hidden = true;
    selectedId = null;
    for (const { row } of requests.values()) row.classList.remove('selected');
  });
  detail.append(close);

  detail.append(el('h4', '', 'Request'));
  const kv = (label, value) => {
    const div = el('div', 'nd-kv');
    div.append(el('b', '', label + ': '), el('span', '', value ?? '—'));
    return div;
  };
  detail.append(kv('Method', data.method), kv('URL', data.url), kv('Via', data.api === 'xhr' ? 'XMLHttpRequest' : 'fetch'));

  detail.append(el('h4', '', 'Response'));
  if (data.cancelled) {
    detail.append(el('div', 'net-status-err', 'cancelled — the frame was replaced by a new run before the response arrived'));
    return;
  }
  if (data.status === undefined) {
    detail.append(el('div', 'net-status-pending', 'pending…'));
    return;
  }
  detail.append(
    kv('Status', `${data.status} ${data.statusText || ''}`),
    kv('Duration', `${data.ms} ms`),
    kv('Size', data.size != null ? fmtSize(data.size) : '—'),
    kv('Content-Type', data.contentType || '—'),
  );

  detail.append(el('h4', '', 'Body'));
  if (!data.preview) {
    detail.append(el('div', 't-preview', data.failed ? `request failed: ${data.statusText}` : '(no text preview)'));
    return;
  }
  if ((data.contentType || '').includes('json') || /^[\[{]/.test(data.preview.trim())) {
    try {
      const parsed = JSON.parse(data.preview);
      const node = toNode(parsed, 0);
      detail.append(enhance(node) ?? renderValue(node));
      return;
    } catch { /* fall through to raw text */ }
  }
  const pre = el('pre', '', data.preview.slice(0, 5000));
  pre.style.whiteSpace = 'pre-wrap';
  detail.append(pre);
}

// Convert a parsed JSON value into the harness's serialized-node format
// so the inspector/tree renderer can be reused on network bodies.
function toNode(v, depth) {
  if (v === null) return { t: 'null' };
  switch (typeof v) {
    case 'string': return { t: 'str', v };
    case 'number': return { t: 'num', v };
    case 'boolean': return { t: 'bool', v };
    case 'undefined': return { t: 'undef' };
  }
  if (depth >= 6) return { t: 'maxdepth', v: Array.isArray(v) ? `Array(${v.length})` : '{…}' };
  if (Array.isArray(v)) {
    return { t: 'arr', n: v.length, items: v.slice(0, 100).map((x) => toNode(x, depth + 1)), trunc: v.length > 100 };
  }
  const keys = Object.keys(v);
  return {
    t: 'obj', cls: 'Object',
    keys: keys.slice(0, 100).map((k) => [k, toNode(v[k], depth + 1)]),
    trunc: keys.length > 100,
  };
}

function clear() {
  requests.clear();
  selectedId = null;
  document.getElementById('network-rows').textContent = '';
  document.getElementById('network-detail').hidden = true;
  document.getElementById('network-badge').hidden = true;
}

function applyApiFilter() {
  for (const { row } of requests.values()) applyApiFilterTo(row);
}
function applyApiFilterTo(row) {
  const apiOnly = document.getElementById('chk-api-only').checked;
  row.classList.toggle('hidden-api', apiOnly && !row.classList.contains('is-api'));
}

export function markRun() {
  // The outgoing frame is about to be destroyed, so its in-flight
  // requests can never complete — settle them as cancelled instead of
  // leaving them pending forever.
  for (const { row, data } of requests.values()) {
    if (data.status !== undefined || data.cancelled) continue;
    data.cancelled = true;
    const tdStatus = row.children[2];
    tdStatus.textContent = '✕ cancelled';
    tdStatus.className = 'net-status net-status-err';
    row.classList.remove('net-pending');
    if (selectedId === data.id) renderDetail(data);
  }

  // Keep requests from prior runs but visually separate them.
  const tbody = document.getElementById('network-rows');
  if (tbody.children.length) {
    const sep = el('tr', 'net-run-sep');
    const td = el('td', '', '— new run —');
    td.colSpan = 5;
    td.style.color = 'var(--fg-faint)';
    td.style.textAlign = 'center';
    sep.append(td);
    tbody.append(sep);
  }
}
