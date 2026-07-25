// Console panel: renders harness console/error events, run dividers,
// groups, console.table, filters, the REPL input line, and clickable
// user-JS stack frames.

import { renderValue, renderTable, el } from './inspect/tree-view.js';
import { enhance } from './inspect/sp-shapes.js';

let out, groupStack, replHistory, replIndex;
let deps = {};

const FILTER_DEBOUNCE_MS = 150;

export function initConsolePanel({ evalInFrame, mapSrcdocLine, gotoJsLine, isConsoleVisible }) {
  deps = { evalInFrame, mapSrcdocLine, gotoJsLine, isConsoleVisible };
  out = document.getElementById('console-out');
  groupStack = [];
  replHistory = [];
  replIndex = -1;

  document.getElementById('btn-clear-console').addEventListener('click', clear);

  // Level + text filters.
  for (const btn of document.querySelectorAll('.lvl-filter')) {
    btn.addEventListener('click', () => {
      btn.classList.toggle('active');
      applyFilters();
    });
  }
  // Debounced: each keystroke otherwise re-filtered every entry in the
  // panel synchronously, which stutters once a session has logged a lot.
  let filterTimer = null;
  document.getElementById('console-filter-text').addEventListener('input', () => {
    clearTimeout(filterTimer);
    filterTimer = setTimeout(applyFilters, FILTER_DEBOUNCE_MS);
  });
  refreshFilterState();

  // REPL input.
  const input = document.getElementById('console-input');
  input.addEventListener('keydown', async (e) => {
    if (e.key === 'Enter' && input.value.trim()) {
      const code = input.value;
      input.value = '';
      replHistory.push(code);
      replIndex = replHistory.length;
      addEntry('log', [el('span', '', code)], { cls: 'repl-echo' });
      const res = await deps.evalInFrame(code);
      const body = renderNodeSmart(res.value);
      if (res.awaited) {
        const tag = el('span', 'sp-badge', 'awaited');
        addEntry(res.ok ? 'log' : 'error', [tag, body], { cls: 'repl-result' });
      } else {
        addEntry(res.ok ? 'log' : 'error', [body], { cls: 'repl-result' });
      }
    } else if (e.key === 'ArrowUp') {
      if (replIndex > 0) { replIndex--; input.value = replHistory[replIndex]; e.preventDefault(); }
    } else if (e.key === 'ArrowDown') {
      if (replIndex < replHistory.length - 1) { replIndex++; input.value = replHistory[replIndex]; }
      else { replIndex = replHistory.length; input.value = ''; }
      e.preventDefault();
    }
  });

  return { handlers: makeHandlers(), clear, runDivider };
}

// Route a serialized node through the SP-aware inspector first; it either
// returns augmented DOM or defers to the generic tree.
function renderNodeSmart(node) {
  return enhance(node) ?? renderValue(node);
}

function makeHandlers() {
  return {
    console: (d) => {
      const parts = d.args.map((a) =>
        a.t === 'str' ? el('span', '', a.v) : renderNodeSmart(a));
      addEntry(d.level, parts);
    },
    table: (d) => addEntry('log', [renderTable(d.data, d.columns)]),
    group: (d) => {
      const header = el('div', 'console-entry group-header' + (d.collapsed ? ' collapsed' : ''));
      header.append(el('span', 'twist', '▶'));
      header.append(el('span', 'entry-body', d.label));
      const container = el('div', 'console-group');
      header.addEventListener('click', () => header.classList.toggle('collapsed'));
      currentContainer().append(header, container);
      groupStack.push(container);
      scrollIfPinned();
    },
    groupEnd: () => { groupStack.pop(); },
    clear: () => clear(),
    error: (d) => {
      const parts = [];
      let msg = d.message || 'Error';
      if (d.rejection) msg = 'Uncaught (in promise): ' + msg;
      parts.push(el('span', '', msg));
      if (d.rejection && d.reason && d.reason.t !== 'str' && d.reason.t !== 'err') {
        parts.push(renderNodeSmart(d.reason));
      }
      const stackEl = renderStack(d);
      if (stackEl) parts.push(stackEl);
      addEntry('error', parts, { badge: true });
    },
  };
}

// Parse a stack/source for about:srcdoc line refs and produce clickable
// links that jump to the corresponding user-JS editor line.
function renderStack(d) {
  const lines = [];
  if (d.stack) lines.push(...d.stack.split('\n').slice(1, 8));
  else if (d.source && d.line) lines.push(`    at ${d.source}:${d.line}:${d.col ?? 0}`);
  if (!lines.length) return null;

  const wrap = el('div', 'stack-frame');
  for (const line of lines) {
    const div = el('div');
    const m = line.match(/(about:srcdoc):(\d+):(\d+)/);
    if (m) {
      const userLine = deps.mapSrcdocLine(Number(m[2]));
      if (userLine) {
        const pre = line.slice(0, m.index);
        const link = el('a', '', `js:${userLine}:${m[3]}`);
        link.addEventListener('click', () => deps.gotoJsLine(userLine));
        div.append(el('span', '', pre), link);
        wrap.append(div);
        continue;
      }
    }
    div.textContent = line;
    wrap.append(div);
  }
  return wrap;
}

function currentContainer() {
  return groupStack.length ? groupStack[groupStack.length - 1] : out;
}

function addEntry(level, parts, { cls, badge } = {}) {
  const entry = el('div', `console-entry lvl-${level} new-entry${cls ? ' ' + cls : ''}`);
  entry.dataset.lvl = level === 'info' || level === 'debug' ? 'log' : level;
  const ts = new Date();
  entry.append(el('span', 'entry-ts',
    `${String(ts.getHours()).padStart(2, '0')}:${String(ts.getMinutes()).padStart(2, '0')}:${String(ts.getSeconds()).padStart(2, '0')}`));
  const body = el('span', 'entry-body');
  for (const p of parts) body.append(p, ' ');
  entry.append(body);
  applyFilterTo(entry);
  currentContainer().append(entry);
  scrollIfPinned();

  if (badge && !deps.isConsoleVisible()) {
    const b = document.getElementById('console-badge');
    b.hidden = false;
    b.textContent = String(Number(b.textContent || 0) + 1);
  }
}

function runDivider(runNumber) {
  groupStack = [];
  const div = el('div', 'run-divider new-entry');
  const ts = new Date().toLocaleTimeString();
  div.append(el('span', 'rd-mark', '▞ ▶'), el('span', '', `run #${runNumber} · ${ts}`));
  out.append(div);
  scrollIfPinned(true);
}

function clear() {
  out.textContent = '';
  groupStack = [];
  const b = document.getElementById('console-badge');
  b.hidden = true; b.textContent = '';
}

// The active levels and the filter text are read once per filter change
// and cached: applyFilterTo() runs for every entry in the panel *and*
// on the hot path in addEntry(), so querying them per entry made a
// chatty run pay a document query per logged line.
let filterState = { lvls: new Set(), text: '' };

function refreshFilterState() {
  filterState = {
    lvls: new Set([...document.querySelectorAll('.lvl-filter.active')].map((b) => b.dataset.lvl)),
    text: document.getElementById('console-filter-text').value.trim().toLowerCase(),
  };
}

function applyFilters() {
  refreshFilterState();
  for (const entry of out.querySelectorAll('.console-entry')) applyFilterTo(entry);
}

function applyFilterTo(entry) {
  if (!entry.dataset.lvl) return;
  // err button governs 'error'; warn → 'warn'; log → everything else.
  const lvl = entry.dataset.lvl === 'error' ? 'error' : entry.dataset.lvl === 'warn' ? 'warn' : 'log';
  entry.classList.toggle('hidden-lvl', !filterState.lvls.has(lvl));
  // Match the logged payload only — entry.textContent would also cover
  // the injected timestamp, so filtering for "10" hit every line logged
  // at :10 seconds.
  const body = entry.querySelector('.entry-body');
  const bodyText = body ? body.textContent.toLowerCase() : '';
  entry.classList.toggle('hidden-txt', !!filterState.text && !bodyText.includes(filterState.text));
}

function scrollIfPinned(force) {
  const nearBottom = out.scrollHeight - out.scrollTop - out.clientHeight < 60;
  if (nearBottom || force) out.scrollTop = out.scrollHeight;
}
