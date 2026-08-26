// Workbench data grid: sortable, filterable table over plain row objects.
//
// Plain DOM, no framework. Columns declare how to read and format values;
// the grid owns sorting, the filter box, the count badge, per-cell
// click-to-copy, and the export menu (visible rows → CSV/JSON/Markdown).
// "Copy as script" actions plug into the same toolbar actions slot (M4).

import {
  copyText, toCsv, toJson, downloadCsv, downloadJson,
} from './export.js';
import { toPnpjs2, toRestFetch, toPnpPowerShell } from './scriptgen.js';
import { showFailure } from './denied.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export { copyText };

// Decoded server-relative path → href path, one encode per segment.
// encodeURI would leave '#' unescaped, turning a legal SharePoint file
// name like "A#B.docx" into a URL fragment.
export const encodeSpPath = (path) =>
  String(path).split('/').map(encodeURIComponent).join('/');

// New-tab anchors: modern SharePoint pages intercept link clicks at the
// document level for SPA routing and can swallow same-origin
// target="_blank" navigations. Opening explicitly from our own handler is
// deterministic; the target/rel attributes stay as semantics + fallback.
export function bindNewTab(a) {
  a.target = '_blank';
  a.rel = 'noopener';
  a.addEventListener('click', (e) => {
    // Hiding the click from the host's router is enough for modified
    // clicks — let the browser keep its native ctrl/cmd/shift semantics
    // (background tab, new window). Plain clicks open explicitly, which
    // the host page cannot cancel.
    e.stopPropagation();
    if (e.ctrlKey || e.metaKey || e.shiftKey || e.altKey || e.button === 1) return;
    e.preventDefault();
    window.open(a.href, '_blank', 'noopener');
  });
  return a;
}

const cellValue = (row, col) =>
  typeof col.value === 'function' ? col.value(row) : row[col.key];

function displayValue(row, col) {
  const v = cellValue(row, col);
  if (typeof col.format === 'function') return col.format(v, row);
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'Yes' : '';
  return String(v);
}

// columns: [{ key, label, value?(row), format?(v,row), render?(v,row)=>Node,
//             link?(v,row)=>href, copyable?, mono?, num?, width? }]
// link renders the cell text as an anchor opening href in a new tab (URL
// columns stay real URLs); copyable then adds a ⧉ copy glyph beside it.
// action marks a column that exists only to host per-row controls: it renders
// like any other, but every export drops it (see export.js dataColumns).
// exportName enables the toolbar export menu; it's the download file stem.
// descriptor { path, options, webUrl } enables the "Copy as…" script menu.
// toolbarExtras: a Node adopted into the toolbar after the filter box, so a
// view's own controls share the grid's single toolbar row (Items tab).
// exportExtras: [[label, run]] entries prepended to the Export menu.
// selectable: adds a leading checkbox column; clicking a row (when the grid
// has no onOpen) toggles it, and every export/copy operates on the selected
// rows when any are selected, the visible rows otherwise.
// subject: what this grid lists ('subwebs', 'the groups on this web') — used
// to say what a denied read could not show. See setError and denied.js.
export function createGrid({
  columns,
  rowKey = 'Id',
  onOpen = null,
  emptyText = 'No rows.',
  filterPlaceholder = 'Filter…',
  exportName = '',
  descriptor = null,
  toolbarExtras = null,
  exportExtras = [],
  selectable = false,
  subject = '',
} = {}) {
  let rows = [];
  let visible = [];
  let sortKey = null;
  let sortDir = 1;
  let filterText = '';
  const selectedKeys = new Set();
  const keyOf = (row) => String(row?.[rowKey] ?? '');

  // Selected rows (in current view order) when a selection exists, else the
  // visible rows — every export/copy path funnels through this.
  const exportRows = () => {
    const chosen = visible.filter((row) => selectedKeys.has(keyOf(row)));
    return chosen.length ? chosen : visible;
  };

  const root = el('div', 'wb-grid');
  const toolbar = el('div', 'wb-grid-toolbar');
  const count = el('span', 'wb-grid-count', '—');
  const filter = el('input', 'wb-grid-filter');
  filter.type = 'search';
  filter.placeholder = filterPlaceholder;
  filter.setAttribute('aria-label', 'Filter rows');
  const actions = el('span', 'wb-grid-actions');
  toolbar.append(count, filter, ...(toolbarExtras ? [toolbarExtras] : []), actions);

  function menuButton(label, title, items) {
    const wrap = el('span', 'wb-menu-wrap');
    const btn = el('button', 'btn btn-xs', label);
    btn.type = 'button';
    btn.title = title;
    const menu = el('div', 'wb-menu');
    menu.hidden = true;
    for (const [itemLabel, run] of items) {
      const item = el('button', 'wb-menu-item', itemLabel);
      item.type = 'button';
      item.addEventListener('click', () => { menu.hidden = true; run(btn); });
      menu.append(item);
    }
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      menu.hidden = !menu.hidden;
    });
    document.addEventListener('click', () => { menu.hidden = true; });
    wrap.append(btn, menu);
    actions.append(wrap);
  }

  if (descriptor) {
    menuButton('Copy as ▾', 'Copy this query as a runnable script', [
      ['PnPjs 2 (DCSPad pane)', (btn) => copyText(toPnpjs2(descriptor), btn)],
      ['REST fetch', (btn) => copyText(toRestFetch(descriptor, descriptor.webUrl), btn)],
      ['PnP.PowerShell', (btn) => copyText(toPnpPowerShell(descriptor, descriptor.webUrl), btn)],
    ]);
  }

  if (exportName) {
    // No markdown-table entry here: the list-content .md document (Items
    // tab exportExtras) is the useful markdown shape; pipe tables aren't.
    menuButton('Export ▾', 'Export the selected rows, or all visible rows', [
      ...exportExtras,
      ['Download CSV', () => downloadCsv(exportName, exportRows(), columns)],
      ['Download JSON', () => downloadJson(exportName, exportRows(), columns)],
      ['Copy CSV', (btn) => copyText(toCsv(exportRows(), columns), btn)],
      ['Copy JSON', (btn) => copyText(toJson(exportRows(), columns), btn)],
    ]);
  }

  const scroller = el('div', 'wb-grid-scroll');
  const table = el('table', 'wb-table');
  const thead = el('thead');
  const headRow = el('tr');

  // ---- column resizing ----
  // Dragging a header edge freezes every column at its current width and
  // switches the table to fixed layout; further drags adjust one column and
  // the table's total width, overflowing into the scroller.
  let widthsFrozen = false;
  function freezeWidths() {
    if (widthsFrozen) return;
    widthsFrozen = true;
    for (const th of headRow.children) {
      th.style.width = `${Math.round(th.getBoundingClientRect().width)}px`;
    }
    table.style.tableLayout = 'fixed';
    syncTableWidth();
  }
  function syncTableWidth() {
    let total = 0;
    for (const th of headRow.children) total += parseFloat(th.style.width) || 0;
    if (total) table.style.width = `${total}px`;
  }
  function attachResizer(th) {
    const handle = el('span', 'wb-col-resize');
    handle.title = 'Drag to resize';
    handle.addEventListener('click', (e) => e.stopPropagation());
    handle.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      freezeWidths();
      const startX = e.clientX;
      const startW = parseFloat(th.style.width) || th.getBoundingClientRect().width;
      handle.classList.add('dragging');
      handle.setPointerCapture(e.pointerId);
      const move = (ev) => {
        th.style.width = `${Math.max(48, Math.round(startW + (ev.clientX - startX)))}px`;
        syncTableWidth();
      };
      const up = () => {
        handle.classList.remove('dragging');
        handle.removeEventListener('pointermove', move);
        handle.removeEventListener('pointerup', up);
      };
      handle.addEventListener('pointermove', move);
      handle.addEventListener('pointerup', up);
    });
    th.append(handle);
  }

  let selectAll = null;
  if (selectable) {
    const th = el('th', 'wb-select-col');
    selectAll = el('input');
    selectAll.type = 'checkbox';
    selectAll.className = 'wb-select-all';
    selectAll.setAttribute('aria-label', 'Select all visible rows');
    selectAll.title = 'Select all visible rows';
    selectAll.addEventListener('change', () => {
      if (selectAll.checked) visible.forEach((row) => selectedKeys.add(keyOf(row)));
      else visible.forEach((row) => selectedKeys.delete(keyOf(row)));
      render();
    });
    th.append(selectAll);
    headRow.append(th);
  }

  columns.forEach((col, colIndex) => {
    const th = el('th', '', col.label ?? col.key);
    th.dataset.colIndex = String(colIndex);
    if (col.num) th.classList.add('wb-num');
    if (col.width) th.style.width = col.width;
    th.tabIndex = 0;
    th.title = `Sort by ${col.label ?? col.key}`;
    const arrow = el('span', 'wb-sort-arrow', '');
    th.append(arrow);
    const sortBy = () => {
      if (sortKey === col.key) sortDir = -sortDir;
      else { sortKey = col.key; sortDir = 1; }
      render();
    };
    th.addEventListener('click', sortBy);
    th.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(); } });
    attachResizer(th);
    headRow.append(th);
  });
  thead.append(headRow);
  const tbody = el('tbody');
  table.append(thead, tbody);
  scroller.append(table);

  const notice = el('div', 'wb-grid-notice');
  notice.hidden = true;
  const status = el('div', 'wb-grid-status');
  status.hidden = true;

  root.append(toolbar, notice, scroller, status);

  filter.addEventListener('input', () => {
    filterText = filter.value.trim().toLowerCase();
    render();
  });

  function matches(row) {
    if (!filterText) return true;
    return columns.some((col) => displayValue(row, col).toLowerCase().includes(filterText));
  }

  function compare(a, b) {
    const col = columns.find((c) => c.key === sortKey);
    if (!col) return 0;
    const va = cellValue(a, col);
    const vb = cellValue(b, col);
    if (va === vb) return 0;
    if (va === null || va === undefined) return 1;
    if (vb === null || vb === undefined) return -1;
    if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * sortDir;
    return String(va).localeCompare(String(vb), undefined, { sensitivity: 'base' }) * sortDir;
  }

  function render() {
    visible = rows.filter(matches);
    if (sortKey) visible = [...visible].sort(compare);

    const selectedVisible = selectable
      ? visible.filter((row) => selectedKeys.has(keyOf(row))).length
      : 0;
    const base = filterText || visible.length !== rows.length
      ? `${visible.length} / ${rows.length}`
      : String(rows.length);
    count.textContent = selectedVisible ? `${base} · ${selectedVisible} selected` : base;
    if (selectAll) {
      selectAll.checked = visible.length > 0 && selectedVisible === visible.length;
      selectAll.indeterminate = selectedVisible > 0 && selectedVisible < visible.length;
    }

    for (const th of headRow.children) {
      if (th.dataset.colIndex === undefined) continue;
      const col = columns[Number(th.dataset.colIndex)];
      const selected = Boolean(col && col.key === sortKey);
      th.querySelector('.wb-sort-arrow').textContent =
        selected ? (sortDir === 1 ? ' ▲' : ' ▼') : '';
      th.classList.toggle('is-sorted', selected);
      th.setAttribute('aria-sort', selected ? (sortDir === 1 ? 'ascending' : 'descending') : 'none');
    }

    tbody.textContent = '';
    if (!visible.length) {
      const tr = el('tr');
      const td = el('td', 'wb-empty', rows.length ? 'No rows match the filter.' : emptyText);
      td.colSpan = columns.length + (selectable ? 1 : 0);
      tr.append(td);
      tbody.append(tr);
      return;
    }
    for (const row of visible) {
      const tr = el('tr');
      if (onOpen) {
        tr.className = 'wb-row-openable';
        tr.tabIndex = 0;
        tr.addEventListener('click', () => onOpen(row));
        // Enter on a focused cell link must open the link, not also drill
        // into the row — only act when the row itself has focus.
        tr.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target === tr) onOpen(row); });
      }
      tr.dataset.key = String(row[rowKey] ?? '');
      if (selectable) {
        const key = keyOf(row);
        tr.classList.toggle('wb-row-selected', selectedKeys.has(key));
        const toggle = () => {
          if (selectedKeys.has(key)) selectedKeys.delete(key);
          else selectedKeys.add(key);
          render();
        };
        const td = el('td', 'wb-select-cell');
        const box = el('input');
        box.type = 'checkbox';
        box.className = 'wb-row-check';
        box.checked = selectedKeys.has(key);
        box.setAttribute('aria-label', 'Select row');
        box.addEventListener('click', (e) => e.stopPropagation());
        box.addEventListener('change', toggle);
        // On a grid that also drills down, the whole checkbox cell is a
        // selection target — a click a few pixels off the box must not open
        // the row instead of ticking it.
        if (onOpen) {
          td.addEventListener('click', (e) => {
            e.stopPropagation();
            if (e.target !== box) toggle();
          });
        }
        td.append(box);
        tr.append(td);
        if (!onOpen) {
          // No drill-down on this grid — the whole row is a selection
          // target, except its interactive bits (links, copy glyphs…).
          tr.classList.add('wb-row-selectable');
          tr.addEventListener('click', (e) => {
            if (e.target.closest('a, button, input, .sp-copy')) return;
            toggle();
          });
        }
      }
      for (const col of columns) {
        const td = el('td', [col.mono ? 'wb-mono' : '', col.num ? 'wb-num' : ''].filter(Boolean).join(' '));
        const text = displayValue(row, col);
        if (typeof col.render === 'function') {
          // Custom cell node (e.g. anchors); displayValue still drives
          // filter/sort/export, so renders stay data-consistent.
          const node = col.render(cellValue(row, col), row);
          if (node) td.append(node);
          tr.append(td);
          continue;
        }
        const href = typeof col.link === 'function' && text
          ? String(col.link(cellValue(row, col), row) || '')
          : '';
        if (href) {
          const a = el('a', 'wb-cell-url', text);
          a.href = href;
          a.title = 'Open in a new tab';
          bindNewTab(a);
          td.append(a);
          if (col.copyable) {
            const glyph = el('span', 'sp-copy wb-cell-copy', '⧉');
            glyph.title = 'Click to copy';
            glyph.addEventListener('click', (e) => { e.stopPropagation(); copyText(text, glyph); });
            td.append(glyph);
          }
        } else if (col.copyable && text) {
          const span = el('span', 'sp-copy', text);
          span.title = 'Click to copy';
          span.addEventListener('click', (e) => { e.stopPropagation(); copyText(text, span); });
          td.append(span);
        } else {
          td.textContent = text;
        }
        tr.append(td);
      }
      tbody.append(tr);
    }
  }

  return {
    el: root,
    actionsEl: actions,
    setRows(next, { partial = false } = {}) {
      rows = Array.isArray(next) ? next : [];
      selectedKeys.clear();   // new data — a stale selection must not scope exports
      status.hidden = true;
      notice.hidden = !partial;
      if (partial) {
        notice.textContent =
          '⚠ Partial result set — the server returned more pages than the workbench cap.';
      }
      render();
    },
    setLoading(message = 'Loading…') {
      status.textContent = message;
      status.className = 'wb-grid-status';
      status.hidden = false;
    },
    // A denial is reported, not alarmed about: neutral register, plain
    // sentence, SharePoint's own words on the tooltip. Everything else stays
    // loud. See denied.js for why.
    setError(err) {
      status.className = 'wb-grid-status';
      showFailure(status, err, subject);
    },
    getVisibleRows: () => [...visible],
    getExportRows: () => [...exportRows()],
    getColumns: () => columns,
  };
}
