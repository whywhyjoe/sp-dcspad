// Workbench data grid: sortable, filterable table over plain row objects.
//
// Plain DOM, no framework. Columns declare how to read and format values;
// the grid owns sorting, the filter box, the count badge, and per-cell
// click-to-copy. Export and "copy as script" actions plug into the toolbar
// actions slot so every view gets them uniformly.

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export function copyText(text, flashEl) {
  const done = () => {
    if (!flashEl) return;
    flashEl.classList.add('copied');
    setTimeout(() => flashEl.classList.remove('copied'), 900);
  };
  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(text).then(done, done);
  } else {
    const ta = document.createElement('textarea');
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch { /* best effort */ }
    ta.remove();
    done();
  }
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

// columns: [{ key, label, value?(row), format?(v,row), copyable?, mono?, width? }]
export function createGrid({
  columns,
  rowKey = 'Id',
  onOpen = null,
  emptyText = 'No rows.',
  filterPlaceholder = 'Filter…',
} = {}) {
  let rows = [];
  let visible = [];
  let sortKey = null;
  let sortDir = 1;
  let filterText = '';

  const root = el('div', 'wb-grid');
  const toolbar = el('div', 'wb-grid-toolbar');
  const count = el('span', 'wb-grid-count', '—');
  const filter = el('input', 'wb-grid-filter');
  filter.type = 'search';
  filter.placeholder = filterPlaceholder;
  filter.setAttribute('aria-label', 'Filter rows');
  const actions = el('span', 'wb-grid-actions');
  toolbar.append(count, filter, actions);

  const scroller = el('div', 'wb-grid-scroll');
  const table = el('table', 'wb-table');
  const thead = el('thead');
  const headRow = el('tr');
  for (const col of columns) {
    const th = el('th', '', col.label ?? col.key);
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
    headRow.append(th);
  }
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

    count.textContent = filterText || visible.length !== rows.length
      ? `${visible.length} / ${rows.length}`
      : String(rows.length);

    for (const th of headRow.children) {
      const col = columns[[...headRow.children].indexOf(th)];
      th.querySelector('.wb-sort-arrow').textContent =
        col && col.key === sortKey ? (sortDir === 1 ? ' ▲' : ' ▼') : '';
    }

    tbody.textContent = '';
    if (!visible.length) {
      const tr = el('tr');
      const td = el('td', 'wb-empty', rows.length ? 'No rows match the filter.' : emptyText);
      td.colSpan = columns.length;
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
        tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') onOpen(row); });
      }
      tr.dataset.key = String(row[rowKey] ?? '');
      for (const col of columns) {
        const td = el('td', col.mono ? 'wb-mono' : '');
        const text = displayValue(row, col);
        if (col.copyable && text) {
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
    setError(err) {
      status.textContent = err?.message || String(err);
      status.className = 'wb-grid-status wb-error';
      status.hidden = false;
    },
    getVisibleRows: () => [...visible],
    getColumns: () => columns,
  };
}
