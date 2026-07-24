// Expandable tree renderer for the harness's type-tagged value nodes.
// Generic layer only — SP-specific smart views live in sp-shapes.js and
// wrap/augment what this module produces.

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

export function renderValue(node, opts = {}) {
  if (!node) return el('span', 't-undef', 'undefined');
  switch (node.t) {
    case 'str': {
      const s = el('span', 't-str', opts.bare ? node.v : JSON.stringify(node.v));
      if (node.trunc) s.append(el('span', 't-truncated', ` …(${node.trunc} chars)`));
      return s;
    }
    case 'num': return el('span', 't-num', String(node.v));
    case 'bool': return el('span', 't-bool', String(node.v));
    case 'null': return el('span', 't-null', 'null');
    case 'undef': return el('span', 't-undef', 'undefined');
    case 'sym': return el('span', 't-str', node.v);
    case 'fn': return el('span', 't-fn', `ƒ ${node.v}()`);
    case 'date': return el('span', 't-node', node.v);
    case 'regex': return el('span', 't-str', node.v);
    case 'node': return el('span', 't-node', node.v);
    case 'circ': return el('span', 't-circular', '[circular]');
    case 'maxdepth': return el('span', 't-preview', node.v);
    case 'err': return renderError(node);
    case 'arr': return renderExpandable(node, `Array(${node.n})`, node.items.map((item, i) => [String(i), item]), opts);
    case 'obj': return renderExpandable(node, node.cls === 'Object' ? '' : node.cls, node.keys, opts);
    default: return el('span', 't-preview', JSON.stringify(node));
  }
}

function renderError(node) {
  const wrap = el('span');
  let head = `${node.name}: ${node.msg}`;
  if (node.status !== undefined) head += ` (HTTP ${node.status}${node.statusText ? ' ' + node.statusText : ''})`;
  wrap.append(el('span', 't-err', head));
  if (node.stack) {
    const stack = el('div', 'stack-frame');
    stack.textContent = node.stack.split('\n').slice(1, 6).join('\n');
    wrap.append(stack);
  }
  return wrap;
}

function previewOf(node) {
  switch (node.t) {
    case 'str': { const v = node.v.length > 24 ? node.v.slice(0, 24) + '…' : node.v; return JSON.stringify(v); }
    case 'num': case 'bool': return String(node.v);
    case 'null': return 'null';
    case 'undef': return 'undefined';
    case 'fn': return 'ƒ';
    case 'arr': return `Array(${node.n})`;
    case 'obj': return node.cls === 'Object' ? '{…}' : `${node.cls}`;
    case 'err': return node.name;
    case 'node': return node.v;
    case 'date': return node.v;
    case 'maxdepth': return node.v;
    case 'circ': return '[circular]';
    default: return '…';
  }
}

function renderExpandable(node, label, entries, opts = {}) {
  const wrap = el('span', 'tree-node');
  const row = el('span', 'tree-row expandable');
  row.append(el('span', 'twist', '▶'));
  if (label) row.append(el('span', '', label + ' '));

  // Inline preview: first few key: value pairs.
  const parts = entries.slice(0, 5).map(([k, v]) =>
    (node.t === 'arr' ? '' : `${k}: `) + previewOf(v));
  const openBrace = node.t === 'arr' ? '[' : '{';
  const closeBrace = node.t === 'arr' ? ']' : '}';
  const more = entries.length > 5 || node.trunc ? ', …' : '';
  row.append(el('span', 't-preview', `${openBrace}${parts.join(', ')}${more}${closeBrace}`));
  wrap.append(row);

  const children = el('div', 'tree-children');
  wrap.append(children);
  let built = false;

  row.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = row.classList.toggle('open');
    if (open && !built) {
      built = true;
      for (const [key, val] of entries) {
        const line = el('div');
        const keySpan = el('span', 'tree-key' + (opts.dimKeys?.has?.(key) ? ' dim-key' : ''), key);
        line.append(keySpan, el('span', '', ': '), renderValue(val, opts));
        children.append(line);
      }
      if (node.trunc) children.append(el('div', 't-truncated', '… truncated'));
    }
  });
  return wrap;
}

// console.table / SP list-item table rendering.
// dataNode: serialized arr-of-obj (or obj-of-obj); columns: optional list.
export function renderTable(dataNode, columns) {
  if (!dataNode || (dataNode.t !== 'arr' && dataNode.t !== 'obj')) {
    return renderValue(dataNode);
  }
  const rows = dataNode.t === 'arr'
    ? dataNode.items.map((item, i) => [String(i), item])
    : dataNode.keys;

  // Collect column set from row objects.
  let cols = columns ? [...columns] : [];
  if (!cols.length) {
    const seen = new Set();
    for (const [, v] of rows) {
      if (v.t === 'obj') for (const [k] of v.keys) seen.add(k);
      else if (v.t === 'arr') v.items.forEach((_, i) => seen.add(String(i)));
      else seen.add('Value');
    }
    cols = [...seen].slice(0, 20);
  }

  const wrap = el('div', 'console-table-wrap');
  const table = el('table', 'console-table');
  const thead = el('thead');
  const hr = el('tr');
  hr.append(el('th', '', '(index)'));
  cols.forEach((c) => hr.append(el('th', '', c)));
  thead.append(hr);
  table.append(thead);

  const tbody = el('tbody');
  for (const [key, v] of rows) {
    const tr = el('tr');
    tr.append(el('td', '', key));
    for (const c of cols) {
      const td = el('td');
      let cell;
      if (v.t === 'obj') cell = v.keys.find(([k]) => k === c)?.[1];
      else if (v.t === 'arr') cell = v.items[Number(c)];
      else if (c === 'Value') cell = v;
      td.textContent = cell ? previewOf(cell) : '';
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  wrap.append(table);
  return wrap;
}

export { el, previewOf };
