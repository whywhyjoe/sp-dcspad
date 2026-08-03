// Quick query builder: pick a list (or type any /_api endpoint), compose
// $select / $filter / $orderby / $expand / $top visually, run it into the
// standard grid. The composed query IS a scriptgen descriptor, so "Copy as
// PnPjs 2 / REST / PnP.PowerShell" flows for free.
//
// The last query per web is remembered in sessionStorage (same per-tab
// convention as the route and site keys — invariant 6 covers localStorage).

import { createGrid } from '../grid.js?v=2';

const QUERY_KEY = 'dcspad.workbench.query';
const DEFAULT_TOP = 100;
const MAX_TOP = 5000;   // sp-rest PAGE_CAP

const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind',
  'Hidden', 'ReadOnlyField', 'Choices',
];

const NUMERIC_TYPES = new Set(['Number', 'Currency', 'Counter', 'Integer']);
const EXPANDABLE_TYPES = new Set(['User', 'UserMulti', 'Lookup', 'LookupMulti']);

const OPERATORS = [
  ['eq', '='], ['ne', '≠'], ['gt', '>'], ['ge', '≥'], ['lt', '<'], ['le', '≤'],
  ['startswith', 'starts with'], ['substringof', 'contains'],
];

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

// ---- pure helpers (exported for tests) ------------------------------------

// One $filter clause with per-type literal quoting.
export function filterClause({ field, type, op, value }) {
  const name = String(field || '').trim();
  if (!name) return '';
  const raw = String(value ?? '').trim();
  let literal;
  if (NUMERIC_TYPES.has(type)) literal = raw === '' ? '0' : String(Number(raw.replace(',', '.')));
  else if (type === 'Boolean') literal = /^(1|true|yes)$/i.test(raw) ? '1' : '0';
  else if (type === 'DateTime') {
    const d = new Date(raw);
    literal = `datetime'${Number.isNaN(d.getTime()) ? raw : d.toISOString()}'`;
  } else literal = `'${raw.replaceAll("'", "''")}'`;
  if (op === 'startswith') return `startswith(${name},${literal})`;
  if (op === 'substringof') return `substringof(${literal},${name})`;
  return `${name} ${op} ${literal}`;
}

// Compose the full $filter from builder rows (left-to-right, no parens — the
// raw editor covers grouping).
export function composeFilter(rows) {
  let out = '';
  for (const row of rows) {
    const clause = filterClause(row);
    if (!clause) continue;
    out = out ? `${out} ${row.join === 'or' ? 'or' : 'and'} ${clause}` : clause;
  }
  return out;
}

// Render the descriptor as the /_api-relative URL shown in the raw editor.
export function descriptorToRaw(descriptor) {
  const parts = [];
  const join = (v) => (Array.isArray(v) ? v.join(',') : String(v));
  const o = descriptor.options || {};
  if (o.select) parts.push(`$select=${join(o.select)}`);
  if (o.expand) parts.push(`$expand=${join(o.expand)}`);
  if (o.filter) parts.push(`$filter=${o.filter}`);
  if (o.orderby) parts.push(`$orderby=${join(o.orderby)}`);
  if (o.top) parts.push(`$top=${o.top}`);
  return `${descriptor.path}${parts.length ? `?${parts.join('&')}` : ''}`;
}

// Parse a raw /_api-relative URL back into { path, options } — the inverse of
// descriptorToRaw for round-trippable strings. Returns null when the query
// string contains anything the descriptor model can't express.
export function rawToDescriptor(raw) {
  const s = String(raw || '').trim().replace(/^\/+/, '').replace(/^_api\//, '');
  if (!s) return null;
  const q = s.indexOf('?');
  if (q === -1) return { path: s, options: {} };
  const path = s.slice(0, q);
  const options = {};
  for (const pair of s.slice(q + 1).split('&')) {
    const eq = pair.indexOf('=');
    if (eq === -1) return null;
    const key = pair.slice(0, eq);
    const value = pair.slice(eq + 1);
    if (key === '$select') options.select = value.split(',');
    else if (key === '$expand') options.expand = value.split(',');
    else if (key === '$filter') {
      try { options.filter = decodeURIComponent(value); }
      catch { return null; }
    }
    else if (key === '$orderby') options.orderby = value;
    else if (key === '$top') options.top = Number(value) || undefined;
    else return null;
  }
  return { path, options };
}

// Column defs generated from a $select list; 'A/B' paths read nested props.
export function columnsForSelect(select) {
  return (select || []).map((entry) => {
    const path = String(entry).split('/');
    return {
      key: entry,
      label: entry,
      value: (row) => {
        let v = row;
        for (const seg of path) v = v?.[seg];
        if (v === null || v === undefined) return '';
        if (typeof v === 'object') {
          const arr = Array.isArray(v) ? v : v.results;
          if (Array.isArray(arr)) return arr.map((x) => (typeof x === 'object' ? JSON.stringify(x) : x)).join(', ');
          return JSON.stringify(v);
        }
        return v;
      },
    };
  });
}

// ---- persisted state ------------------------------------------------------

function readSaved(webUrl) {
  try {
    const all = JSON.parse(sessionStorage.getItem(QUERY_KEY) || '{}');
    return all[webUrl.toLowerCase()] || null;
  } catch { return null; }
}

function writeSaved(webUrl, state) {
  try {
    const all = JSON.parse(sessionStorage.getItem(QUERY_KEY) || '{}');
    all[webUrl.toLowerCase()] = state;
    sessionStorage.setItem(QUERY_KEY, JSON.stringify(all));
  } catch { /* private mode */ }
}

// ---- view -----------------------------------------------------------------

export function createQueryView({ client }) {
  const root = el('section', 'wb-view wb-view-query');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Query builder</h2>'
    + '<p class="wb-view-hint">Compose an OData query against any list — or any '
    + '/_api endpoint — and run it. “Copy as” turns the query into a script.</p>';

  // --- composer ---
  const composer = el('div', 'wb-qb');

  const targetRow = el('div', 'wb-qb-row');
  const listSelect = el('select', 'wb-qb-list');
  listSelect.setAttribute('aria-label', 'Query target list');
  const endpointInput = el('input', 'wb-qb-endpoint');
  endpointInput.type = 'text';
  endpointInput.placeholder = 'web/currentuser — path after /_api/';
  endpointInput.hidden = true;
  targetRow.append(el('span', 'wb-qb-label', 'Target'), listSelect, endpointInput);

  const fieldsBox = el('div', 'wb-qb-fields');
  const fieldsList = el('div', 'wb-qb-fieldlist');
  fieldsBox.append(el('span', 'wb-qb-label', '$select'), fieldsList);

  const filtersBox = el('div', 'wb-qb-filters');
  const filterRows = el('div', 'wb-qb-filterrows');
  const addFilter = el('button', 'btn btn-xs', '+ Filter');
  addFilter.type = 'button';
  filtersBox.append(el('span', 'wb-qb-label', '$filter'), filterRows, addFilter);

  const optionsRow = el('div', 'wb-qb-row');
  const orderSelect = el('select', 'wb-qb-order');
  const orderDir = el('select', 'wb-qb-orderdir');
  for (const [v, label] of [['asc', 'ascending'], ['desc', 'descending']]) {
    const opt = el('option', '', label);
    opt.value = v;
    orderDir.append(opt);
  }
  const topInput = el('input', 'wb-qb-top');
  topInput.type = 'number';
  topInput.min = '1';
  topInput.max = String(MAX_TOP);
  topInput.value = String(DEFAULT_TOP);
  const expandInput = el('input', 'wb-qb-expand');
  expandInput.type = 'text';
  expandInput.placeholder = 'extra $expand (comma-separated)';
  optionsRow.append(
    el('span', 'wb-qb-label', '$orderby'), orderSelect, orderDir,
    el('span', 'wb-qb-label', '$top'), topInput,
    el('span', 'wb-qb-label', '$expand'), expandInput,
  );

  const rawRow = el('div', 'wb-qb-rawrow');
  const rawArea = el('textarea', 'wb-qb-raw');
  rawArea.spellcheck = false;
  rawArea.setAttribute('aria-label', 'Raw query');
  const rawNote = el('span', 'wb-qb-rawnote', 'Editing the raw query overrides the builder.');
  rawNote.hidden = true;
  const runBtn = el('button', 'btn btn-xs wb-qb-run wb-primary', 'Run ▶');
  runBtn.type = 'button';
  const backToBuilder = el('button', 'btn btn-xs', 'Back to builder');
  backToBuilder.type = 'button';
  backToBuilder.hidden = true;
  rawRow.append(rawArea, rawNote, runBtn, backToBuilder);

  composer.append(targetRow, fieldsBox, filtersBox, optionsRow, rawRow);

  const results = el('div', 'wb-qb-results');
  root.append(head, composer, results);

  // --- state ---
  let lists = [];
  let fields = [];              // field entities of the picked list
  let rawMode = false;
  let loadedForWeb = '';

  const fieldByName = (name) => fields.find((f) => f.InternalName === name);
  const guidPath = (listId) => `web/lists(guid'${listId}')/items`;

  function pickedListId() {
    return listSelect.value === '::endpoint' ? '' : listSelect.value;
  }

  // --- filter rows ---
  function addFilterRow(saved = {}) {
    const row = el('div', 'wb-qb-filterrow');
    const join = el('select', 'wb-qb-join');
    for (const [v, label] of [['and', 'AND'], ['or', 'OR']]) {
      const opt = el('option', '', label);
      opt.value = v;
      join.append(opt);
    }
    join.value = saved.join || 'and';
    if (!filterRows.childElementCount) join.classList.add('wb-qb-join-first');

    const fieldSel = el('select', 'wb-qb-field');
    for (const f of fields) {
      const opt = el('option', '', f.InternalName);
      opt.value = f.InternalName;
      fieldSel.append(opt);
    }
    if (saved.field) fieldSel.value = saved.field;

    const opSel = el('select', 'wb-qb-op');
    for (const [v, label] of OPERATORS) {
      const opt = el('option', '', label);
      opt.value = v;
      opSel.append(opt);
    }
    if (saved.op) opSel.value = saved.op;

    const valueInput = el('input', 'wb-qb-value');
    valueInput.type = 'text';
    valueInput.placeholder = 'value';
    valueInput.value = saved.value || '';

    const remove = el('button', 'btn btn-xs', '×');
    remove.type = 'button';
    remove.title = 'Remove this filter';
    remove.addEventListener('click', () => {
      row.remove();
      onBuilderChange();
    });

    for (const control of [join, fieldSel, opSel]) {
      control.addEventListener('change', onBuilderChange);
    }
    valueInput.addEventListener('input', onBuilderChange);

    row.append(join, fieldSel, opSel, valueInput, remove);
    filterRows.append(row);
  }

  function readFilterRows() {
    return [...filterRows.children].map((row) => ({
      join: row.querySelector('.wb-qb-join').value,
      field: row.querySelector('.wb-qb-field').value,
      type: fieldByName(row.querySelector('.wb-qb-field').value)?.TypeAsString || 'Text',
      op: row.querySelector('.wb-qb-op').value,
      value: row.querySelector('.wb-qb-value').value,
    }));
  }

  // --- composing ---
  function selectedFields() {
    return [...fieldsList.querySelectorAll('input:checked')].map((box) => box.value);
  }

  function composeDescriptor() {
    if (rawMode) {
      const parsed = rawToDescriptor(rawArea.value);
      return parsed ? { ...parsed, webUrl: client.webUrl() } : null;
    }
    const listId = pickedListId();
    const select = selectedFields();
    const expand = new Set(
      expandInput.value.split(',').map((s) => s.trim()).filter(Boolean),
    );
    // 'A/B' selections need their root expanded.
    for (const entry of select) {
      if (entry.includes('/')) expand.add(entry.split('/')[0]);
    }
    const options = {};
    if (select.length) options.select = select;
    if (expand.size) options.expand = [...expand];
    const filter = composeFilter(readFilterRows());
    if (filter) options.filter = filter;
    if (orderSelect.value) {
      options.orderby = orderDir.value === 'desc'
        ? `${orderSelect.value} desc` : orderSelect.value;
    }
    const top = Math.min(Math.max(Number(topInput.value) || DEFAULT_TOP, 1), MAX_TOP);
    options.top = top;
    const path = listId ? guidPath(listId) : String(endpointInput.value || '').trim().replace(/^\/+/, '');
    if (!path) return null;
    return { path, options, webUrl: client.webUrl() };
  }

  function onBuilderChange() {
    if (rawMode) return;
    const descriptor = composeDescriptor();
    rawArea.value = descriptor ? descriptorToRaw(descriptor) : '';
  }

  function enterRawMode() {
    if (rawMode) return;
    rawMode = true;
    composer.classList.add('wb-qb-rawmode');
    rawNote.hidden = false;
    backToBuilder.hidden = false;
  }

  function leaveRawMode() {
    rawMode = false;
    composer.classList.remove('wb-qb-rawmode');
    rawNote.hidden = true;
    backToBuilder.hidden = true;
    onBuilderChange();
  }

  rawArea.addEventListener('input', enterRawMode);
  backToBuilder.addEventListener('click', leaveRawMode);

  // --- field checkboxes ---
  function renderFieldList(savedSelect = null) {
    fieldsList.textContent = '';
    orderSelect.textContent = '';
    const blank = el('option', '', '(no ordering)');
    blank.value = '';
    orderSelect.append(blank);
    const wanted = new Set(savedSelect || ['Id', 'Title']);
    for (const f of fields) {
      const entry = EXPANDABLE_TYPES.has(f.TypeAsString)
        ? `${f.InternalName}/Title` : f.InternalName;
      const label = el('label', 'wb-qb-fieldopt');
      const box = el('input');
      box.type = 'checkbox';
      box.value = entry;
      box.checked = wanted.has(entry);
      box.addEventListener('change', onBuilderChange);
      label.append(box, document.createTextNode(entry));
      label.append(el('span', 'wb-qb-fieldtype', f.TypeAsString));
      fieldsList.append(label);

      const opt = el('option', '', f.InternalName);
      opt.value = f.InternalName;
      orderSelect.append(opt);
    }
  }

  async function loadFieldsForList(listId, saved = null) {
    fields = [];
    fieldsList.textContent = '';
    filterRows.textContent = '';
    if (!listId) { renderFieldList(); onBuilderChange(); return; }
    fieldsList.append(el('div', 'wb-qb-loading', 'Loading fields…'));
    try {
      const { items } = await client.getAll(`web/lists(guid'${listId}')/fields`, {
        select: FIELD_SELECT,
      });
      fields = items.filter((f) => !f.Hidden);
      renderFieldList(saved?.select);
      for (const savedRow of saved?.filters || []) addFilterRow(savedRow);
      if (saved?.orderby) orderSelect.value = saved.orderby;
      if (saved?.orderdir) orderDir.value = saved.orderdir;
      onBuilderChange();
    } catch (err) {
      fieldsList.textContent = '';
      fieldsList.append(el('div', 'wb-qb-loading wb-error', err?.message || String(err)));
    }
  }

  // --- target picker ---
  function renderListPicker(savedListId = '') {
    listSelect.textContent = '';
    for (const list of lists) {
      const opt = el('option', '', list.Hidden ? `${list.Title} (hidden)` : list.Title);
      opt.value = list.Id;
      listSelect.append(opt);
    }
    const endpoint = el('option', '', '— arbitrary endpoint —');
    endpoint.value = '::endpoint';
    listSelect.append(endpoint);
    if (savedListId && lists.some((l) => l.Id === savedListId)) listSelect.value = savedListId;
    else if (savedListId === '::endpoint') listSelect.value = '::endpoint';
    endpointInput.hidden = listSelect.value !== '::endpoint';
  }

  listSelect.addEventListener('change', () => {
    endpointInput.hidden = listSelect.value !== '::endpoint';
    loadFieldsForList(pickedListId());
  });
  endpointInput.addEventListener('input', onBuilderChange);
  addFilter.addEventListener('click', () => {
    if (!fields.length) return;
    addFilterRow();
    onBuilderChange();
  });
  for (const control of [orderSelect, orderDir, topInput, expandInput]) {
    control.addEventListener('change', onBuilderChange);
    control.addEventListener('input', onBuilderChange);
  }

  // --- running ---
  let grid = null;

  async function run() {
    const descriptor = composeDescriptor();
    if (!descriptor || !descriptor.path) return;

    writeSaved(client.webUrl(), {
      listId: listSelect.value,
      endpoint: endpointInput.value,
      select: selectedFields(),
      filters: readFilterRows().map(({ join, field, op, value }) => ({ join, field, op, value })),
      orderby: orderSelect.value,
      orderdir: orderDir.value,
      top: topInput.value,
      expand: expandInput.value,
      raw: rawMode ? rawArea.value : '',
    });

    const select = descriptor.options?.select;
    results.textContent = '';
    grid = createGrid({
      columns: Array.isArray(select) && select.length
        ? columnsForSelect(select)
        : [{ key: '__json', label: 'Result', value: (row) => JSON.stringify(row), mono: true }],
      rowKey: 'Id',
      emptyText: 'The query returned no rows.',
      filterPlaceholder: 'Filter results…',
      exportName: 'sp-query',
      // Raw-mode strings that don't round-trip get no Copy-as menu — a
      // wrong script is worse than none.
      descriptor: rawMode && !rawToDescriptor(rawArea.value) ? null : descriptor,
    });
    results.append(grid.el);
    grid.setLoading('Running query…');
    try {
      const { items, partial } = await client.getAll(descriptor.path, descriptor.options);
      const rows = Array.isArray(select) && select.length
        ? items
        : items.map((item, i) => ({ Id: item?.Id ?? i, ...item }));
      grid.setRows(rows, { partial });
    } catch (err) {
      grid.setError(err);
    }
  }

  runBtn.addEventListener('click', run);

  // --- load ---
  async function load() {
    const webUrl = client.webUrl();
    if (loadedForWeb === webUrl) return;
    loadedForWeb = webUrl;
    try {
      const { items } = await client.getAll('web/lists', {
        select: ['Id', 'Title', 'Hidden', 'BaseTemplate'],
        orderby: 'Title',
        top: 5000,
      });
      lists = items;
    } catch (err) {
      lists = [];
      results.textContent = '';
      results.append(el('div', 'wb-grid-status wb-error', err?.message || String(err)));
    }
    const saved = readSaved(webUrl);
    renderListPicker(saved?.listId || '');
    endpointInput.value = saved?.endpoint || '';
    if (saved?.top) topInput.value = saved.top;
    if (saved?.expand) expandInput.value = saved.expand;
    await loadFieldsForList(pickedListId(), saved);
    if (saved?.raw) {
      rawArea.value = saved.raw;
      enterRawMode();
    }
  }

  return { el: root, load };
}
