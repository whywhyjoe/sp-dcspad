// SP-aware object inspector: recognizes the shapes SharePoint developers
// stare at all day — OData envelopes, list-item collections, SP entities,
// PnPjs results — and renders smart views on top of the generic tree.
// Returns null for anything unrecognized (caller falls back to renderValue).
//
// Standalone by design: this module is the rendering engine the future
// Site Inspector panel will reuse.

import { renderValue, renderTable, el, previewOf } from './tree-view.js';

// Plumbing keys that get dimmed in trees and dropped from table views.
const NOISE_KEYS = new Set(['__metadata', '__deferred', 'odata.metadata', 'odata.type', 'odata.id', 'odata.etag', 'odata.editLink', '@odata.context', '@odata.type', '@odata.id', '@odata.etag', '@odata.editLink', 'FirstUniqueAncestorSecurableObject', 'RoleAssignments']);

const key = (node, k) => node?.t === 'obj' ? node.keys.find(([n]) => n === k)?.[1] : undefined;
const keyNames = (node) => node?.t === 'obj' ? node.keys.map(([n]) => n) : [];
const str = (node) => (node && (node.t === 'str' || node.t === 'num')) ? String(node.v) : undefined;

export function enhance(node) {
  if (!node || (node.t !== 'obj' && node.t !== 'arr')) return null;

  // --- OData verbose envelope: { d: {...} } ---
  const d = key(node, 'd');
  if (d && node.keys.length === 1) {
    return envelope('OData verbose', d, node);
  }

  // --- OData nometadata/minimal envelope: { value: [...] } ---
  const value = key(node, 'value');
  if (value?.t === 'arr' && keyNames(node).every((k) => k === 'value' || k.startsWith('odata.') || k.startsWith('@odata.'))) {
    return envelope('OData', value, node);
  }

  // --- results collection: { results: [...] } (verbose collections) ---
  const results = key(node, 'results');
  if (results?.t === 'arr') {
    return collection(results, node);
  }

  // --- bare arrays of SP items (PnPjs unwraps to plain arrays) ---
  if (node.t === 'arr' && node.items.length && node.items.every(looksLikeSpObject)) {
    return collection(node, null);
  }

  // --- single SP entity ---
  if (looksLikeSpObject(node)) {
    return entity(node);
  }

  return null;
}

function looksLikeSpObject(node) {
  if (node?.t !== 'obj') return false;
  if (key(node, '__metadata')) return true;
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  return has('InternalName', 'TypeAsString')                  // Field
    || has('BaseTemplate', 'EntityTypeName')                  // List
    || has('ServerRelativeUrl', 'WebTemplate')                // Web
    || has('LoginName', 'PrincipalType')                      // User/Group
    || names.includes('odata.type') || names.includes('@odata.type');
}

function spType(node) {
  const meta = key(node, '__metadata');
  return str(key(meta, 'type')) || str(key(node, 'odata.type')) || str(key(node, '@odata.type')) || detectShape(node);
}

function detectShape(node) {
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  if (has('InternalName', 'TypeAsString')) return 'SP.Field';
  if (has('BaseTemplate', 'EntityTypeName')) return 'SP.List';
  if (has('ServerRelativeUrl', 'WebTemplate')) return 'SP.Web';
  if (has('LoginName', 'PrincipalType')) {
    return str(key(node, 'OwnerTitle')) !== undefined ? 'SP.Group' : 'SP.User';
  }
  return null;
}

// ---------------------------------------------------------------
// Envelope: badge + unwrapped payload, metadata folded away
// ---------------------------------------------------------------
function envelope(label, inner, outer) {
  const wrap = el('div', 'tree-node');
  const head = el('div');
  head.append(badge(label));
  wrap.append(head);

  const enhanced = enhance(inner);
  wrap.append(enhanced ?? renderValue(inner, { dimKeys: NOISE_KEYS }));

  // Fold the envelope's own metadata keys behind a dim expandable row.
  const metaKeys = outer.keys.filter(([k]) => k !== 'd' && k !== 'value');
  if (metaKeys.length) {
    const fold = el('div', 'sp-meta-fold');
    fold.append(renderValue({ t: 'obj', cls: 'envelope metadata', keys: metaKeys }, { dimKeys: NOISE_KEYS }));
    wrap.append(fold);
  }
  return wrap;
}

// ---------------------------------------------------------------
// Collection: count badge, tree/table toggle, paging warning
// ---------------------------------------------------------------
function collection(arrNode, parentNode) {
  const wrap = el('div', 'tree-node');
  const head = el('div');
  const type = arrNode.items.length ? spType(arrNode.items[0]) : null;
  head.append(badge(`${arrNode.n} item${arrNode.n === 1 ? '' : 's'}`));
  if (type) head.append(el('span', 'sp-entity-head', shortType(type)));

  const toggle = el('span', 'table-toggle', '⊞ table view');
  head.append(toggle);
  wrap.append(head);

  const treeEl = el('div');
  if (arrNode.items.length && arrNode.items.every((i) => i.t === 'obj')) {
    // Render each item through entity enhancement when applicable.
    const list = el('div');
    arrNode.items.forEach((item, i) => {
      const row = el('div');
      row.append(el('span', 'tree-key dim-key', `${i}: `));
      row.append(enhance(item) ?? renderValue(item, { dimKeys: NOISE_KEYS }));
      list.append(row);
    });
    if (arrNode.trunc) list.append(el('div', 't-truncated', `… showing first ${arrNode.items.length} of ${arrNode.n}`));
    treeEl.append(list);
  } else {
    treeEl.append(renderValue(arrNode, { dimKeys: NOISE_KEYS }));
  }

  const tableEl = el('div');
  tableEl.hidden = true;
  let tableBuilt = false;
  toggle.addEventListener('click', () => {
    const showTable = tableEl.hidden;
    if (showTable && !tableBuilt) {
      tableBuilt = true;
      tableEl.append(renderTable(filterNoise(arrNode)));
    }
    tableEl.hidden = !showTable;
    treeEl.hidden = showTable;
    toggle.textContent = showTable ? '≡ tree view' : '⊞ table view';
  });
  wrap.append(treeEl, tableEl);

  // Paging: surface __next / nextLink prominently.
  const next = str(key(parentNode, '__next')) || str(key(parentNode, 'odata.nextLink')) || str(key(parentNode, '@odata.nextLink'));
  if (next) {
    const warn = el('div', 'sp-next-link');
    warn.append(el('span', '', '⚠ partial result set — next page: '));
    warn.append(copySpan(next, next.length > 80 ? next.slice(0, 80) + '…' : next));
    wrap.append(warn);
  }
  return wrap;
}

function filterNoise(arrNode) {
  return {
    ...arrNode,
    items: arrNode.items.map((item) =>
      item.t === 'obj'
        ? { ...item, keys: item.keys.filter(([k]) => !NOISE_KEYS.has(k)) }
        : item),
  };
}

// ---------------------------------------------------------------
// Entity: compact header with the fields devs actually need
// ---------------------------------------------------------------
const ENTITY_FIELDS = {
  'SP.List': [
    ['Title', false], ['Id', true], ['EntityTypeName', true], ['BaseTemplate', false], ['ItemCount', false],
  ],
  'SP.Field': [
    ['Title', false], ['InternalName', true], ['TypeAsString', false], ['Required', false],
  ],
  'SP.Web': [
    ['Title', false], ['ServerRelativeUrl', true], ['WebTemplate', false],
  ],
  'SP.User': [
    ['Title', false], ['LoginName', true], ['Email', true],
  ],
  'SP.Group': [
    ['Title', false], ['Id', true], ['OwnerTitle', false],
  ],
  'SP.ListItem': [
    ['Title', false], ['Id', false],
  ],
};

function shortType(type) {
  // "SP.Data.TasksListItem" -> badge-friendly names
  if (!type) return '';
  if (type.startsWith('SP.Data.') && type.endsWith('Item')) return 'SP.ListItem · ' + type.slice(8);
  return type;
}

function entityKind(type) {
  if (!type) return null;
  if (ENTITY_FIELDS[type]) return type;
  for (const known of Object.keys(ENTITY_FIELDS)) {
    if (type.startsWith(known)) return known;
  }
  if (type.startsWith('SP.Data.')) return 'SP.ListItem';
  return null;
}

function entity(node) {
  const type = spType(node);
  const kind = entityKind(type);
  const wrap = el('div', 'tree-node');
  const head = el('div', 'sp-entity-head');
  head.append(badge(shortType(type) || 'SP'));

  if (kind) {
    for (const [field, copyable] of ENTITY_FIELDS[kind]) {
      const v = key(node, field);
      if (v === undefined) continue;
      const fieldEl = el('span', 'sp-field');
      fieldEl.append(el('span', 'dim-key tree-key', `${field}: `));
      const text = v.t === 'str' || v.t === 'num' || v.t === 'bool' ? String(v.v) : previewOf(v);
      fieldEl.append(copyable ? copySpan(text, text) : el('span', '', text));
      head.append(fieldEl);
    }
  }
  wrap.append(head);
  wrap.append(renderValue(node, { dimKeys: NOISE_KEYS }));
  return wrap;
}

// ---------------------------------------------------------------
// helpers
// ---------------------------------------------------------------
function badge(text) {
  return el('span', 'sp-badge', text);
}

function copySpan(copyText, displayText) {
  const s = el('span', 'sp-copy', displayText);
  s.title = 'Click to copy';
  s.addEventListener('click', async (e) => {
    e.stopPropagation();
    try {
      await navigator.clipboard.writeText(copyText);
      s.classList.add('copied');
      setTimeout(() => s.classList.remove('copied'), 800);
    } catch { /* clipboard denied — ignore */ }
  });
  return s;
}
