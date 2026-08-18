// Lists & libraries view: every list in the web (hidden ones included — the
// thing the SP UI won't show), with drill-down into fields, views, content
// types, and the raw entity rendered through the SP-aware inspector.

import { createGrid, encodeSpPath, bindNewTab } from '../grid.js?v=2';
import { copyText, downloadMarkdown } from '../export.js';
import {
  buildItemsMarkdown, contentFields, viewColumnFields,
  fieldText, itemTitle, personText,
} from '../item-export.js';
import { LIST_SETTINGS, linkUrl } from '../config-links.js';
import { principalTypeName } from '../perm-kinds.js';
import { enhance } from '../../inspect/sp-shapes.js';
import { renderValue } from '../../inspect/tree-view.js';
import { toNode } from '../../inspect/to-node.js';

// Friendly names for the templates that actually show up in day-to-day work;
// anything else renders as its number.
export const BASE_TEMPLATE_NAMES = {
  100: 'Generic list',
  101: 'Document library',
  102: 'Survey',
  103: 'Links',
  104: 'Announcements',
  105: 'Contacts',
  106: 'Events',
  107: 'Tasks (classic)',
  108: 'Discussion board',
  109: 'Picture library',
  110: 'Data sources',
  112: 'User information',
  116: 'Master page gallery',
  119: 'Site pages',
  120: 'Custom grid',
  140: 'Workflow history',
  160: 'Access requests',
  171: 'Tasks',
  850: 'Publishing pages',
};

const LIST_SELECT = [
  'Id', 'Title', 'BaseTemplate', 'BaseType', 'ItemCount', 'Hidden', 'IsCatalog',
  'Created', 'LastItemModifiedDate', 'EntityTypeName', 'Description',
  'DefaultViewUrl', 'RootFolder/ServerRelativeUrl',
];

// Deep-internal plumbing SharePoint keeps for itself, collapsed behind the
// grid's bottom expander. Detection layers, strongest first:
//   1. Hidden        — SharePoint marks nearly all system lists hidden
//                      (appdata, TaxonomyHiddenList, Sharing Links, …).
//   2. IsCatalog     — SharePoint's own flag on gallery lists; locale-proof.
//   3. /_catalogs, /FormServerTemplates paths + gallery/social templates —
//                      belt and braces for older shapes (112 user info,
//                      113 web part, 114 list template, 116 master page,
//                      121 solutions, 122 no-code public, 123 theme,
//                      124 design, 125 app data, 175 maintenance logs,
//                      544 MicroFeed).
//   4. INTERNAL_TITLES — visible classic-publishing infrastructure no flag
//                      catches. English titles only; the mechanical signals
//                      above carry non-English sites. Curated — trim or grow
//                      deliberately, one line per decision.
// Deliberately NOT internal: Workflow Tasks (real user tasks live there),
// XML form libraries (115 — users create InfoPath libraries as content),
// publishing Pages (850) and Images (851).
export const INTERNAL_TEMPLATES = new Set([112, 113, 114, 116, 121, 122, 123, 124, 125, 175, 544]);
export const INTERNAL_TITLES = new Set([
  'TaxonomyHiddenList',
  'Style Library',
  'Form Templates',
  'Cache Profiles',
  'Device Channels',
  'Quick Deploy Items',
  'Reusable Content',
  'Content and Structure Reports',
  'Site Collection Documents',
  'Site Collection Images',
  'Suggested Content Browser Locations',
]);
export function isInternalList(list) {
  const path = String(list?.RootFolder?.ServerRelativeUrl || '').toLowerCase();
  return Boolean(list?.Hidden)
    || Boolean(list?.IsCatalog)
    || INTERNAL_TEMPLATES.has(list?.BaseTemplate)
    || path.includes('/_catalogs')
    || path.endsWith('/formservertemplates')
    || INTERNAL_TITLES.has(list?.Title);
}

const FIELD_SELECT = [
  'Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'Required',
  'Hidden', 'ReadOnlyField', 'Group', 'DefaultValue', 'Choices', 'Description',
  'EnforceUniqueValues', 'Indexed',
];

const VIEW_SELECT = [
  'Id', 'Title', 'DefaultView', 'PersonalView', 'Hidden', 'ServerRelativeUrl',
  'RowLimit', 'Paged', 'ViewQuery',
];

const CT_SELECT = ['Id', 'Name', 'Group', 'Hidden', 'ReadOnly', 'Sealed', 'Description'];

const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');
const choicesText = (v) => {
  const arr = Array.isArray(v) ? v : v?.results;
  return Array.isArray(arr) ? arr.join(' | ') : '';
};
const fileStem = (s) =>
  String(s || 'list').toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '') || 'list';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;

// Raw OData input for the Items tab → { filter, orderby, error }. Only the
// two clauses the tab doesn't already own are accepted; anything else
// $-prefixed is rejected so a typo'd $select/$top can't silently fight the
// controls. Splitting respects OData single-quoted literals (with doubled
// apostrophes), so a filter like Title eq 'R&D' survives intact.
export function parseItemsQuery(text) {
  const out = { filter: '', orderby: '', error: '' };
  const raw = String(text || '').trim().replace(/^\?/, '');
  if (!raw) return out;

  const parts = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] === "'") {
      if (quoted && raw[i + 1] === "'") { i++; continue; }   // '' = escaped '
      quoted = !quoted;
    } else if (raw[i] === '&' && !quoted) {
      parts.push(raw.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(raw.slice(start));

  const seen = new Set();
  for (const part of parts) {
    const piece = part.trim();
    if (!piece) continue;
    const clause = /^\$?(filter|orderby)\s*=\s*(.+)$/i.exec(piece);
    if (clause) {
      const name = clause[1].toLowerCase();
      if (seen.has(name)) {
        out.error = `Duplicate $${name} clause — combine them into one.`;
        return out;
      }
      seen.add(name);
      out[name] = clause[2].trim();
    } else if (piece.startsWith('$')) {
      out.error = 'Only $filter and $orderby are supported here — columns and max items have their own controls.';
      return out;
    } else {
      // A bare expression is the common case: treat it as the filter.
      out.filter = out.filter ? `${out.filter} and ${piece}` : piece;
    }
  }
  return out;
}

export function createListsView({ client, navigate }) {
  const root = el('section', 'wb-view wb-view-lists');

  // Server-relative paths become real links: absolute against the inspected
  // web's origin, opened in a new tab (the cell text stays the relative path).
  const webOrigin = () => {
    try { return new URL(client.webUrl()).origin; } catch { return ''; }
  };
  const absUrl = (rel) => (rel ? `${webOrigin()}${encodeSpPath(rel)}` : '');

  // ---- all-lists grid pane ----
  const gridPane = el('div', 'wb-pane');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Lists &amp; libraries</h2>'
    + '<p class="wb-view-hint">Every list in this web — SharePoint-internal '
    + 'plumbing sits behind the expander below the grid. Click a row for '
    + 'fields, views, content types, and items.</p>';

  const grid = createGrid({
    columns: [
      { key: 'Title', label: 'Title' },
      { key: 'BaseTemplate', label: 'Template', format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? '') },
      { key: 'ItemCount', label: 'Items' },
      { key: 'Hidden', label: 'Hidden' },
      { key: 'Url', label: 'Url', value: (row) => row.RootFolder?.ServerRelativeUrl || '', mono: true, copyable: true, link: absUrl },
      { key: 'Id', label: 'Id', mono: true, copyable: true },
      { key: 'LastItemModifiedDate', label: 'Modified', format: fmtDate },
      // Appended last: tests address earlier columns positionally.
      {
        key: 'Settings',
        label: '',
        value: (row) => row.Id,
        format: () => '',   // keep filter/sort/export free of the glyph
        render: (id) => {
          const a = document.createElement('a');
          a.className = 'wb-cell-link';
          a.href = linkUrl(client.webUrl(), LIST_SETTINGS, { guid: id });
          a.title = 'Open list settings in a new tab';
          a.textContent = '⚙';
          bindNewTab(a);
          return a;
        },
      },
    ],
    onOpen: (row) => navigate({ view: 'lists', listId: row.Id, listTitle: row.Title }),
    emptyText: 'No lists in this web.',
    filterPlaceholder: 'Filter lists…',
    exportName: 'sp-lists',
    descriptor: {
      path: 'web/lists',
      options: { select: LIST_SELECT, expand: 'RootFolder', orderby: 'Title', top: 5000 },
      webUrl: client.webUrl(),
    },
  });
  const moreBtn = el('button', 'btn btn-xs wb-lists-more');
  moreBtn.type = 'button';
  moreBtn.hidden = true;
  gridPane.append(head, grid.el, moreBtn);

  // ---- detail pane (rebuilt per list) ----
  const detailPane = el('div', 'wb-pane');
  detailPane.hidden = true;

  root.append(gridPane, detailPane);

  let listsLoaded = false;
  let allLists = [];
  let listsPartial = false;
  let showInternal = false;
  const tabCache = new Map();   // `${listId}::${tab}` -> Promise<rows|json>

  function renderLists() {
    const internal = allLists.filter(isInternalList);
    grid.setRows(
      showInternal ? allLists : allLists.filter((l) => !isInternalList(l)),
      { partial: listsPartial },
    );
    moreBtn.hidden = internal.length === 0;
    const label = `${internal.length} internal list${internal.length === 1 ? '' : 's'}`;
    moreBtn.textContent = showInternal ? `Hide ${label} ▴` : `Show ${label} ▾`;
    moreBtn.title = 'SharePoint-internal plumbing: hidden lists, galleries, the taxonomy cache';
  }
  moreBtn.addEventListener('click', () => {
    showInternal = !showInternal;
    renderLists();
  });

  async function loadLists() {
    if (listsLoaded) return;
    grid.setLoading('Loading lists…');
    try {
      const { items, partial } = await client.getAll('web/lists', {
        select: LIST_SELECT,
        expand: 'RootFolder',
        orderby: 'Title',
        top: 5000,
      });
      allLists = items;
      listsPartial = partial;
      renderLists();
      listsLoaded = true;
    } catch (err) {
      grid.setError(err);
    }
  }

  function cached(listId, tab, fetcher) {
    const key = `${listId}::${tab}`;
    if (!tabCache.has(key)) {
      tabCache.set(key, fetcher().catch((err) => {
        tabCache.delete(key);   // allow retry after a failure
        throw err;
      }));
    }
    return tabCache.get(key);
  }

  // Items tab: list-content markdown export. A view supplies the content
  // columns when chosen; otherwise every content field exports, ordered
  // Title, ID, content columns as returned, then the Created/Modified pairs.
  // Ordering is ID desc, capped by the max-items input. The grid is a
  // preview of the same rows — filter it and the export follows.
  function buildItemsPane(wrap, listId, listTitle) {
    // Every control shares the grid's own single toolbar row (count · filter
    // · controls · actions) so the tab reads like every other grid, with the
    // filter box in its standard slot. The controls node is persistent and
    // adopted by each rebuilt grid via toolbarExtras.
    const controls = el('span', 'wb-items-controls');

    const viewLabel = el('label', 'wb-items-label', 'Columns ');
    const viewSel = el('select', 'wb-items-view');
    viewSel.setAttribute('aria-label', 'Column source: a view or all columns');
    viewLabel.append(viewSel);

    const maxLabel = el('label', 'wb-items-label', 'Max ');
    const maxIn = el('input', 'wb-items-max');
    maxIn.type = 'number';
    maxIn.min = '1';
    maxIn.max = '5000';
    maxIn.value = '500';
    maxIn.setAttribute('aria-label', 'Maximum items to fetch');
    maxLabel.append(maxIn);

    // Advanced, so collapsed by default: a quiet toggle that expands the raw
    // OData input in place. While a query is applied, the toggle stays
    // accent-marked — a hidden active query must never silently shape rows.
    const queryToggle = el('button', 'wb-items-querytoggle', 'Query ▾');
    queryToggle.type = 'button';
    const QUERY_HINT = 'Advanced: raw OData clauses for the item query — '
      + '$filter=… and/or $orderby=…, joined with &. A bare expression '
      + 'counts as $filter. Order defaults to ID desc.';
    queryToggle.title = QUERY_HINT;
    queryToggle.setAttribute('aria-expanded', 'false');

    const queryWrap = el('span', 'wb-items-querywrap');
    queryWrap.hidden = true;
    const queryIn = el('input', 'wb-items-query');
    queryIn.type = 'text';
    queryIn.placeholder = "$filter=Status eq 'Active'&$orderby=DueDate desc";
    queryIn.setAttribute('aria-label', 'OData $filter and $orderby for the item query');
    queryIn.title = QUERY_HINT;
    const queryErr = el('span', 'wb-items-error');
    queryWrap.append(queryIn, queryErr);

    controls.append(viewLabel, maxLabel, queryToggle, queryWrap);

    const gridBox = el('div', 'wb-items-grid');
    wrap.append(gridBox);

    function setQueryOpen(open) {
      queryWrap.hidden = !open;
      queryToggle.textContent = open ? 'Query ▴' : 'Query ▾';
      queryToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) queryIn.focus();
    }
    queryToggle.addEventListener('click', () => {
      const opening = queryWrap.hidden;
      if (!opening && queryErr.textContent) {
        // Collapsing an erroneous (never-applied) query discards it, so no
        // error can hide behind a closed toggle.
        queryIn.value = '';
        queryErr.textContent = '';
        reload();
      }
      setQueryOpen(opening);
    });

    let itemsGrid = null;
    let current = null;   // { fields, viewFieldNames, viewTitle } of the loaded rows
    let viewsFilled = false;
    let loadSeq = 0;
    // Rows are cached by the parameters that shape the *data* — a view
    // change only re-picks columns, so it never refetches items.
    let itemsCache = { key: '', items: null, partial: false };

    const clampMax = () => {
      const n = Math.floor(Number(maxIn.value));
      const max = Number.isFinite(n) ? Math.min(Math.max(n, 1), 5000) : 500;
      maxIn.value = String(max);
      return max;
    };

    const exportDoc = () => (itemsGrid && current
      ? buildItemsMarkdown({
        listTitle,
        webUrl: client.webUrl(),
        viewTitle: current.viewTitle,
        items: itemsGrid.getExportRows(),   // selected rows when any, else visible
        fields: current.fields,
        viewFieldNames: current.viewFieldNames,
        filter: current.filter,
        orderby: current.orderby,
      })
      : '');

    function markApplied() {
      const applied = Boolean(current && (current.filter || current.orderby));
      queryToggle.classList.toggle('wb-applied', applied);
      queryToggle.title = applied ? `Active query: ${queryIn.value.trim()}` : QUERY_HINT;
    }

    async function reload() {
      // Claim the sequence before validating: an invalid input must also
      // orphan any in-flight load, or its rows would land under the error.
      const seq = ++loadSeq;
      const parsed = parseItemsQuery(queryIn.value);
      queryErr.textContent = parsed.error;
      if (parsed.error) return;   // keep whatever is loaded until the input is fixed
      const max = clampMax();
      // Keep the current grid (and the controls riding its toolbar) in place
      // while fetching; the first load shows a bare status instead.
      let status = null;
      if (itemsGrid) {
        itemsGrid.setLoading('Loading items…');
      } else {
        gridBox.textContent = '';
        status = el('div', 'wb-grid-status', 'Loading items…');
        gridBox.append(status);
      }
      // Hoisted so the catch can still build the grid shell (and mount the
      // controls) when only the ITEMS request fails — a first-load failure
      // must never leave the user without the view/max/query line to fix it.
      let fields = null;
      let viewFieldNames = null;
      let query = null;
      try {
        const [fieldsResult, { items: views }] = await Promise.all([
          cached(listId, 'fields', () =>
            client.getAll(guidPath(listId, '/fields'), { select: FIELD_SELECT })),
          cached(listId, 'views', () =>
            client.getAll(guidPath(listId, '/views'), { select: VIEW_SELECT })),
        ]);
        fields = fieldsResult.items;
        if (!viewsFilled) {
          viewsFilled = true;
          const none = el('option', '', 'All columns');
          none.value = '';
          viewSel.append(none);
          for (const view of views.filter((v) => !v.PersonalView)) {
            const opt = el('option', '', view.DefaultView ? `${view.Title} (default)` : view.Title);
            opt.value = view.Id;
            viewSel.append(opt);
          }
        }
        let viewTitle = '';
        if (viewSel.value) {
          const vf = await cached(listId, `viewfields::${viewSel.value}`, () =>
            client.get(guidPath(listId, `/views(guid'${viewSel.value}')/viewfields`)));
          viewFieldNames = vf?.Items?.results || vf?.Items || [];
          viewTitle = views.find((v) => v.Id === viewSel.value)?.Title || '';
        }

        const hasAttachments = fields.some((f) => f.TypeAsString === 'Attachments');
        const expand = ['FieldValuesAsText', ...(hasAttachments ? ['AttachmentFiles'] : [])];
        // $select=* deliberately — do NOT "optimize" this into an explicit
        // field projection. User/Lookup fields 400 when selected by bare
        // internal name ("The query to field 'X' is not valid… $expand must
        // contain X"); '*' returns every scalar plus lookup ids, and
        // FieldValuesAsText carries the display text. Cost one live outage.
        const query = {
          path: guidPath(listId, '/items'),
          options: {
            select: ['*', ...expand],
            expand,
            ...(parsed.filter ? { filter: parsed.filter } : {}),
            orderby: parsed.orderby || 'ID desc',
            top: max,
          },
        };
        const dataKey = JSON.stringify([max, parsed.filter, parsed.orderby]);
        let items;
        let partial;
        if (itemsCache.items && itemsCache.key === dataKey) {
          ({ items, partial } = itemsCache);
        } else {
          ({ items, partial } = await client.getAll(query.path, query.options, { cap: max }));
          if (seq !== loadSeq) return;
          // Default ordering gets a defensive client-side sort; a typed
          // $orderby is the server's to honor (re-sorting would undo it).
          if (!parsed.orderby) {
            items.sort((a, b) => (Number(b.ID ?? b.Id) || 0) - (Number(a.ID ?? a.Id) || 0));
          }
          itemsCache = { key: dataKey, items, partial };
        }
        if (seq !== loadSeq) return;

        mountItemsGrid(buildColumns(fields, viewFieldNames, items), query);
        itemsGrid.setRows(items, { partial });
        current = {
          fields, viewFieldNames, viewTitle,
          filter: parsed.filter, orderby: parsed.orderby,
        };
        markApplied();
      } catch (err) {
        if (seq !== loadSeq) return;
        const raw = err?.message || String(err);
        // Large-list throttling reads as a cryptic server error — translate.
        const message = /SPQueryThrottledException|list view threshold/i.test(raw)
          ? `SharePoint throttled this query — filter/order by an indexed column and keep the matched set under 5,000. (${raw})`
          : raw;
        if (!itemsGrid && fields) {
          // The items request failed but the metadata is here: mount the
          // grid shell anyway so the controls line exists and the user can
          // adjust the view/max/query and retry.
          mountItemsGrid(buildColumns(fields, viewFieldNames, []), query);
        }
        if (itemsGrid) {
          itemsGrid.setError({ message });
        } else if (status) {
          status.textContent = message;
          status.classList.add('wb-error');
        }
      }
    }

    function buildColumns(fields, viewFieldNames, items) {
      const content = viewFieldNames
        ? viewColumnFields(fields, viewFieldNames)
        : contentFields(fields);
      const filesOf = (row) => (Array.isArray(row.AttachmentFiles)
        ? row.AttachmentFiles : row.AttachmentFiles?.results || []);
      const anyAttachments = items.some((row) => filesOf(row).length);
      return [
        { key: 'Title', label: 'Title', value: (row) => itemTitle(row) },
        { key: 'ID', label: 'ID', num: true, value: (row) => row.ID ?? row.Id },
        ...content.map((f) => ({
          key: f.InternalName,
          label: f.Title || f.InternalName,
          value: (row) => fieldText(row, f),
        })),
        ...(anyAttachments ? [{
          key: 'Attachments',
          label: 'Attachments',
          value: (row) => filesOf(row).map((f) => f?.FileName || '').filter(Boolean).join(', '),
        }] : []),
        { key: 'Created', label: 'Created', format: fmtDate },
        { key: 'CreatedBy', label: 'Created By', value: (row) => personText(row, 'Author') },
        { key: 'Modified', label: 'Modified', format: fmtDate },
        { key: 'ModifiedBy', label: 'Modified By', value: (row) => personText(row, 'Editor') },
      ];
    }

    function mountItemsGrid(columns, query) {
      const newGrid = createGrid({
        columns,
        rowKey: 'ID',
        emptyText: 'No items in this list.',
        filterPlaceholder: 'Filter items…',
        exportName: `items-${fileStem(listTitle)}`,
        descriptor: query ? { ...query, webUrl: client.webUrl() } : null,
        selectable: true,
        toolbarExtras: controls,
        exportExtras: [
          ['Download .md', () => {
            const md = exportDoc();
            if (md) downloadMarkdown(`items-${fileStem(listTitle)}`, md);
          }],
          ['Copy .md', (btn) => {
            const md = exportDoc();
            if (md) copyText(md, btn);
          }],
        ],
      });
      // Swapping grids re-parents the controls; don't drop the user's focus.
      const focused = controls.contains(document.activeElement) ? document.activeElement : null;
      gridBox.textContent = '';
      gridBox.append(newGrid.el);
      itemsGrid = newGrid;
      if (focused) focused.focus();
    }

    viewSel.addEventListener('change', reload);
    maxIn.addEventListener('change', reload);
    queryIn.addEventListener('change', reload);
    queryIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') reload(); });
    reload();
  }

  const TABS = [
    {
      id: 'fields',
      label: 'Fields',
      grid: (listId, title) => ({
        columns: [
          { key: 'Title', label: 'Title' },
          { key: 'InternalName', label: 'Internal name', mono: true, copyable: true },
          { key: 'TypeAsString', label: 'Type' },
          { key: 'Required', label: 'Required' },
          { key: 'Hidden', label: 'Hidden' },
          { key: 'ReadOnlyField', label: 'Read-only' },
          { key: 'Choices', label: 'Choices', format: choicesText },
          { key: 'DefaultValue', label: 'Default' },
          { key: 'Group', label: 'Group' },
        ],
        exportName: `fields-${fileStem(title)}`,
        query: { path: guidPath(listId, '/fields'), options: { select: FIELD_SELECT } },
      }),
    },
    {
      id: 'views',
      label: 'Views',
      grid: (listId, title) => ({
        columns: [
          { key: 'Title', label: 'Title' },
          { key: 'DefaultView', label: 'Default' },
          { key: 'Hidden', label: 'Hidden' },
          { key: 'PersonalView', label: 'Personal' },
          { key: 'RowLimit', label: 'Row limit' },
          { key: 'ServerRelativeUrl', label: 'Url', mono: true, copyable: true, link: absUrl },
          { key: 'ViewQuery', label: 'CAML query', mono: true, copyable: true },
        ],
        exportName: `views-${fileStem(title)}`,
        query: { path: guidPath(listId, '/views'), options: { select: VIEW_SELECT } },
      }),
    },
    {
      id: 'contenttypes',
      label: 'Content types',
      grid: (listId, title) => ({
        columns: [
          { key: 'Name', label: 'Name' },
          { key: 'Id', label: 'Id', value: (row) => row.Id?.StringValue || String(row.Id ?? ''), mono: true, copyable: true },
          { key: 'Group', label: 'Group' },
          { key: 'Hidden', label: 'Hidden' },
          { key: 'ReadOnly', label: 'Read-only' },
          { key: 'Sealed', label: 'Sealed' },
          { key: 'Description', label: 'Description' },
        ],
        exportName: `contenttypes-${fileStem(title)}`,
        query: { path: guidPath(listId, '/contenttypes'), options: { select: CT_SELECT } },
      }),
    },
    {
      id: 'permissions',
      label: 'Permissions',
      grid: (listId, title) => ({
        columns: [
          { key: 'Member', label: 'Principal', value: (row) => row.Member?.Title || '' },
          { key: 'LoginName', label: 'Login', value: (row) => row.Member?.LoginName || '', mono: true, copyable: true },
          { key: 'PrincipalType', label: 'Type', value: (row) => row.Member?.PrincipalType, format: principalTypeName },
          {
            key: 'Roles',
            label: 'Roles',
            value: (row) => (row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || [])
              .map((r) => r.Name).filter(Boolean).join(', '),
          },
        ],
        exportName: `permissions-${fileStem(title)}`,
        query: {
          path: guidPath(listId, '/roleassignments'),
          options: {
            expand: ['Member', 'RoleDefinitionBindings'],
            select: [
              'PrincipalId', 'Member/Id', 'Member/Title', 'Member/LoginName',
              'Member/PrincipalType', 'RoleDefinitionBindings/Id', 'RoleDefinitionBindings/Name',
            ],
          },
        },
      }),
    },
    { id: 'items', label: 'Items' },
    { id: 'raw', label: 'Raw' },
  ];

  function showDetail(route) {
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = '';

    const listId = route.listId;
    const back = el('button', 'btn btn-xs wb-back', '← All lists');
    back.type = 'button';
    back.addEventListener('click', () => navigate({ view: 'lists' }));

    const title = el('h2', '', route.listTitle || 'List');
    const sub = el('span', 'wb-detail-id sp-copy', listId);
    sub.title = 'Click to copy the list id';
    sub.addEventListener('click', () => copyText(listId, sub));

    const settingsLink = el('a', 'btn btn-xs wb-detail-settings', 'List settings ↗');
    settingsLink.href = linkUrl(client.webUrl(), LIST_SETTINGS, { guid: listId });
    bindNewTab(settingsLink);
    settingsLink.title = 'Open this list’s settings page in a new tab';

    const headRow = el('div', 'wb-detail-head');
    headRow.append(back, title, sub, settingsLink);

    const tabsBar = el('div', 'wb-tabs');
    tabsBar.setAttribute('role', 'tablist');
    const body = el('div', 'wb-tab-body');

    const panes = new Map();
    let activeTab = null;

    function activate(tab) {
      activeTab = tab.id;
      for (const btn of tabsBar.children) {
        btn.classList.toggle('active', btn.dataset.tab === tab.id);
        btn.setAttribute('aria-selected', btn.dataset.tab === tab.id ? 'true' : 'false');
      }
      body.textContent = '';
      body.append(pane(tab));
    }

    function pane(tab) {
      if (panes.has(tab.id)) return panes.get(tab.id);
      const wrap = el('div', 'wb-tab-pane');
      panes.set(tab.id, wrap);

      if (tab.id === 'items') {
        buildItemsPane(wrap, listId, route.listTitle || 'List');
        return wrap;
      }

      if (tab.id === 'raw') {
        const status = el('div', 'wb-grid-status', 'Loading raw list entity…');
        wrap.append(status);
        cached(listId, 'raw', () => client.get(guidPath(listId)))
          .then((json) => {
            status.remove();
            const node = toNode(json, 0, { maxDepth: 8, maxItems: 250 });
            const inspector = el('div', 'wb-raw');
            inspector.append(enhance(node) ?? renderValue(node));
            wrap.append(inspector);
          })
          .catch((err) => {
            status.textContent = err?.message || String(err);
            status.classList.add('wb-error');
          });
        return wrap;
      }

      const spec = tab.grid(listId, route.listTitle);
      const tabGrid = createGrid({
        columns: spec.columns,
        emptyText: 'Nothing here.',
        filterPlaceholder: `Filter ${tab.label.toLowerCase()}…`,
        exportName: spec.exportName,
        descriptor: { ...spec.query, webUrl: client.webUrl() },
      });
      wrap.append(tabGrid.el);
      tabGrid.setLoading(`Loading ${tab.label.toLowerCase()}…`);
      cached(listId, tab.id, () => client.getAll(spec.query.path, spec.query.options))
        .then(({ items, partial }) => tabGrid.setRows(items, { partial }))
        .catch((err) => tabGrid.setError(err));
      return wrap;
    }

    for (const tab of TABS) {
      const btn = el('button', 'wb-tab', tab.label);
      btn.type = 'button';
      btn.dataset.tab = tab.id;
      btn.setAttribute('role', 'tab');
      btn.addEventListener('click', () => activate(tab));
      tabsBar.append(btn);
    }

    detailPane.append(headRow, tabsBar, body);
    activate(TABS.find((t) => t.id === route.tab) || TABS[0]);
    void activeTab;
  }

  function load(route) {
    if (route?.listId) {
      showDetail(route);
    } else {
      detailPane.hidden = true;
      gridPane.hidden = false;
      loadLists();
    }
  }

  return { el: root, load, grid };
}
