// Lists & libraries view: every list in the web (hidden ones included — the
// thing the SP UI won't show), with drill-down into fields, views, content
// types, and the raw entity rendered through the SP-aware inspector.

import { createGrid } from '../grid.js?v=2';
import { copyText } from '../export.js';
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
  'Id', 'Title', 'BaseTemplate', 'BaseType', 'ItemCount', 'Hidden',
  'Created', 'LastItemModifiedDate', 'EntityTypeName', 'Description',
  'DefaultViewUrl', 'RootFolder/ServerRelativeUrl',
];

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

export function createListsView({ client, navigate }) {
  const root = el('section', 'wb-view wb-view-lists');

  // ---- all-lists grid pane ----
  const gridPane = el('div', 'wb-pane');
  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Lists &amp; libraries</h2>'
    + '<p class="wb-view-hint">Every list in this web, hidden ones included. '
    + 'Click a row for fields, views, and content types.</p>';

  const grid = createGrid({
    columns: [
      { key: 'Title', label: 'Title' },
      { key: 'BaseTemplate', label: 'Template', format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? '') },
      { key: 'ItemCount', label: 'Items' },
      { key: 'Hidden', label: 'Hidden' },
      { key: 'Url', label: 'Url', value: (row) => row.RootFolder?.ServerRelativeUrl || '', mono: true, copyable: true },
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
          a.target = '_blank';
          a.rel = 'noopener';
          a.title = 'Open list settings in a new tab';
          a.textContent = '⚙';
          a.addEventListener('click', (e) => e.stopPropagation());
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
  gridPane.append(head, grid.el);

  // ---- detail pane (rebuilt per list) ----
  const detailPane = el('div', 'wb-pane');
  detailPane.hidden = true;

  root.append(gridPane, detailPane);

  let listsLoaded = false;
  const tabCache = new Map();   // `${listId}::${tab}` -> Promise<rows|json>

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
      grid.setRows(items, { partial });
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
          { key: 'ServerRelativeUrl', label: 'Url', mono: true, copyable: true },
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
    settingsLink.target = '_blank';
    settingsLink.rel = 'noopener';
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
