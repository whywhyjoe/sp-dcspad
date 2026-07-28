// Site overview: web + site property sheets, activated features (both
// scopes), subwebs, the property bag (OData-encoded keys decoded for
// display), regional settings, and the current user.

import { createGrid } from '../grid.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

// SharePoint property-bag keys arrive OData-encoded (`_x005f_` for `_`,
// `_x0020_` for space, and friends). Decode for display; copy the raw key.
const decodeODataKey = (key) =>
  String(key).replace(/_x([0-9a-f]{4})_/gi, (match, hex) =>
    String.fromCharCode(parseInt(hex, 16)));

const flatten = (v) => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
};

const entityRows = (entity) =>
  Object.entries(entity || {})
    .filter(([k]) => !k.startsWith('odata.') && !k.startsWith('@odata') && k !== '__metadata')
    .map(([k, v]) => ({ Property: k, Value: flatten(v) }));

const WEB_SELECT = [
  'Id', 'Title', 'Description', 'Url', 'ServerRelativeUrl', 'WebTemplate',
  'Configuration', 'Created', 'LastItemModifiedDate', 'Language', 'UIVersion',
  'QuickLaunchEnabled', 'MembersCanShare',
];
const SITE_SELECT = ['Id', 'Url', 'ServerRelativeUrl', 'ReadOnly', 'ShareByEmailEnabled'];

export function createSiteView({ client }) {
  const root = el('section', 'wb-view wb-view-site');

  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Site overview</h2>'
    + '<p class="wb-view-hint">Web and site collection properties, features, '
    + 'subwebs, and the property bag.</p>';

  const tabsBar = el('div', 'wb-tabs');
  const body = el('div', 'wb-tab-body');
  root.append(head, tabsBar, body);

  const panes = new Map();

  function sheetPane(query, extraSections = []) {
    const wrap = el('div', 'wb-tab-pane');
    const grid = createGrid({
      rowKey: 'Property',
      columns: [
        { key: 'Property', label: 'Property', mono: true, copyable: true },
        { key: 'Value', label: 'Value', copyable: true },
      ],
      emptyText: 'Nothing returned.',
      filterPlaceholder: 'Filter properties…',
      exportName: query.exportName,
      descriptor: { path: query.path, options: query.options, webUrl: client.webUrl() },
    });
    wrap.append(grid.el);
    grid.setLoading('Loading…');
    client.get(query.path, query.options)
      .then((entity) => grid.setRows(entityRows(entity)))
      .catch((err) => grid.setError(err));

    for (const extra of extraSections) {
      const box = el('div', 'wb-subpanel');
      box.hidden = false;
      box.append(el('h3', 'wb-subpanel-title', extra.title));
      const hostEl = el('div', 'wb-subpanel-body');
      box.append(hostEl);
      const extraGrid = createGrid({
        rowKey: 'Property',
        columns: [
          { key: 'Property', label: 'Property', mono: true, copyable: true },
          { key: 'Value', label: 'Value', copyable: true },
        ],
        emptyText: 'Nothing returned.',
        filterPlaceholder: 'Filter…',
        exportName: extra.exportName,
      });
      hostEl.append(extraGrid.el);
      extraGrid.setLoading('Loading…');
      client.get(extra.path, extra.options)
        .then((entity) => extraGrid.setRows(entityRows(extra.map ? extra.map(entity) : entity)))
        .catch((err) => extraGrid.setError(err));
      wrap.append(box);
    }
    return wrap;
  }

  function featuresPane() {
    const wrap = el('div', 'wb-tab-pane');
    const grid = createGrid({
      rowKey: 'DefinitionId',
      columns: [
        { key: 'Scope', label: 'Scope' },
        { key: 'DisplayName', label: 'Feature', format: (v) => v || '(no display name)' },
        { key: 'DefinitionId', label: 'Definition id', mono: true, copyable: true },
      ],
      emptyText: 'No activated features.',
      filterPlaceholder: 'Filter features…',
      exportName: 'sp-features',
      descriptor: {
        path: 'web/features',
        options: { select: ['DefinitionId', 'DisplayName'] },
        webUrl: client.webUrl(),
      },
    });
    wrap.append(grid.el);
    grid.setLoading('Loading features (site + web scope)…');
    const options = { select: ['DefinitionId', 'DisplayName'] };
    Promise.all([
      client.getAll('site/features', options),
      client.getAll('web/features', options),
    ])
      .then(([site, web]) => {
        const rows = [
          ...site.items.map((f) => ({ ...f, Scope: 'Site' })),
          ...web.items.map((f) => ({ ...f, Scope: 'Web' })),
        ];
        grid.setRows(rows, { partial: site.partial || web.partial });
      })
      .catch((err) => grid.setError(err));
    return wrap;
  }

  function subwebsPane() {
    const wrap = el('div', 'wb-tab-pane');
    const query = {
      path: 'web/webs',
      options: { select: ['Id', 'Title', 'ServerRelativeUrl', 'WebTemplate', 'Created', 'Language'] },
    };
    const grid = createGrid({
      columns: [
        { key: 'Title', label: 'Title' },
        { key: 'ServerRelativeUrl', label: 'Url', mono: true, copyable: true },
        { key: 'WebTemplate', label: 'Template' },
        { key: 'Language', label: 'Language' },
        { key: 'Created', label: 'Created', format: (v) => (v ? String(v).slice(0, 10) : '') },
        { key: 'Id', label: 'Id', mono: true, copyable: true },
      ],
      emptyText: 'No subwebs.',
      filterPlaceholder: 'Filter subwebs…',
      exportName: 'sp-subwebs',
      descriptor: { ...query, webUrl: client.webUrl() },
    });
    wrap.append(grid.el);
    grid.setLoading('Loading subwebs…');
    client.getAll(query.path, query.options)
      .then(({ items, partial }) => grid.setRows(items, { partial }))
      .catch((err) => grid.setError(err));
    return wrap;
  }

  function propertyBagPane() {
    const wrap = el('div', 'wb-tab-pane');
    const grid = createGrid({
      rowKey: 'RawKey',
      columns: [
        { key: 'Key', label: 'Key (decoded)', mono: true },
        { key: 'RawKey', label: 'Raw key', mono: true, copyable: true },
        { key: 'Value', label: 'Value', copyable: true },
      ],
      emptyText: 'Empty property bag.',
      filterPlaceholder: 'Filter keys…',
      exportName: 'sp-propertybag',
      descriptor: { path: 'web/allproperties', options: {}, webUrl: client.webUrl() },
    });
    wrap.append(grid.el);
    grid.setLoading('Loading property bag…');
    client.get('web/allproperties')
      .then((bag) => {
        const rows = Object.entries(bag || {})
          .filter(([k]) => !k.startsWith('odata.') && !k.startsWith('@odata') && k !== '__metadata')
          .map(([k, v]) => ({ Key: decodeODataKey(k), RawKey: k, Value: flatten(v) }));
        grid.setRows(rows);
      })
      .catch((err) => grid.setError(err));
    return wrap;
  }

  const TABS = [
    {
      id: 'web',
      label: 'Web',
      build: () => sheetPane(
        { path: 'web', options: { select: WEB_SELECT }, exportName: 'sp-web' },
        [
          {
            title: 'Regional settings',
            path: 'web/regionalsettings',
            options: { expand: 'TimeZone' },
            exportName: 'sp-regionalsettings',
          },
          {
            title: 'Current user',
            path: 'web/currentuser',
            options: {},
            exportName: 'sp-currentuser',
          },
        ],
      ),
    },
    {
      id: 'site',
      label: 'Site collection',
      build: () => sheetPane({ path: 'site', options: { select: SITE_SELECT }, exportName: 'sp-site' }),
    },
    { id: 'features', label: 'Features', build: featuresPane },
    { id: 'subwebs', label: 'Subwebs', build: subwebsPane },
    { id: 'propertybag', label: 'Property bag', build: propertyBagPane },
  ];

  function activate(tab) {
    for (const btn of tabsBar.children) {
      btn.classList.toggle('active', btn.dataset.tab === tab.id);
    }
    if (!panes.has(tab.id)) panes.set(tab.id, tab.build());
    body.textContent = '';
    body.append(panes.get(tab.id));
  }

  for (const tab of TABS) {
    const btn = el('button', 'wb-tab', tab.label);
    btn.type = 'button';
    btn.dataset.tab = tab.id;
    btn.addEventListener('click', () => activate(tab));
    tabsBar.append(btn);
  }

  function load() {
    if (!tabsBar.querySelector('.wb-tab.active')) activate(TABS[0]);
  }

  return { el: root, load };
}
