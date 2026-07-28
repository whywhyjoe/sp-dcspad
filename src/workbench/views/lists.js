// Lists & libraries view: every list in the web (hidden ones included — the
// thing the SP UI won't show), with drill-down detail arriving in M2.

import { createGrid } from '../grid.js';

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

const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');

export function createListsView({ client, navigate }) {
  const root = document.createElement('section');
  root.className = 'wb-view wb-view-lists';

  const head = document.createElement('div');
  head.className = 'wb-view-head';
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
    ],
    onOpen: (row) => navigate({ view: 'lists', listId: row.Id, listTitle: row.Title }),
    emptyText: 'No lists in this web.',
    filterPlaceholder: 'Filter lists…',
  });

  root.append(head, grid.el);

  let loaded = false;

  async function load(route) {
    // M2 will route { listId } to a detail pane; for now always show the grid.
    void route;
    if (loaded) return;
    grid.setLoading('Loading lists…');
    try {
      const { items, partial } = await client.getAll('web/lists', {
        select: LIST_SELECT,
        expand: 'RootFolder',
        orderby: 'Title',
        top: 5000,
      });
      grid.setRows(items, { partial });
      loaded = true;
    } catch (err) {
      grid.setError(err);
    }
  }

  return { el: root, load, grid };
}
