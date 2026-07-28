// Security view: site groups (with lazy membership), role definitions
// (decoded BasePermissions), web role assignments, and an explicit
// broken-inheritance scan across lists.

import { createGrid } from '../grid.js';
import { decodeBasePermissions, principalTypeName } from '../perm-kinds.js';
import { BASE_TEMPLATE_NAMES } from './lists.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const roleNames = (row) =>
  (row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || [])
    .map((r) => r.Name).filter(Boolean).join(', ');

export function createSecurityView({ client }) {
  const root = el('section', 'wb-view wb-view-security');

  const head = el('div', 'wb-view-head');
  head.innerHTML = '<h2>Users, groups &amp; permissions</h2>'
    + '<p class="wb-view-hint">Site groups, role definitions, and who holds '
    + 'what on this web. The inheritance scan is on-demand — it makes '
    + 'SharePoint evaluate security per list.</p>';

  const tabsBar = el('div', 'wb-tabs');
  const body = el('div', 'wb-tab-body');
  root.append(head, tabsBar, body);

  const panes = new Map();

  // ---- Groups + lazy membership ----
  function groupsPane() {
    const wrap = el('div', 'wb-tab-pane');
    const groupsQuery = {
      path: 'web/sitegroups',
      options: { select: ['Id', 'Title', 'Description', 'OwnerTitle', 'PrincipalType', 'OnlyAllowMembersViewMembership'] },
    };
    const grid = createGrid({
      columns: [
        { key: 'Title', label: 'Group' },
        { key: 'Id', label: 'Id' },
        { key: 'OwnerTitle', label: 'Owner' },
        { key: 'OnlyAllowMembersViewMembership', label: 'Members-only view' },
        { key: 'Description', label: 'Description' },
      ],
      onOpen: openMembers,
      emptyText: 'No site groups.',
      filterPlaceholder: 'Filter groups…',
      exportName: 'sp-groups',
      descriptor: { ...groupsQuery, webUrl: client.webUrl() },
    });

    const membersBox = el('div', 'wb-subpanel');
    membersBox.hidden = true;
    const membersTitle = el('h3', 'wb-subpanel-title', '');
    const membersHost = el('div', 'wb-subpanel-body');
    membersBox.append(membersTitle, membersHost);

    wrap.append(grid.el, membersBox);

    grid.setLoading('Loading site groups…');
    client.getAll(groupsQuery.path, groupsQuery.options)
      .then(({ items, partial }) => grid.setRows(items, { partial }))
      .catch((err) => grid.setError(err));

    function openMembers(group) {
      membersBox.hidden = false;
      membersTitle.textContent = `Members of ${group.Title}`;
      membersHost.textContent = '';
      const membersQuery = {
        path: `web/sitegroups(${group.Id})/users`,
        options: { select: ['Id', 'Title', 'LoginName', 'Email', 'IsSiteAdmin', 'PrincipalType'] },
      };
      const membersGrid = createGrid({
        columns: [
          { key: 'Title', label: 'Name' },
          { key: 'LoginName', label: 'Login', mono: true, copyable: true },
          { key: 'Email', label: 'Email', copyable: true },
          { key: 'IsSiteAdmin', label: 'Site admin' },
          { key: 'PrincipalType', label: 'Type', format: principalTypeName },
        ],
        emptyText: 'No members.',
        filterPlaceholder: 'Filter members…',
        exportName: `members-${group.Id}`,
        descriptor: { ...membersQuery, webUrl: client.webUrl() },
      });
      membersHost.append(membersGrid.el);
      membersGrid.setLoading('Loading members…');
      client.getAll(membersQuery.path, membersQuery.options)
        .then(({ items }) => membersGrid.setRows(items))
        // Locked-down groups 403 for non-owners: show it inline, don't fail the view.
        .catch((err) => membersGrid.setError(err));
    }

    return wrap;
  }

  // ---- Role definitions + decoded permissions ----
  function roleDefsPane() {
    const wrap = el('div', 'wb-tab-pane');
    const grid = createGrid({
      columns: [
        { key: 'Name', label: 'Role' },
        { key: 'RoleTypeKind', label: 'Kind' },
        { key: 'Hidden', label: 'Hidden' },
        {
          key: 'BasePermissions',
          label: 'Permissions',
          value: (row) => decodeBasePermissions(row.BasePermissions).flags.length,
          format: (v, row) => {
            const d = decodeBasePermissions(row.BasePermissions);
            if (d.isFullControl) return 'Full control';
            if (d.isEmpty) return 'None';
            return `${d.flags.length} flags`;
          },
        },
        { key: 'Description', label: 'Description' },
      ],
      onOpen: openDecode,
      emptyText: 'No role definitions.',
      filterPlaceholder: 'Filter roles…',
      exportName: 'sp-roledefinitions',
      descriptor: {
        path: 'web/roledefinitions',
        options: { select: ['Id', 'Name', 'Description', 'RoleTypeKind', 'Hidden', 'BasePermissions'] },
        webUrl: client.webUrl(),
      },
    });

    const decodeBox = el('div', 'wb-subpanel');
    decodeBox.hidden = true;
    const decodeTitle = el('h3', 'wb-subpanel-title', '');
    const decodeBody = el('div', 'wb-subpanel-body wb-flags');
    decodeBox.append(decodeTitle, decodeBody);
    wrap.append(grid.el, decodeBox);

    grid.setLoading('Loading role definitions…');
    client.getAll('web/roledefinitions', {
      select: ['Id', 'Name', 'Description', 'RoleTypeKind', 'Hidden', 'BasePermissions'],
    })
      .then(({ items, partial }) => grid.setRows(items, { partial }))
      .catch((err) => grid.setError(err));

    function openDecode(role) {
      decodeBox.hidden = false;
      const d = decodeBasePermissions(role.BasePermissions);
      decodeTitle.textContent = `${role.Name} — ${d.isFullControl ? 'full control' : `${d.flags.length} permission flags`}`;
      decodeBody.textContent = '';
      for (const flag of d.flags) decodeBody.append(el('span', 'wb-flag', flag));
      if (d.isEmpty) decodeBody.append(el('span', 'wb-view-hint', 'No permission bits set.'));
    }

    return wrap;
  }

  // ---- Web role assignments ----
  function assignmentsPane() {
    const wrap = el('div', 'wb-tab-pane');
    const grid = createGrid({
      rowKey: 'PrincipalId',
      columns: [
        { key: 'Member', label: 'Principal', value: (row) => row.Member?.Title || '' },
        { key: 'LoginName', label: 'Login', value: (row) => row.Member?.LoginName || '', mono: true, copyable: true },
        { key: 'PrincipalType', label: 'Type', value: (row) => row.Member?.PrincipalType, format: principalTypeName },
        { key: 'Roles', label: 'Roles', value: roleNames },
      ],
      emptyText: 'No role assignments.',
      filterPlaceholder: 'Filter assignments…',
      exportName: 'sp-roleassignments',
      descriptor: {
        path: 'web/roleassignments',
        options: { expand: ['Member', 'RoleDefinitionBindings'] },
        webUrl: client.webUrl(),
      },
    });
    wrap.append(grid.el);
    grid.setLoading('Loading role assignments…');
    client.getAll('web/roleassignments', {
      expand: ['Member', 'RoleDefinitionBindings'],
      select: [
        'PrincipalId', 'Member/Id', 'Member/Title', 'Member/LoginName',
        'Member/PrincipalType', 'RoleDefinitionBindings/Id', 'RoleDefinitionBindings/Name',
      ],
    })
      .then(({ items, partial }) => grid.setRows(items, { partial }))
      .catch((err) => grid.setError(err));
    return wrap;
  }

  // ---- Broken-inheritance scan (explicit) ----
  function inheritancePane() {
    const wrap = el('div', 'wb-tab-pane');
    const bar = el('div', 'wb-scan-bar');
    const btn = el('button', 'btn', 'Scan lists for unique permissions');
    btn.type = 'button';
    const hint = el('span', 'wb-view-hint',
      'Asks SharePoint for HasUniqueRoleAssignments on every list — slow on large sites, so it only runs on demand.');
    bar.append(btn, hint);

    const grid = createGrid({
      columns: [
        { key: 'Title', label: 'List' },
        { key: 'HasUniqueRoleAssignments', label: 'Unique permissions' },
        { key: 'BaseTemplate', label: 'Template', format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? '') },
        { key: 'Hidden', label: 'Hidden' },
        { key: 'Id', label: 'Id', mono: true, copyable: true },
      ],
      emptyText: 'Run the scan to see results.',
      filterPlaceholder: 'Filter results…',
      exportName: 'sp-unique-permissions',
      descriptor: {
        path: 'web/lists',
        options: { select: ['Id', 'Title', 'Hidden', 'BaseTemplate', 'HasUniqueRoleAssignments'], top: 5000 },
        webUrl: client.webUrl(),
      },
    });
    wrap.append(bar, grid.el);
    grid.setRows([]);   // show the "run the scan" empty state up front

    btn.addEventListener('click', async () => {
      btn.disabled = true;
      grid.setLoading('Scanning… (per-list security evaluation)');
      try {
        const { items, partial } = await client.getAll('web/lists', {
          select: ['Id', 'Title', 'Hidden', 'BaseTemplate', 'HasUniqueRoleAssignments'],
          top: 5000,
        });
        const broken = items.filter((l) => l.HasUniqueRoleAssignments);
        grid.setRows(broken.length ? broken : items, { partial });
        hint.textContent = broken.length
          ? `${broken.length} of ${items.length} lists break inheritance (showing them).`
          : `No list breaks inheritance (showing all ${items.length} scanned).`;
      } catch (err) {
        grid.setError(err);
      } finally {
        btn.disabled = false;
      }
    });

    return wrap;
  }

  const TABS = [
    { id: 'groups', label: 'Groups', build: groupsPane },
    { id: 'roledefs', label: 'Role definitions', build: roleDefsPane },
    { id: 'assignments', label: 'Role assignments', build: assignmentsPane },
    { id: 'inheritance', label: 'Inheritance scan', build: inheritancePane },
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
