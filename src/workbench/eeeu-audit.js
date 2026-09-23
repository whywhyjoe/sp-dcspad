// EEEU / broad-access audit for the Permissions view. Finds where a site
// grants access to everyone in the organization — the "Everyone except
// external users" (EEEU) and "Everyone" claims, groups that contain them, and
// "People in your organization" sharing links — on the site itself, its
// lists and libraries, and (optionally) the items that carry their own
// permissions.
//
// Ported loosely from Joe's standalone DCS.SecurityGroups content audit (bmo
// FCUPortal, 2026-09). What changed, and why:
//   - Detection is by claim as well as by name. The reference matched titles
//     only ("Everyone except external users", "Everyone", and BMO's
//     "SharePoint EEEU Visitors" group convention); titles are localizable
//     and a group need not follow the convention. The named list is kept
//     (editable in the form), and every site group whose MEMBERS include
//     EEEU/Everyone is found at the start of the scan and matched too.
//   - Organization-wide sharing links are flagged: they grant the same
//     breadth through a SharingLinks.*.Organization* group, never naming EEEU.
//   - Limited Access-only grants are dropped. SharePoint adds them to every
//     parent of an item shared with EEEU; they are noise, not access.
//   - Only securables with their own permissions are read. An inheriting
//     list or subsite has exactly its parent's grants, already reported on
//     the parent's row.
//   - Bounded concurrency over the REST client's own 429/503 retry, and a
//     cooperative stop that keeps what was found — instead of one request at
//     a time behind a ten-minute timeout.
//
// No DOM. All SharePoint access goes through the injected client(s), so the
// rules below are unit-tested against fixtures (tests/workbench.mjs).

import { isInternalList } from './views/lists.js';

export const DEFAULT_TARGETS = [
  'Everyone except external users',
  'Everyone',
  'SharePoint EEEU Visitors',
];

// Claims, not names: stable across tenants' display languages.
//   EEEU      c:0-.f|rolemanager|spo-grid-all-users/<tenant id>
//   Everyone  c:0(.s|true
const EEEU_CLAIM = /^c:0-\.f\|rolemanager\|spo-grid-all-users\//i;
const EVERYONE_CLAIM = /^c:0\(\.s\|true$/i;
// "People in your organization" links: SharingLinks.<doc>.OrganizationView.<link>
// (or OrganizationEdit). Anonymous and specific-people links are other kinds.
const ORG_LINK = /^SharingLinks\.[0-9a-f-]+\.Organization(View|Edit)\b/i;

export const SCOPES = ['site', 'pages', 'lists', 'documents'];

const lower = (s) => String(s ?? '').trim().toLowerCase();

// 'Everyone except external users' | 'Everyone' | null
export function broadPrincipalKind(principal) {
  const login = String(principal?.LoginName || '');
  if (EEEU_CLAIM.test(login)) return 'Everyone except external users';
  if (EVERYONE_CLAIM.test(login)) return 'Everyone';
  return null;
}

export function isOrgSharingLink(principal) {
  return ORG_LINK.test(String(principal?.LoginName || '')) || ORG_LINK.test(String(principal?.Title || ''));
}

// broadGroups: Map(lower group title -> kind it contains). Returns the reason
// a principal counts as broad access, or null.
export function makeMatcher({ targets = DEFAULT_TARGETS, broadGroups = new Map() } = {}) {
  const names = new Set((targets || []).map(lower).filter(Boolean));
  return (member) => {
    if (!member) return null;
    const direct = broadPrincipalKind(member);
    if (direct) return direct;
    if (isOrgSharingLink(member)) return 'Org-wide sharing link';
    const group = broadGroups.get(lower(member.Title)) || broadGroups.get(lower(member.LoginName));
    if (group) return `Group containing ${group}`;
    if (names.has(lower(member.Title)) || names.has(lower(member.LoginName))) return 'Named principal';
    return null;
  };
}

const bindingsOf = (a) => a?.RoleDefinitionBindings?.results || a?.RoleDefinitionBindings || [];
export const isLimitedAccess = (role) =>
  Number(role?.RoleTypeKind) === 1 || /limited access/i.test(String(role?.Name || ''));

// Role assignments -> the broad ones: [{ principal, login, permission, via }].
export function matchAssignments(assignments, matcher, { hideLimitedAccess = true } = {}) {
  const out = [];
  for (const a of assignments || []) {
    const via = matcher(a?.Member);
    if (!via) continue;
    const roles = bindingsOf(a);
    const real = roles.filter((r) => !isLimitedAccess(r));
    if (hideLimitedAccess && roles.length && !real.length) continue;
    out.push({
      principal: a.Member?.Title || '',
      login: a.Member?.LoginName || '',
      permission: (hideLimitedAccess ? real : roles).map((r) => r.Name).filter(Boolean).join(', '),
      via,
    });
  }
  return out;
}

// Which scope checkbox a list belongs to; null = hidden or system (skipped).
export function listScopeOf(list) {
  if (!list || isInternalList(list)) return null;
  if (list.BaseTemplate === 119 || list.BaseTemplate === 850) return 'pages';
  return list.BaseType === 1 ? 'documents' : 'lists';
}

const LIST_TYPE = { pages: 'Pages library', documents: 'Library', lists: 'List' };
const itemType = (scope, item) => {
  if (Number(item?.FSObjType) === 1) return 'Folder';
  return scope === 'pages' ? 'Page' : scope === 'documents' ? 'File' : 'Item';
};

const ASSIGNMENT_OPTIONS = {
  expand: ['Member', 'RoleDefinitionBindings'],
  select: [
    'PrincipalId', 'Member/Title', 'Member/LoginName', 'Member/PrincipalType',
    'RoleDefinitionBindings/Name', 'RoleDefinitionBindings/RoleTypeKind',
  ],
};

async function pool(items, limit, fn, shouldStop) {
  let next = 0;
  const worker = async () => {
    while (next < items.length && !shouldStop()) {
      const item = items[next++];
      await fn(item);
    }
  };
  await Promise.all(Array.from({ length: Math.max(1, Math.min(limit, items.length)) }, worker));
}

const originOf = (url) => { try { return new URL(url).origin; } catch { return ''; } };
// Segment by segment, decode-then-encode (the same rule as pages.js): encodeURI
// leaves '#' and '?' alone, so /Docs/Budget#2026.xlsx would open /Docs/Budget
// with the rest as a fragment. A literal '%' that is not an escape survives
// via the catch.
const encodePath = (path) => String(path).split('/').map((segment) => {
  try { return encodeURIComponent(decodeURIComponent(segment)); } catch { return encodeURIComponent(segment); }
}).join('/');
const absolute = (webUrl, path) => (path ? `${originOf(webUrl)}${encodePath(path)}` : '');

// Runs the audit. `client` is the inspected (root) web; `openWeb(url)`
// resolves a client for another web on the same tenant (subsites).
// `subsites` are the direct subsites the operator ticked — each is scanned
// with its whole tree below it. Resolves to { rows, problems, webs, stopped }.
export async function runEeeuAudit({
  client,
  openWeb,
  subsites = [],
  scopes = { site: true, pages: true, lists: true, documents: true },
  includeItems = true,
  targets = DEFAULT_TARGETS,
  hideLimitedAccess = true,
  concurrency = 3,
  maxDepth = 12,
  onProgress = () => {},
  onRow = () => {},
  onProblem = () => {},
  shouldStop = () => false,
} = {}) {
  const rows = [];
  const problems = [];
  const websScanned = [];
  let seq = 0;

  const addRow = (row) => {
    const full = { Key: `r${++seq}`, ...row };
    rows.push(full);
    onRow(full, rows);
  };
  const addProblem = (webTitle, scope, location, operation, err) => {
    const full = {
      Key: `p${++seq}`,
      Web: webTitle,
      Scope: scope,
      Location: location,
      Operation: operation,
      Status: err?.status ? String(err.status) : '',
      Message: err?.message || String(err),
    };
    problems.push(full);
    onProblem(full, problems);
  };

  // 1. Site groups that contain EEEU/Everyone. Site groups belong to the
  //    site collection, so one read serves every web in the scan.
  onProgress('Finding groups that contain everyone…');
  const broadGroups = new Map();
  let rootWeb = { Title: '', Url: client.webUrl() };
  try {
    rootWeb = await client.get('web', { select: ['Title', 'Url', 'HasUniqueRoleAssignments'] });
  } catch (err) {
    addProblem('', 'Site', client.webUrl(), 'Read the site', err);
  }
  if (shouldStop()) return { rows, problems, webs: websScanned, broadGroups, stopped: true };
  try {
    const { items: groups } = await client.getAll('web/sitegroups', { select: ['Id', 'Title', 'LoginName'] });
    let done = 0;
    await pool(groups, concurrency, async (group) => {
      try {
        const { items: members } = await client.getAll(`web/sitegroups(${group.Id})/users`, {
          select: ['Title', 'LoginName', 'PrincipalType'],
        });
        const kind = members.map(broadPrincipalKind).find(Boolean);
        if (kind) broadGroups.set(lower(group.Title), kind);
      } catch (err) {
        addProblem(rootWeb.Title, 'Group', group.Title, 'Read group membership', err);
      }
      done += 1;
      onProgress(`Checking group membership… ${done} of ${groups.length}`);
    }, shouldStop);
  } catch (err) {
    addProblem(rootWeb.Title, 'Site', rootWeb.Url || client.webUrl(), 'List site groups', err);
  }
  const matcher = makeMatcher({ targets, broadGroups });

  // 2. One web: its own grants (only when it has its own), then its lists.
  async function scanWeb(webClient, web, { isRoot }) {
    if (shouldStop()) return;
    const webTitle = web.Title || web.Url || '';
    const webUrl = web.Url || webClient.webUrl();
    websScanned.push({ Title: webTitle, Url: webUrl });

    if (scopes.site && (isRoot || web.HasUniqueRoleAssignments !== false)) {
      onProgress(`${webTitle}: site permissions…`);
      try {
        const { items } = await webClient.getAll('web/roleassignments', ASSIGNMENT_OPTIONS);
        for (const m of matchAssignments(items, matcher, { hideLimitedAccess })) {
          addRow({
            Web: webTitle, Type: 'Site', Location: webTitle, Name: webTitle,
            SharedWith: m.principal, Login: m.login, Permission: m.permission, Via: m.via,
            Unique: isRoot ? 'Site' : 'Yes', Url: webUrl,
          });
        }
      } catch (err) {
        addProblem(webTitle, 'Site', webUrl, 'Read site permissions', err);
      }
    }

    if (!scopes.pages && !scopes.lists && !scopes.documents) return;
    if (shouldStop()) return;
    onProgress(`${webTitle}: lists…`);
    let lists = [];
    try {
      ({ items: lists } = await webClient.getAll('web/lists', {
        select: [
          'Id', 'Title', 'BaseTemplate', 'BaseType', 'Hidden', 'IsCatalog',
          'HasUniqueRoleAssignments', 'ItemCount', 'RootFolder/ServerRelativeUrl',
        ],
        expand: 'RootFolder',
        top: 5000,
      }));
    } catch (err) {
      addProblem(webTitle, 'Lists', webUrl, 'List the site’s lists', err);
      return;
    }

    const chosen = lists
      .map((list) => ({ list, scope: listScopeOf(list) }))
      .filter(({ scope }) => scope && scopes[scope]);

    let listNo = 0;
    for (const { list, scope } of chosen) {
      if (shouldStop()) return;
      listNo += 1;
      const listPath = list.RootFolder?.ServerRelativeUrl || '';
      const listLabel = list.Title || listPath;
      const base = `web/lists(guid'${list.Id}')`;
      onProgress(`${webTitle}: ${listLabel} (${listNo} of ${chosen.length})…`);

      if (list.HasUniqueRoleAssignments) {
        try {
          const { items } = await webClient.getAll(`${base}/roleassignments`, ASSIGNMENT_OPTIONS);
          for (const m of matchAssignments(items, matcher, { hideLimitedAccess })) {
            addRow({
              Web: webTitle, Type: LIST_TYPE[scope], Location: listLabel, Name: listLabel,
              SharedWith: m.principal, Login: m.login, Permission: m.permission, Via: m.via,
              Unique: 'Yes', Url: absolute(webUrl, listPath),
            });
          }
        } catch (err) {
          addProblem(webTitle, LIST_TYPE[scope], listLabel, 'Read list permissions', err);
        }
      }

      if (!includeItems || shouldStop()) continue;
      let uniqueItems = [];
      try {
        const { items, partial } = await webClient.getAll(`${base}/items`, {
          select: ['Id', 'Title', 'FileRef', 'FileLeafRef', 'FSObjType', 'HasUniqueRoleAssignments'],
          top: 5000,
        }, { allowLargeCap: true });
        uniqueItems = items.filter((it) => it.HasUniqueRoleAssignments === true);
        // The client's ceiling (100,000) was reached: items past it were never
        // looked at. That must read as an incomplete audit, not a clean one.
        if (partial) {
          addProblem(webTitle, LIST_TYPE[scope], listLabel, 'List items with their own permissions', {
            message: `Only the first ${items.length.toLocaleString('en-US')} items were checked — `
              + 'the list is larger than the audit can read, so grants on later items are not in '
              + 'these results.',
          });
        }
      } catch (err) {
        addProblem(webTitle, LIST_TYPE[scope], listLabel, 'List items with their own permissions', err);
        continue;
      }
      if (shouldStop()) return;
      let itemNo = 0;
      await pool(uniqueItems, concurrency, async (item) => {
        const name = item.FileLeafRef || item.Title || `Item ${item.Id}`;
        try {
          const { items } = await webClient.getAll(`${base}/items(${item.Id})/roleassignments`, ASSIGNMENT_OPTIONS);
          for (const m of matchAssignments(items, matcher, { hideLimitedAccess })) {
            addRow({
              Web: webTitle, Type: itemType(scope, item), Location: listLabel, Name: name,
              SharedWith: m.principal, Login: m.login, Permission: m.permission, Via: m.via,
              Unique: 'Yes', Url: absolute(webUrl, item.FileRef),
            });
          }
        } catch (err) {
          addProblem(webTitle, itemType(scope, item), item.FileRef || name, 'Read item permissions', err);
        }
        itemNo += 1;
        onProgress(`${webTitle}: ${listLabel} — item ${itemNo} of ${uniqueItems.length} with its own permissions`);
      }, shouldStop);
    }
  }

  await scanWeb(client, rootWeb, { isRoot: true });

  // 3. Ticked subsites, each with its whole tree (depth-first, deduped).
  const visited = new Set([lower(rootWeb.Url || client.webUrl())]);
  async function scanTree(url, depth) {
    if (shouldStop() || depth > maxDepth) return;
    let webClient;
    let web;
    try {
      webClient = await openWeb(url);
      if (shouldStop()) return;
      web = await webClient.get('web', { select: ['Title', 'Url', 'HasUniqueRoleAssignments'] });
    } catch (err) {
      addProblem('', 'Site', url, 'Open subsite', err);
      return;
    }
    const key = lower(web.Url || webClient.webUrl());
    if (visited.has(key)) return;
    visited.add(key);
    await scanWeb(webClient, web, { isRoot: false });
    if (shouldStop()) return;
    let children = [];
    try {
      ({ items: children } = await webClient.getAll('web/webs', { select: ['Title', 'Url', 'ServerRelativeUrl'] }));
    } catch (err) {
      addProblem(web.Title || '', 'Site', web.Url || url, 'List subsites', err);
      return;
    }
    for (const child of children) {
      await scanTree(child.Url || child.ServerRelativeUrl, depth + 1);
    }
  }
  for (const sub of subsites) {
    await scanTree(sub.url, 1);
  }

  return { rows, problems, webs: websScanned, broadGroups, stopped: shouldStop() };
}
