// Script generator: turn a view's query descriptor — the same { path, options }
// object handed to sp-rest — into paste-ready code in three dialects:
// PnPjs 2 (runs unmodified in the DCSPad JS pane), raw REST fetch, and
// PnP.PowerShell. One descriptor, three outputs; no second source of truth.

const join = (v) => (Array.isArray(v) ? v.join(',') : String(v));

function queryString({ select, expand, filter, orderby, top } = {}) {
  const parts = [];
  if (select) parts.push(`$select=${join(select)}`);
  if (expand) parts.push(`$expand=${join(expand)}`);
  if (filter) parts.push(`$filter=${encodeURIComponent(String(filter))}`);
  if (orderby) parts.push(`$orderby=${join(orderby)}`);
  if (top) parts.push(`$top=${top}`);
  return parts.length ? `?${parts.join('&')}` : '';
}

// ---- PnPjs 2 --------------------------------------------------------------

// Route table: regex over the descriptor path -> pnpjs fluent root.
// Group 1 (when present) is the id captured from the path.
const PNPJS_ROUTES = [
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)$/i, (id, m) => `sp.web.lists.getById("${id}").items.getById(${m[2]})`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/items$/i, (id) => `sp.web.lists.getById("${id}").items`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/fields$/i, (id) => `sp.web.lists.getById("${id}").fields`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/views$/i, (id) => `sp.web.lists.getById("${id}").views`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/contenttypes$/i, (id) => `sp.web.lists.getById("${id}").contentTypes`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/roleassignments$/i, (id) => `sp.web.lists.getById("${id}").roleAssignments`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)$/i, (id) => `sp.web.lists.getById("${id}")`],
  [/^web\/lists$/i, () => 'sp.web.lists'],
  [/^web\/sitegroups\((\d+)\)\/users$/i, (id) => `sp.web.siteGroups.getById(${id}).users`],
  [/^web\/sitegroups$/i, () => 'sp.web.siteGroups'],
  [/^web\/roledefinitions$/i, () => 'sp.web.roleDefinitions'],
  [/^web\/roleassignments$/i, () => 'sp.web.roleAssignments'],
  [/^web\/webs$/i, () => 'sp.web.webs'],
  [/^web\/features$/i, () => 'sp.web.features'],
  [/^site\/features$/i, () => 'sp.site.features'],
  [/^web\/allproperties$/i, () => 'sp.web.allProperties'],
  [/^web\/regionalsettings$/i, () => 'sp.web.regionalSettings'],
  [/^web\/currentuser$/i, () => 'sp.web.currentUser'],
  [/^web$/i, () => 'sp.web'],
  [/^site$/i, () => 'sp.site'],
];

export function toPnpjs2({ path, options = {} }) {
  const clean = String(path).replace(/^\/+/, '');
  const route = PNPJS_ROUTES.find(([re]) => re.test(clean));
  if (!route) {
    // No fluent equivalent — keep it runnable through pnpjs anyway.
    return [
      '// No direct PnPjs 2 fluent route for this endpoint; raw call:',
      `const data = await sp.web.getParentWeb(); // placeholder — see REST tab`,
      `// REST: /_api/${clean}${queryString(options)}`,
    ].join('\n');
  }
  const [re, root] = route;
  const match = clean.match(re);
  let chain = root(match?.[1], match);
  if (options.select) chain += `\n  .select(${join(options.select).split(',').map((s) => `"${s}"`).join(', ')})`;
  if (options.expand) chain += `\n  .expand(${join(options.expand).split(',').map((s) => `"${s}"`).join(', ')})`;
  if (options.filter) chain += `\n  .filter("${String(options.filter).replaceAll('"', '\\"')}")`;
  if (options.orderby) chain += `\n  .orderBy("${join(options.orderby)}")`;
  if (options.top) chain += `\n  .top(${options.top})`;
  return [
    '// PnPjs 2.x — paste into the DCSPad JS pane (pnpjs2 framework enabled)',
    `const data = await ${chain}\n  .get();`,
    'console.table(data);',
  ].join('\n');
}

// ---- raw REST fetch -------------------------------------------------------

export function toRestFetch({ path, options = {} }, webUrl = '') {
  const clean = String(path).replace(/^\/+/, '');
  const base = webUrl
    ? `"${webUrl}/_api/${clean}${queryString(options)}"`
    : `\`\${_spPageContextInfo.webAbsoluteUrl}/_api/${clean}${queryString(options)}\``;
  return [
    '// Raw SharePoint REST (GET) — same-origin cookies authenticate',
    `const response = await fetch(${base}, {`,
    "  credentials: 'same-origin',",
    "  headers: { Accept: 'application/json;odata=nometadata' },",
    '});',
    'const data = await response.json();',
    'console.table(data.value ?? data);',
  ].join('\n');
}

// ---- PnP.PowerShell -------------------------------------------------------

const POWERSHELL_ROUTES = [
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/fields$/i, (id) => `Get-PnPField -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/views$/i, (id) => `Get-PnPView -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/contenttypes$/i, (id) => `Get-PnPContentType -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)$/i, (id) => `Get-PnPList -Identity "${id}" -Includes HasUniqueRoleAssignments`],
  [/^web\/lists$/i, () => 'Get-PnPList -Includes Hidden, ItemCount'],
  [/^web\/sitegroups\((\d+)\)\/users$/i, (id) => `Get-PnPGroupMember -Group (Get-PnPGroup -Identity ${id})`],
  [/^web\/sitegroups$/i, () => 'Get-PnPGroup'],
  [/^web\/roledefinitions$/i, () => 'Get-PnPRoleDefinition'],
  [/^web\/webs$/i, () => 'Get-PnPSubWeb'],
  [/^web\/features$/i, () => 'Get-PnPFeature -Scope Web'],
  [/^site\/features$/i, () => 'Get-PnPFeature -Scope Site'],
  [/^web\/allproperties$/i, () => 'Get-PnPPropertyBag'],
  [/^web$/i, () => 'Get-PnPWeb'],
  [/^site$/i, () => 'Get-PnPSite'],
];

export function toPnpPowerShell({ path, options = {} }, webUrl = '') {
  const clean = String(path).replace(/^\/+/, '');
  const connect = `Connect-PnPOnline -Url "${webUrl || 'https://tenant.sharepoint.com/sites/yoursite'}" -Interactive`;
  const route = POWERSHELL_ROUTES.find(([re]) => re.test(clean));
  if (route) {
    const [re, cmd] = route;
    const id = clean.match(re)?.[1];
    return [`# PnP.PowerShell`, connect, cmd(id)].join('\n');
  }
  return [
    '# PnP.PowerShell — no direct cmdlet; raw REST via Invoke-PnPSPRestMethod',
    connect,
    `Invoke-PnPSPRestMethod -Url "/_api/${clean}${queryString(options).replaceAll('"', '`"')}"`,
  ].join('\n');
}

export const FORMATS = [
  ['Copy as PnPjs 2', toPnpjs2],
  ['Copy as REST fetch', toRestFetch],
  ['Copy as PnP.PowerShell', toPnpPowerShell],
];
