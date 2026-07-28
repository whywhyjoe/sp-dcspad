// SP Workbench suite: shell + mock-mode grid behavior, then the live path
// simulated with an injected host context and stubbed /_api routes.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const WB_URL = process.env.DCSPAD_WORKBENCH_URL
  || APP_URL.replace(/index\.html.*$/, 'workbench.html');

const browser = await launchBrowser();

// ---- mock mode ------------------------------------------------------------

const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto(WB_URL);
await page.waitForSelector('.wb-table tbody tr');

await check('mock: chip reads SP: Mock', async () =>
  (await page.locator('#wb-chip-text').textContent()) === 'SP: Mock');

await check('mock: rail renders the Lists section', async () =>
  (await page.locator('.wb-rail-btn.active .wb-rail-label').textContent()) === 'Lists');

await check('mock: lists grid shows the mock lists', async () =>
  (await page.locator('.wb-table tbody tr').count()) === 8);

await check('mock: hidden lists are included', async () =>
  (await page.locator('.wb-table tbody tr', { hasText: 'User Information List' }).count()) === 1);

await check('mock: template numbers render as names', async () =>
  (await page.locator('.wb-table tbody tr', { hasText: 'Document library' }).count()) >= 1);

await check('mock: filter narrows rows and updates the count', async () => {
  await page.fill('.wb-grid-filter', 'gallery');
  const rows = await page.locator('.wb-table tbody tr').count();
  const count = await page.locator('.wb-grid-count').textContent();
  await page.fill('.wb-grid-filter', '');
  return rows === 1 && count === '1 / 8';
});

await check('mock: sorting by Items toggles asc/desc', async () => {
  const itemsHeader = page.locator('.wb-table th', { hasText: 'Items' });
  await itemsHeader.click();
  const asc = await page.locator('.wb-table tbody tr td:nth-child(3)').first().textContent();
  await itemsHeader.click();
  const desc = await page.locator('.wb-table tbody tr td:nth-child(3)').first().textContent();
  return Number(asc) < Number(desc);
});

await check('mock: Id cells are click-to-copy targets', async () =>
  (await page.locator('.wb-table tbody .sp-copy').count()) > 0);

await check('mock: route survives a reload via sessionStorage', async () => {
  await page.reload();
  await page.waitForSelector('.wb-table tbody tr');
  return (await page.locator('.wb-rail-btn.active').count()) === 1;
});

// ---- drilldown (M2) ----

await check('drill: opening a list shows the detail tabs', async () => {
  await page.locator('.wb-table tbody tr', { hasText: 'Projects' }).locator('td').first().click();
  await page.waitForSelector('.wb-tab');
  const tabs = await page.locator('.wb-tab').allTextContents();
  return tabs.join(',') === 'Fields,Views,Content types,Permissions,Raw';
});

await check('drill: fields grid lists internal names and joined choices', async () => {
  await page.waitForSelector('.wb-tab-pane .wb-table tbody tr');
  const rows = await page.locator('.wb-tab-pane .wb-table tbody tr').count();
  const text = await page.locator('.wb-tab-pane .wb-table').textContent();
  return rows === 7 && text.includes('ProjectStatus') && text.includes('Planned | Active | Blocked | Done');
});

await check('drill: raw tab renders the SP.List smart view', async () => {
  await page.locator('.wb-tab', { hasText: 'Raw' }).click();
  await page.waitForSelector('.wb-raw');
  const text = await page.locator('.wb-raw').textContent();
  return text.includes('SP.List') && text.includes('Projects');
});

await check('drill: back returns to the lists grid', async () => {
  await page.locator('.wb-back').click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  return (await page.locator('.wb-pane:not([hidden]) .wb-table tbody tr').count()) === 8;
});

// ---- export (M2) ----

await check('export: toolbar menu offers CSV/JSON/Markdown', async () => {
  await page.locator('.wb-pane:not([hidden]) .wb-menu-wrap .btn', { hasText: 'Export' }).first().click();
  const items = await page.locator('.wb-pane:not([hidden]) .wb-menu:not([hidden]) .wb-menu-item').allTextContents();
  await page.keyboard.press('Escape');
  await page.locator('body').click();
  return items.join(',') === 'Download CSV,Download JSON,Copy CSV,Copy JSON,Copy Markdown';
});

await check('export: toCsv follows RFC 4180 with a BOM', async () =>
  page.evaluate(async () => {
    const { toCsv } = await import('/src/workbench/export.js');
    const csv = toCsv(
      [{ a: 'plain', b: 'has "quotes", commas' }, { a: 'x', b: null }],
      [{ key: 'a', label: 'A' }, { key: 'b', label: 'B' }],
    );
    return csv.startsWith('﻿')
      && csv.includes('A,B')
      && csv.includes('"has ""quotes"", commas"')
      && csv.endsWith('x,');
  }));

await check('export: toMarkdown escapes pipes', async () =>
  page.evaluate(async () => {
    const { toMarkdown } = await import('/src/workbench/export.js');
    const md = toMarkdown([{ a: 'x|y' }], [{ key: 'a', label: 'A' }]);
    return md.includes('| A |') && md.includes('x\\|y');
  }));

// ---- security view (M3) ----

await check('security: tabs render and groups load', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Security' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const tabs = await page.locator('.wb-tab').allTextContents();
  const rows = await page.locator('.wb-tab-body .wb-table tbody tr').count();
  return tabs.join(',') === 'Groups,Role definitions,Role assignments,Inheritance scan' && rows === 3;
});

await check('security: opening a group loads membership lazily', async () => {
  await page.locator('.wb-tab-body .wb-table tbody tr td').first().click();
  await page.waitForSelector('.wb-subpanel .wb-table tbody tr');
  const title = await page.locator('.wb-subpanel-title').textContent();
  return title.includes('Members of Mock Site Owners');
});

await check('security: role definitions decode BasePermissions', async () => {
  await page.locator('.wb-tab', { hasText: 'Role definitions' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  return text.includes('Full control') && text.includes('flags');
});

await check('security: decoder matches documented Contribute mask', async () =>
  page.evaluate(async () => {
    const { decodeBasePermissions } = await import('/src/workbench/perm-kinds.js');
    const { flags } = decodeBasePermissions({ High: '432', Low: '1011028719' });
    return flags.includes('AddListItems') && flags.includes('EditListItems')
      && !flags.includes('ManageWeb') && !flags.includes('ManagePermissions');
  }));

await check('security: inheritance scan runs on demand only', async () => {
  await page.locator('.wb-tab', { hasText: 'Inheritance scan' }).click();
  const before = await page.locator('.wb-tab-body .wb-table tbody tr').allTextContents();
  await page.locator('.wb-scan-bar .btn').click();
  await page.waitForSelector('.wb-scan-bar .wb-view-hint:has-text("break inheritance")');
  const rows = await page.locator('.wb-tab-body .wb-table tbody tr').count();
  return before.join('').includes('Run the scan') && rows === 1;
});

await check('security: list detail has a Permissions tab', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  await page.locator('.wb-table tbody tr', { hasText: 'Projects' }).locator('td').first().click();
  await page.locator('.wb-tab', { hasText: 'Permissions' }).click();
  await page.waitForSelector('.wb-tab-pane .wb-table tbody tr');
  const text = await page.locator('.wb-tab-pane .wb-table').textContent();
  return text.includes('Mock Site Owners') && text.includes('Full Control');
});

// ---- site switcher ----

await check('switch: inspecting another same-tenant site reloads the views', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  await page.fill('#wb-site-input', '/sites/OtherSite');
  await page.locator('#wb-site-open').click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  const status = await page.locator('#wb-status-context').textContent();
  const saved = await page.evaluate(() => sessionStorage.getItem('dcspad.workbench.site'));
  return status.includes('inspecting') && status.includes('/sites/OtherSite')
    && saved.includes('/sites/OtherSite');
});

await check('switch: cross-tenant URLs are rejected inline', async () => {
  await page.fill('#wb-site-input', 'https://evil.example.com/sites/x');
  await page.locator('#wb-site-open').click();
  await page.waitForSelector('#wb-site-error:not([hidden])');
  const text = await page.locator('#wb-site-error').textContent();
  return text.includes('different tenant');
});

await check('switch: blank input returns to the host web', async () => {
  await page.fill('#wb-site-input', '');
  await page.locator('#wb-site-open').click();
  await page.waitForFunction(() =>
    !document.getElementById('wb-status-context').textContent.includes('inspecting'));
  const saved = await page.evaluate(() => sessionStorage.getItem('dcspad.workbench.site'));
  return saved === null;
});

// ---- site overview + script generator (M4) ----

await check('site: tabs render and the web sheet loads', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Site' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const tabs = await page.locator('.wb-tab').allTextContents();
  const text = await page.locator('.wb-tab-body').textContent();
  return tabs.join(',') === 'Web,Site collection,Features,Subwebs,Property bag'
    && text.includes('WebTemplate') && text.includes('Regional settings') && text.includes('Current user');
});

await check('site: features merge site + web scope', async () => {
  await page.locator('.wb-tab', { hasText: 'Features' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  return text.includes('Site') && text.includes('Web')
    && text.includes('(no display name)') && text.includes('00bfea71-4ea5-48d4-a4ad-7ea5c011abe5');
});

await check('site: property bag decodes OData keys, keeps raw copyable', async () => {
  await page.locator('.wb-tab', { hasText: 'Property bag' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  return text.includes('vti_defaultlanguage') && text.includes('vti_x005f_defaultlanguage');
});

await check('scriptgen: pnpjs2 output chains select/top on the fluent route', async () =>
  page.evaluate(async () => {
    const { toPnpjs2 } = await import('/src/workbench/scriptgen.js');
    const code = toPnpjs2({ path: 'web/lists', options: { select: ['Id', 'Title'], top: 5000 } });
    return code.includes('sp.web.lists') && code.includes('.select("Id", "Title")')
      && code.includes('.top(5000)') && code.includes('.get()');
  }));

await check('scriptgen: guid paths route to getById', async () =>
  page.evaluate(async () => {
    const { toPnpjs2 } = await import('/src/workbench/scriptgen.js');
    const code = toPnpjs2({ path: "web/lists(guid'5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03')/fields", options: {} });
    return code.includes('sp.web.lists.getById("5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03").fields');
  }));

await check('scriptgen: REST output carries the nometadata header and query', async () =>
  page.evaluate(async () => {
    const { toRestFetch } = await import('/src/workbench/scriptgen.js');
    const code = toRestFetch({ path: 'web/lists', options: { select: 'Id' } }, 'https://t.sharepoint.com/sites/x');
    return code.includes('https://t.sharepoint.com/sites/x/_api/web/lists?$select=Id')
      && code.includes('application/json;odata=nometadata');
  }));

await check('scriptgen: PowerShell falls back to Invoke-PnPSPRestMethod', async () =>
  page.evaluate(async () => {
    const { toPnpPowerShell } = await import('/src/workbench/scriptgen.js');
    const mapped = toPnpPowerShell({ path: 'web/roledefinitions', options: {} }, 'https://t.sharepoint.com/sites/x');
    const fallback = toPnpPowerShell({ path: 'web/regionalsettings', options: {} }, 'https://t.sharepoint.com/sites/x');
    return mapped.includes('Get-PnPRoleDefinition') && fallback.includes('Invoke-PnPSPRestMethod');
  }));

await check('scriptgen: grids expose the Copy as menu', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  const buttons = await page.locator('.wb-pane:not([hidden]) .wb-menu-wrap .btn').allTextContents();
  return buttons.some((b) => b.includes('Copy as')) && buttons.some((b) => b.includes('Export'));
});

await page.close();

// ---- live path (injected context + stubbed /_api) -------------------------

const live = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const seenHeaders = [];

const listPage1 = {
  value: [
    { Id: '11111111-0000-0000-0000-000000000001', Title: 'Alpha', BaseTemplate: 100, ItemCount: 1, Hidden: false, RootFolder: { ServerRelativeUrl: '/Lists/Alpha' } },
    { Id: '11111111-0000-0000-0000-000000000002', Title: 'Beta', BaseTemplate: 101, ItemCount: 2, Hidden: false, RootFolder: { ServerRelativeUrl: '/Beta' } },
  ],
  'odata.nextLink': '',
};
const listPage2 = {
  value: [
    { Id: '11111111-0000-0000-0000-000000000003', Title: 'Gamma', BaseTemplate: 119, ItemCount: 3, Hidden: true, RootFolder: { ServerRelativeUrl: '/SitePages' } },
  ],
};

await live.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = {
    webAbsoluteUrl: location.origin,
    userDisplayName: 'Stub User',
  };
});

await live.route('**/_api/**', async (route) => {
  const url = route.request().url();
  seenHeaders.push(route.request().headers().accept || '');
  if (url.includes('/_api/web/lists')) {
    if (url.includes('page=2')) return route.fulfill({ json: listPage2 });
    const first = { ...listPage1 };
    first['odata.nextLink'] = `${new URL(url).origin}/_api/web/lists?page=2`;
    return route.fulfill({ json: first });
  }
  return route.fulfill({ json: { value: [] } });
});

await live.goto(WB_URL);
await live.waitForSelector('.wb-table tbody tr');

await check('live: chip reads SP: Live from the host contract', async () =>
  (await live.locator('#wb-chip-text').textContent()) === 'SP: Live');

await check('live: status bar names the web and user', async () =>
  (await live.locator('#wb-status-context').textContent()).includes('Stub User'));

await check('live: /_api requests send the nometadata Accept header', () =>
  seenHeaders.length > 0
  && seenHeaders.every((h) => h.includes('application/json;odata=nometadata')));

await check('live: paging links are followed across pages', async () =>
  (await live.locator('.wb-table tbody tr').count()) === 3);

await check('live: rows from the second page render', async () =>
  (await live.locator('.wb-table tbody tr', { hasText: 'Gamma' }).count()) === 1);

await check('live: switching sites re-targets every /_api request', async () => {
  seenHeaders.length = 0;
  const seenUrls = [];
  await live.route('**/sites/other/_api/**', async (route) => {
    const url = route.request().url();
    seenUrls.push(url);
    if (url.includes('/_api/web/lists')) return route.fulfill({ json: listPage2 });
    return route.fulfill({ json: { Title: 'Other Site', Url: `${new URL(url).origin}/sites/other` } });
  });
  await live.fill('#wb-site-input', '/sites/other');
  await live.locator('#wb-site-open').click();
  await live.waitForFunction(() =>
    document.getElementById('wb-status-context').textContent.includes('/sites/other'));
  await live.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  return seenUrls.some((u) => u.includes('/sites/other/_api/web?'))
    && seenUrls.some((u) => u.includes('/sites/other/_api/web/lists'));
});

await live.close();
await browser.close();
exitWithResult();
