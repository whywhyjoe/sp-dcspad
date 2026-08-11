// SP Workbench suite: shell + mock-mode grid behavior, then the live path
// simulated with an injected host context and stubbed /_api routes.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const WB_URL = process.env.DCSPAD_WORKBENCH_URL
  || APP_URL.replace(/index\.html.*$/, 'workbench.html');

const browser = await launchBrowser();

// ---- mock mode ------------------------------------------------------------

const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto(WB_URL);
await page.waitForSelector('.wb-home-cards');

await check('mock: chip reads SP: Mock', async () =>
  (await page.locator('#wb-chip-text').textContent()) === 'SP: Mock');

await check('mock: Site landing is the default view with web + user cards', async () => {
  const active = await page.locator('.wb-rail-btn.active .wb-rail-label').textContent();
  const cards = await page.locator('.wb-home-card').count();
  const role = await page.locator('.wb-home-card .wb-role-chip').textContent();
  const facts = await page.locator('.wb-home-card').first().textContent();
  return active === 'Site' && cards === 2 && role === 'Site admin'
    && facts.includes('Mock Web');
});

await check('mock: status bar shows the current user and a role chip', async () => {
  const user = await page.locator('#wb-status-user').textContent();
  const chip = await page.locator('#wb-status-role').textContent();
  const cls = await page.locator('#wb-status-role').getAttribute('class');
  return user.includes('Mock Developer') && chip === 'Site admin'
    && cls.includes('wb-role-admin');
});

await check('mock: subwebs offer one-click Inspect', async () => {
  await page.waitForSelector('.wb-home-subwebs .wb-table tbody tr');
  const rows = await page.locator('.wb-home-subwebs .wb-table tbody tr').count();
  const inspect = await page.locator('.wb-home-subwebs .btn', { hasText: 'Inspect' }).count();
  return rows === 1 && inspect === 1;
});

await check('mock: rail renders the Lists section', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  return (await page.locator('.wb-rail-btn.active .wb-rail-label').textContent()) === 'Lists';
});

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
  await page.locator('.wb-rail-btn', { hasText: 'Permissions' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const tabs = await page.locator('.wb-tab').allTextContents();
  const rows = await page.locator('.wb-tab-body .wb-table tbody tr').count();
  return tabs.join(',') === 'Groups,Members,Role definitions,Role assignments,Inheritance scan' && rows === 3;
});

await check('perm: head links jump to the SP permission panels', async () => {
  const links = await page.locator('.wb-head-link').evaluateAll((nodes) =>
    nodes.map((a) => `${a.textContent}=${a.getAttribute('href')}`));
  return links.length === 3
    && links[0].includes('user.aspx')
    && links.some((l) => l.includes('people.aspx'))
    && links.some((l) => l.includes('groups.aspx'));
});

await check('perm: members roster flattens users by group, exportable', async () => {
  await page.locator('.wb-tab', { hasText: 'Members' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const rows = await page.locator('.wb-tab-body .wb-table tbody tr').count();
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  const exportBtn = await page.locator('.wb-tab-body .wb-menu-wrap .btn', { hasText: 'Export' }).count();
  return rows === 4 && text.includes('Mock Site Members')
    && text.includes('dev@mock.local') && exportBtn === 1;
});

await check('perm: adding a user posts the claims login to the group', async () => {
  await page.fill('.wb-members-login', 'new.user@mock.local');
  await page.locator('.wb-members-add .btn', { hasText: 'Add to group' }).click();
  await page.waitForSelector('.wb-view-security .wb-consent:not([hidden])');
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const add = writes.find((w) => w.url.includes('/sitegroups(3)/users') && !w.url.includes('removebyid'));
  return Boolean(add)
    && JSON.parse(add.body).LoginName === 'i:0#.f|membership|new.user@mock.local';
});

await check('perm: removing a member asks first, then posts removebyid', async () => {
  await page.locator('.wb-view-security .wb-consent .btn', { hasText: 'Dismiss' }).click();
  await page.locator('.wb-tab-body .wb-table tbody tr .wb-cell-copylink').first().click();
  await page.waitForSelector('.wb-view-security .wb-consent:not([hidden])');
  const question = await page.locator('.wb-view-security .wb-consent').textContent();
  const before = await page.evaluate(() =>
    (globalThis.__DCSPAD_WB_WRITES__ || []).filter((w) => w.url.includes('removebyid')).length);
  await page.locator('.wb-view-security .wb-consent .btn', { hasText: 'Remove' }).click();
  await page.waitForFunction((n) =>
    (globalThis.__DCSPAD_WB_WRITES__ || []).filter((w) => w.url.includes('removebyid')).length > n, before);
  return question.includes('Remove') && before === 0;
});

await check('security: opening a group loads membership lazily', async () => {
  await page.locator('.wb-tab', { hasText: 'Groups' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
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

await check('advanced: tabs render and the web sheet loads', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Advanced' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const tabs = await page.locator('.wb-tab').allTextContents();
  const text = await page.locator('.wb-tab-body').textContent();
  return tabs.join(',') === 'Web,Site collection,Features,Subwebs,Property bag'
    && text.includes('WebTemplate') && text.includes('Regional settings') && text.includes('Current user');
});

await check('advanced: features merge site + web scope', async () => {
  await page.locator('.wb-tab', { hasText: 'Features' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  return text.includes('Site') && text.includes('Web')
    && text.includes('(no display name)') && text.includes('00bfea71-4ea5-48d4-a4ad-7ea5c011abe5');
});

await check('advanced: property bag decodes OData keys, keeps raw copyable', async () => {
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

// ---- config links (T2) ----

await check('links: rail lists all eight views in the agreed order', async () => {
  const labels = await page.locator('.wb-rail-btn .wb-rail-label').allTextContents();
  const seps = await page.locator('.wb-rail-sep').count();
  const groups = await page.locator('.wb-rail-group').allTextContents();
  return labels.join(',') === 'Site,Permissions,Lists,Pages,Files,Query,Panels,Advanced'
    && groups.join(',') === 'Site,Content,Tools'
    && seps === 2;
});

await check('links: Panels renders the curated quick jumps only', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Panels' }).click();
  await page.waitForSelector('.wb-linkgroup');
  const groups = await page.locator('.wb-linkgroup h3').allTextContents();
  const text = await page.locator('.wb-links').textContent();
  const first = page.locator('.wb-link').first();
  const href = await first.getAttribute('href');
  const target = await first.getAttribute('target');
  return groups.join(',') === 'General,Permissions & people,Recycle bins,Galleries,Search'
    && href.endsWith('/_layouts/15/settings.aspx')
    && href.startsWith(new URL(WB_URL).origin)
    && target === '_blank'
    // The 2026-07-31 curation pass removed these — they must stay gone.
    && !text.includes('Change the look') && !text.includes('Site features')
    && !text.includes('Term store') && !text.includes('Master page');
});

await check('links: lists grid rows carry a list-settings link', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  const href = await page.locator('.wb-table tbody .wb-cell-link').first().getAttribute('href');
  return href.includes('/_layouts/15/listedit.aspx?List=');
});

await check('links: list drilldown header links to list settings', async () => {
  await page.locator('.wb-table tbody tr', { hasText: 'Projects' }).locator('td').first().click();
  await page.waitForSelector('.wb-detail-settings');
  const href = await page.locator('.wb-detail-settings').getAttribute('href');
  await page.locator('.wb-back').click();
  await page.waitForSelector('.wb-pane:not([hidden]) .wb-table tbody tr');
  return href.includes('listedit.aspx?List=5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03');
});

// ---- site favorites (T2) ----

await check('fav: star toggle persists through the state.js seam', async () => {
  await page.locator('#wb-site-fav').click();
  const star = await page.locator('#wb-site-fav').textContent();
  const doc = await page.evaluate(() => JSON.parse(localStorage.getItem('dcspad.v2.wbsites')));
  return star === '★' && doc.kind === 'dcspad-workbench-sites'
    && doc.items.length === 1 && doc.items[0].url === '';
});

await check('fav: menu lists favorites and earlier recents', async () => {
  await page.locator('#wb-site-favlist').click();
  await page.waitForSelector('#wb-site-menu:not([hidden])');
  const text = await page.locator('#wb-site-menu').textContent();
  await page.locator('body').click();
  return text.includes('Favorites') && text.includes('(host web)')
    && text.includes('Recent') && text.includes('/sites/OtherSite');
});

await check('fav: favorites survive a reload', async () => {
  await page.reload();
  await page.waitForSelector('.wb-table tbody tr');
  return (await page.locator('#wb-site-fav').textContent()) === '★';
});

await check('fav: removing the favorite empties the stored doc', async () => {
  await page.locator('#wb-site-fav').click();
  const doc = await page.evaluate(() => JSON.parse(localStorage.getItem('dcspad.v2.wbsites')));
  return (await page.locator('#wb-site-fav').textContent()) === '☆' && doc.items.length === 0;
});

// ---- query builder (T2) ----

await check('query: list picker and field checkboxes load', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Query' }).click();
  // <option> elements never satisfy waitForSelector's visibility check.
  await page.waitForFunction(() =>
    document.querySelectorAll('.wb-qb-list option').length > 0);
  const options = await page.locator('.wb-qb-list option').count();
  await page.selectOption('.wb-qb-list', {
    value: '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03',
  });
  await page.waitForSelector('.wb-qb-fieldopt');
  const fields = await page.locator('.wb-qb-fieldopt').count();
  return options === 9 && fields === 6;
});

await check('query: composed filter quotes by type and runs into the grid', async () => {
  await page.locator('.wb-qb-fieldopt input[value="ProjectStatus"]').check();
  await page.locator('.wb-qb .btn', { hasText: '+ Filter' }).click();
  await page.selectOption('.wb-qb-filterrow .wb-qb-field', 'ProjectStatus');
  await page.fill('.wb-qb-filterrow .wb-qb-value', 'Active');
  const raw = await page.locator('.wb-qb-raw').inputValue();
  await page.locator('.wb-qb-run').click();
  await page.waitForSelector('.wb-qb-results .wb-table tbody tr');
  const rows = await page.locator('.wb-qb-results .wb-table tbody tr').count();
  const copyAs = await page.locator('.wb-qb-results .wb-menu-wrap .btn', { hasText: 'Copy as' }).count();
  return raw.includes("$filter=ProjectStatus eq 'Active'")
    && raw.includes('$select=') && rows === 6 && copyAs === 1;
});

await check('query: last query is restored per web after a reload', async () => {
  await page.reload();
  await page.waitForSelector('.wb-qb-raw');
  await page.waitForFunction(() => document.querySelector('.wb-qb-raw')?.value.length > 0);
  const raw = await page.locator('.wb-qb-raw').inputValue();
  return raw.includes('ProjectStatus');
});

await check('query: pure helpers quote, compose, and round-trip', async () =>
  page.evaluate(async () => {
    const {
      filterClause, composeFilter, rawToDescriptor, descriptorToRaw, columnsForSelect,
    } = await import('/src/workbench/views/query.js');
    const dateClause = filterClause({ field: 'DueDate', type: 'DateTime', op: 'ge', value: '2026-01-01' });
    const quoteClause = filterClause({ field: 'Title', type: 'Text', op: 'substringof', value: "O'Brien" });
    const combined = composeFilter([
      { field: 'A', type: 'Number', op: 'gt', value: '2' },
      { field: 'B', type: 'Text', op: 'eq', value: 'x', join: 'or' },
    ]);
    const parsed = rawToDescriptor('web/lists?$select=Id,Title&$top=5');
    const round = descriptorToRaw(parsed);
    const rejected = rawToDescriptor('web/lists?$skip=5');
    const malformed = rawToDescriptor('web/lists?$filter=%');
    const col = columnsForSelect(['Editor/Title'])[0];
    return dateClause.startsWith("DueDate ge datetime'2026-01-01")
      && quoteClause === "substringof('O''Brien',Title)"
      && combined === "A gt 2 or B eq 'x'"
      && parsed.options.select.join(',') === 'Id,Title' && parsed.options.top === 5
      && round === 'web/lists?$select=Id,Title&$top=5'
      && rejected === null
      && malformed === null
      && col.value({ Editor: { Title: 'Pat' } }) === 'Pat';
  }));

// ---- page inspector (T2) ----

await check('canvas: JSON format parses into typed controls, never throwing', async () =>
  page.evaluate(async () => {
    const { parseCanvasContent, webPartName } = await import('/src/workbench/canvas.js');
    const parsed = parseCanvasContent(JSON.stringify([
      { controlType: 4, position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12 }, innerHTML: '<p>hey</p>' },
      { controlType: 3, position: { zoneIndex: 2, sectionIndex: 1, controlIndex: 1, sectionFactor: 6 }, webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', webPartData: { title: 'QL', serverProcessedContent: { searchablePlainTexts: { a: 'x' } } } },
      { utterly: 'malformed' },
    ]));
    const empty = parseCanvasContent(null);
    const broken = parseCanvasContent('not json at all');
    return parsed.ok
      && parsed.controls.map((c) => c.kind).join(',') === 'text,webpart,unknown'
      && parsed.errors.length === 1
      && webPartName('C70391EA-0B10-4EE9-B2B4-006D3FCAD0CD') === 'Quick links'
      && webPartName('00000000-dead-beef-0000-000000000000') === '00000000-dead-beef-0000-000000000000'
      && empty.ok && empty.controls.length === 0
      && broken.ok === false && broken.errors.length === 1;
  }));

await check('canvas: legacy HTML format and section math decode', async () =>
  page.evaluate(async () => {
    const { parseCanvasContent, buildSectionTree, sanitizeHtml, textOfControl } =
      await import('/src/workbench/canvas.js');
    const html = '<div><div data-sp-canvascontrol="" data-sp-controldata='
      + '"{&quot;controlType&quot;:4,&quot;position&quot;:{&quot;zoneIndex&quot;:1,'
      + '&quot;sectionIndex&quot;:1,&quot;controlIndex&quot;:1,&quot;sectionFactor&quot;:12}}">'
      + '<div data-sp-rte=""><p>legacy body</p></div></div></div>';
    const parsed = parseCanvasContent(html);
    const tree = buildSectionTree(parseCanvasContent(JSON.stringify([
      { controlType: 4, position: { zoneIndex: 2, sectionIndex: 1, controlIndex: 1, sectionFactor: 6 }, innerHTML: 'b' },
      { controlType: 4, position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12 }, innerHTML: 'a' },
      { controlType: 4, innerHTML: 'stray' },
    ])).controls);
    const clean = sanitizeHtml('<p onclick="x()">a</p><script>bad()</script><a href="javascript:1">l</a>');
    return parsed.ok && parsed.controls.length === 1
      && parsed.controls[0].kind === 'text'
      && textOfControl(parsed.controls[0]) === 'legacy body'
      && tree.sections.length === 2
      && tree.sections[0].zoneIndex === 1
      && tree.sections[1].columns[0].sectionFactor === 6
      && tree.unplaced.length === 1
      && !clean.includes('script') && !clean.includes('onclick') && !clean.includes('javascript:');
  }));

await check('pages: master grid lists pages with folders and promoted badges', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Pages' }).click();
  await page.waitForSelector('.wb-view-pages .wb-table tbody tr');
  const rows = await page.locator('.wb-view-pages .wb-table tbody tr').count();
  const text = await page.locator('.wb-view-pages .wb-table').textContent();
  const libHref = await page.locator('.wb-view-pages .wb-head-link').getAttribute('href');
  return rows === 5 && text.includes('News-Update.aspx') && text.includes('News')
    && text.includes('/news') && text.includes('/news/fr')
    && libHref.includes('/SitePages');
});

await check('pages: drilldown opens on Extract with the reordered tabs and URL copy', async () => {
  await page.locator('.wb-view-pages .wb-table tbody tr', { hasText: 'Home.aspx' })
    .locator('td').first().click();
  await page.waitForSelector('.wb-text-rendered');
  const tabs = await page.locator('.wb-view-pages .wb-tab').allTextContents();
  const active = await page.locator('.wb-view-pages .wb-tab.active').textContent();
  const frag = page.locator('.wb-view-pages .wb-detail-id');
  const fragText = await frag.textContent();
  const fragTitle = await frag.getAttribute('title');
  const actions = await page.locator('.wb-detail-actions .btn').allTextContents();
  return tabs.join(',') === 'Extract,Metadata,Structure,Web parts,Raw'
    && active === 'Extract'
    && fragText === '/SitePages/Home.aspx'
    && fragTitle.includes(`${new URL(WB_URL).origin}/SitePages/Home.aspx`)
    && actions.join(',').includes('Export content')
    && actions.join(',').includes('Export raw');
});

await check('pages: structure tab parses the canvas and flags the malformed entry', async () => {
  await page.locator('.wb-view-pages .wb-tab', { hasText: 'Structure' }).click();
  await page.waitForSelector('.wb-canvas-tree');
  const tree = await page.locator('.wb-canvas-tree').textContent();
  const notice = await page.locator('.wb-view-pages .wb-pane:not([hidden]) .wb-grid-notice').textContent();
  return tree.includes('2 columns') && tree.includes('6/12') && tree.includes('Quick links')
    && tree.includes('Unplaced entries (1)')
    && notice.includes('could not be fully parsed');
});

await check('pages: web part inventory names known ids, degrades unknown ones', async () => {
  await page.locator('.wb-view-pages .wb-tab', { hasText: 'Web parts' }).click();
  await page.waitForSelector('.wb-tab-body .wb-table tbody tr');
  const text = await page.locator('.wb-tab-body .wb-table').textContent();
  return text.includes('Quick links')
    && text.includes('c70391ea-0b10-4ee9-b2b4-006d3fcad0cd')
    && text.includes('ffff0000-1111-2222-3333-444455556666')
    && text.includes('Docs · Pad');
});

await check('pages: Extract merges every part into one content box plus one HTML box', async () => {
  await page.locator('.wb-view-pages .wb-tab', { hasText: 'Extract' }).click();
  await page.waitForSelector('.wb-text-rendered');
  const boxes = await page.locator('.wb-view-pages .wb-text-block').count();
  const titles = await page.locator('.wb-view-pages .wb-subpanel-title').allTextContents();
  const headings = await page.locator('.wb-text-part').allTextContents();
  const rendered = await page.locator('.wb-text-rendered').textContent();
  const raw = await page.locator('.wb-text-raw').textContent();
  return boxes === 2 && titles.join(',') === 'Content,HTML'
    // one heading per non-empty part; 'Mystery part' carries no text and is skipped
    && headings.join(',') === 'Text,Quick links'
    && rendered.includes('Welcome to the mock intranet')
    && rendered.includes('Docs') && rendered.includes('Pad')
    && !rendered.includes('Mystery part')
    && !rendered.toLowerCase().includes('web part')
    && raw.includes('<h2>Welcome</h2>') && raw.includes('<!-- Text -->');
});

await check('page-export: content export merges metadata and content per spec', async () =>
  page.evaluate(async () => {
    const { buildContentExport, pageLocation, exportFileStem } =
      await import('/src/workbench/page-export.js');
    const { parseCanvasContent } = await import('/src/workbench/canvas.js');
    const { mockResolver } = await import('/src/workbench/mock-data.js');
    const item = mockResolver(
      `${location.origin}/_api/web/lists(guid'5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02')/items(1)`,
    );
    const md = buildContentExport({
      item,
      controls: parseCanvasContent(item.CanvasContent1).controls,
      siteTitle: 'Mock Web',
      webUrl: location.origin,
      libraryTitle: 'Site Pages',
      libraryRootPath: '/SitePages',
    });
    return md.startsWith('# Home')
      && md.includes('> Mock landing page.')
      && md.includes('Created 2026-05-02 by Mock Developer')
      && md.includes('Location: Mock Web | Site Pages')
      && md.includes('Welcome to the mock intranet home page')
      && md.includes('## Text')
      && md.includes('## Quick links')
      && md.includes('- Docs')
      // empty parts are skipped and the technical framing is gone
      && !md.includes('Mystery part')
      && !md.includes('Web part:')
      && md.includes('1 part could not be read')
      && md.includes('## Metadata')
      && md.includes(`- URL: ${location.origin}/SitePages/Home.aspx`)
      && md.indexOf('## Metadata') > md.indexOf('Welcome to the mock intranet')
      && pageLocation({
        siteTitle: 'FCUPortal', libraryTitle: 'SitePages',
        fileDirRef: '/SitePages/News/fr', libraryRootPath: '/SitePages',
      }) === 'FCUPortal | SitePages | News/fr'
      && exportFileStem({ FileLeafRef: 'News-Update.aspx' }) === 'news-update';
  }));

await check('pages: metadata tab maps field types to editors and guards content fields', async () => {
  await page.locator('.wb-view-pages .wb-tab', { hasText: 'Metadata' }).click();
  await page.waitForSelector('.wb-editor-row');
  const kinds = await page.evaluate(() =>
    [...document.querySelectorAll('.wb-editor-row')].map((row) => {
      const readonly = row.classList.contains('wb-editor-readonly');
      const control = row.querySelector('select') ? 'select'
        : row.querySelector('textarea') ? 'textarea'
          : row.querySelector('input')?.type || 'static';
      return `${row.dataset.internal}:${control}:${readonly ? 'ro' : 'edit'}`;
    }).join('|'));
  return kinds.includes('PageCategory:select:edit')
    && kinds.includes('ReviewDate:datetime-local:edit')
    && kinds.includes('ShowInNav:checkbox:edit')
    && kinds.includes('RelatedLink:text:edit')
    && kinds.includes('CanvasContent1:static:ro')
    && kinds.includes('Editor:static:ro');
});

await check('pages: saving metadata posts ValidateUpdateListItem through the mock writer', async () => {
  await page.fill('.wb-editor-row[data-internal="Title"] input', 'Home v2');
  await page.locator('.wb-editor-bar .btn').click();
  await page.waitForSelector('.wb-editor-status.wb-editor-saved');
  // Earlier checks record group-membership writes — assert on the VULI one.
  const writes = await page.evaluate(() =>
    (globalThis.__DCSPAD_WB_WRITES__ || []).filter((w) => w.url.includes('ValidateUpdateListItem')));
  const write = writes[writes.length - 1];
  const body = JSON.parse(write.body);
  return writes.length === 1
    && write.url.includes("lists(guid'5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02')/items(1)/ValidateUpdateListItem")
    && body.formValues.length === 1
    && body.formValues[0].FieldName === 'Title'
    && body.formValues[0].FieldValue === 'Home v2'
    && body.bNewDocumentUpdate === false;
});

await check('field-editor: FieldValue conventions match ValidateUpdateListItem', async () =>
  page.evaluate(async () => {
    const { toFormValue, fromItemValue, isEditable } = await import('/src/workbench/field-editor.js');
    return toFormValue({ TypeAsString: 'MultiChoice' }, ['A', 'B']) === ';#A;#B;#'
      && toFormValue({ TypeAsString: 'MultiChoice' }, []) === ''
      && toFormValue({ TypeAsString: 'Boolean' }, true) === '1'
      && toFormValue({ TypeAsString: 'Boolean' }, false) === '0'
      && toFormValue({ TypeAsString: 'URL' }, { url: 'https://x', description: 'desc' }) === 'https://x, desc'
      && toFormValue({ TypeAsString: 'URL' }, { url: '', description: 'desc' }) === ''
      && toFormValue({ TypeAsString: 'Number' }, '3,5') === '3.5'
      && toFormValue({ TypeAsString: 'DateTime' }, '2026-07-30T14:00').includes('2026-07-30')
      && fromItemValue({ TypeAsString: 'MultiChoice' }, ';#A;#B;#').join(',') === 'A,B'
      && fromItemValue({ TypeAsString: 'Boolean' }, 'Yes') === true
      && fromItemValue({ TypeAsString: 'URL' }, { Url: 'https://x', Description: 'd' }).url === 'https://x'
      && isEditable({ TypeAsString: 'Choice', InternalName: 'C' }) === true
      && isEditable({ TypeAsString: 'Note', InternalName: 'CanvasContent1' }) === false
      && isEditable({ TypeAsString: 'Text', InternalName: 'T', ReadOnlyField: true }) === false
      && isEditable({ TypeAsString: 'User', InternalName: 'U' }) === false;
  }));

await page.close();

// ---- live path (injected context + stubbed /_api) -------------------------

const live = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const seenHeaders = [];
let pageDetailUrl = '';

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
  if (url.includes("lists(guid'11111111-0000-0000-0000-000000000003')/items(7)")) {
    pageDetailUrl = url;
    const expand = new URL(url).searchParams.get('$expand') || '';
    if (!expand.split(',').includes('Author') || !expand.split(',').includes('Editor')) {
      return route.fulfill({
        status: 400,
        json: { 'odata.error': { message: { value: 'Author must be included in $expand.' } } },
      });
    }
    return route.fulfill({ json: {
      Id: 7,
      Title: 'Live page',
      FileLeafRef: 'Live.aspx',
      FileRef: '/SitePages/Live.aspx',
      FileDirRef: '/SitePages',
      Description: 'Live detail fixture',
      Created: '2026-07-01T00:00:00Z',
      Modified: '2026-08-01T00:00:00Z',
      Author: { Title: 'Page Author' },
      Editor: { Title: 'Page Editor' },
      CanvasContent1: JSON.stringify([{
        controlType: 4,
        position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12 },
        innerHTML: '<p>Live page body</p>',
      }]),
      LayoutWebpartsContent: '',
    } });
  }
  if (url.includes("lists(guid'11111111-0000-0000-0000-000000000003')/items")) {
    return route.fulfill({ json: { value: [{
      Id: 7,
      Title: 'Live page',
      FileLeafRef: 'Live.aspx',
      FileRef: '/SitePages/Live.aspx',
      FileDirRef: '/SitePages',
      PromotedState: 0,
      Modified: '2026-08-01T00:00:00Z',
      UniqueId: '77777777-0000-0000-0000-000000000007',
      Editor: { Title: 'Page Editor' },
    }] } });
  }
  if (url.includes('/_api/web/lists')) {
    if (url.includes('page=2')) return route.fulfill({ json: listPage2 });
    const first = { ...listPage1 };
    first['odata.nextLink'] = `${new URL(url).origin}/_api/web/lists?page=2`;
    return route.fulfill({ json: first });
  }
  return route.fulfill({ json: { value: [] } });
});

await live.goto(WB_URL);
// The Site landing view renders first; the paged-list checks drive Lists.
await live.waitForSelector('.wb-home-cards');
await live.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
await live.locator('.wb-table tbody tr', { hasText: 'Alpha' }).first().waitFor();

// The topbar design pass shortened the live chip to just 'SP' (mock keeps
// the explicit 'SP: Mock'); the status bar carries the connection detail.
await check('live: chip reads SP from the host contract', async () =>
  (await live.locator('#wb-chip-text').textContent()) === 'SP');

await check('live: status bar names the web and user', async () =>
  (await live.locator('#wb-status-context').textContent()).includes('Stub User'));

await check('live: /_api requests send the nometadata Accept header', () =>
  seenHeaders.length > 0
  && seenHeaders.every((h) => h.includes('application/json;odata=nometadata')));

await check('live: paging links are followed across pages', async () =>
  (await live.locator('.wb-table tbody tr').count()) === 3);

await check('live: rows from the second page render', async () =>
  (await live.locator('.wb-table tbody tr', { hasText: 'Gamma' }).count()) === 1);

await check('live: page detail expands Author and Editor lookup fields', async () => {
  await live.locator('.wb-rail-btn', { hasText: 'Pages' }).click();
  await live.waitForSelector('.wb-view-pages .wb-table tbody tr', { hasText: 'Live.aspx' });
  await live.locator('.wb-view-pages .wb-table tbody tr', { hasText: 'Live.aspx' })
    .locator('td').first().click();
  await live.waitForSelector('.wb-view-pages .wb-text-rendered', { hasText: 'Live page body' });
  const expand = new URL(pageDetailUrl).searchParams.get('$expand') || '';
  return expand.split(',').includes('Author') && expand.split(',').includes('Editor');
});

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
