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

await live.close();
await browser.close();
exitWithResult();
