// Exercises the exact boot.js + dcspad.app.js path used by the SharePoint
// Modern Script Editor, without requiring a tenant for every regression run.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const origin = new URL(APP_URL).origin;
const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
const requests = [];
const pageErrors = [];

page.on('request', (request) => requests.push(request.url()));
page.on('pageerror', (error) => pageErrors.push(error.message));

await page.goto(`${origin}/tests/hosted-fixture.html`);
await page.waitForSelector('#dcspad-mount .monaco-editor');

await check('hosted boot mounts one Monaco workbench', async () =>
  (await page.locator('#dcspad-mount .app').count()) === 1
  && (await page.locator('#dcspad-mount .monaco-editor').count()) === 1);

await check('hosted mode flag is applied', () =>
  page.evaluate(() => document.documentElement.classList.contains('dcspad-hosted')));

await check('hosted runtime uses the boot script folder as its asset base', () =>
  page.evaluate((expected) => window.__DCSPAD_ASSET_BASE__ === expected, `${origin}/`));

await page.waitForFunction(() => document.documentElement.dataset.monacoReady === 'true');
await check('hosted app and Monaco runtime URLs are versioned', () => {
  const app = requests.find((url) => url.includes('/dcspad.app.js?'));
  const monaco = requests.find((url) => url.includes('/vendor/monaco/monaco.js?'));
  return !!app && !!monaco
    && new URL(app).searchParams.has('v')
    && new URL(monaco).searchParams.has('v');
});

const pnpRow = page.locator('.lib-item', { hasText: 'PnPjs v2 (classic)' });
await pnpRow.locator('input[type="checkbox"]').check();
await page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'ready');
await check('hosted PnPjs declarations use the versioned vendor URL', () => {
  const types = requests.find((url) => url.includes('/vendor/monaco/pnpjs-types.json?'));
  return !!types && new URL(types).searchParams.has('v');
});

await page.waitForFunction(() =>
  performance.getEntriesByType('resource')
    .some((entry) => entry.name.includes('/vendor/monaco/ts.worker.js?')));
await check('hosted TypeScript worker is same-origin and versioned', () => {
  const worker = requests.find((url) => url.includes('/vendor/monaco/ts.worker.js?'));
  return !!worker
    && new URL(worker).origin === origin
    && new URL(worker).searchParams.has('v');
});

await check('hosted bundle produces no uncaught page errors',
  pageErrors.filter((message) => message !== 'Canceled').length === 0);

await browser.close();
exitWithResult();
