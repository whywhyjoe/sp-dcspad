// Runtime config suite: relative asset URL resolution, explicit intelligence
// metadata, and ordered local-to-CDN framework fallback.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const origin = new URL(APP_URL).origin;
const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1400, height: 820 } });
const primaryUrl = `${origin}/tests/fixtures/missing-local-framework.js`;
const fallbackUrl = `${origin}/tests/fixtures/testlib.js`;

await page.route('**/dcspad.config.json*', (route) => route.fulfill({
  contentType: 'application/json',
  body: JSON.stringify({
    version: 1,
    frameworks: {
      prefer: 'local',
      fallbackToCdn: true,
      items: {
        pnpjs2: {
          localUrl: `${origin}/custom/pnp-rollup.js`,
          cdnUrl: '',
          probeGlobal: 'pnp',
          intelligence: ['pnpjs-2.15.0'],
        },
        alpine: {
          localUrl: primaryUrl,
          cdnUrl: fallbackUrl,
          probeGlobal: 'testlib',
          intelligence: ['alpine-3'],
        },
      },
    },
    assets: {
      designSystem: {
        prefer: 'local',
        localBaseUrl: './bsp-design-system/',
        hostedBaseUrl: '',
        intelligence: ['bsp-design'],
        files: { components: 'components.css' },
      },
    },
  }),
}));

await page.goto(APP_URL);
await page.waitForSelector('.monaco-editor');

await check('relative configured asset folders resolve from dcspad.config.json', () =>
  page.evaluate(async (expected) => {
    const { getAppConfig } = await import('/src/config.js');
    return getAppConfig().assets.designSystem.localBaseUrl === expected;
  }, `${origin}/bsp-design-system/`));

await check('asset intelligence packs activate independently of framework checkboxes', () =>
  page.evaluate(async () => {
    const { getEnabledIntelligence } = await import('/src/libraries.js');
    return getEnabledIntelligence().includes('bsp-design');
  }));

const pnpRow = page.locator('.lib-item', { hasText: 'PnPjs v2 (classic)' });
await pnpRow.locator('input[type="checkbox"]').check();
await check('explicit PnP intelligence survives an opaque custom rollup URL', () =>
  page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'ready')
    .then(() => true, () => false));
await pnpRow.locator('input[type="checkbox"]').uncheck();

const alpineRow = page.locator('.lib-item', { hasText: 'Alpine.js' });
await check('framework tooltip shows primary and fallback sources', async () => {
  const title = await alpineRow.locator('.lib-name').getAttribute('title');
  return title?.includes(primaryUrl) && title.includes(`Fallback: ${fallbackUrl}`);
});
await alpineRow.locator('input[type="checkbox"]').check();

await page.evaluate(async () => {
  const monaco = await import('/vendor/monaco/monaco.js');
  monaco.editor.getModel(monaco.Uri.parse('file:///dcspad/script.js'))
    .setValue('console.log("configured-fallback", testlib.hello());');
});
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('configured-fallback'));

await check('missing local framework falls back before user JavaScript', async () =>
  (await page.locator('#console-out').textContent()).includes('configured-fallback hi'));

await check('fallback remains parser-ordered at the original catalog position', () =>
  page.evaluate(({ primary, fallback }) => {
    const frame = document.querySelector('#preview-host iframe');
    const scripts = [...frame.contentDocument.querySelectorAll('script[src]')]
      .map((script) => script.src);
    return JSON.stringify(scripts) === JSON.stringify([primary, fallback]);
  }, { primary: primaryUrl, fallback: fallbackUrl }));

await check('successful fallback is explained in the captured console', async () =>
  (await page.locator('#console-out').textContent())
    .includes('loading configured fallback'));

await browser.close();
exitWithResult();
