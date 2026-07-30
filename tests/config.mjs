// Runtime config suite: relative asset URL resolution, explicit intelligence
// metadata, and ordered local-to-CDN framework fallback.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const origin = new URL(APP_URL).origin;
const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1400, height: 820 } });
const primaryUrl = `${origin}/tests/fixtures/missing-local-framework.js`;
const fallbackUrl = `${origin}/tests/fixtures/testlib.js`;
let notesFetches = 0;

await page.route('**/docs/design-reference.html', (route) => route.fulfill({
  contentType: 'text/plain',
  body: '<!doctype html><html><head><link rel="stylesheet" href="./docs.css"></head><body><h1>Design reference</h1><a href="guide.md">Guide</a><script src="./page-script.js"></script><script>document.body.dataset.scriptReady = "yes";</script></body></html>',
}));
await page.route('**/docs/page-script.js', (route) => route.fulfill({
  contentType: 'application/octet-stream',
  headers: {
    'Content-Disposition': 'attachment; filename="page-script.js"',
    'X-Content-Type-Options': 'nosniff',
  },
  body: 'document.body.dataset.externalScriptReady = "yes";',
}));
await page.route('**/docs/guide.md', (route) => route.fulfill({
  contentType: 'text/plain',
  body: '# Authoring guide\n\nUse **semantic HTML**.\n\n```js\nconsole.log("docs");\n```\n\n| Item | Value |\n| --- | --- |\n| Mode | Markdown |',
}));
await page.route('**/docs/notes.txt', (route) => route.fulfill({
  contentType: 'text/plain',
  body: `Design tokens\n=============\nUse semantic aliases.\nRefresh ${++notesFetches}`,
}));
await page.route('**/docs/history-*.txt', (route) => route.fulfill({
  contentType: 'text/plain',
  body: `History ${new URL(route.request().url()).pathname}`,
}));
await page.route('**/docs/browser.css', (route) => route.fulfill({
  contentType: 'text/css',
  body: '.browser-marker { color: teal; }',
}));
await page.route('**/docs/browser.js', (route) => route.fulfill({
  contentType: 'text/javascript',
  body: 'window.browserCodeExecuted = true;\nconst browserMarker = "js";',
}));
await page.route('**/docs/browser.json', (route) => route.fulfill({
  contentType: 'application/json',
  body: '{"browserMarker":"json"}',
}));
await page.route('**/docs/browser.csv', (route) => route.fulfill({
  contentType: 'text/csv',
  body: 'name,value\nbrowserMarker,csv',
}));
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
    docs: [
      {
        id: 'design-reference',
        title: 'Design reference',
        url: './docs/design-reference.html',
        type: 'html',
      },
      {
        id: 'authoring-guide',
        title: 'Authoring guide',
        url: './docs/guide.md',
        type: 'markdown',
      },
      {
        id: 'plain-notes',
        title: 'Plain notes',
        url: './docs/notes.txt',
        type: 'txt',
      },
      {
        id: 'token-data',
        title: 'Token data',
        url: './docs/browser.json',
        type: 'json',
      },
    ],
    copilot: {
      enabled: true,
      url: 'https://m365.cloud.microsoft/chat',
    },
    assets: {
      designSystem: {
        prefer: 'local',
        localBaseUrl: './bsp-design-system/',
        hostedBaseUrl: '',
        intelligence: ['bsp-design'],
        files: { components: 'components.css' },
      },
      fluentIcons: {
        prefer: 'local',
        localBaseUrl: './bsp-fluent-icon-lib/',
        hostedBaseUrl: '',
        intelligence: ['fluent-icons'],
        files: {
          regularCss: 'fonts/FluentSystemIcons-Regular.css',
          filledCss: 'fonts/FluentSystemIcons-Filled.css',
          lightCss: 'fonts/FluentSystemIcons-Light.css',
        },
        runtime: {
          enabled: true,
          cssFiles: ['regularCss', 'filledCss', 'lightCss'],
          fluentIconElement: true,
        },
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

await check('Browser bookmarks and Copilot URLs normalize from dcspad.config.json', () =>
  page.evaluate(async (expected) => {
    const { getAppConfig } = await import('/src/config.js');
    const config = getAppConfig();
    return config.docs.length === 4
      && config.docs[0].url === expected.doc
      && config.docs[1].type === 'markdown'
      && config.docs[2].type === 'text'
      && config.docs[3].type === 'text'
      && config.copilot.enabled
      && config.copilot.url === 'https://m365.cloud.microsoft/chat';
  }, { doc: `${origin}/docs/design-reference.html` }));

await page.click('#btn-docs');
await page.click('#docs-menu-items [data-doc-id="design-reference"]');
await page.waitForFunction(() =>
  document.getElementById('docs-frame')?.srcdoc.includes('Design reference'));
await page.frameLocator('#docs-frame').locator('h1').waitFor();
await page.frameLocator('#docs-frame').locator('body[data-script-ready="yes"]').waitFor();
await page.frameLocator('#docs-frame').locator('body[data-external-script-ready="yes"]').waitFor();
await check('HTML docs are fetched as text and rendered through srcdoc with a base URL', () =>
  Promise.all([
    page.evaluate((expected) => {
      const frame = document.getElementById('docs-frame');
      return !frame.hidden && frame.srcdoc.includes(`<base href="${expected}">`);
    }, `${origin}/docs/design-reference.html`),
    page.frameLocator('#docs-frame').locator('h1').textContent()
      .then((text) => text === 'Design reference'),
    page.frameLocator('#docs-frame').locator('body').getAttribute('data-script-ready')
      .then((value) => value === 'yes'),
    page.frameLocator('#docs-frame').locator('body').getAttribute('data-external-script-ready')
      .then((value) => value === 'yes'),
  ]).then((values) => values.every(Boolean)));

await page.frameLocator('#docs-frame').getByRole('link', { name: 'Guide' }).click();
await page.frameLocator('#docs-frame').locator('h1', { hasText: 'Authoring guide' }).waitFor();
await check('Browser links reuse the fetch-and-srcdoc loader instead of navigating to a download', () =>
  Promise.all([
    page.locator('#browser-address-input').inputValue()
      .then((value) => value === `${origin}/docs/guide.md`),
    page.evaluate(() =>
      document.getElementById('docs-frame').srcdoc.includes('Authoring guide')),
  ]).then((values) => values.every(Boolean)));
await check('Markdown docs render headings, code fences, and tables in the same viewer', () =>
  Promise.all([
    page.frameLocator('#docs-frame').locator('strong').textContent()
      .then((text) => text === 'semantic HTML'),
    page.frameLocator('#docs-frame').locator('pre code').textContent()
      .then((text) => text.includes('console.log')),
    page.frameLocator('#docs-frame').locator('table tbody td').first().textContent()
      .then((text) => text === 'Mode'),
    page.locator('#docs-frame').getAttribute('sandbox')
      .then((value) => !value.includes('allow-scripts')),
  ]).then((values) => values.every(Boolean)));

await page.fill('#browser-address-input', `${origin}/docs/notes.txt`);
await page.keyboard.press('Enter');
await page.frameLocator('#docs-frame').locator('pre', { hasText: 'Design tokens' }).waitFor();
await check('plain-text resources render safely in the Browser pane', () =>
  page.frameLocator('#docs-frame').locator('pre').textContent()
    .then((text) => text.includes('Use semantic aliases.')));
await check('Browser history records unique URLs most-recent-first', () =>
  page.evaluate(async (expected) => {
    const { getState } = await import('/src/state.js');
    const history = getState().settings.browserHistory;
    return history.length === 3
      && history[0] === expected.notes
      && history[1] === expected.guide
      && history[2] === expected.design;
  }, {
    notes: `${origin}/docs/notes.txt`,
    guide: `${origin}/docs/guide.md`,
    design: `${origin}/docs/design-reference.html`,
  }));

await page.click('#browser-refresh');
await page.frameLocator('#docs-frame').locator('pre', { hasText: 'Refresh 2' }).waitFor();
await check('refresh bypasses the Browser cache without duplicating history', () =>
  page.evaluate(async (expected) => {
    const { getState } = await import('/src/state.js');
    const history = getState().settings.browserHistory;
    return history.length === 3 && history[0] === expected;
  }, `${origin}/docs/notes.txt`).then((result) => result && notesFetches === 2));

for (let index = 0; index < 11; index += 1) {
  const url = `${origin}/docs/history-${index}.txt`;
  await page.fill('#browser-address-input', url);
  await page.keyboard.press('Enter');
  await page.frameLocator('#docs-frame').locator('pre', { hasText: `/docs/history-${index}.txt` }).waitFor();
}
await check('Browser history persists only the last 10 URLs', () =>
  page.evaluate(async (expected) => {
    const { getState, saveNow } = await import('/src/state.js');
    saveNow();
    const history = getState().settings.browserHistory;
    const stored = JSON.parse(localStorage.getItem('dcspad.v2.workspace')).settings.browserHistory;
    return history.length === 10
      && history[0] === expected.first
      && history[9] === expected.last
      && stored.length === 10
      && document.querySelectorAll('#browser-history option').length === 11;
  }, {
    first: `${origin}/docs/history-10.txt`,
    last: `${origin}/docs/history-1.txt`,
  }));

await page.selectOption('#browser-history', `${origin}/docs/history-5.txt`);
await page.frameLocator('#docs-frame').locator('pre', { hasText: '/docs/history-5.txt' }).waitFor();
await check('history dropdown reopens a selected recent URL', () =>
  page.locator('#browser-address-input').inputValue()
    .then((value) => value === `${origin}/docs/history-5.txt`));

for (const resource of [
  { url: `${origin}/docs/browser.css`, marker: '.browser-marker { color: teal; }' },
  { url: `${origin}/docs/browser.js`, marker: 'window.browserCodeExecuted = true;' },
  { url: `${origin}/docs/browser.json`, marker: '{"browserMarker":"json"}' },
  { url: `${origin}/docs/browser.csv`, marker: 'browserMarker,csv' },
]) {
  await page.fill('#browser-address-input', resource.url);
  await page.keyboard.press('Enter');
  await page.frameLocator('#docs-frame').locator('pre', { hasText: resource.marker }).waitFor();
}
await check('CSS, JavaScript, JSON, and CSV resources render as safe source text', () =>
  Promise.all([
    page.frameLocator('#docs-frame').locator('pre').textContent()
      .then((text) => text.includes('browserMarker,csv')),
    page.locator('#docs-frame').getAttribute('sandbox')
      .then((value) => !value.includes('allow-scripts')),
  ]).then((values) => values.every(Boolean)));

await page.fill('#browser-address-input', 'https://example.com/guide.html');
await page.keyboard.press('Enter');
await check('pasted URLs outside the current SharePoint tenant are rejected', async () =>
  (await page.locator('#docs-state').textContent()).includes(`limited to ${origin}`));

await page.fill('#browser-address-input', `${origin}/docs/guide.md`);
await page.keyboard.press('Enter');
await page.frameLocator('#docs-frame').locator('h1', { hasText: 'Authoring guide' }).waitFor();
await page.click('#btn-max-docs');
await check('Browser can temporarily maximize over the development panes', async () => {
  const box = await page.locator('#extras-docs').boundingBox();
  return (await page.locator('#main').evaluate((element) =>
    element.classList.contains('max-docs')))
    && box.width > 1300;
});
await page.keyboard.press('Escape');
await page.click('#extras-tabs [data-extra="resources"]');
await check('Resources and Browser switch without disturbing the editor/runtime panes', async () =>
  (await page.locator('#panel-frameworks').isVisible())
  && (await page.locator('#extras-docs').isHidden())
  && (await page.locator('#editors').isVisible())
  && (await page.locator('#runtime').isVisible()));

await page.evaluate(() => {
  window.__docsOpened = null;
  window.open = (url, target) => {
    window.__docsOpened = { url, target };
    return { focus() {} };
  };
});
await page.click('#btn-copilot');
await check('approved AI help opens in one reusable external tab rather than an iframe', () =>
  page.evaluate(() =>
    window.__docsOpened?.url === 'https://m365.cloud.microsoft/chat'
    && window.__docsOpened?.target === 'dcspad-copilot'
    && !document.querySelector('iframe[src*="m365.cloud.microsoft"]')));

await check('asset intelligence packs activate independently of framework checkboxes', () =>
  page.evaluate(async () => {
    const { getEnabledIntelligence } = await import('/src/libraries.js');
    const packs = getEnabledIntelligence();
    return packs.includes('bsp-design') && packs.includes('fluent-icons');
  }));

await check('configured Fluent runtime resolves all local CSS and bridge URLs', () =>
  page.evaluate(async (expected) => {
    const { getEnabledLibraries } = await import('/src/libraries.js');
    const runtime = getEnabledLibraries()
      .find((entry) => entry.name === 'fluentIcons configured assets');
    return JSON.stringify(runtime?.css) === JSON.stringify(expected.css)
      && runtime?.js === expected.js
      && runtime?.cssText.includes('FluentSystemIcons-Regular');
  }, {
    css: [
      `${origin}/bsp-fluent-icon-lib/fonts/FluentSystemIcons-Regular.css`,
      `${origin}/bsp-fluent-icon-lib/fonts/FluentSystemIcons-Filled.css`,
      `${origin}/bsp-fluent-icon-lib/fonts/FluentSystemIcons-Light.css`,
    ],
    js: `${origin}/src/bridge/fluent-icon-font.js`,
  }));

const pnpRow = page.locator('.lib-item', { hasText: 'PnPjs 2.15 (pnp2 bundle)' });
await pnpRow.locator('input[type="checkbox"]').check();
await check('explicit PnP intelligence survives an opaque custom rollup URL', () =>
  page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'ready')
    .then(() => true, () => false));
await pnpRow.locator('input[type="checkbox"]').uncheck();

const alpineRow = page.locator('.lib-item', { hasText: 'Alpine.js 3.15.2' });
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
      .map((script) => script.src)
      .filter((src) => src === primary || src === fallback);
    return JSON.stringify(scripts) === JSON.stringify([primary, fallback]);
  }, { primary: primaryUrl, fallback: fallbackUrl }));

await check('successful fallback is explained in the captured console', async () =>
  (await page.locator('#console-out').textContent())
    .includes('loading configured fallback'));

await browser.close();
exitWithResult();
