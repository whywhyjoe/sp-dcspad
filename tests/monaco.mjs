// Monaco-specific integration suite: model switching, keyboard action,
// diagnostics, PnPjs 2.15.0 completion, asset routing, and worker failure.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
const assetRequests = [];
const pageErrors = [];
page.on('request', (request) => {
  if (request.url().includes('/vendor/monaco/')) assetRequests.push(request.url());
});
page.on('pageerror', (error) => pageErrors.push(error.message));

await page.goto(APP_URL);
await page.context().grantPermissions(
  ['clipboard-read', 'clipboard-write'],
  { origin: new URL(APP_URL).origin },
);
await page.waitForSelector('.monaco-editor');

const focusEditor = () =>
  page.locator('#pane-editor .view-lines').click({ position: { x: 80, y: 10 } });

const setDoc = async (name, code) => {
  await page.click(`#editor-tabs .tab[data-editor="${name}"]`);
  await focusEditor();
  await page.keyboard.press('Control+a');
  await page.evaluate((text) => navigator.clipboard.writeText(text), code);
  await page.keyboard.press('Control+v');
};

const visibleEditorText = async () =>
  (await page.locator('#pane-editor .view-lines').textContent()).replaceAll('\u00a0', ' ');

await check('one Monaco editor owns three typed models', async () => {
  const result = await page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    return monaco.editor.getModels().map((model) => ({
      uri: model.uri.toString(),
      language: model.getLanguageId(),
    }));
  });
  return JSON.stringify(result) === JSON.stringify([
    { uri: 'file:///dcspad/index.html', language: 'html' },
    { uri: 'file:///dcspad/styles.css', language: 'css' },
    { uri: 'file:///dcspad/script.js', language: 'javascript' },
  ]);
});

await check('PnPjs 2.15 runtime detection survives custom catalog names', () =>
  page.evaluate(async () => {
    const { isPnpjs215Runtime } = await import('/src/libraries.js');
    return isPnpjs215Runtime({
      name: 'Pnpjs JSD',
      js: 'https://cdn.jsdelivr.net/npm/@pnp/pnpjs@2.15.0/dist/pnp.min.js',
    }) && !isPnpjs215Runtime({
      name: 'PnPjs current',
      js: 'https://cdn.jsdelivr.net/npm/@pnp/pnpjs@4.13.0/dist/pnp.min.js',
    });
  }));

await setDoc('html', '<main>HTML_MODEL_MARKER</main>');
await setDoc('css', 'body { color: papayawhip; } /* CSS_MODEL_MARKER */');
await setDoc('js', 'console.log("JS_MODEL_MARKER");');
await page.click('#editor-tabs .tab[data-editor="html"]');
await check('HTML model survives tab swaps', (await visibleEditorText()).includes('HTML_MODEL_MARKER'));
await page.click('#editor-tabs .tab[data-editor="css"]');
await check('CSS model survives tab swaps', (await visibleEditorText()).includes('CSS_MODEL_MARKER'));
await page.click('#editor-tabs .tab[data-editor="js"]');
await check('JS model survives tab swaps', (await visibleEditorText()).includes('JS_MODEL_MARKER'));

await focusEditor();
await page.keyboard.press('Control+Enter');
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('JS_MODEL_MARKER'));
await check('Ctrl/Cmd+Enter runs from Monaco', true);

await setDoc('js', 'console.log(new Date().toISOString());');
await page.waitForTimeout(750);
await check('standard browser globals have no false diagnostics', async () => {
  const markers = await page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    const resource = monaco.Uri.parse('file:///dcspad/script.js');
    return monaco.editor.getModelMarkers({ resource });
  });
  return markers.length === 0;
});

await setDoc('js', 'const count = "not a number";\ncount.toFixed(2);');
await check('JavaScript semantic diagnostics render', () =>
  page.waitForFunction(() => document.querySelectorAll('.squiggly-error').length > 0)
    .then(() => true, () => false));

const pnpRow = page.locator('.lib-item', { hasText: 'PnPjs v2 (classic)' });
await pnpRow.locator('input[type="checkbox"]').check();
await page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'ready');
await setDoc('js', 'pnp.sp.w');
await page.waitForTimeout(500);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('PnPjs fluent completion includes web', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => /^web(?:\b|\s|\()/.test(text)));
await page.keyboard.press('Escape');
await pnpRow.locator('input[type="checkbox"]').uncheck();
await check('PnPjs declarations unload with the runtime library', () =>
  page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'disabled')
    .then(() => true, () => false));

await page.click('#editor-tabs .tab[data-editor="html"]');
await page.waitForTimeout(300);
await page.click('#editor-tabs .tab[data-editor="css"]');
await page.waitForTimeout(300);

await check('Monaco assets are same-origin .js/CSS/font files', () => {
  const origin = new URL(APP_URL).origin;
  return assetRequests.length > 0
    && assetRequests.every((url) => url.startsWith(origin) && !url.startsWith('blob:'))
    && assetRequests.some((url) => /editor\.worker\.js/.test(url))
    && assetRequests.some((url) => /ts\.worker\.js/.test(url))
    && assetRequests.some((url) => /html\.worker\.js/.test(url))
    && assetRequests.some((url) => /css\.worker\.js/.test(url));
});
await check('Monaco integration produces no uncaught page errors', () => {
  // Monaco restarts language workers when extra libraries change and reports
  // the deliberately-abandoned requests as "Canceled" promises.
  const unexpected = pageErrors.filter((message) => message !== 'Canceled');
  if (unexpected.length) console.log(`      page errors: ${unexpected.join(' | ')}`);
  return unexpected.length === 0;
});

const blocked = await browser.newPage({ viewport: { width: 1200, height: 760 } });
await blocked.route('**/ts.worker.js*', (route) => route.abort());
await blocked.goto(APP_URL);
await blocked.waitForSelector('.monaco-editor');
await check('blocked language worker is surfaced without losing the editor', () =>
  blocked.waitForFunction(() =>
    (document.documentElement.dataset.monacoWorkerError === 'javascript'
    || document.documentElement.dataset.monacoWorkerError === 'typescript')
    && document.querySelector('#status-run')?.textContent.includes('worker unavailable'))
    .then(() => true, () => false));
await blocked.close();

await browser.close();
exitWithResult();
