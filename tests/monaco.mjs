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
page.on('pageerror', (error) => pageErrors.push(error.stack || error.message));

await page.goto(APP_URL);
await page.context().grantPermissions(
  ['clipboard-read', 'clipboard-write'],
  { origin: new URL(APP_URL).origin },
);
try {
  await page.waitForSelector('.monaco-editor');
} catch (error) {
  console.log(`      startup page errors: ${pageErrors.join(' | ') || '(none)'}`);
  console.log('      Alpine module import:', await page.evaluate(() =>
    import('/src/intelligence/alpine.js')
      .then(() => 'ok', (failure) => `${failure.message}\n${failure.stack || ''}`)));
  console.log(`      startup body: ${(await page.locator('body').innerText()).slice(0, 1200)}`);
  throw error;
}

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
const jsModelText = () =>
  page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    return monaco.editor.getModel(monaco.Uri.parse('file:///dcspad/script.js')).getValue();
  });

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

await check('Alpine v3 intelligence survives catalogs saved before pack metadata', () =>
  page.evaluate(async () => {
    const { isAlpine3Runtime } = await import('/src/libraries.js');
    return isAlpine3Runtime({
      id: 'alpine',
      name: 'Alpine.js',
      js: 'https://cdn.jsdelivr.net/npm/alpinejs@3/dist/cdn.min.js',
    }) && isAlpine3Runtime({
      id: 'custom-alpine',
      name: 'Alpine local copy',
      js: 'https://cdn.example.test/alpinejs/3.14.9/cdn.min.js',
    }) && !isAlpine3Runtime({
      id: 'alpine2',
      name: 'Alpine.js v2',
      js: 'https://cdn.jsdelivr.net/npm/alpinejs@2/dist/alpine.min.js',
    });
  }));

await check('maintained PnPjs and Alpine frameworks resolve only to lib-mirror', () =>
  page.evaluate(async () => {
    const [{ getAppConfig, applyFrameworkConfig }, { PRESETS }] = await Promise.all([
      import('/src/config.js'),
      import('/src/libraries.js'),
    ]);
    const config = getAppConfig();
    const pnp = applyFrameworkConfig(PRESETS.find((entry) => entry.id === 'pnpjs2'), config);
    const alpine = applyFrameworkConfig(PRESETS.find((entry) => entry.id === 'alpine'), config);
    return pnp.js.endsWith('/lib-mirror/pnp2.bundle.js')
      && alpine.js.endsWith('/lib-mirror/alpine.js')
      && !pnp.fallbackJs
      && !alpine.fallbackJs
      && pnp.configuredSources.cdn === ''
      && alpine.configuredSources.cdn === '';
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

await page.waitForFunction(() =>
  document.documentElement.dataset.bspIntelligence === 'ready');
await page.waitForFunction(() =>
  document.documentElement.dataset.fluentIconIntelligence === 'ready');
await check('generated BMO design intelligence includes documented tokens and classes', () =>
  page.evaluate(async () => {
    const response = await fetch('/vendor/intelligence/bsp-design.json');
    const data = await response.json();
    const token = data.tokens.find((item) => item.name === '--fg-primary');
    const classItem = data.classes.find((item) => item.name === 'btn--primary');
    return data.pack === 'bsp-design'
      && data.tokens.length > 150
      && data.classes.length > 350
      && token?.value === 'var(--bmo-slate)'
      && token.description
      && classItem?.base === 'btn'
      && classItem.description;
  }));

await check('generated Fluent intelligence covers the complete SVG catalog', () =>
  page.evaluate(async () => {
    const [artifactResponse, manifestResponse] = await Promise.all([
      fetch('/vendor/intelligence/fluent-icons.json'),
      fetch('/vendor/intelligence/manifest.json'),
    ]);
    const artifact = await artifactResponse.json();
    const manifest = await manifestResponse.json();
    const home = artifact.icons.find((item) => item.idBase === 'home');
    const directional = artifact.icons.some((item) =>
      item.variants.some((variant) => /-(?:ltr|rtl)$/.test(variant)));
    const color = artifact.icons.some((item) =>
      item.variants.some((variant) => variant.includes('-color')));
    return artifact.pack === 'fluent-icons'
      && artifact.icons.length > 2600
      && manifest.fluentIcons.counts.variants === 18681
      && home?.variants.includes('24-regular')
      && directional
      && color;
  }));

await setDoc('css', 'main { color: var(--fg-p');
await page.waitForTimeout(300);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('BMO CSS completion includes documented custom properties', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => text.startsWith('--fg-primary')));
await page.keyboard.press('Escape');

await setDoc('html', '<button class="btn btn--p');
await page.waitForTimeout(300);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('BMO HTML class completion includes canonical BEM modifiers', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => text.startsWith('btn--primary')));
await page.keyboard.press('Escape');

await setDoc('html', '<button class="btn btn--primary">Save</button>');
await page.evaluate(async () => {
  const monaco = await import('/vendor/monaco/monaco.js');
  const editor = monaco.editor.getEditors()[0];
  editor.setPosition({ lineNumber: 1, column: 24 });
  editor.trigger('dcspad-test', 'editor.action.showHover', {});
});
await page.waitForSelector('.monaco-hover-content');
await check('BMO HTML class hover explains BEM composition and source', async () => {
  const hover = (await page.locator('.monaco-hover-content').allTextContents()).join(' ');
  return hover.includes('BEM modifier')
    && hover.includes('Compose with .btn')
    && hover.includes('components.css');
});
await page.keyboard.press('Escape');

await setDoc('html', '<fluent-icon name="home-24-r');
await page.waitForTimeout(250);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('Fluent custom-element completion includes real font variants', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => text.startsWith('home-24-regular')));
await page.keyboard.press('Escape');

await setDoc('html', '<svg><use href="#ic_fluent_home_24_r');
await page.waitForTimeout(250);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('Fluent sprite completion includes exact symbol ids', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => text.startsWith('ic_fluent_home_24_regular')));
await page.keyboard.press('Escape');

await setDoc('html', '<i class="icon-ic_fluent_home_24_r');
await page.waitForTimeout(250);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('Fluent font completion includes generated CSS classes', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => text.startsWith('icon-ic_fluent_home_24_regular')));
await page.keyboard.press('Escape');

await setDoc('html', '<fluent-icon name="home-24-regular"></fluent-icon>');
await page.evaluate(async () => {
  const monaco = await import('/vendor/monaco/monaco.js');
  const editor = monaco.editor.getEditors()[0];
  editor.setPosition({ lineNumber: 1, column: 27 });
  editor.trigger('dcspad-test', 'editor.action.showHover', {});
});
await page.waitForSelector('.monaco-hover-content');
await check('Fluent hover documents element, font, and sprite forms', async () => {
  const hover = (await page.locator('.monaco-hover-content').allTextContents()).join(' ');
  return hover.includes('home-24-regular')
    && hover.includes('icon-ic_fluent_home_24_regular')
    && hover.includes('ic_fluent_home_24_regular');
});
await page.keyboard.press('Escape');

await setDoc('html', '<fluent-icon name="definitely-not-a-fluent-icon"></fluent-icon>');
await page.waitForFunction(async () => {
  const monaco = await import('/vendor/monaco/monaco.js');
  const resource = monaco.Uri.parse('file:///dcspad/index.html');
  return monaco.editor.getModelMarkers({ resource })
    .some((marker) => marker.message.includes('Unknown Fluent icon token'));
});
await check('unknown Fluent tokens receive a focused Monaco warning', true);

await setDoc('html', '<fluent-icon name="home-24-regular"></fluent-icon>');
await page.waitForFunction(async () => {
  const monaco = await import('/vendor/monaco/monaco.js');
  const resource = monaco.Uri.parse('file:///dcspad/index.html');
  return !monaco.editor.getModelMarkers({ resource })
    .some((marker) => marker.message.includes('Fluent icon'));
});
await check('valid Fluent tokens do not receive custom diagnostics', true);

const alpineRow = page.locator('.lib-item', { hasText: 'Alpine.js' });
await alpineRow.locator('input[type="checkbox"]').check();
await page.waitForFunction(() =>
  document.documentElement.dataset.alpineIntelligence === 'ready');
await setDoc('js', 'Alpine.data\nconsole.log(Alpine.data);');
await page.waitForTimeout(750);
await check('Alpine global has no false JavaScript diagnostics', async () => {
  const markers = await page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    const resource = monaco.Uri.parse('file:///dcspad/script.js');
    return monaco.editor.getModelMarkers({ resource });
  });
  if (markers.length) {
    console.log(`      Alpine markers: ${markers.map((marker) => marker.message).join(' | ')}`);
  }
  return markers.length === 0;
});
await check('mirrored Alpine runtime is exactly 3.15.2', async () => {
  await setDoc('js', 'console.log("alpine-version", Alpine.version);');
  await focusEditor();
  await page.keyboard.press('Control+Enter');
  await page.waitForFunction(() =>
    document.querySelector('#console-out')?.textContent.includes('alpine-version'));
  return (await page.locator('#console-out').textContent()).includes('alpine-version 3.15.2');
});
await check('Alpine JavaScript completion includes data/store/plugin', async () => {
  for (const [prefix, name] of [['d', 'data'], ['s', 'store'], ['p', 'plugin']]) {
    await setDoc('js', `Alpine.${prefix}`);
    await page.waitForTimeout(300);
    await page.keyboard.press('Control+Space');
    await page.waitForSelector('.suggest-widget.visible');
    const suggestions = await page.locator('.suggest-widget .monaco-list-row').allTextContents();
    await page.keyboard.press('Escape');
    if (!suggestions.some((text) =>
      new RegExp(`^${name}(?:\\b|\\s|\\()`).test(text))) return false;
  }
  return true;
});

const pnpRow = page.locator('.lib-item', { hasText: 'PnPjs 2.15 (pnp2 bundle)' });
await pnpRow.locator('input[type="checkbox"]').check();
await page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'ready');
await check('mirrored PnPjs 2.15 bundle exposes pnp2 and the pnp alias', async () => {
  await setDoc('js', 'console.log("pnp-aliases", pnp === pnp2, pnp2.version);');
  await focusEditor();
  await page.keyboard.press('Control+Enter');
  await page.waitForFunction(() =>
    document.querySelector('#console-out')?.textContent.includes('pnp-aliases'));
  return (await page.locator('#console-out').textContent()).includes('pnp-aliases true v2');
});
await check('PnPjs fluent completion includes web for pnp and pnp2', async () => {
  for (const globalName of ['pnp', 'pnp2']) {
    await setDoc('js', `${globalName}.sp.w`);
    await page.waitForTimeout(500);
    await page.keyboard.press('Control+Space');
    await page.waitForSelector('.suggest-widget.visible');
    const suggestions = await page.locator('.suggest-widget .monaco-list-row').allTextContents();
    await page.keyboard.press('Escape');
    if (!suggestions.some((text) => /^web(?:\b|\s|\()/.test(text))) return false;
  }
  return true;
});
await pnpRow.locator('input[type="checkbox"]').uncheck();
await check('PnPjs declarations unload with the runtime library', () =>
  page.waitForFunction(() => document.documentElement.dataset.pnpTypes === 'disabled')
    .then(() => true, () => false));

await check('disabling PnPjs does not erase Alpine declarations', () =>
  page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    const libs = monaco.typescript.javascriptDefaults.getExtraLibs();
    const paths = Object.keys(libs);
    return paths.some((path) => path.includes('@types/dcspad-alpine'))
      && !paths.some((path) => path.includes('@pnp/'));
  }));

await setDoc('html', '<div x-d');
await page.waitForTimeout(500);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('Alpine HTML completion includes core directives', async () =>
  (await page.locator('.suggest-widget .monaco-list-row').allTextContents())
    .some((text) => /^x-data(?:\b|\s|=)/.test(text)));
await page.keyboard.press('Escape');

await setDoc('html', '<button x-data @click="$');
await page.waitForTimeout(250);
await page.keyboard.press('Control+Space');
await page.waitForSelector('.suggest-widget.visible');
await check('Alpine HTML expressions include magic properties', async () => {
  const suggestions = await page.locator('.suggest-widget .monaco-list-row').allTextContents();
  return ['$dispatch', '$refs', '$store'].every((name) =>
    suggestions.some((text) => text.startsWith(name)));
});
await page.keyboard.press('Escape');

await alpineRow.locator('input[type="checkbox"]').uncheck();
await page.waitForFunction(() =>
  document.documentElement.dataset.alpineIntelligence === 'disabled');
await check('Alpine declarations unload with the runtime library', () =>
  page.evaluate(async () => {
    const monaco = await import('/vendor/monaco/monaco.js');
    const paths = Object.keys(
      monaco.typescript.javascriptDefaults.getExtraLibs(),
    );
    return !paths.some((path) => path.includes('@types/dcspad-alpine'))
      && document.documentElement.dataset.bspIntelligence === 'ready';
  }));

await setDoc('js', 'SNIPPET_UNDO_MARKER();');
await page.click('#btn-snippet-add');
await page.fill('#snippet-name-input', 'undo-boundary');
await page.click('#snippet-name-save');
await setDoc('js', 'const KEEP_BEFORE_SNIPPET = true;');
await focusEditor();
await page.keyboard.press('Control+End');
await page.keyboard.type('\nconst TYPED_BEFORE_SNIPPET = true;');
await page.locator('#snippet-list .snippet-item', { hasText: 'undo-boundary' }).click();
await page.waitForFunction(() =>
  document.querySelector('#pane-editor .view-lines')?.textContent.includes('SNIPPET_UNDO_MARKER'));
await focusEditor();
await page.keyboard.press('Control+z');
await check('snippet insertion is one isolated undo action', async () => {
  const text = await jsModelText();
  const ok = text.includes('KEEP_BEFORE_SNIPPET')
    && text.includes('TYPED_BEFORE_SNIPPET')
    && !text.includes('SNIPPET_UNDO_MARKER');
  if (!ok) console.log(`      JS after undo: ${JSON.stringify(text)}`);
  return ok;
});
await page.keyboard.press('Control+y');
await check('isolated snippet insertion can be redone', () =>
  page.waitForFunction(() =>
    document.querySelector('#pane-editor .view-lines')?.textContent.includes('SNIPPET_UNDO_MARKER'))
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
  const unexpected = pageErrors.filter((message) => !message.startsWith('Canceled'));
  if (unexpected.length) console.log(`      page errors: ${unexpected.join(' | ')}`);
  return unexpected.length === 0;
});

const blocked = await browser.newPage({ viewport: { width: 1200, height: 760 } });
await blocked.route('**/ts.worker.js*', (route) => route.abort());
await blocked.goto(APP_URL);
await blocked.waitForSelector('.monaco-editor');
await blocked.waitForFunction(() =>
    document.documentElement.dataset.monacoWorkerError === 'javascript'
    || document.documentElement.dataset.monacoWorkerError === 'typescript');
await blocked.click('#btn-run');
await blocked.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('ran in'));
await check('worker warning persists after Run without losing the editor', () =>
  blocked.evaluate(() =>
    document.querySelectorAll('.monaco-editor').length === 1
    && !!document.documentElement.dataset.monacoWorkerError
    && document.querySelector('#status-run')?.textContent.includes('ran in')));
await blocked.close();

await browser.close();
exitWithResult();
