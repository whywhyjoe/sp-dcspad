// DCSPad smoke suite: execution model, console capture, SP inspector,
// network monitor, REPL, run isolation, rerun lifecycle (cross-run
// network history, cancelled requests/evals), library injection, autosave.
// Setup: see tests/README.md (app server on 8642, fixtures on 8643).

import { readFileSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { launchBrowser, check, exitWithResult, APP_URL, FIXTURES_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
const starterSnippetDoc = JSON.parse(readFileSync(
  join('..', 'examples', 'dcspad-starter-snippets.json'),
  'utf8',
));

// The locally installed browser may use a persistent profile. Clear this
// origin before the app boots so "new project" and default-library checks are
// deterministic without clearing again on the persistence reload near EOF.
await page.goto(new URL('/__dcspad-smoke-reset__', APP_URL).href);
await page.evaluate(() => localStorage.clear());
await page.goto(APP_URL);
await page.waitForTimeout(1200);
await page.context().grantPermissions(
  ['clipboard-read', 'clipboard-write'],
  { origin: new URL(APP_URL).origin },
);

await check('Monaco editor renders', async () =>
  (await page.locator('.monaco-editor').count()) === 1
  && await page.evaluate(() => document.documentElement.dataset.monacoReady === 'true'));

await check('SP chip shows Mock', async () =>
  (await page.locator('#sp-chip-text').textContent()) === 'SP: Mock');

await check('new projects start with a labeled untitled name control', async () =>
  (await page.locator('.project-name__label').textContent()) === 'PRJ'
  && (await page.locator('#project-name-text').textContent()) === '(untitled)');
await check('a fresh project is visibly unsaved to a project file', async () =>
  (await page.locator('#project-file-state').textContent()) === 'unsaved'
  && (await page.locator('#project-file-state').getAttribute('title')).includes('browser autosave'));

const setDoc = async (name, code) => {
  await page.click(`#editor-tabs .tab[data-editor="${name}"]`);
  await page.locator('#pane-editor .view-lines').click({ position: { x: 80, y: 10 } });
  await page.keyboard.press('Control+a');
  await page.evaluate((text) => navigator.clipboard.writeText(text), code);
  await page.keyboard.press('Control+v');
};
const setJs = (code) => setDoc('js', code);

await setDoc('html', `
<button id="fluent-button" aria-label="Home">
  <fluent-icon name="home-24-regular"></fluent-icon>
</button>
<i id="fluent-filled" class="icon-ic_fluent_home_24_filled" aria-hidden="true"></i>
`);
await setJs('');
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('ran in'));
await check('disabled configured Fluent bridge is not injected into the preview', () =>
  page.evaluate(() => {
    const frame = document.querySelector('#preview-host iframe');
    const previewDocument = frame.contentDocument;
    const links = [...previewDocument.querySelectorAll('link[rel="stylesheet"]')]
      .map((link) => new URL(link.href).pathname);
    const icon = previewDocument.querySelector('fluent-icon');
    return !links.some((path) => path.includes('/bsp-fluent-icon-lib/fonts/'))
      && icon.childElementCount === 0;
  }));

await setJs(`
window.counter = (window.counter || 0) + 1;
console.log("counter", window.counter);
console.warn("a warning");
console.info("some info");
console.group("my group");
console.log("inside group");
console.groupEnd();
console.table([{a:1,b:2},{a:3,b:4}]);
console.log({ d: { results: [
  { __metadata: { type: "SP.List" }, Title: "Tasks", Id: "abc-123", EntityTypeName: "TasksList", BaseTemplate: 100, ItemCount: 42 }
] } });
fetch("/_api/web").catch(()=>{});
var x = new XMLHttpRequest(); x.open("GET", "/index.html"); x.send();
Promise.reject(new Error("boom-rejection"));
setTimeout(() => { throw new Error("boom-throw"); }, 50);
`);

await page.click('#btn-run');
await page.waitForTimeout(1500);

await check('preview iframe created', async () =>
  (await page.locator('#preview-host iframe').count()) === 1);

const consoleText = await page.locator('#console-out').textContent();
await check('console.log captured', consoleText.includes('counter'));
await check('console.warn captured', consoleText.includes('a warning'));
await check('group captured', consoleText.includes('my group') && consoleText.includes('inside group'));
await check('console.table rendered', async () =>
  (await page.locator('#console-out .console-table').count()) >= 1);
await check('SP inspector badge for d.results', async () =>
  (await page.locator('#console-out .sp-badge').count()) >= 1);
await check('SP entity header shows List fields', async () =>
  (await page.locator('#console-out .sp-entity-head').first().textContent()).includes('SP.List'));
await check('rejection captured', consoleText.includes('boom-rejection'));
await check('thrown error captured', consoleText.includes('boom-throw'));
await check('status bar shows ran-in', async () =>
  (await page.locator('#status-run').textContent()).includes('ran in'));

// network panel
await page.click('#diag-tabs .tab[data-diag="network"]');
await page.waitForTimeout(200);
await check('network rows captured (fetch + xhr)', async () =>
  (await page.locator('#network-rows .network-row').count()) >= 2);
const netText = await page.locator('#network-rows').textContent();
await check('fetch to _api listed', netText.includes('/_api/web'));
await check('xhr listed', netText.includes('/index.html'));

await page.locator('#network-rows .network-row').nth(1).click();
await check('network detail opens', async () =>
  !(await page.locator('#network-detail').isHidden()));

// run isolation
await page.click('#diag-tabs .tab[data-diag="console"]');
await page.click('#btn-run');
await page.waitForTimeout(1200);
const consoleText2 = await page.locator('#console-out').textContent();
await check('iframe state fully reset (counter still 1)',
  consoleText2.includes('counter 1') && !consoleText2.includes('counter 2'));

// REPL
await page.fill('#console-input', 'window.counter + 100');
await page.press('#console-input', 'Enter');
await page.waitForTimeout(400);
await check('REPL evaluates in frame', async () =>
  (await page.locator('#console-out').textContent()).includes('101'));

await page.fill('#console-input', 'Promise.resolve({hello:"world"})');
await page.press('#console-input', 'Enter');
await page.waitForTimeout(400);
await check('REPL awaits promises', async () =>
  (await page.locator('#console-out').textContent()).includes('awaited'));

await check('_spPageContextInfo injected into iframe', async () =>
  await page.evaluate(() => {
    const f = document.querySelector('#preview-host iframe');
    return !!f.contentWindow._spPageContextInfo?.webAbsoluteUrl;
  }));

// --- rerun lifecycle ---
// Regressions covered: per-frame net ids reused across runs overwrote the
// panel's cross-run history; requests still in flight when the frame was
// replaced stayed "pending" forever; REPL evals awaiting a promise at
// rerun never settled. Uses waitForFunction, not fixed sleeps.

// Route a URL into limbo so its request is still pending at the next run.
await page.route('**/hang-forever*', () => { /* never fulfil */ });

const waitForNetRow = (marker) => page.waitForFunction((m) =>
  [...document.querySelectorAll('#network-rows .network-row')].some((r) =>
    r.textContent.includes(m) && !r.querySelector('.net-status-pending')), marker);

await setJs('fetch("/run-a-marker.json").catch(()=>{}); fetch("/hang-forever").catch(()=>{});');
await page.click('#btn-run');
await waitForNetRow('run-a-marker');

// A REPL eval that can never settle in this frame.
await page.fill('#console-input', 'new Promise(() => {})');
await page.press('#console-input', 'Enter');

await setJs('fetch("/run-b-marker.json").catch(()=>{});');
await page.click('#btn-run');
await waitForNetRow('run-b-marker');

await page.click('#diag-tabs .tab[data-diag="network"]');
await check('network history keeps rows from both runs', async () => {
  const t = await page.locator('#network-rows').textContent();
  return t.includes('run-a-marker') && t.includes('run-b-marker');
});
await check('old row detail shows the old request (ids namespaced per run)', async () => {
  await page.locator('#network-rows .network-row', { hasText: 'run-a-marker' }).click();
  return (await page.locator('#network-detail').textContent()).includes('run-a-marker');
});
await check('exactly one row selected across runs', async () =>
  (await page.locator('#network-rows .network-row.selected').count()) === 1);
await check('request pending at rerun is marked cancelled', async () =>
  (await page.locator('#network-rows .network-row', { hasText: 'hang-forever' }).textContent())
    .includes('cancelled'));
await check('abandoned REPL eval settles with a cancellation result', async () =>
  (await page.locator('#console-out').textContent())
    .includes('cancelled — a new run replaced the frame'));
await page.click('#diag-tabs .tab[data-diag="console"]');

// --- fragment links + console text filter ---
// The <base href> that makes "#foo" navigate away only exists on a live
// tenant (the local mock deliberately sets baseHref: null), so these are
// regression guards on the interceptor itself: an in-page link must
// scroll and fire hashchange, and user preventDefault() must still win.
const inFrame = (fn) => page.evaluate(fn);
const frameScrollY = () => inFrame(() =>
  document.querySelector('#preview-host iframe').contentWindow.scrollY);

await setDoc('html', `<div id="head-marker">HEADMARKER</div>
<a id="jump" href="#target">jump</a>
<a id="noop" href="#">noop</a>
<div style="height:1500px"></div>
<div id="target">TARGETMARKER</div>`);
await setJs(`
window.addEventListener('hashchange', function () { console.log('hashchange-fired'); });
document.getElementById('noop').addEventListener('click', function (e) {
  e.preventDefault();
  console.log('noop-handler-ran');
});
`);
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('ran in'));

// The status flips to "ran in" when the harness posts back, which can beat
// the outside world's view of the srcdoc document swap (seen on Edge on
// Windows) — wait until the new document is actually reachable.
await page.waitForFunction(() =>
  document.querySelector('#preview-host iframe')?.contentDocument?.getElementById('jump'));
await inFrame(() =>
  document.querySelector('#preview-host iframe').contentDocument.getElementById('jump').click());
await page.waitForFunction(() =>
  document.querySelector('#console-out').textContent.includes('hashchange-fired'));

await check('fragment link scrolls the preview instead of leaving it', async () => {
  const still = await inFrame(() => {
    const w = document.querySelector('#preview-host iframe').contentWindow;
    return !!w.document.getElementById('target') && !!w.document.getElementById('head-marker');
  });
  return still && (await frameScrollY()) > 100;
});
await check('fragment navigation still fires hashchange', async () =>
  (await page.locator('#console-out').textContent()).includes('hashchange-fired'));

const scrollBefore = await frameScrollY();
await inFrame(() =>
  document.querySelector('#preview-host iframe').contentDocument.getElementById('noop').click());
await page.waitForFunction(() =>
  document.querySelector('#console-out').textContent.includes('noop-handler-ran'));
await check('user preventDefault beats the interceptor (no scroll-to-top)', async () =>
  scrollBefore > 100 && (await frameScrollY()) === scrollBefore);

// Text filter: the injected timestamp must not be searchable.
const FILTER_SETTLE = 300;    // filter input is debounced at 150 ms
const stamp = await page.locator('#console-out .entry-ts').first().textContent();
await page.fill('#console-filter-text', stamp);
await page.waitForTimeout(FILTER_SETTLE);
await check('text filter ignores entry timestamps', async () =>
  (await page.locator('#console-out .console-entry:not(.hidden-txt)').count()) === 0);

await page.fill('#console-filter-text', 'noop-handler');
await page.waitForTimeout(FILTER_SETTLE);
await check('text filter still matches logged content', async () =>
  (await page.locator('#console-out .console-entry:not(.hidden-txt)').count()) === 1);
await page.fill('#console-filter-text', '');
await page.waitForTimeout(FILTER_SETTLE);

// --- framework catalog: add via form, injection, 404, drag reorder, delete ---
// (local fixture: sandbox blocks public CDNs; the mechanism — ordered
// blocking <script src> — is identical)
const addFramework = async (name, url) => {
  // The add form is a collapsed pinned footer by default.
  if (await page.locator('#lib-custom-form').isHidden()) await page.click('#btn-add-framework');
  await page.fill('#lib-custom-name', name);
  await page.fill('#lib-custom-url', url);
  await page.click('#lib-add-submit');
};
const libRow = (text) => page.locator('#lib-list .lib-item', { hasText: text });
// The assertion IS the wait: poll until the iframe's <script src> order
// matches (or time out and fail the check).
const scriptOrderBecomes = (expected) => page.waitForFunction((exp) => {
  const f = document.querySelector('#preview-host iframe');
  if (!f || !f.contentDocument) return false;
  const srcs = [...f.contentDocument.querySelectorAll('script[src]')].map((s) => s.getAttribute('src'));
  const catalogScripts = srcs.filter((src) => exp.includes(src));
  return JSON.stringify(catalogScripts) === JSON.stringify(exp);
}, expected, { timeout: 8000 }).then(() => true, () => false);

const LIB_A = `${FIXTURES_URL}/fixtures/testlib.js`;
const LIB_B = `${FIXTURES_URL}/fixtures/testlib.js?b`;

await addFramework('', LIB_A);   // no name — falls back to filename
await setJs('console.log("testlib-type", typeof testlib, testlib && testlib.hello());');
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('testlib-type'));
await check('added framework loads before user JS', async () =>
  (await page.locator('#console-out').textContent()).includes('testlib-type object hi'));

await addFramework('', `${FIXTURES_URL}/fixtures/does-not-exist.js`);
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('Failed to load resource'));
await check('framework 404 surfaces as console error', true);
page.once('dialog', (d) => d.accept());
await libRow('does-not-exist.js').locator('.lib-del').click();
await check('framework removed from catalog', async () =>
  (await libRow('does-not-exist.js').count()) === 0);

await addFramework('testlib-b', LIB_B);
await page.click('#btn-run');
await check('enabled frameworks inject in catalog order', () =>
  scriptOrderBecomes([LIB_A, LIB_B]));
await libRow('testlib-b').locator('.lib-drag').dragTo(libRow('testlib.js'), {
  targetPosition: { x: 24, y: 2 },
});
await page.click('#btn-run');
await check('drag reorder changes injection order', () =>
  scriptOrderBecomes([LIB_B, LIB_A]));
await check('drag reorder persists explicit catalog order', () =>
  page.evaluate(() => import('/src/libraries.js').then(({ getCatalogDoc }) => {
    const items = getCatalogDoc().items;
    const first = items.findIndex((item) => item.name === 'testlib-b');
    const second = items.findIndex((item) => item.name === 'testlib.js');
    return first !== -1
      && first < second
      && items.every((item, index) => item.order === index + 1);
  })));
page.once('dialog', (d) => d.accept());
await libRow('testlib-b').locator('.lib-del').click();

// --- catalog file: export → import-without-entry (prunes) → restore ---
const [catDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#btn-catalog-export'),
]);
const catPath = await catDownload.path();
const catJson = JSON.parse(readFileSync(catPath, 'utf8'));
const testlibEntry = catJson.items.find((i) => i.name === 'testlib.js');
await check('exported catalog file has the right shape', () =>
  catJson.kind === 'dcspad-framework-catalog'
  && Array.isArray(catJson.items) && !!testlibEntry);

// A snippet document used to pass the catalog's shallow { items } check and
// replace every framework. The signed wrong type must now be rejected before
// confirmation or mutation.
const wrongCatalogPath = join(tmpdir(), 'snippets-mistaken-for-frameworks.json');
writeFileSync(wrongCatalogPath, JSON.stringify({
  kind: 'dcspad-snippet-library',
  v: 1,
  items: [{ id: 'wrong-type', name: 'Wrong type', lang: 'js', code: 'void 0;' }],
}));
const catalogCountBeforeWrongImport = await page.locator('#panel-frameworks .lib-item').count();
const wrongCatalogDialogPromise = page.waitForEvent('dialog');
await page.setInputFiles('#import-catalog-file', wrongCatalogPath);
const wrongCatalogDialog = await wrongCatalogDialogPromise;
const wrongCatalogMessage = wrongCatalogDialog.message();
await wrongCatalogDialog.dismiss();
await check('snippet files are rejected by the framework importer without mutation', async () =>
  wrongCatalogMessage.includes('snippet library, not a framework catalog')
  && await page.locator('#panel-frameworks .lib-item').count() === catalogCountBeforeWrongImport);

// Import a copy with testlib removed while testlib is still enabled:
// the row must disappear AND its dead id must be pruned from the
// workspace's enabled list (regression: ids accumulated forever).
const prunedCatalogPath = join(tmpdir(), 'dcspad-catalog-without-testlib.json');
writeFileSync(prunedCatalogPath, JSON.stringify(
  { v: 1, items: catJson.items.filter((i) => i !== testlibEntry) }));
page.once('dialog', (d) => d.accept());
await page.setInputFiles('#import-catalog-file', prunedCatalogPath);
await check('catalog import replaces the catalog', () =>
  page.waitForFunction(() =>
    !document.querySelector('#lib-list').textContent.includes('testlib.js'))
    .then(() => true, () => false));
await check('catalog import prunes dead enabled ids from the workspace', () =>
  page.waitForFunction((id) =>
    !JSON.parse(localStorage.getItem('dcspad.v2.workspace')).libraries.enabled.includes(id),
    testlibEntry.id).then(() => true, () => false));

page.once('dialog', (d) => d.accept());
await page.setInputFiles('#import-catalog-file', catPath);
await check('catalog file round-trip restores entries', () =>
  page.waitForFunction(() =>
    document.querySelector('#lib-list').textContent.includes('testlib.js'))
    .then(() => true, () => false));

// --- snippets: default pack, guarded files, save from selection, insert ---
await check('new users receive the maintained starter snippet library', async () =>
  (await page.locator('#snippet-list .snippet-item').count()) === starterSnippetDoc.items.length
  && (await page.locator('#snippets-count').textContent()) === String(starterSnippetDoc.items.length));

// Isolate the custom-snippet checks from the starter pack while retaining a
// correctly signed empty personal library.
await page.evaluate(() => localStorage.setItem(
  'dcspad.v2.snippets',
  JSON.stringify({ kind: 'dcspad-snippet-library', v: 1, items: [] }),
));
await page.reload();
await page.waitForSelector('.monaco-editor');
await page.waitForFunction(() => document.documentElement.dataset.monacoReady === 'true');

await setJs('var SNIPPET_MARKER = 42;');
await page.locator('#pane-editor .view-lines').click({ position: { x: 80, y: 10 } });
await page.keyboard.press('Control+a');
await page.click('#btn-snippet-add');
await check('snippet naming dialog opens', () =>
  page.locator('#snippet-name-dialog').evaluate((dialog) => dialog.open));
await page.fill('#snippet-name-input', 'my-snip');
await page.click('#snippet-name-save');
await check('snippet saved from selection', async () =>
  (await page.locator('#snippet-list .snippet-item', { hasText: 'my-snip' }).count()) === 1);

const saveSnippet = async (pane, code, name) => {
  await setDoc(pane, code);
  await page.locator('#pane-editor .view-lines').click({ position: { x: 80, y: 10 } });
  await page.keyboard.press('Control+a');
  await page.click('#btn-snippet-add');
  await page.fill('#snippet-name-input', name);
  await page.click('#snippet-name-save');
};
await saveSnippet('html', '<p>zulu</p>', 'Zulu markup');
await saveSnippet('css', '.alpha { color: teal; }', 'alpha styles');
await check('snippets display alphabetically regardless of file type', async () => {
  const names = await page.locator('#snippet-list .snippet-item .lib-name').allTextContents();
  return JSON.stringify(names) === JSON.stringify(['alpha styles', 'my-snip', 'Zulu markup']);
});

const [snippetDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#btn-snippets-export'),
]);
const snippetPath = await snippetDownload.path();
const snippetJson = JSON.parse(readFileSync(snippetPath, 'utf8'));
await check('exported snippet library carries its file signature', () =>
  snippetJson.kind === 'dcspad-snippet-library'
  && snippetJson.items.length === 3);

const snippetCountBeforeWrongImport = await page.locator('#snippet-list .snippet-item').count();
const wrongSnippetDialogPromise = page.waitForEvent('dialog');
await page.setInputFiles('#import-snippets-file', catPath);
const wrongSnippetDialog = await wrongSnippetDialogPromise;
const wrongSnippetMessage = wrongSnippetDialog.message();
await wrongSnippetDialog.dismiss();
await check('framework files are rejected by the snippet importer without mutation', async () =>
  wrongSnippetMessage.includes('framework catalog, not a snippet library')
  && await page.locator('#snippet-list .snippet-item').count() === snippetCountBeforeWrongImport);

await setJs('// cleared\n');
await page.locator('#snippet-list .snippet-item', { hasText: 'my-snip' }).click();
await page.waitForFunction(
  () => document.querySelector('#pane-editor .view-lines')?.textContent
    .replaceAll('\u00a0', ' ')
    .includes('SNIPPET_MARKER = 42'),
  null,
  { timeout: 2000 },
).catch(() => {});
await check('snippet inserts into the JS editor at the cursor', async () =>
  (await page.locator('#pane-editor .view-lines').textContent())
    .replaceAll('\u00a0', ' ')
    .includes('SNIPPET_MARKER = 42'));

// --- project name + project file: require name, save → load round-trip ---
await setJs('console.log("round-trip-original");');
await page.click('#btn-file');
await page.click('#mi-save-project');
await check('project JSON save requires an inline project name', async () =>
  await page.locator('#project-name-form').isVisible()
  && (await page.locator('#project-name-error').textContent()).includes('Name this project'));
await page.fill('#project-name-input', 'This Is an Example');
const [projDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#project-name-save'),
]);
const projPath = await projDownload.path();
const projJson = JSON.parse(readFileSync(projPath, 'utf8'));
await check('saved project file has the right shape', () =>
  projJson.kind === 'project' && projJson.docs.js.includes('round-trip-original')
  && Array.isArray(projJson.libraries.enabled) && projJson.name === 'This Is an Example');
await check('project JSON uses the project slug and .dcspad.json extension', () =>
  projDownload.suggestedFilename() === 'this-is-an-example.dcspad.json');
await check('downloading a project marks its current contents saved', async () =>
  (await page.locator('#project-file-state').textContent()) === 'saved');
await check('project import picker prefers .dcspad.json files', () =>
  page.locator('#import-project-file').getAttribute('accept')
    .then((accept) => accept.startsWith('.dcspad.json')));

await setJs('console.log("changed after save");');
await page.waitForFunction(() =>
  document.querySelector('#project-file-state')?.textContent === 'unsaved');
await check('editing after a project-file save marks the project unsaved', async () =>
  (await page.locator('#project-file-state').textContent()) === 'unsaved');
await page.setInputFiles('#import-project-file', projPath);
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('project loaded'));
await page.waitForFunction(() =>
  document.querySelector('#pane-editor .view-lines')?.textContent.includes('round-trip-original'));
await check('loading the project file restores the JS pane', async () =>
  (await page.locator('#pane-editor .view-lines').textContent()).includes('round-trip-original'));
await check('loading the project file restores its title', async () =>
  (await page.locator('#project-name-text').textContent()) === 'This Is an Example');
await check('loading a project establishes a saved baseline', async () =>
  (await page.locator('#project-file-state').textContent()) === 'saved');

// File > New Project only prompts when the current project differs from its
// last file. Cancel preserves the work; Save downloads it before resetting.
await setJs('console.log("save-before-new");');
await page.click('#btn-file');
await page.click('#mi-new-project');
await check('New Project prompts to save changed work', () =>
  page.locator('#new-project-dialog').evaluate((dialog) => dialog.open));
await page.click('#new-project-cancel');
await check('canceling New Project preserves the current project', async () =>
  (await page.locator('#project-name-text').textContent()) === 'This Is an Example'
  && (await page.locator('#pane-editor .view-lines').textContent()).includes('save-before-new'));

await page.click('#btn-file');
await page.click('#mi-new-project');
const [beforeNewDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#new-project-save'),
]);
await page.waitForFunction(() =>
  document.querySelector('#project-name-text')?.textContent === '(untitled)'
  && document.querySelector('#status-run')?.textContent.includes('new project'));
await check('saving from the New Project prompt downloads the changed project', async () =>
  beforeNewDownload.suggestedFilename() === 'this-is-an-example.dcspad.json'
  && readFileSync(await beforeNewDownload.path(), 'utf8').includes('save-before-new'));
await page.click('#editor-tabs .tab[data-editor="js"]');
await page.waitForFunction(() =>
  document.querySelector('#pane-editor .view-lines')?.textContent
    .replaceAll('\u00a0', ' ')
    .includes('DCSPad ready'));
await check('New Project restores fresh untitled editors and clears the old preview', async () =>
  (await page.locator('#project-name-text').textContent()) === '(untitled)'
  && (await page.locator('#project-file-state').textContent()) === 'unsaved'
  && await page.locator('#preview-empty').isVisible()
  && (await page.locator('#pane-editor .view-lines').textContent())
    .replaceAll('\u00a0', ' ')
    .includes('DCSPad ready'));

// Restore the named fixture for the title-derived export checks below.
await page.setInputFiles('#import-project-file', projPath);
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('project loaded'));

await setJs('console.log("discard-before-new");');
await page.waitForFunction(() =>
  document.querySelector('#project-file-state')?.textContent === 'unsaved');
await page.click('#btn-file');
await page.click('#mi-new-project');
await page.click('#new-project-discard');
await page.waitForFunction(() =>
  document.querySelector('#project-name-text')?.textContent === '(untitled)');
await check('Don’t save starts a fresh project without downloading', async () =>
  (await page.locator('#project-file-state').textContent()) === 'unsaved'
  && await page.locator('#preview-empty').isVisible());

// Restore once more for the title-derived export checks below.
await page.setInputFiles('#import-project-file', projPath);
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('project loaded'));

// Named pane exports share the title-derived base. Export-all skips the
// deliberately emptied CSS pane and downloads the two remaining types.
await setDoc('css', '');
await page.click('#btn-file');
const [namedJsDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#mi-export-js'),
]);
await check('named pane export uses the project slug', () =>
  namedJsDownload.suggestedFilename() === 'this-is-an-example.js');

const allDownloads = [];
const captureAllDownload = (download) => allDownloads.push(download);
page.on('download', captureAllDownload);
await page.click('#btn-file');
await page.click('#mi-export-all');
await page.waitForTimeout(300);
page.off('download', captureAllDownload);
await check('export all downloads every non-empty pane and skips empty panes', () =>
  JSON.stringify(allDownloads.map((d) => d.suggestedFilename()).sort())
    === JSON.stringify(['this-is-an-example.html', 'this-is-an-example.js']));

// A project referencing a framework missing from the catalog loads
// tolerantly but warns by name.
const ghostProject = join(tmpdir(), 'dcspad-ghost-project.json');
writeFileSync(ghostProject, JSON.stringify({
  app: 'dcspad', kind: 'project', v: 1,
  docs: { html: '', css: '', js: '// ghost-lib project' },
  libraries: { enabled: ['ghost-lib-id'] }, jsAsModule: false,
}));
await page.setInputFiles('#import-project-file', ghostProject);
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('not in your catalog'));
await check('missing-framework project loads with a named warning', async () =>
  (await page.locator('#console-out').textContent()).includes('ghost-lib-id'));

// --- pane export ---
await setJs('var EXPORT_MARKER = 1;\n');
await page.click('#btn-file');
const [jsDownload] = await Promise.all([
  page.waitForEvent('download'),
  page.click('#mi-export-js'),
]);
await check('export JS pane downloads the pane contents', async () =>
  readFileSync(await jsDownload.path(), 'utf8').includes('EXPORT_MARKER')
  && jsDownload.suggestedFilename() === 'dcspad.js');

// --- storage failure is surfaced, not silent ---
// Stub setItem to throw (the only way to hit quota deterministically);
// the status bar must show an error instead of sitting on "saving…".
await page.evaluate(() => {
  window.__origSetItem = Storage.prototype.setItem;
  Storage.prototype.setItem = function () { throw new DOMException('quota', 'QuotaExceededError'); };
});
await setJs('var QUOTA_PROBE = 1;');
await check('failed autosave surfaces an error in the status bar', () =>
  page.waitForFunction(() =>
    document.querySelector('#status-save').textContent.includes('save failed'))
    .then(() => true, () => false));
await page.evaluate(() => { Storage.prototype.setItem = window.__origSetItem; });
// Re-establish a known, successfully-saved doc for the reload checks.
await setJs('var EXPORT_MARKER = 1;\n');
await page.waitForFunction(() =>
  document.querySelector('#status-save').textContent.includes('✓ saved'));

// --- autosave/restore across reload (workspace, catalog, snippets) ---
await page.reload();
await page.waitForTimeout(1200);
await check('JS doc restored after reload', async () =>
  (await page.locator('#pane-editor .view-lines').textContent()).includes('EXPORT_MARKER'));
await check('catalog framework persisted after reload', async () =>
  (await libRow('testlib.js').count()) === 1);
await check('snippet library persisted after reload', async () =>
  (await page.locator('#snippet-list .snippet-item', { hasText: 'my-snip' }).count()) === 1);

// Explicit recovery controls restore the same state as a fresh install.
page.once('dialog', (dialog) => dialog.accept());
await page.click('#btn-catalog-reset');
await check('framework reset restores PRESETS and clears workspace selections', () =>
  page.evaluate(async () => {
    const { getCatalogDoc, PRESETS } = await import('/src/libraries.js');
    const { getState } = await import('/src/state.js');
    const workspace = getState();
    const stored = JSON.parse(localStorage.getItem('dcspad.v2.catalog'));
    return stored.kind === 'dcspad-framework-catalog'
      && getCatalogDoc().items.length === PRESETS.length
      && stored.items.length === PRESETS.length
      && stored.items[2].id === 'dcs-standard'
      && stored.items[2].order === 3
      && stored.items.every((item, index) => item.order === index + 1)
      && workspace.libraries.enabled.length === 0
      && JSON.stringify(workspace.libraries.pinned) === JSON.stringify(['pnpjs2']);
  }));

// A catalog saved before explicit order fields existed should adopt the
// built-in DCS position once, then persist concrete order values.
await page.evaluate(() => {
  const stored = JSON.parse(localStorage.getItem('dcspad.v2.catalog'));
  const dcsIndex = stored.items.findIndex((item) => item.id === 'dcs-standard');
  const [dcs] = stored.items.splice(dcsIndex, 1);
  stored.items.unshift(dcs);
  stored.items.forEach((item) => { delete item.order; });
  localStorage.setItem('dcspad.v2.catalog', JSON.stringify(stored));
});
await page.reload();
await check('legacy framework catalog adopts explicit preset order', () =>
  page.evaluate(() => {
    const stored = JSON.parse(localStorage.getItem('dcspad.v2.catalog'));
    return stored.items[2].id === 'dcs-standard'
      && stored.items.every((item, index) => item.order === index + 1);
  }));

page.once('dialog', (dialog) => dialog.accept());
await page.click('#btn-snippets-reset');
await page.waitForFunction((expected) =>
  Number(document.querySelector('#snippets-count')?.textContent) === expected,
  starterSnippetDoc.items.length);
await check('snippet reset restores the maintained starter pack', () =>
  page.evaluate((expected) => {
    const stored = JSON.parse(localStorage.getItem('dcspad.v2.snippets'));
    return stored.kind === 'dcspad-snippet-library'
      && stored.items.length === expected;
  }, starterSnippetDoc.items.length));

await browser.close();
exitWithResult();
