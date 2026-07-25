// DCSPad smoke suite: execution model, console capture, SP inspector,
// network monitor, REPL, run isolation, rerun lifecycle (cross-run
// network history, cancelled requests/evals), library injection, autosave.
// Setup: see tests/README.md (app server on 8642, fixtures on 8643).

import { launchBrowser, check, exitWithResult, APP_URL, FIXTURES_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });

await page.goto(APP_URL);
await page.waitForTimeout(1200);

await check('editors render (3 CM instances)', async () =>
  (await page.locator('.cm-editor').count()) === 3);

await check('SP chip shows Mock', async () =>
  (await page.locator('#sp-chip-text').textContent()) === 'SP: Mock');

await page.evaluate(() => localStorage.clear());

const setDoc = async (name, code) => {
  await page.click(`#editor-tabs .tab[data-editor="${name}"]`);
  await page.click(`#pane-${name} .cm-content`);
  await page.keyboard.press('Control+a');
  await page.keyboard.insertText(code);
};
const setJs = (code) => setDoc('js', code);

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

// library injection via local fixture (sandbox blocks public CDNs;
// the mechanism — ordered blocking <script src> — is identical)
await page.fill('#lib-custom-url', `${FIXTURES_URL}/fixtures/testlib.js`);
await page.click('#lib-custom-form button');
await page.waitForTimeout(200);
await setJs('console.log("testlib-type", typeof testlib, testlib && testlib.hello());');
await page.click('#btn-run');
await page.waitForTimeout(1500);
await check('custom library loads before user JS', async () =>
  (await page.locator('#console-out').textContent()).includes('testlib-type object hi'));

await page.fill('#lib-custom-url', `${FIXTURES_URL}/fixtures/does-not-exist.js`);
await page.click('#lib-custom-form button');
await page.waitForTimeout(200);
await page.click('#btn-run');
await page.waitForTimeout(1500);
await check('library 404 surfaces as console error', async () =>
  (await page.locator('#console-out').textContent()).includes('Failed to load resource'));
await page.locator('#lib-custom-list .lib-del').nth(1).click();
await page.waitForTimeout(300);

// autosave/restore
await page.reload();
await page.waitForTimeout(1200);
await check('JS doc restored after reload', async () =>
  (await page.locator('#pane-js .cm-content').textContent()).includes('testlib-type'));
await check('custom library persisted after reload', async () =>
  (await page.locator('#lib-custom-list').textContent()).includes('testlib.js'));

await browser.close();
exitWithResult();
