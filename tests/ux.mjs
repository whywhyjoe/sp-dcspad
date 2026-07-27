// Phase-2 UX additions suite: topbar pane toggles (visibility, persistence,
// keyboard chords through Monaco), in-context text-size steppers, error
// count pills, the REPL Eval button, add-framework footer validation, and
// the persisted Frameworks/Snippets split.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
await page.goto(APP_URL);
await page.waitForTimeout(400);
await page.evaluate(() => localStorage.clear());
await page.reload();
await page.waitForTimeout(3200);   // splash
await page.waitForSelector('.monaco-editor');

// --- pane toggles ---
await page.click('#seg-resources');
await check('resources pane hides', () => page.locator('#sidebar').isHidden());

await page.click('#seg-preview');
await check('preview pane hides (new capability)', () => page.locator('#preview-panel').isHidden());

await page.click('#seg-console');
await check('all three hidden — editors own the full width', async () => {
  const hidden = await page.locator('#diag-panel').isHidden();
  const box = await page.locator('#editors').boundingBox();
  return hidden && box.width > 1400;
});

await page.reload();
await page.waitForTimeout(1600);
await page.waitForSelector('.monaco-editor');
await check('pane visibility persists across reload', async () =>
  (await page.locator('#sidebar').isHidden())
  && (await page.locator('#preview-panel').isHidden())
  && (await page.locator('#diag-panel').isHidden()));

await page.click('#seg-resources');
await page.click('#seg-preview');
await page.click('#seg-console');
await check('segments restore every pane', async () =>
  (await page.locator('#sidebar').isVisible())
  && (await page.locator('#preview-panel').isVisible())
  && (await page.locator('#diag-panel').isVisible()));

// The same chord must work while Monaco has focus (it swallows document
// keydown for bound keys — the app registers editor actions for these).
await page.locator('#pane-editor .view-lines').click();
await page.keyboard.press('Control+j');
await check('Ctrl+J toggles console with the editor focused', () =>
  page.locator('#diag-panel').isHidden());
await page.keyboard.press('Control+j');
await check('Ctrl+J restores the console', () => page.locator('#diag-panel').isVisible());

// --- editor text-size stepper (persisted; line height locks to 1.7×) ---
await page.click('#btn-editor-font-inc');
await page.click('#btn-editor-font-inc');
await page.reload();
await page.waitForTimeout(1600);
await page.waitForSelector('.monaco-editor');
await check('editor font size persists and reaches Monaco', async () =>
  (await page.evaluate(() =>
    getComputedStyle(document.querySelector('#pane-editor .view-lines')).fontSize)) === '15px');

// --- REPL Eval button + error count pills ---
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('ran in'));
await page.fill('#console-input', '6 * 7');
await page.click('#btn-repl-eval');
await page.waitForFunction(() =>
  document.querySelector('#console-out')?.textContent.includes('42'));
await check('Eval button shares the Enter submit path', true);

await page.fill('#console-input', 'no_such_fn_dcspad()');
await page.click('#btn-repl-eval');
await page.waitForFunction(() =>
  !document.getElementById('console-badge').hidden);
await check('console error pill shows a count, not a dot', async () =>
  (await page.locator('#console-badge').textContent()) === '1');
await page.fill('#console-input', 'no_such_fn_dcspad()');
await page.click('#btn-repl-eval');
await page.waitForFunction(() =>
  document.getElementById('console-badge').textContent === '2');
await check('count accumulates', true);
await page.click('#btn-clear-console');
await check('clear resets the count pill', () =>
  page.locator('#console-badge').isHidden());

// --- add-framework footer: collapsed by default, inline validation ---
await check('add-framework form starts collapsed', () =>
  page.locator('#lib-custom-form').isHidden());
await page.click('#btn-add-framework');
await page.fill('#lib-custom-url', 'not a url at all');
await page.click('#lib-add-submit');
await check('invalid URL shows an inline error (no dialog)', () =>
  page.locator('#lib-custom-error').isVisible());
await page.fill('#lib-custom-url', 'https://example.com/page.html');
await page.click('#lib-add-submit');
await check('non-.js/.css URL rejected inline', () =>
  page.locator('#lib-custom-error').isVisible());
await page.click('#lib-add-cancel');
await check('cancel collapses the footer again', () =>
  page.locator('#lib-custom-form').isHidden());

// --- Frameworks/Snippets split: drag persists, double-click resets ---
const splitBox = await page.locator('#split-side').boundingBox();
await page.mouse.move(splitBox.x + splitBox.width / 2, splitBox.y + 2);
await page.mouse.down();
await page.mouse.move(splitBox.x + splitBox.width / 2, splitBox.y - 60, { steps: 4 });
await page.mouse.up();
const draggedH = await page.evaluate(() =>
  parseFloat(getComputedStyle(document.getElementById('panel-snippets')).height));
await page.reload();
await page.waitForTimeout(1600);
await check('sidebar split persists across reload', async () => {
  const h = await page.evaluate(() =>
    parseFloat(getComputedStyle(document.getElementById('panel-snippets')).height));
  return Math.abs(h - draggedH) < 3 && draggedH > 240;
});
await page.dblclick('#split-side');
await check('double-click resets the split to 210', async () =>
  (await page.evaluate(() =>
    parseFloat(getComputedStyle(document.getElementById('panel-snippets')).height))) === 210);

await browser.close();
exitWithResult();
