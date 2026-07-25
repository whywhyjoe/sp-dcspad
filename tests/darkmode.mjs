// Preview dark-mode suite: default dark canvas, toggle both ways,
// user-CSS-wins layering, persistence.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
await page.goto(APP_URL);
await page.waitForTimeout(400);
await page.evaluate(() => localStorage.clear());
await page.reload();
await page.waitForTimeout(3200);   // splash

const frameBg = () => page.evaluate(() => {
  const f = document.querySelector('#preview-host iframe');
  return getComputedStyle(f.contentDocument.documentElement).backgroundColor;
});
const bodyBg = () => page.evaluate(() => {
  const f = document.querySelector('#preview-host iframe');
  return getComputedStyle(f.contentDocument.body).backgroundColor;
});

await page.click('#btn-run');
await page.waitForTimeout(800);
await check('default preview is dark', (await frameBg()) === 'rgb(29, 32, 38)');
await check('toggle shows sun in dark mode',
  (await page.locator('#btn-preview-theme').textContent()) === '☀');
await check('host has dark class (no white flash)', await page.evaluate(() =>
  document.getElementById('preview-host').classList.contains('dark')));

await page.click('#btn-preview-theme');
await page.waitForTimeout(800);
const lightBg = await frameBg();
await check('toggle to light re-runs with default rendering',
  lightBg === 'rgba(0, 0, 0, 0)' || lightBg === 'rgb(255, 255, 255)');
await check('toggle shows moon in light mode',
  (await page.locator('#btn-preview-theme').textContent()) === '🌙');

await page.click('#btn-preview-theme');
await page.waitForTimeout(800);
await page.click('#editor-tabs .tab[data-editor="css"]');
await page.click('#pane-css .cm-content');
await page.keyboard.press('Control+a');
await page.keyboard.insertText('body { background: papayawhip; }');
await page.click('#btn-run');
await page.waitForTimeout(800);
await check('user CSS beats dark chrome style', (await bodyBg()) === 'rgb(255, 239, 213)');
await check('html stays dark where user did not style it', (await frameBg()) === 'rgb(29, 32, 38)');

await page.click('#btn-preview-theme');
await page.waitForTimeout(600);
await page.reload();
await page.waitForTimeout(1500);
await check('preview theme persists across reload',
  (await page.locator('#btn-preview-theme').textContent()) === '🌙');

await browser.close();
exitWithResult();
