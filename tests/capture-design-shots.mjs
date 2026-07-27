// Captures the screenshot set for design/DESIGN-BRIEF.md into
// design/screenshots/. Not a test — a documentation tool. Two servers as
// per README (app on 8642), then: node capture-design-shots.mjs
//
// States captured: main workbench after a run, both dropdown menus, an
// error state (console + network dots lit), and the collapsed
// diagnostics/sidebar layout.

import { launchBrowser, APP_URL } from './lib.mjs';
import { mkdirSync } from 'fs';

const OUT = '../design/screenshots';
mkdirSync(OUT, { recursive: true });

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
await page.goto(APP_URL);
await page.evaluate(() => localStorage.clear());
await page.reload();
await page.waitForTimeout(1400);
// skip the splash if still up
await page.evaluate(() => document.getElementById('splash')?.click());
await page.waitForTimeout(500);

const shot = (name) => page.screenshot({ path: `${OUT}/${name}.png` });

// 1. main workbench after running the default project
await page.click('#btn-run');
await page.waitForFunction(() =>
  document.querySelector('#status-run')?.textContent.includes('ran in'));
await page.waitForTimeout(300);
await shot('01-main-after-run');

// 2. menus
await page.click('#btn-file');
await shot('02-file-menu');
await page.keyboard.press('Escape');
await page.click('body', { position: { x: 700, y: 400 } });
await page.click('#btn-settings');
await shot('03-settings-menu');
await page.click('body', { position: { x: 700, y: 400 } });

// 3. error state: console error + failing network request -> both dots
await page.evaluate(() => {
  const set = (id, v) => { /* type into CM via its api is heavy; use REPL */ };
});
await page.fill('#console-input', 'fetch("/definitely-not-real-404")');
await page.press('#console-input', 'Enter');
await page.waitForTimeout(500);
await page.fill('#console-input', 'nonexistentVariable.foo');
await page.press('#console-input', 'Enter');
await page.waitForTimeout(600);
await shot('04-error-state-dots');

// 4. hidden console + hidden resources (topbar pane toggles)
await page.click('#seg-console');
await page.click('#seg-resources');
await page.waitForTimeout(300);
await shot('05-collapsed-panels');

await browser.close();
console.log('captured 5 shots into design/screenshots/');
