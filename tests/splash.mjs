// Splash lifecycle suite: full reveal auto-removes, app interactive
// afterwards, revisit gets the quick shimmer.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
await page.goto(APP_URL);
await page.evaluate(() => localStorage.clear());
await page.reload();
await page.waitForTimeout(3700);

await check('splash auto-removes after first-visit reveal', async () =>
  await page.evaluate(() => !document.getElementById('splash')));

await page.click('#btn-run');
await page.waitForTimeout(800);
await check('app interactive after splash', async () =>
  (await page.locator('#status-run').textContent()).includes('ran in'));

await page.reload();
await page.waitForTimeout(200);
const visible = await page.evaluate(() =>
  !!document.getElementById('splash') && !document.getElementById('splash').hidden);
await page.waitForTimeout(1400);
const gone = await page.evaluate(() => !document.getElementById('splash'));
await check('short shimmer on revisit', visible && gone);

await browser.close();
exitWithResult();
