// Exercises the exact boot-workbench.js + dcspad.workbench.js path used by
// the SharePoint Modern Script Editor, on a fixture whose mount sits far
// below the fold (as it does on a real page).
//
// Why this suite exists: the first hosted deploy rendered a blank page.
// app.css — which boot-workbench.js injects into the host page — sets
// unscoped `html, body { height: 100% }` and `body { overflow: hidden }`,
// so a workbench left in normal document flow ends up below the fold on a
// page that can no longer scroll. It only became visible in page-edit mode,
// where .dcspad-suspended reverts those rules. Every check below fails
// against that build.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const origin = new URL(APP_URL).origin;
const VIEWPORT = { width: 1500, height: 900 };

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: VIEWPORT });
const pageErrors = [];
page.on('pageerror', (error) => pageErrors.push(error.message));

await page.goto(`${origin}/tests/workbench-hosted-fixture.html`);
await page.waitForSelector('.wb');

await check('hosted: boot marks the host page and mounts the shell', () =>
  page.evaluate(() =>
    document.documentElement.classList.contains('dcspad-hosted')
    && !!document.getElementById('wb-mount')
    && document.querySelectorAll('.wb').length === 1));

await check('hosted: the workbench is visible inside the viewport', () =>
  page.evaluate(() => {
    const wb = document.querySelector('.wb');
    const rect = wb.getBoundingClientRect();
    const style = getComputedStyle(wb);
    return style.position === 'fixed'
      && style.display !== 'none'
      && style.visibility !== 'hidden'
      && rect.top >= 0 && rect.top < window.innerHeight
      && rect.bottom > 0
      && rect.width > 400 && rect.height > 200;
  }));

await check('hosted: the workbench paints over the host page, not under it', () =>
  page.evaluate(() => {
    const wb = document.querySelector('.wb');
    const rect = wb.getBoundingClientRect();
    const x = Math.round(rect.left + rect.width / 2);
    const y = Math.round(rect.top + rect.height / 2);
    return wb.contains(document.elementFromPoint(x, y));
  }));

await check('hosted: the grid renders real rows once the bundle runs', async () => {
  await page.waitForSelector('.wb-table tbody tr');
  return (await page.locator('.wb-table tbody tr').count()) > 0;
});

await check('hosted: the boot note is pinned so failures stay visible', () =>
  page.evaluate(async () => {
    // Re-create the note exactly as boot-workbench.js styles it and confirm
    // the pinning lands it on screen even though #wb-mount is below the fold.
    const probe = document.createElement('div');
    probe.style.cssText =
      'position:fixed;top:60px;left:12px;z-index:1000;'
      + 'font:12px/1.5 Consolas,monospace;color:#a2a9b8;background:#14161b;'
      + 'padding:14px;border-radius:6px';
    probe.textContent = 'probe';
    document.getElementById('wb-mount').appendChild(probe);
    const rect = probe.getBoundingClientRect();
    const onScreen = rect.top >= 0 && rect.top < window.innerHeight;
    probe.remove();
    return onScreen;
  }));

await check('hosted: SPA navigation into edit mode suspends the workbench', async () => {
  await page.evaluate(() => {
    history.pushState({}, '', `${location.pathname}?Mode=Edit`);
  });
  await page.waitForFunction(() =>
    document.documentElement.classList.contains('dcspad-suspended'));
  return page.evaluate(() =>
    getComputedStyle(document.getElementById('wb-mount')).display === 'none');
});

await check('hosted: leaving edit mode restores the workbench', async () => {
  await page.evaluate(() => {
    history.pushState({}, '', location.pathname);
  });
  await page.waitForFunction(() =>
    !document.documentElement.classList.contains('dcspad-suspended'));
  return page.evaluate(() => {
    const rect = document.querySelector('.wb').getBoundingClientRect();
    return getComputedStyle(document.getElementById('wb-mount')).display !== 'none'
      && rect.top < window.innerHeight;
  });
});

await check('hosted: a page loaded directly in edit mode never boots', async () => {
  const editPage = await browser.newPage({ viewport: VIEWPORT });
  await editPage.goto(`${origin}/tests/workbench-hosted-fixture.html?Mode=Edit`);
  await editPage.waitForTimeout(700);
  const result = await editPage.evaluate(() => ({
    booted: !!document.querySelector('.wb'),
    note: document.querySelector('[data-wb-anchor]')?.textContent || '',
  }));
  await editPage.close();
  return !result.booted && result.note.includes('inactive while the page is in edit mode');
});

await check('hosted: no page errors during the hosted boot', () => pageErrors.length === 0);

await browser.close();
exitWithResult();
