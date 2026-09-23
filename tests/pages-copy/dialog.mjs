// SP Workbench — page copy: the dialog in mock mode (design/PAGE-COPY.md §6).
// Drives the real dialog markup (src/workbench/page-copy-dialog.js class
// hooks) through Playwright against the mock /sites/pagesrc + /sites/pagedst
// webs (src/workbench/mock-pagecopy.js). House style: tests/pages-copy/
// runner.mjs (module shape), the reference flow in quick-e2e.mjs.
//
// Every check opens its own Workbench page, resets the page-copy mock state
// and the write log, then drives the dialog through real user actions and
// reads the rendered DOM back.

const SOURCE_SITE = '/sites/pagesrc';
const DEST_SITE = '/sites/pagedst';

async function openDialog(browser, WB_URL, pageName) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards');
  await page.fill('#wb-site-input', SOURCE_SITE);
  await page.press('#wb-site-input', 'Enter');
  await page.waitForFunction(
    (site) => document.getElementById('wb-status-context').textContent.includes(site),
    SOURCE_SITE,
  );
  await page.evaluate(() => {
    window.__PAGECOPY_MOCK__?.reset();
    window.__DCSPAD_WB_WRITES__ = [];
  });
  await page.evaluate(() => { location.hash = ''; });
  await page.locator('.wb-nav-item, [data-view]').filter({ hasText: 'Pages' }).first().click();
  await page.getByText(pageName).first().click();
  await page.waitForSelector('.wb-page-copy-btn');
  await page.click('.wb-page-copy-btn');
  await page.waitForSelector('.wb-page-copy-dialog');
  return page;
}

async function waitCheckEnabled(page) {
  await page.waitForFunction(
    () => !document.querySelector('.wb-pc-check')?.disabled,
    null, { timeout: 10000 },
  );
}

async function waitReportVisible(page) {
  await page.waitForFunction(
    () => !document.querySelector('.wb-pc-report')?.hidden,
    null, { timeout: 10000 },
  );
}

export async function run({ browser, check, WB_URL }) {
  await check('dialog: opens for Quarterly-Update.aspx with This site pre-selected', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    const thisChecked = await page.$eval('.wb-pc-dest-this', (el) => el.checked);
    await page.close();
    return thisChecked === true;
  });

  await check('dialog: auto-connects to This site; Check fills Quarterly-Update-copy.aspx', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const fileName = await page.inputValue('.wb-pc-filename');
    await page.close();
    return fileName === 'Quarterly-Update-copy.aspx';
  });

  await check('dialog: Copy is disabled before Check, enabled after a clean Check, '
    + 'and disabled immediately after editing Title / toggling Publish', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    const beforeCheck = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (!beforeCheck) { await page.close(); throw new Error('Copy enabled before any Check'); }

    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const afterCheck = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (afterCheck) { await page.close(); throw new Error('Copy still disabled after a clean Check'); }

    await page.fill('.wb-pc-title', 'Edited title');
    const afterTitleEdit = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (!afterTitleEdit) { await page.close(); throw new Error('Copy stayed enabled after editing Title'); }

    // Re-check to re-arm Copy, then confirm toggling Publish independently
    // invalidates it too (not just a leftover from the Title edit).
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const reArmed = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (reArmed) { await page.close(); throw new Error('Copy did not re-arm after a fresh Check'); }

    await page.click('.wb-pc-publish');
    const afterPublishToggle = await page.$eval('.wb-pc-go', (el) => el.disabled);
    await page.close();
    return afterPublishToggle === true;
  });

  await check('dialog: Cancel before Copy leaves zero writes and closes the dialog', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    await page.click('.wb-pc-cancel');
    await page.waitForFunction(() => !document.querySelector('.wb-page-copy-dialog'), null, { timeout: 5000 });
    const writes = await page.evaluate(() => (window.__DCSPAD_WB_WRITES__ || []).length);
    await page.close();
    return writes === 0;
  });

  await check('dialog: double-clicking Copy sends exactly one create POST', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    await Promise.all([page.click('.wb-pc-go'), page.click('.wb-pc-go')]);
    await page.waitForSelector('.wb-pc-outcome:not([hidden])', { timeout: 15000 });
    const createCount = await page.evaluate(() => (window.__DCSPAD_WB_WRITES__ || [])
      .filter((w) => /\/_api\/sitepages\/pages(\?.*)?$/i.test(w.url)).length);
    await page.close();
    return createCount === 1;
  });

  await check('dialog: a full same-web run finishes done, no failed step, Open copy link present; '
    + 'the copy then shows in the Pages list', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    await page.click('.wb-pc-go');
    await page.waitForSelector('.wb-pc-outcome:not([hidden])', { timeout: 15000 });
    const outcome = await page.getAttribute('.wb-pc-outcome', 'data-outcome');
    const failedSteps = await page.$$eval('.wb-pc-step', (rows) => rows.filter((r) => r.dataset.status === 'failed').length);
    const openHidden = await page.$eval('.wb-pc-open', (el) => el.hidden);
    if (outcome !== 'done' || failedSteps !== 0 || openHidden) {
      await page.close();
      throw new Error(`outcome=${outcome} failedSteps=${failedSteps} openHidden=${openHidden}`);
    }
    await page.click('.wb-pc-close');
    await page.waitForSelector('.wb-back', { timeout: 5000 });
    await page.click('.wb-back');
    await page.waitForFunction(
      () => document.body.textContent.includes('Quarterly-Update-copy.aspx'),
      null, { timeout: 10000 },
    );
    await page.close();
    return true;
  });

  await check('dialog: closing a finished run via ✕ also closes it and the grid shows the copy '
    + '(not treated as a cancel)', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    await page.click('.wb-pc-go');
    await page.waitForSelector('.wb-pc-outcome:not([hidden])', { timeout: 15000 });
    await page.click('.app-dialog__head button'); // the ✕ close icon
    await page.waitForFunction(() => !document.querySelector('.wb-page-copy-dialog'), null, { timeout: 5000 });
    await page.waitForSelector('.wb-back', { timeout: 5000 });
    await page.click('.wb-back');
    await page.waitForFunction(
      () => document.body.textContent.includes('Quarterly-Update-copy.aspx'),
      null, { timeout: 10000 },
    );
    await page.close();
    return true;
  });

  await check('dialog: Legacy-News.aspx allows a This-site Check but refuses Another site '
    + 'with a legacy-HTML reason and Copy stays disabled', async () => {
    const page = await openDialog(browser, WB_URL, 'Legacy-News.aspx');
    await waitCheckEnabled(page); // This site (auto-connected) allows the legacy-HTML page
    await page.click('.wb-pc-dest-other');
    await page.fill('.wb-pc-url', DEST_SITE);
    await page.click('.wb-pc-connect');
    await page.waitForFunction(
      () => (document.querySelector('.wb-pc-target-status')?.textContent || '').trim().length > 0,
      null, { timeout: 10000 },
    );
    const statusText = await page.textContent('.wb-pc-target-status');
    const goDisabled = await page.$eval('.wb-pc-go', (el) => el.disabled);
    await page.close();
    if (!/legacy html/i.test(statusText)) throw new Error(`target status did not mention legacy HTML: ${statusText}`);
    return goDisabled === true;
  });

  await check('dialog: Repost.aspx is refused on open, names the RepostPage layout, and has no '
    + 'usable Check button', async () => {
    const page = await openDialog(browser, WB_URL, 'Repost.aspx');
    await page.waitForSelector('.sp-files-error', { timeout: 10000 });
    const reasonText = await page.textContent('.sp-files-error');
    const checkBtnCount = await page.locator('.wb-pc-check').count();
    await page.close();
    if (!/RepostPage/.test(reasonText)) throw new Error(`refusal did not name RepostPage: ${reasonText}`);
    return checkBtnCount === 0;
  });

  await check('dialog: on /sites/pagedst the report shows the Audience required gap and Copy '
    + 'stays disabled; filling it and Check again enables Copy', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-dest-other');
    await page.fill('.wb-pc-url', DEST_SITE);
    await page.click('.wb-pc-connect');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const reportText = await page.textContent('.wb-pc-report');
    const goDisabledBefore = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (!/Audience/.test(reportText) || !goDisabledBefore) {
      await page.close();
      throw new Error(`report="${reportText.slice(0, 300)}" goDisabledBefore=${goDisabledBefore}`);
    }

    const gapInput = page.locator('.wb-pc-gaps input, .wb-pc-gaps textarea, .wb-pc-gaps select').first();
    await gapInput.fill('All employees');
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const goDisabledAfter = await page.$eval('.wb-pc-go', (el) => el.disabled);
    await page.close();
    return goDisabledAfter === false;
  });

  await check('dialog: on /sites/pagedst the report lists Owner as not carried and PageCategory '
    + '“IT” as skipped for a missing choice', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-dest-other');
    await page.fill('.wb-pc-url', DEST_SITE);
    await page.click('.wb-pc-connect');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    const reportText = await page.textContent('.wb-pc-report');
    await page.close();
    const ownerOk = /Owner/.test(reportText) && /people, lookup and managed metadata columns are not copied/i.test(reportText);
    const categoryOk = /Page ?Category/.test(reportText) && /does not exist on the destination/i.test(reportText);
    if (!ownerOk || !categoryOk) throw new Error(`report="${reportText.slice(0, 400)}"`);
    return true;
  });

  await check('dialog: a failed move step reports outcome failed with Discard copy visible; '
    + 'discarding recycles the staged page and the source page 7 survives', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await page.evaluate(() => {
      window.__PAGECOPY_MOCK_FAIL__ = [{
        match: 'MoveFileByPath', code: 'conflict', status: 400,
        message: 'The destination file already exists.', once: true,
      }];
    });
    await waitCheckEnabled(page);
    await page.click('.wb-pc-check');
    await waitReportVisible(page);
    await page.click('.wb-pc-go');
    await page.waitForSelector('.wb-pc-outcome:not([hidden])', { timeout: 15000 });
    const outcome = await page.getAttribute('.wb-pc-outcome', 'data-outcome');
    const discardHidden = await page.$eval('.wb-pc-discard', (el) => el.hidden);
    if (outcome !== 'failed' || discardHidden) {
      await page.close();
      throw new Error(`outcome=${outcome} discardHidden=${discardHidden}`);
    }

    await page.click('.wb-pc-discard');
    await page.waitForFunction(
      () => document.querySelector('.wb-pc-discard')?.hidden === true,
      null, { timeout: 10000 },
    );
    const state = await page.evaluate(() => {
      const pages = [...window.__PAGECOPY_MOCK__.state['/sites/pagesrc'].pages.values()];
      return {
        hasStagedCopy: pages.some((p) => String(p.dto?.FileName || '').includes('~copy-')),
        hasSourcePage: pages.some((p) => p.dto?.FileName === 'Quarterly-Update.aspx'),
      };
    });
    await page.close();
    if (state.hasStagedCopy || !state.hasSourcePage) {
      throw new Error(`hasStagedCopy=${state.hasStagedCopy} hasSourcePage=${state.hasSourcePage}`);
    }
    return true;
  });
}
