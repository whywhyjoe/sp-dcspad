// SP Workbench — page copy: the per-part drop control (design/PAGE-COPY.md
// §1 decision 4, §5.2 "dynamic" and "unavailable"). Same recipe as
// tests/pages-copy/dialog.mjs: drives the real dialog markup
// (src/workbench/page-copy-dialog.js class hooks) through Playwright against
// the mock /sites/pagesrc + /sites/pagedst webs (src/workbench/mock-pagecopy.js).

const SOURCE_SITE = '/sites/pagesrc';
const DEST_SITE = '/sites/pagedst';

// Fixture instance ids (src/workbench/mock-pagecopy.js, Quarterly-Update.aspx
// canvas) this suite drops or asserts against.
const TEXT_INSTANCE = 'd15e0000-0000-4000-8000-000000000001';
const TITLE_AREA_INSTANCE = 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788';
const VIDEO_INSTANCE = 'd15e0000-0000-4000-8000-00000000000d';
const DOC_LIBRARY_INSTANCE = 'd15e0000-0000-4000-8000-00000000000a';
const FILE_VIEWER_INSTANCE = 'd15e0000-0000-4000-8000-000000000009';
const QUICK_CHART_INSTANCE = 'd15e0000-0000-4000-8000-000000000007';

async function openDialog(browser, WB_URL, pageName) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
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

async function checkAndWait(page) {
  await page.click('.wb-pc-check');
  await waitReportVisible(page);
}

async function dropInstance(page, instanceId) {
  await page.click(`.wb-pc-drop[data-instance="${instanceId}"]`);
}

export async function run({ browser, check, WB_URL }) {
  await check('drop: every web-part row gets a Drop checkbox; Text and the title area get none', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await checkAndWait(page);
    const counts = await page.evaluate(({ text, title, video }) => ({
      video: document.querySelectorAll(`.wb-pc-drop[data-instance="${video}"]`).length,
      text: document.querySelectorAll(`.wb-pc-drop[data-instance="${text}"]`).length,
      titleArea: document.querySelectorAll(`.wb-pc-drop[data-instance="${title}"]`).length,
      totalWebparts: document.querySelectorAll('.wb-pc-drop').length,
    }), { text: TEXT_INSTANCE, title: TITLE_AREA_INSTANCE, video: VIDEO_INSTANCE });
    await page.close();
    if (counts.video !== 1 || counts.text !== 0 || counts.titleArea !== 0) {
      throw new Error(`counts=${JSON.stringify(counts)}`);
    }
    return counts.totalWebparts > 0;
  });

  await check('drop: toggling Drop darkens Copy synchronously, keeps the report visible, and shows the recheck note', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await checkAndWait(page);
    const goEnabledAfterCheck = await page.$eval('.wb-pc-go', (el) => !el.disabled);

    await dropInstance(page, VIDEO_INSTANCE);
    const state = await page.evaluate(() => ({
      goDisabled: document.querySelector('.wb-pc-go')?.disabled,
      reportHidden: document.querySelector('.wb-pc-report')?.hidden,
      noteHidden: document.querySelector('.wb-pc-recheck-note')?.hidden,
    }));
    await page.close();
    if (!goEnabledAfterCheck) throw new Error('Copy was not enabled after the first clean Check');
    return state.goDisabled === true && state.reportHidden === false && state.noteHidden === false;
  });

  await check('drop: re-Check then Copy removes exactly the dropped control from the saved canvas and keeps every other one', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await checkAndWait(page);
    const sourceIds = await page.evaluate(() => {
      const src = window.__PAGECOPY_MOCK__.state['/sites/pagesrc'];
      const dto = src.pages.get(7).dto;
      return JSON.parse(dto.CanvasContent1).filter((c) => c && c.id).map((c) => c.id);
    });

    await dropInstance(page, VIDEO_INSTANCE);
    await checkAndWait(page);
    const goDisabled = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (goDisabled) { await page.close(); throw new Error('Copy stayed disabled after re-Check with no pending consumers'); }

    await page.click('.wb-pc-go');
    await page.waitForSelector('.wb-pc-outcome:not([hidden])', { timeout: 15000 });
    const outcome = await page.getAttribute('.wb-pc-outcome', 'data-outcome');
    const destIds = await page.evaluate(() => {
      const src = window.__PAGECOPY_MOCK__.state['/sites/pagesrc'];
      const copy = [...src.pages.values()].find((p) => p.dto?.FileLeafRef === 'Quarterly-Update-copy.aspx');
      return JSON.parse(copy.dto.CanvasContent1).filter((c) => c && c.id).map((c) => c.id);
    });
    await page.close();

    if (outcome !== 'done') throw new Error(`outcome=${outcome}`);
    const expected = sourceIds.filter((id) => id !== VIDEO_INSTANCE).sort();
    const actual = [...destIds].sort();
    if (destIds.includes(VIDEO_INSTANCE)) throw new Error('dropped control survived into the saved canvas');
    return JSON.stringify(actual) === JSON.stringify(expected);
  });

  await check('drop: dropping the Document library provider lists the File viewer as a consumer; Copy stays disabled until it is confirmed', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await checkAndWait(page);

    await dropInstance(page, DOC_LIBRARY_INSTANCE);
    await checkAndWait(page);

    const consumerText = await page.evaluate((fileViewer) => {
      const cb = document.querySelector(`.wb-pc-drop-consumer[data-instance="${fileViewer}"]`);
      return cb ? cb.closest('label')?.textContent || '' : null;
    }, FILE_VIEWER_INSTANCE);
    const goDisabledBefore = await page.$eval('.wb-pc-go', (el) => el.disabled);
    if (consumerText === null || !/File viewer/.test(consumerText) || !goDisabledBefore) {
      await page.close();
      throw new Error(`consumerText=${consumerText} goDisabledBefore=${goDisabledBefore}`);
    }

    await page.click(`.wb-pc-drop-consumer[data-instance="${FILE_VIEWER_INSTANCE}"]`);
    const goDisabledAfter = await page.$eval('.wb-pc-go', (el) => el.disabled);
    await page.close();
    return goDisabledAfter === false;
  });

  await check('drop: an unavailable part (Quick chart on /sites/pagedst) can be dropped; its warning disappears after re-Check', async () => {
    const page = await openDialog(browser, WB_URL, 'Quarterly-Update.aspx');
    await waitCheckEnabled(page);
    await page.click('.wb-pc-dest-other');
    await page.fill('.wb-pc-url', DEST_SITE);
    await page.click('.wb-pc-connect');
    await waitCheckEnabled(page);
    await checkAndWait(page);

    const gapInput = page.locator('.wb-pc-gaps input, .wb-pc-gaps textarea, .wb-pc-gaps select').first();
    await gapInput.fill('All employees');
    await checkAndWait(page);

    const before = await page.evaluate(() => ({
      dropLabel: document.querySelector('.wb-pc-drop[data-instance="' + 'd15e0000-0000-4000-8000-000000000007' + '"]')
        ?.closest('label')?.textContent || '',
      warnings: document.querySelector('.wb-pc-report')?.textContent || '',
    }));
    if (!/Quick chart/.test(before.warnings) || !/not available on the destination/.test(before.dropLabel)) {
      await page.close();
      throw new Error(`before=${JSON.stringify(before)}`);
    }

    await dropInstance(page, QUICK_CHART_INSTANCE);
    await checkAndWait(page);
    const after = await page.evaluate(() => document.querySelector('.wb-pc-report')?.textContent || '');
    const goDisabled = await page.$eval('.wb-pc-go', (el) => el.disabled);
    await page.close();
    if (/Quick chart is not available/.test(after)) throw new Error(`warning still present: ${after.slice(0, 300)}`);
    return goDisabled === false;
  });
}
