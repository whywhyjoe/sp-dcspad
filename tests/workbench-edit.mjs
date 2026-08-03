// SP Workbench edit-path suite: the Files browser and the sp-write client —
// mock-mode listing/metadata/upload first, then the live path with an
// injected host context and stubbed /_api write endpoints (the files.mjs
// pattern: contextinfo digests, AddUsingPath capture, ValidateUpdateListItem
// capture, and negative switches).

import { writeFileSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const WB_URL = process.env.DCSPAD_WORKBENCH_URL
  || APP_URL.replace(/index\.html.*$/, 'workbench.html');

// Poll a Node-side condition (e.g. a stubbed-request log) — page.waitFor*
// can't see these, and the UI often repaints before the request lands.
async function until(fn, timeout = 8000) {
  const start = Date.now();
  while (!fn()) {
    if (Date.now() - start > timeout) return false;
    await new Promise((r) => setTimeout(r, 50));
  }
  return true;
}

const browser = await launchBrowser();

// ---- mock mode ------------------------------------------------------------

const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto(WB_URL);
await page.waitForSelector('.wb-home-cards');

await check('mock: Files view lists every file type, folders first', async () => {
  await page.locator('.wb-rail-btn', { hasText: 'Files' }).click();
  await page.waitForSelector('.wb-view-files .wb-table tbody tr');
  const rows = await page.locator('.wb-view-files .wb-table tbody tr').allTextContents();
  return rows.length === 7
    && rows[0].includes('Reports')
    && rows.some((r) => r.includes('archive.zip'))
    && rows.some((r) => r.includes('logo.png'))
    && rows.some((r) => r.includes('proposal.docx'));
});

await check('mock: size column sorts numerically', async () => {
  await page.locator('.wb-view-files th', { hasText: 'Size' }).click();
  const first = await page.locator('.wb-view-files .wb-table tbody tr').first().textContent();
  await page.locator('.wb-view-files th', { hasText: 'Size' }).click();
  const afterDesc = await page.locator('.wb-view-files .wb-table tbody tr').first().textContent();
  // Reset sort back to name order for later checks.
  await page.locator('.wb-view-files th', { hasText: 'Name' }).click();
  return first.includes('notes.txt') && afterDesc.includes('archive.zip');
});

await check('mock: breadcrumbs navigate into and out of folders', async () => {
  await page.locator('.wb-view-files .wb-table tbody tr', { hasText: 'Reports' })
    .locator('td').first().click();
  await page.waitForSelector('.wb-view-files .wb-table tbody tr', { hasText: 'q1-report' });
  const inFolder = await page.locator('.wb-view-files .wb-table tbody tr').count();
  await page.locator('.wb-crumb', { hasText: 'Shared Documents' }).click();
  await page.waitForSelector('.wb-view-files .wb-table tbody tr', { hasText: 'proposal.docx' });
  const backOut = await page.locator('.wb-view-files .wb-table tbody tr').count();
  return inFolder === 2 && backOut === 7;
});

await check('mock: library picker jumps to a library root', async () => {
  const options = await page.locator('.wb-lib-select option').allTextContents();
  return options.includes('Documents') && options.includes('Site Assets');
});

await check('mock: every file row offers download and copy-direct-URL actions', async () => {
  const downloads = await page.locator('.wb-view-files .wb-file-actions a').count();
  const copies = await page.locator('.wb-view-files .wb-cell-copylink').count();
  const firstTitle = await page.locator('.wb-view-files .wb-cell-copylink').first()
    .getAttribute('title');
  return downloads === 6 && copies === 6 && firstTitle.includes('direct URL');
});

await check('mock: file metadata panel renders per-type editors with values', async () => {
  await page.locator('.wb-view-files .wb-table tbody tr', { hasText: 'proposal.docx' })
    .locator('td').first().click();
  await page.waitForSelector('.wb-file-meta .wb-editor-row');
  const choice = await page.locator('.wb-file-meta .wb-editor-row[data-internal="DocCategory"] select')
    .inputValue();
  const bool = await page.locator('.wb-file-meta .wb-editor-row[data-internal="Confidential"] input')
    .isChecked();
  const url = await page.locator('.wb-file-meta .wb-editor-row[data-internal="SourceLink"] input')
    .first().inputValue();
  const authorRo = await page.locator('.wb-file-meta .wb-editor-row[data-internal="Author"]')
    .getAttribute('class');
  return choice === 'Contract' && bool === true
    && url === 'https://example.com/spec' && authorRo.includes('readonly');
});

await check('mock: metadata save posts the file-path ValidateUpdateListItem', async () => {
  await page.fill('.wb-file-meta .wb-editor-row[data-internal="Title"] input', 'Retitled');
  await page.locator('.wb-file-meta .wb-editor-bar .btn').click();
  await page.waitForSelector('.wb-file-meta .wb-editor-status.wb-editor-saved');
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const write = writes[writes.length - 1];
  const body = JSON.parse(write.body);
  return write.url.includes('GetFileByServerRelativePath(')
    && write.url.includes('/ListItemAllFields/ValidateUpdateListItem')
    && body.bNewDocumentUpdate === true
    && body.formValues[0].FieldName === 'Title';
});

await check('mock: uploading a new file routes through the mock writer', async () => {
  await page.setInputFiles('.wb-view-files input[type=file]', {
    name: 'hello.txt', mimeType: 'text/plain', buffer: Buffer.from('hello'),
  });
  await page.waitForSelector('.wb-file-meta .wb-subpanel-title', { hasText: 'Uploaded' });
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const upload = writes.find((w) => w.url.includes('AddUsingPath'));
  return Boolean(upload)
    && upload.url.includes("overwrite=false")
    && upload.contentType === 'application/octet-stream';
});

await check('mock: same-name upload asks for overwrite consent first', async () => {
  const before = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('x'),
  });
  await page.waitForSelector('.wb-consent:not([hidden])');
  const consentText = await page.locator('.wb-consent').textContent();
  const during = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await page.waitForFunction(
    (n) => (globalThis.__DCSPAD_WB_WRITES__ || []).length > n, during,
  );
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const replaced = writes[writes.length - 1];
  return consentText.includes('already exists')
    && during === before   // no write until consent
    && replaced.url.includes('overwrite=true');
});

await check('mock: oversized uploads are rejected client-side with no write', async () => {
  const before = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  // Playwright caps in-memory buffers at 50 MB — stage the file on disk.
  const hugePath = join(tmpdir(), 'dcspad-huge-upload.bin');
  writeFileSync(hugePath, Buffer.alloc(50 * 1024 * 1024 + 1));
  try {
    await page.setInputFiles('.wb-view-files input[type=file]', hugePath);
    await page.waitForSelector('.wb-consent-error');
  } finally {
    rmSync(hugePath, { force: true });
  }
  const text = await page.locator('.wb-consent').textContent();
  const after = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return text.includes('upload limit') && after === before;
});

await page.close();

// ---- live path (injected context + stubbed /_api writes) ------------------

const live = await browser.newPage({ viewport: { width: 1400, height: 900 } });

const LIB_ID = 'ab12cd34-0000-4000-8000-00000000aa01';
const uploads = [];        // { url, digest, bodyLength }
const vuliCalls = [];      // { url, digest, body }
const libraryLookups = [];
const flags = { failMetadata: false, racyConflictOnce: true };

const LIVE_FILES = [
  {
    Name: 'proposal.docx',
    ServerRelativeUrl: '/Shared Documents/proposal.docx',
    Length: 48230,
    TimeLastModified: '2026-07-10T09:00:00Z',
    UIVersionLabel: '2.0',
    CheckOutType: 2,
  },
];

const LIVE_FIELDS = [
  { Id: 'f1', Title: 'Title', InternalName: 'Title', TypeAsString: 'Text', FieldTypeKind: 2, Required: false, Hidden: false, ReadOnlyField: false },
  { Id: 'f2', Title: 'Category', InternalName: 'DocCategory', TypeAsString: 'Choice', FieldTypeKind: 6, Required: false, Hidden: false, ReadOnlyField: false, Choices: ['Contract', 'Report'] },
];

await live.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = {
    webAbsoluteUrl: location.origin,
    userDisplayName: 'Stub User',
  };
});

await live.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();

  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: {
        FormDigestValue: 'WB-DIGEST',
        FormDigestTimeoutSeconds: 1800,
        WebFullUrl: new URL(url).origin,
      },
    });
  }
  if (url.includes('/Files/AddUsingPath(')) {
    const record = {
      url,
      digest: request.headers()['x-requestdigest'] || '',
      bodyLength: request.postDataBuffer()?.length ?? 0,
    };
    uploads.push(record);
    if (url.includes('racy.bin') && url.includes('overwrite=false') && flags.racyConflictOnce) {
      flags.racyConflictOnce = false;
      return route.fulfill({
        status: 409,
        json: { 'odata.error': { message: { value: 'The file already exists.' } } },
      });
    }
    const name = /AddUsingPath\(decodedUrl='([^']*)'/.exec(url)?.[1] || 'file';
    return route.fulfill({
      json: { ServerRelativeUrl: `/Shared Documents/${decodeURIComponent(name)}` },
    });
  }
  if (url.includes('/ValidateUpdateListItem')) {
    const body = JSON.parse(request.postData() || '{}');
    vuliCalls.push({ url, digest: request.headers()['x-requestdigest'] || '', body });
    const results = (body.formValues || []).map((fv) => ({
      FieldName: fv.FieldName,
      HasException: flags.failMetadata,
      ErrorMessage: flags.failMetadata ? 'The server said no.' : null,
    }));
    return route.fulfill({ json: { value: results } });
  }
  if (url.includes('/fields') && url.includes("lists(guid'")) {
    return route.fulfill({ json: { value: LIVE_FIELDS } });
  }
  if (url.includes('/_api/web/GetList(@listUrl)')) {
    libraryLookups.push(url);
    return route.fulfill({ json: { Id: LIB_ID } });
  }
  if (url.includes('GetFileByServerRelativePath(') && url.includes('/ListItemAllFields')) {
    return route.fulfill({ json: { Id: 7, Title: 'Proposal', DocCategory: 'Report' } });
  }
  if (url.includes('GetFolderByServerRelativePath(')) {
    if (url.includes('/Folders')) {
      return route.fulfill({ json: { value: [] } });
    }
    if (url.includes('/Files')) {
      return route.fulfill({ json: { value: LIVE_FILES } });
    }
    return route.fulfill({
      json: { ListItemAllFields: null },
    });
  }
  if (url.includes('/_api/web/lists')) {
    return route.fulfill({
      json: {
        value: [{
          Id: LIB_ID,
          Title: 'Documents',
          BaseType: 1,
          Hidden: false,
          BaseTemplate: 101,
          RootFolder: { ServerRelativeUrl: '/Shared Documents' },
        }],
      },
    });
  }
  return route.fulfill({ json: { value: [] } });
});

await live.goto(WB_URL);
await live.waitForSelector('.wb-home-cards');
await live.locator('.wb-rail-btn', { hasText: 'Files' }).click();
await live.waitForSelector('.wb-view-files .wb-table tbody tr', { hasText: 'proposal.docx' });

await check('live: root-library files resolve metadata through GetList', async () => {
  await live.locator('.wb-view-files .wb-table tbody tr', { hasText: 'proposal.docx' })
    .locator('td').first().click();
  await live.waitForSelector('.wb-file-meta .wb-editor-row[data-internal="Title"]');
  const title = await live.locator('.wb-file-meta .wb-editor-row[data-internal="Title"] input')
    .inputValue();
  await live.locator('.wb-file-meta .wb-file-meta-head .btn', { hasText: 'Close' }).click();
  return title === 'Proposal'
    && libraryLookups.length === 1
    && new URL(libraryLookups[0]).searchParams.get('@listUrl') === "'/Shared Documents'";
});

await check('live: binary upload posts AddUsingPath with a digest and the raw bytes', async () => {
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'new.bin', mimeType: 'application/octet-stream', buffer: Buffer.from([1, 2, 3]),
  });
  await live.waitForSelector('.wb-file-meta .wb-subpanel-title', { hasText: 'Uploaded' });
  const upload = uploads[0];
  return uploads.length === 1
    && upload.url.includes("AddUsingPath(decodedUrl='new.bin',overwrite=false)")
    && upload.digest === 'WB-DIGEST'
    && upload.bodyLength === 3;
});

await check('live: metadata failure keeps the file and retry re-posts only metadata', async () => {
  flags.failMetadata = true;
  await live.fill('.wb-file-meta .wb-editor-row[data-internal="Title"] input', 'New title');
  await live.locator('.wb-file-meta .wb-editor-bar .btn').click();
  await live.waitForSelector('.wb-file-meta .wb-editor-status.wb-editor-failed');
  const fieldError = await live
    .locator('.wb-file-meta .wb-editor-row[data-internal="Title"] .wb-editor-error')
    .textContent();
  const uploadsAfterFail = uploads.length;
  flags.failMetadata = false;
  await live.locator('.wb-file-meta .wb-editor-bar .btn').click();
  await live.waitForSelector('.wb-file-meta .wb-editor-status.wb-editor-saved');
  const retryCall = vuliCalls[vuliCalls.length - 1];
  return fieldError.includes('The server said no')
    && vuliCalls.length === 2
    && uploads.length === uploadsAfterFail   // retry never re-uploads
    && retryCall.url.includes('GetFileByServerRelativePath(')
    && retryCall.body.bNewDocumentUpdate === true
    && retryCall.digest === 'WB-DIGEST';
});

await check('live: keep-without-metadata dismisses the panel', async () => {
  const closeLabel = await live.locator('.wb-file-meta .wb-file-meta-head .btn').textContent();
  await live.locator('.wb-file-meta .wb-file-meta-head .btn').click();
  const hidden = await live.locator('.wb-file-meta').isHidden();
  return closeLabel.includes('Keep without metadata') && hidden;
});

await check('live: a same-name upload asks before replacing', async () => {
  const before = uploads.length;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('zz'),
  });
  await live.waitForSelector('.wb-consent:not([hidden])');
  const during = uploads.length;
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  const landed = await until(() => uploads.length > during);
  const replaced = uploads[uploads.length - 1];
  return landed && during === before && replaced.url.includes('overwrite=true');
});

await check('live: a 409 race surfaces the same consent and retries with overwrite', async () => {
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'racy.bin', mimeType: 'application/octet-stream', buffer: Buffer.from('r'),
  });
  await live.waitForSelector('.wb-consent:not([hidden])');
  const consentText = await live.locator('.wb-consent').textContent();
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  const landed = await until(() =>
    uploads.filter((u) => u.url.includes('racy.bin')).length === 2);
  const attempts = uploads.filter((u) => u.url.includes('racy.bin'));
  return landed && consentText.includes('already exists')
    && attempts[0].url.includes('overwrite=false')
    && attempts[1].url.includes('overwrite=true');
});

await live.close();
await browser.close();
exitWithResult();
