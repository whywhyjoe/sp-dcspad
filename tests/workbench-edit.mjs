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

await check('mock: picker, crumbs, filter and actions share one toolbar row', async () => {
  // The location bar is adopted into the grid toolbar, so all four sit on a
  // single line — asserted by geometry, since that is the point of it.
  const boxes = await page.evaluate(() => {
    const view = document.querySelector('.wb-view-files');
    const pick = (sel) => {
      const node = view.querySelector(sel);
      return node ? node.getBoundingClientRect().top : null;
    };
    return {
      inToolbar: Boolean(view.querySelector('.wb-grid-toolbar .wb-crumbs-bar')),
      select: pick('.wb-lib-select'),
      crumb: pick('.wb-crumb'),
      filter: pick('.wb-grid-filter'),
      upload: pick('.wb-grid-actions .wb-primary'),
    };
  });
  const tops = [boxes.select, boxes.crumb, boxes.filter, boxes.upload];
  return boxes.inToolbar && tops.every((t) => t !== null)
    && Math.max(...tops) - Math.min(...tops) < 8;
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

await check('mock: upload shows the pad-style dialog, greys the unavailable field', async () => {
  await page.setInputFiles('.wb-view-files input[type=file]', {
    name: 'hello.txt', mimeType: 'text/plain', buffer: Buffer.from('hello'),
  });
  await page.waitForSelector('.wb-upload-metadata');
  // Mock Documents library: Title and DocVersion exist, _ExtendedDescription
  // doesn't — exactly one greyed field with a disabled input.
  const unavailable = await page.locator('.wb-upload-metadata .sp-metadata-field.unavailable').count();
  const descDisabled = await page.locator('.wb-upload-metadata .wb-upload-meta-description').isDisabled();
  await page.fill('.wb-upload-metadata .wb-upload-meta-title', 'Hello note');
  await page.locator('.wb-upload-meta-go').click();
  await page.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “hello.txt” ✓'));
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const upload = writes.find((w) => w.url.includes('AddUsingPath') && w.url.includes('hello.txt'));
  const meta = writes.find((w) =>
    w.url.includes('hello.txt') && w.url.includes('ValidateUpdateListItem'));
  const metaBody = meta ? JSON.parse(meta.body) : null;
  await page.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return unavailable === 1 && descDisabled
    && upload?.url.includes('overwrite=false')
    && upload?.contentType === 'application/octet-stream'
    && metaBody?.formValues.length === 1
    && metaBody.formValues[0].FieldName === 'Title'
    && metaBody.formValues[0].FieldValue === 'Hello note'
    && metaBody.bNewDocumentUpdate === true;
});

await check('mock: cancelling the dialog before upload produces zero writes', async () => {
  const before = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.setInputFiles('.wb-view-files input[type=file]', {
    name: 'never.txt', mimeType: 'text/plain', buffer: Buffer.from('n'),
  });
  await page.waitForSelector('.wb-upload-metadata');
  await page.locator('.wb-upload-metadata .btn', { hasText: 'Cancel' }).click();
  const after = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  return (await page.locator('.wb-upload-metadata').count()) === 0 && after === before;
});

await check('mock: same-name upload asks consent, then prefills the dialog', async () => {
  const before = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('x'),
  });
  await page.waitForSelector('.wb-consent:not([hidden])');
  const consentText = await page.locator('.wb-consent').textContent();
  const during = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await page.waitForSelector('.wb-upload-metadata');
  const title = await page.locator('.wb-upload-metadata .wb-upload-meta-title').inputValue();
  const docVersion = await page.locator('.wb-upload-metadata .wb-upload-meta-docVersion').inputValue();
  await page.locator('.wb-upload-meta-go').click();
  await page.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “proposal.docx” ✓'));
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const uploadIndex = writes.findLastIndex((w) =>
    w.url.includes('AddUsingPath') && w.url.includes('proposal.docx'));
  const afterUpload = writes.slice(uploadIndex + 1);
  await page.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return consentText.includes('already exists')
    && during === before   // no write until consent
    && title === 'Project proposal' && docVersion === '1.4'   // prefilled
    && writes[uploadIndex].url.includes('overwrite=true')
    // untouched prefill → no metadata write after the upload
    && !afterUpload.some((w) => w.url.includes('ValidateUpdateListItem'));
});

await check('mock: new-folder prompt validates, then posts Folders/AddUsingPath', async () => {
  await page.locator('.wb-view-files .wb-newfolder').click();
  const before = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.fill('.wb-consent input', 'bad:name');
  await page.locator('.wb-consent .btn', { hasText: 'Create' }).click();
  await page.waitForSelector('.wb-consent-error');
  const rejected = await page.evaluate(() => (globalThis.__DCSPAD_WB_WRITES__ || []).length);
  await page.locator('.wb-view-files .wb-newfolder').click();
  await page.fill('.wb-consent input', 'Reports 2026');
  await page.locator('.wb-consent .btn', { hasText: 'Create' }).click();
  await page.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Created folder'));
  const writes = await page.evaluate(() => globalThis.__DCSPAD_WB_WRITES__ || []);
  const created = writes[writes.length - 1];
  await page.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return rejected === before   // invalid name never reaches the writer
    // odataPathLiteral percent-encodes the path inside the literal
    && decodeURIComponent(created.url)
      .includes("/_api/web/Folders/AddUsingPath(decodedUrl='/Shared Documents/Reports 2026')");
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
const prefillUrls = [];
const libraryLookups = [];
const flags = {
  failMetadata: false, racyConflictOnce: true, forceCheckout: false, bornCheckedOut: false,
  failNextOverwrite: false,
};
const checkOuts = [];      // { url, digest, order }
const checkIns = [];       // { url, digest, order }
// The stub keeps SharePoint's side of the bargain: a check-out stands — across
// the overwrite too — until CheckIn() releases it.
const STUB_HOLDER = { Id: 7, Title: 'Stub User', LoginName: 'i:0#.f|membership|stub@x' };
let freshFileCheckOutType = 2;
let callSeq = 0;           // orders a check-out against the upload it precedes

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
  // Underscore-prefixed internal names get an OData_ entity property — the
  // prefill $select must use it, not the internal name.
  { Id: 'f3', Title: 'Description', InternalName: '_ExtendedDescription', EntityPropertyName: 'OData__ExtendedDescription', TypeAsString: 'Note', FieldTypeKind: 3, Required: false, Hidden: false, ReadOnlyField: false },
];

await live.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = {
    webAbsoluteUrl: location.origin,
    userDisplayName: 'Stub User',
    // The login claim is what identifies "checked out to me".
    userLoginName: 'i:0#.f|membership|stub@x',
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
      order: ++callSeq,
    };
    uploads.push(record);
    if (url.includes('racy.bin') && url.includes('overwrite=false') && flags.racyConflictOnce) {
      flags.racyConflictOnce = false;
      return route.fulfill({
        status: 409,
        json: { 'odata.error': { message: { value: 'The file already exists.' } } },
      });
    }
    if (flags.failNextOverwrite && url.includes('overwrite=true')) {
      flags.failNextOverwrite = false;
      return route.fulfill({
        status: 500,
        json: { 'odata.error': { message: { value: 'Upload stub failure.' } } },
      });
    }
    const name = decodeURIComponent(/AddUsingPath\(decodedUrl='([^']*)'/.exec(url)?.[1] || 'file');
    let checkOutType = name === 'proposal.docx' ? LIVE_FILES[0].CheckOutType : 2;
    if (flags.bornCheckedOut && url.includes('overwrite=false')) {
      freshFileCheckOutType = 0;
      checkOutType = 0;
    }
    return route.fulfill({
      json: { ServerRelativeUrl: `/Shared Documents/${name}`, CheckOutType: checkOutType },
    });
  }
  if (url.includes('/CheckOut()')) {
    checkOuts.push({
      url,
      digest: request.headers()['x-requestdigest'] || '',
      order: ++callSeq,
    });
    LIVE_FILES[0].CheckOutType = 0;
    LIVE_FILES[0].CheckedOutByUser = STUB_HOLDER;
    return route.fulfill({ json: {} });
  }
  if (url.includes('/CheckIn(')) {
    checkIns.push({
      url,
      digest: request.headers()['x-requestdigest'] || '',
      order: ++callSeq,
    });
    if (url.includes('proposal.docx')) {
      LIVE_FILES[0].CheckOutType = 2;
      delete LIVE_FILES[0].CheckedOutByUser;
    } else {
      freshFileCheckOutType = 2;
    }
    return route.fulfill({ json: {} });
  }
  if (/lists\(guid'/i.test(url)
      && new URL(url).searchParams.get('$select') === 'ForceCheckout') {
    return route.fulfill({ json: { ForceCheckout: flags.forceCheckout } });
  }
  if (url.includes('GetFileByServerRelativePath(')
      && (new URL(url).searchParams.get('$select') || '').includes('CheckOutType')) {
    return route.fulfill({
      json: url.includes('proposal.docx') ? LIVE_FILES[0] : { CheckOutType: freshFileCheckOutType },
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
    prefillUrls.push(url);
    return route.fulfill({ json: {
      Id: 7, Title: 'Proposal', DocCategory: 'Report', OData__ExtendedDescription: 'Live desc',
    } });
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

await check('live: upload dialog shows library availability, uploads with digest', async () => {
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'new.bin', mimeType: 'application/octet-stream', buffer: Buffer.from([1, 2, 3]),
  });
  await live.waitForSelector('.wb-upload-metadata');
  // The live library exposes Title + Description — DocVersion greyed.
  const available = await live.locator('.wb-upload-metadata .sp-metadata-field.available').count();
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “new.bin” ✓'));
  const upload = uploads[0];
  return available === 2
    && uploads.length === 1
    && upload.url.includes("AddUsingPath(decodedUrl='new.bin',overwrite=false)")
    && upload.digest === 'WB-DIGEST'
    && upload.bodyLength === 3
    && vuliCalls.length === 0;   // nothing typed → no metadata write
});

await check('live: metadata failure keeps the file and retry re-posts only metadata', async () => {
  flags.failMetadata = true;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'meta.bin', mimeType: 'application/octet-stream', buffer: Buffer.from('m'),
  });
  await live.waitForSelector('.wb-upload-metadata');
  await live.fill('.wb-upload-metadata .wb-upload-meta-title', 'New title');
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForSelector('.wb-upload-metadata .sp-files-error:not([hidden])');
  const errorText = await live.locator('.wb-upload-metadata .sp-files-error').textContent();
  const retryLabel = await live.locator('.wb-upload-meta-go').textContent();
  const uploadsAfterFail = uploads.length;
  flags.failMetadata = false;
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “meta.bin” ✓'));
  const retryCall = vuliCalls[vuliCalls.length - 1];
  return errorText.includes('could not be saved')
    && retryLabel.includes('Retry metadata')
    && vuliCalls.length === 2
    && uploads.length === uploadsAfterFail   // retry never re-uploads
    && retryCall.url.includes('GetFileByServerRelativePath(')
    && retryCall.body.formValues[0].FieldValue === 'New title'
    && retryCall.body.bNewDocumentUpdate === true
    && retryCall.digest === 'WB-DIGEST';
});

await check('live: keep-without-metadata closes the dialog and keeps the file', async () => {
  flags.failMetadata = true;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'kept.bin', mimeType: 'application/octet-stream', buffer: Buffer.from('k'),
  });
  await live.waitForSelector('.wb-upload-metadata');
  await live.fill('.wb-upload-metadata .wb-upload-meta-title', 'K');
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForSelector('.wb-upload-metadata .wb-upload-meta-keep:not([hidden])');
  // After the upload exists, Esc and ✕ must NOT silently mean "keep" —
  // only the explicit buttons decide.
  await live.keyboard.press('Escape');
  const stillOpen = (await live.locator('.wb-upload-metadata').count()) === 1;
  const closeHidden = await live.locator('.wb-upload-metadata .app-dialog__head .btn').isHidden();
  await live.locator('.wb-upload-meta-keep').click();
  flags.failMetadata = false;
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('kept without metadata'));
  return stillOpen && closeHidden
    && (await live.locator('.wb-upload-metadata').count()) === 0;
});

await check('live: a same-name upload asks first and prefills from the live file', async () => {
  const before = uploads.length;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('zz'),
  });
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent:not([hidden])')?.textContent.includes('already exists'));
  const during = uploads.length;
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  const prefill = await live.locator('.wb-upload-metadata .wb-upload-meta-title').inputValue();
  const descPrefill = await live.locator('.wb-upload-metadata .wb-upload-meta-description').inputValue();
  await live.locator('.wb-upload-meta-go').click();
  const landed = await until(() => uploads.length > during);
  const replaced = uploads[uploads.length - 1];
  // Clear the success notice so the next check's consent-wait can't race it.
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “proposal.docx” ✓'));
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  // The prefill $select must address _ExtendedDescription by its entity
  // property (OData__ExtendedDescription), or live tenants 400 the request.
  const prefillSelect = prefillUrls.find((u) =>
    u.includes('proposal.docx') && u.includes('$select='));
  return landed && during === before
    && prefill === 'Proposal'          // read from the file being replaced
    && descPrefill === 'Live desc'     // via the entity property name
    && Boolean(prefillSelect) && prefillSelect.includes('OData__ExtendedDescription')
    && replaced.url.includes('overwrite=true');
});

await check('live: a 409 race re-consents and the retry keeps the typed values', async () => {
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'racy.bin', mimeType: 'application/octet-stream', buffer: Buffer.from('r'),
  });
  await live.waitForSelector('.wb-upload-metadata');
  await live.fill('.wb-upload-metadata .wb-upload-meta-title', 'Racy T');
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent:not([hidden])')?.textContent.includes('already exists'));
  const consentText = await live.locator('.wb-consent').textContent();
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  const carried = await live.locator('.wb-upload-metadata .wb-upload-meta-title').inputValue();
  await live.locator('.wb-upload-meta-go').click();
  const landed = await until(() =>
    uploads.filter((u) => u.url.includes('racy.bin')).length === 2);
  const attempts = uploads.filter((u) => u.url.includes('racy.bin'));
  // The carried value must not just DISPLAY — it must WRITE. The diff runs
  // against the replaced file's baseline, so 'Racy T' posts as metadata.
  const wrote = await until(() => vuliCalls.some((c) =>
    c.url.includes('racy.bin')
    && c.body.formValues?.some((fv) => fv.FieldName === 'Title' && fv.FieldValue === 'Racy T')));
  return landed && wrote && consentText.includes('already exists')
    && carried === 'Racy T'   // values survive the race retry
    && attempts[0].url.includes('overwrite=false')
    && attempts[1].url.includes('overwrite=true');
});

// ---- forced check-out ------------------------------------------------------
// Same contract as the pad's export: replacing a file in a ForceCheckout
// library is consented to, then checked out, uploaded, and checked back in.
flags.forceCheckout = true;

await check('live: a forced-check-out library gates Replace behind a consent box', async () => {
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('c1'),
  });
  await live.waitForSelector('.wb-consent .wb-consent-checkout');
  const gateText = await live.locator('.wb-consent-gate').textContent();
  const blocked = await live.locator('.wb-consent .btn', { hasText: 'Replace' }).isDisabled();
  await live.check('.wb-consent .wb-consent-checkout');
  const enabled = !(await live.locator('.wb-consent .btn', { hasText: 'Replace' }).isDisabled());
  return gateText.includes('requires check-out')
    && gateText.includes('before replacing it')
    && blocked && enabled;
});

await check('live: consenting checks the file out before the overwrite upload', async () => {
  const before = uploads.length;
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  await live.locator('.wb-upload-meta-go').click();
  const landed = await until(() => uploads.length > before);
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Uploaded “proposal.docx” ✓'));
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  const replaced = uploads[uploads.length - 1];
  return landed
    && checkOuts.length === 1
    && decodeURIComponent(checkOuts[0].url).includes("decodedUrl='/Shared Documents/proposal.docx'")
    && checkOuts[0].digest === 'WB-DIGEST'
    && checkOuts[0].order < replaced.order
    && replaced.url.includes('overwrite=true');
});

await check('live: the replaced file is checked back in after the upload', async () => {
  const replaced = uploads[uploads.length - 1];
  return checkIns.length === 1
    && decodeURIComponent(checkIns[0].url).includes("decodedUrl='/Shared Documents/proposal.docx'")
    && /checkintype=0/i.test(checkIns[0].url)
    && checkIns[0].digest === 'WB-DIGEST'
    && checkIns[0].order > replaced.order
    && LIVE_FILES[0].CheckOutType === 2;
});

await check('live: a file checked out to someone else is refused, not consented to', async () => {
  LIVE_FILES[0].CheckOutType = 1;
  LIVE_FILES[0].CheckedOutByUser = { Id: 42, Title: 'Dana Lee', LoginName: 'i:0#.f|membership|dana@x' };
  await live.locator('.wb-view-files .btn', { hasText: 'Refresh' }).click();
  await live.waitForTimeout(200);
  const before = uploads.length;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('c2'),
  });
  await live.waitForSelector('.wb-consent-error');
  const text = await live.locator('.wb-consent').textContent();
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return text.includes('checked out to Dana Lee')
    && uploads.length === before
    && checkOuts.length === 1;   // and nothing was checked out
});

await check('live: a file you already hold is replaced without a second check-out', async () => {
  LIVE_FILES[0].CheckedOutByUser = {
    Id: 7, Title: 'Stub User', LoginName: 'i:0#.f|membership|stub@x',
  };
  await live.locator('.wb-view-files .btn', { hasText: 'Refresh' }).click();
  await live.waitForTimeout(200);
  const before = uploads.length;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('c3'),
  });
  await live.waitForSelector('.wb-consent .wb-consent-checkout');
  const gateText = await live.locator('.wb-consent-gate').textContent();
  await live.check('.wb-consent .wb-consent-checkout');
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  await live.locator('.wb-upload-meta-go').click();
  const landed = await until(() => uploads.length > before);
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('Checked in.'));
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  return landed
    && gateText.includes('already checked out to you')
    && checkOuts.length === 1    // no second CheckOut()
    && checkIns.length === 2     // but it is checked back in
    && LIVE_FILES[0].CheckOutType === 2;
});

await check('live: a new file born checked out is checked in, with no CheckOut call', async () => {
  flags.bornCheckedOut = true;
  const before = checkIns.length;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'fresh-checkout.bin', mimeType: 'application/octet-stream', buffer: Buffer.from('n1'),
  });
  await live.waitForSelector('.wb-upload-metadata');
  await live.locator('.wb-upload-meta-go').click();
  // A dismissed notice keeps its text, so wait on this file's own notice.
  await live.waitForFunction(() => {
    const text = document.querySelector('.wb-consent')?.textContent || '';
    return text.includes('fresh-checkout.bin') && text.includes('Checked in.');
  });
  flags.bornCheckedOut = false;
  return checkIns.length === before + 1
    && decodeURIComponent(checkIns.at(-1).url).includes('fresh-checkout.bin')
    && checkOuts.length === 1
    && freshFileCheckOutType === 2;
});

// The upload fails after the Workbench's own check-out. The listing must be
// refreshed: read stale, it still says "checked in", and the next attempt
// would post a second CheckOut(), which SharePoint rejects.
await check('live: an upload failing after the check-out is stated, and the retry does not check out twice', async () => {
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();
  const checkOutsBefore = checkOuts.length;
  const checkInsBefore = checkIns.length;
  flags.failNextOverwrite = true;
  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('f1'),
  });
  await live.waitForSelector('.wb-consent .wb-consent-checkout');
  await live.check('.wb-consent .wb-consent-checkout');
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  await live.locator('.wb-upload-meta-go').click();
  await live.waitForFunction(() =>
    document.querySelector('.wb-consent')?.textContent.includes('still checked out to you'));
  const failedOnce = checkOuts.length === checkOutsBefore + 1
    && LIVE_FILES[0].CheckOutType === 0;
  await live.locator('.wb-consent .btn', { hasText: 'Dismiss' }).click();

  await live.setInputFiles('.wb-view-files input[type=file]', {
    name: 'proposal.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('f2'),
  });
  await live.waitForSelector('.wb-consent .wb-consent-checkout');
  const gateText = await live.locator('.wb-consent-gate').textContent();
  await live.check('.wb-consent .wb-consent-checkout');
  await live.locator('.wb-consent .btn', { hasText: 'Replace' }).click();
  await live.waitForSelector('.wb-upload-metadata');
  await live.locator('.wb-upload-meta-go').click();
  await until(() => checkIns.length > checkInsBefore);
  return failedOnce
    && gateText.includes('already checked out to you')
    && checkOuts.length === checkOutsBefore + 1   // no second CheckOut()
    && LIVE_FILES[0].CheckOutType === 2;
});

await live.close();
await browser.close();
exitWithResult();
