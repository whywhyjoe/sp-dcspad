// Local single-file import + SharePoint document-library transfer.

import { writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const browser = await launchBrowser();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
await page.goto(APP_URL);
await page.evaluate(() => localStorage.clear());
await page.reload();
await page.waitForSelector('.monaco-editor');

const editorText = () =>
  page.locator('#pane-editor .view-lines').textContent()
    .then((text) => text.replaceAll('\u00a0', ' '));

const htmlPath = join(tmpdir(), 'DCSPad-Import-Probe.HTML');
const cssPath = join(tmpdir(), 'dcspad-share-export.css');
writeFileSync(htmlPath, '<main id="local-import-marker">Local import</main>\n');
writeFileSync(cssPath, '.share-export-marker { color: teal; }\n');

// Local import: one picker maps by extension and never mutates before consent.
await page.setInputFiles('#import-pane-file', htmlPath);
await check('mixed-case HTML extension opens the shared replacement confirmation', async () =>
  await page.locator('#pane-replace-dialog').evaluate((dialog) => dialog.open)
  && (await page.locator('#pane-replace-title').textContent()) === 'Replace HTML code?'
  && (await page.locator('#pane-replace-context').textContent()).includes('HTML editor'));
await page.click('#pane-replace-cancel');
await page.click('#editor-tabs .tab[data-editor="html"]');
await check('cancel leaves the HTML editor unchanged', async () =>
  !(await editorText()).includes('local-import-marker'));

// The input reset lets the exact same file fire a second change event.
await page.setInputFiles('#import-pane-file', htmlPath);
await page.click('#pane-replace-confirm');
await check('confirm replaces only the mapped editor', async () =>
  (await editorText()).includes('local-import-marker'));
await check('local replacement marks the HTML preview stale', () =>
  page.locator('#unsaved-html').isVisible());

await page.setInputFiles('#import-pane-file', {
  name: 'renamed.txt',
  mimeType: 'text/plain',
  buffer: Buffer.from('not code'),
});
await check('unsupported local file is rejected in-app', async () =>
  await page.locator('#app-toast').isVisible()
  && (await page.locator('#app-toast').textContent()).includes('not an HTML, CSS, or JavaScript'));

await page.setInputFiles('#import-pane-file', {
  name: 'too-large.js',
  mimeType: 'text/javascript',
  buffer: Buffer.alloc((5 * 1024 * 1024) + 1, 32),
});
await check('oversized local file is rejected before confirmation', async () =>
  await page.locator('#app-toast').isVisible()
  && (await page.locator('#app-toast').textContent()).includes('limited to 5 MB'));

// Install a same-origin explicit host adapter and stub SharePoint REST.
const origin = new URL(APP_URL).origin;
const apiRequests = [];
let upload = null;
await page.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  apiRequests.push(url);
  if (url.includes('/_api/contextinfo')) {
    const otherSite = url.includes('/sites/other/_api/contextinfo');
    const browserSite = url.includes('/sites/browser/_api/contextinfo');
    await route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify({
        FormDigestValue: otherSite
          ? 'OTHER-DIGEST'
          : browserSite ? 'BROWSER-DIGEST' : 'TEST-DIGEST',
        FormDigestTimeoutSeconds: 1800,
        WebFullUrl: otherSite
          ? `${origin}/sites/other`
          : browserSite ? `${origin}/sites/browser` : origin,
        SiteFullUrl: origin,
      }),
    });
    return;
  }
  if (url.includes('/Files/AddUsingPath(')) {
    upload = {
      url,
      body: request.postData() || '',
      digest: request.headers()['x-requestdigest'],
    };
    await route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify({ ServerRelativeUrl: '/existing.css' }),
    });
    return;
  }
  if (url.includes('/GetFileByServerRelativePath(')) {
    const otherSite = url.includes('/sites/other/_api/');
    await route.fulfill({
      status: 200,
      contentType: 'text/javascript',
      body: otherSite
        ? 'console.log("other-site-import-marker");\n'
        : 'console.log("sharepoint-import-marker");\n',
    });
    return;
  }
  if (url.includes('/GetFolderByServerRelativePath(')) {
    const otherSite = url.includes('/sites/other/_api/');
    const browserSite = url.includes('/sites/browser/_api/');
    const browserDocsFolder = /sites%2Fbrowser%2FDocs/i.test(url);
    const hashFolder = /Code%23One/i.test(url);
    await route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify(browserDocsFolder ? {
        ServerRelativeUrl: '/sites/browser/Docs',
        Folders: [],
        Files: [
          {
            Name: 'guide.md',
            ServerRelativeUrl: '/sites/browser/Docs/guide.md',
            Length: 72,
          },
          {
            Name: 'notes.txt',
            ServerRelativeUrl: '/sites/browser/Docs/notes.txt',
            Length: 28,
          },
          {
            Name: 'styles.css',
            ServerRelativeUrl: '/sites/browser/Docs/styles.css',
            Length: 34,
          },
        ],
      } : browserSite ? {
        ServerRelativeUrl: '/sites/browser',
        Folders: [
          {
            Name: 'Docs',
            ServerRelativeUrl: '/sites/browser/Docs',
          },
        ],
        Files: [
          {
            Name: 'overview.html',
            ServerRelativeUrl: '/sites/browser/overview.html',
            Length: 112,
          },
          {
            Name: 'app.js',
            ServerRelativeUrl: '/sites/browser/app.js',
            Length: 44,
          },
        ],
      } : otherSite ? {
        ServerRelativeUrl: '/sites/other',
        Folders: [],
        Files: [
          {
            Name: 'existing.css',
            ServerRelativeUrl: '/sites/other/existing.css',
            Length: 34,
          },
          {
            Name: 'other-site.js',
            ServerRelativeUrl: '/sites/other/other-site.js',
            Length: 52,
          },
        ],
      } : hashFolder ? {
        ServerRelativeUrl: '/Code#One',
        Folders: [],
        Files: [],
      } : {
        ServerRelativeUrl: '/',
        Folders: [{ Name: 'Code#One', ServerRelativeUrl: '/Code#One' }],
        Files: [
          { Name: 'existing.css', ServerRelativeUrl: '/existing.css', Length: 34 },
          { Name: 'sample.js', ServerRelativeUrl: '/sample.js', Length: 42 },
          { Name: 'ignore.txt', ServerRelativeUrl: '/ignore.txt', Length: 10 },
        ],
      }),
    });
    return;
  }
  await route.fulfill({ status: 404, body: 'not stubbed' });
});
await page.route('**/sites/browser/Docs/guide.md', (route) => route.fulfill({
  status: 200,
  contentType: 'text/plain',
  body: '# Browser-picked guide\n\nOpened from a different SharePoint subsite.',
}));
await page.addInitScript(() => {
  if (sessionStorage.getItem('dcspad-context-test') === 'modern') {
    window.spModuleLoader = {
      _bundledComponents: {
        'b6917cb1-93a0-4b97-a84d-7cf49975d4ec': {
          PageManager: {
            _instance: {
              pageContext: {
                legacyPageContext: {
                  webAbsoluteUrl: location.origin,
                  userDisplayName: 'Modern Context User',
                },
              },
            },
          },
        },
      },
    };
  } else {
    window.__DCSPAD_SP_CONTEXT__ = {
      webAbsoluteUrl: location.origin,
      userDisplayName: 'Context Test User',
    };
  }
});
await page.reload();
await page.waitForSelector('.monaco-editor');
await check('explicit host adapter enables SharePoint file actions', async () =>
  (await page.locator('#sp-chip-text').textContent()) === 'SP: Live'
  && !(await page.locator('#mi-sp-import').isDisabled())
  && (await page.locator('#sp-chip').getAttribute('data-context')).includes('context: host'));

await page.click('#btn-file');
await page.click('#mi-sp-import');
await page.waitForSelector('#sp-files-dialog[open]');
await check('SharePoint picker filters out unsupported files', async () => {
  const text = await page.locator('#sp-files-list').textContent();
  return text.includes('sample.js') && text.includes('existing.css') && !text.includes('ignore.txt');
});
await check('SharePoint folder query uses supported expanded-property syntax', () => {
  const requestUrl = apiRequests.find((url) =>
    url.includes('/GetFolderByServerRelativePath('));
  if (!requestUrl) return false;
  const query = new URL(requestUrl).searchParams;
  const select = query.get('$select') || '';
  return query.get('$expand') === 'Folders,Files'
    && select.includes('Folders/ServerRelativeUrl')
    && select.includes('Files/TimeLastModified');
});

// ResourcePath folder URLs encode # rather than turning it into a fragment.
await page.locator('.sp-file-row', { hasText: 'Code#One' }).click();
await page.waitForFunction(() =>
  document.querySelector('#sp-folder-path')?.textContent === '/Code#One');
await check('ResourcePath folder request safely encodes #', () =>
  apiRequests.some((url) => /Code%23One/i.test(url) && !new URL(url).hash));
await page.click('#sp-folder-up');
await page.waitForFunction(() =>
  document.querySelector('#sp-folder-path')?.textContent === '/');

await page.fill('#sp-site-url', 'https://different.example.com/sites/nope');
await page.click('#sp-site-open');
await check('cross-origin SharePoint site is rejected before REST browsing', async () =>
  (await page.locator('#sp-files-error').textContent()).includes('on this tenant'));

await page.fill('#sp-site-url', `${origin}/sites/other`);
await page.click('#sp-site-open');
await page.waitForFunction(() =>
  document.querySelector('#sp-folder-path')?.textContent === '/sites/other');
await check('same-tenant site URL switches the picker boundary', async () =>
  (await page.locator('#sp-site-url').inputValue()) === `${origin}/sites/other`
  && (await page.locator('#sp-files-list').textContent()).includes('other-site.js')
  && !(await page.locator('#sp-files-list').textContent()).includes('sample.js'));
await check('selected SharePoint site persists with its last folder', () =>
  page.waitForFunction(() => {
    const settings =
      JSON.parse(localStorage.getItem('dcspad.v2.workspace')).settings;
    return settings.spFilesWebUrl.endsWith('/sites/other')
      && settings.spFilesFolder === '/sites/other';
  }).then(() => true, () => false));

await page.locator('.sp-file-row', { hasText: 'other-site.js' }).click();
await page.click('#sp-files-primary');
await page.waitForSelector('#pane-replace-dialog[open]');
await check('SharePoint import uses the same explicit replacement confirmation', async () =>
  (await page.locator('#pane-replace-title').textContent()) === 'Replace JS code?');
await page.click('#pane-replace-confirm');
await check('confirmed SharePoint import replaces the JS editor', async () =>
  (await editorText()).includes('other-site-import-marker'));
await check('successful SharePoint import closes both dialogs', async () =>
  !(await page.locator('#sp-files-dialog').evaluate((dialog) => dialog.open))
  && !(await page.locator('#pane-replace-dialog').evaluate((dialog) => dialog.open)));

// Seed CSS through the same local import path, then overwrite the matching
// SharePoint file. The first click only arms overwrite; the second writes.
await page.setInputFiles('#import-pane-file', cssPath);
await page.click('#pane-replace-confirm');
await page.click('#btn-file');
await page.click('#mi-sp-export');
await page.waitForSelector('#sp-files-dialog[open]');
await page.selectOption('#sp-export-pane', 'css');
await page.fill('#sp-export-name', 'existing.css');
await page.click('#sp-files-primary');
await check('existing SharePoint file requires a separate overwrite confirmation', async () =>
  upload === null
  && (await page.locator('#sp-files-notice').textContent()).includes('already exists')
  && (await page.locator('#sp-files-primary').textContent()) === 'Overwrite');
await page.click('#sp-files-primary');
await page.waitForFunction(() => !document.getElementById('sp-files-dialog').open);
await check('confirmed SharePoint upload sends pane text and a digest', () =>
  upload?.body.includes('share-export-marker')
  && upload?.digest === 'OTHER-DIGEST'
  && upload?.url.includes('/sites/other/_api/')
  && /overwrite=true/i.test(upload.url));

await page.click('#extras-tabs [data-extra="docs"]');
await page.click('#browser-browse');
await page.waitForSelector('#sp-files-dialog[open]');
await check('Browser uses the shared SharePoint picker in resource mode', async () =>
  (await page.locator('#sp-files-title').textContent()) === 'Browse SharePoint'
  && (await page.locator('#sp-files-primary').textContent()) === 'Open file'
  && (await page.locator('#sp-files-empty').textContent())
    === 'No HTML, Markdown, or text files in this folder.');

await page.fill('#sp-site-url', `${origin}/sites/browser`);
await page.click('#sp-site-open');
await page.waitForFunction(() =>
  document.querySelector('#sp-folder-path')?.textContent === '/sites/browser');
await check('Browser picker accepts a different same-tenant subsite and filters its files', async () => {
  const text = await page.locator('#sp-files-list').textContent();
  return (await page.locator('#sp-site-url').inputValue()) === `${origin}/sites/browser`
    && text.includes('Docs')
    && text.includes('overview.html')
    && !text.includes('app.js');
});

await page.locator('.sp-file-row', { hasText: 'Docs' }).click();
await page.waitForFunction(() =>
  document.querySelector('#sp-folder-path')?.textContent === '/sites/browser/Docs');
await check('Browser picker traverses folders and shows only HTML, Markdown, and text', async () => {
  const text = await page.locator('#sp-files-list').textContent();
  return text.includes('guide.md')
    && text.includes('notes.txt')
    && !text.includes('styles.css');
});

await page.locator('.sp-file-row', { hasText: 'guide.md' }).click();
await page.click('#sp-files-primary');
await page.frameLocator('#docs-frame')
  .locator('h1', { hasText: 'Browser-picked guide' }).waitFor();
await check('picked SharePoint resource opens in Browser and joins URL history', () =>
  page.evaluate(async (expected) => {
    const { getState } = await import('/src/state.js');
    return !document.getElementById('sp-files-dialog').open
      && document.getElementById('browser-address-input').value === expected
      && getState().settings.browserHistory[0] === expected;
  }, `${origin}/sites/browser/Docs/guide.md`));

await page.evaluate(() => sessionStorage.setItem('dcspad-context-test', 'modern'));
await page.reload();
await page.waitForSelector('.monaco-editor');
await check('guarded Modern legacyPageContext works when globals are absent', async () =>
  (await page.locator('#sp-chip-text').textContent()) === 'SP: Live'
  && (await page.locator('#sp-chip').getAttribute('data-context')).includes('context: modern-legacy'));

await browser.close();
exitWithResult();
