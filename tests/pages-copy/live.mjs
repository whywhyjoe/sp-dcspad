// SP Workbench page-copy — stubbed-live connection behaviour
// (design/PAGE-COPY.md §2.3, §8 stubbed-live bullet).
//
// Exercises sp-rest.js / sp-write.js / sp-pages.js's per-web digest handling
// and page-copy-dialog.js's connect()/connectSeq race against a Playwright
// page.route model of SharePoint webs living on the app's own origin:
// /sites/lsrc (source), /sites/ldst (target), and /sites/slow (a target
// whose /_api/web answer is deliberately delayed, for the connect race).
//
// Testing caveat (design/PAGE-COPY.md §2.3): fetchImpl injected into
// sp-write.js does not reach the singleton digest client in sp-files.js —
// only page.route intercepts the real fetch() calls every write here makes.
// House style: tests/workbench-schema.mjs lines 1070+ (addInitScript +
// page.route stubbed-live setup), tests/pages-copy/dialog.mjs
// (waitCheckEnabled/waitReportVisible), tests/pages-copy/runner.mjs (module
// shape, page.evaluate(import(...))).
//
// The digest cache in sp-files.js is a module-level singleton, shared by
// every createSpWriteClient on one page load — so any check whose assertion
// depends on a *fresh* digest fetch (checks 1-3) gets its own page via
// newLivePage(); a page reload is a new module instantiation, i.e. a new
// cache.

const MODERN_SITE_PAGES_FEATURE_ID = 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec';

// ---- route model -----------------------------------------------------------
//
// One handler, shared by every fresh page, closing over a per-page `state`
// object ({ writes, reads, digestCounts, flags }) so assertions never leak
// across pages. `site` is whichever of lsrc/ldst/slow the request's path
// names; anything else falls through to the network untouched.

function siteOf(pathname) {
  if (pathname.includes('/sites/ldst/')) return 'ldst';
  if (pathname.includes('/sites/lsrc/')) return 'lsrc';
  if (pathname.includes('/sites/slow/')) return 'slow';
  return null;
}

async function handleApi(route, state) {
  const request = route.request();
  const url = request.url();
  const method = request.method();
  const u = new URL(url);
  const site = siteOf(u.pathname);
  if (!site) return route.continue();

  const apiIdx = u.pathname.toLowerCase().indexOf('/_api/');
  const relLower = (apiIdx >= 0 ? u.pathname.slice(apiIdx + 6) : u.pathname).toLowerCase();

  // ---- reads -----------------------------------------------------------
  if (method === 'GET') {
    state.reads.push({ url, path: relLower, site });

    if (relLower === 'web') {
      const respond = () => route.fulfill({ json: {
        Id: `${site}-web-id`, Title: `${site.toUpperCase()} Web`,
        Url: `${u.origin}/sites/${site}`, ServerRelativeUrl: `/sites/${site}`,
      } });
      if (site === 'slow') {
        await new Promise((r) => setTimeout(r, 1500));
      }
      return respond();
    }
    if (relLower === 'site') {
      return route.fulfill({ json: {
        Id: `${site}-site-id`, Url: `${u.origin}/sites/${site}`, ServerRelativeUrl: `/sites/${site}`,
      } });
    }
    if (relLower.startsWith('web/features/getbyid(')) {
      return route.fulfill({ json: { DefinitionId: MODERN_SITE_PAGES_FEATURE_ID } });
    }
    if (relLower === 'web/lists') {
      return route.fulfill({ json: { value: [
        { Id: `${site}-pages-list`, Title: 'Site Pages', BaseTemplate: 119, Hidden: false,
          RootFolder: { ServerRelativeUrl: `/sites/${site}/SitePages` } },
      ] } });
    }
    if (relLower.includes("lists(guid'") && relLower.endsWith('/fields')) {
      return route.fulfill({ json: { value: [] } });
    }
    if (relLower === 'web/getclientsidewebparts') {
      return route.fulfill({ json: { value: [] } });
    }
    if (relLower.startsWith('web/getfilebyserverrelativepath(')) {
      return route.fulfill({
        status: 404, json: { 'odata.error': { message: { value: 'File Not Found.' } } },
      });
    }
    if (relLower.startsWith('web/getfolderbyserverrelativepath(')) {
      return route.fulfill({ json: { Exists: false } });
    }
    return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'no fixture' } } } });
  }

  // ---- writes ------------------------------------------------------------
  if (method === 'POST') {
    if (relLower === 'contextinfo') {
      state.digestCounts[site] = (state.digestCounts[site] || 0) + 1;
      return route.fulfill({ json: {
        FormDigestValue: `DIGEST-${site.toUpperCase()}`,
        FormDigestTimeoutSeconds: 1800,
        WebFullUrl: `${u.origin}/sites/${site}`,
      } });
    }

    const record = {
      url, path: relLower, site,
      digest: request.headers()['x-requestdigest'] || '',
      contentType: request.headers()['content-type'] || '',
      ifMatch: request.headers()['if-match'] || '',
      body: request.postData() || '',
    };
    state.writes.push(record);

    if (state.flags[`failNext_${site}`]) {
      state.flags[`failNext_${site}`] = false;
      return route.fulfill({ status: 403, json: { 'odata.error': { message: { value: 'Denied' } } } });
    }
    if (state.flags[`abortNext_${site}`]) {
      state.flags[`abortNext_${site}`] = false;
      return route.abort('failed');
    }

    if (relLower === 'sitepages/pages') {
      const id = (state.nextPageId = (state.nextPageId || 40) + 1);
      return route.fulfill({ json: { Id: id } });
    }
    // savepage / MoveFileByPath / everything else this suite doesn't inspect
    // the body of: a bare success is all sp-write.js's post() needs to
    // resolve.
    return route.fulfill({ json: {} });
  }

  return route.continue();
}

async function newLivePage(browser) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.addInitScript(() => {
    window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
  });
  const state = { writes: [], reads: [], digestCounts: Object.create(null), flags: Object.create(null) };
  await page.route('**/_api/**', (route) => handleApi(route, state));
  return { page, state };
}

export async function run({ browser, check, WB_URL }) {
  const origin = new URL(WB_URL).origin;

  // ---- (1) a write to /sites/ldst carries that web's own digest, and
  // nothing is posted under /sites/lsrc -------------------------------------
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });

    await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const target = createSpRestClient();
      await target.connectWeb(`${origin}/sites/ldst`);
      const write = createSpWriteClient({ client: target });
      await write.postJson('_test/write-a', { x: 1 });
    }, origin);

    await check('live: a write through a client connected to /sites/ldst fetches that web’s own contextinfo, carries X-RequestDigest: DIGEST-LDST, and nothing is POSTed under /sites/lsrc', () => {
      const ldstWrites = state.writes.filter((w) => w.site === 'ldst');
      const lsrcWrites = state.writes.filter((w) => w.site === 'lsrc');
      return ldstWrites.length === 1
        && ldstWrites[0].digest === 'DIGEST-LDST'
        && lsrcWrites.length === 0
        && (state.digestCounts.ldst || 0) === 1
        && !(state.digestCounts.lsrc);
    });

    await page.close();
  }

  // ---- (2) a source write and a target write in one page use their own
  // digests ------------------------------------------------------------------
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });

    await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const source = createSpRestClient();
      await source.connectWeb(`${origin}/sites/lsrc`);
      const sourceWrite = createSpWriteClient({ client: source });
      await sourceWrite.postJson('_test/write-src', { x: 1 });

      const target = createSpRestClient();
      await target.connectWeb(`${origin}/sites/ldst`);
      const targetWrite = createSpWriteClient({ client: target });
      await targetWrite.postJson('_test/write-tgt', { x: 1 });
    }, origin);

    await check('live: a source write and a target write in one page fetch both contextinfo endpoints and each POST carries its own web’s digest', () => {
      const lsrcWrites = state.writes.filter((w) => w.site === 'lsrc');
      const ldstWrites = state.writes.filter((w) => w.site === 'ldst');
      return lsrcWrites.length === 1 && lsrcWrites[0].digest === 'DIGEST-LSRC'
        && ldstWrites.length === 1 && ldstWrites[0].digest === 'DIGEST-LDST'
        && (state.digestCounts.lsrc || 0) === 1
        && (state.digestCounts.ldst || 0) === 1;
    });

    await page.close();
  }

  // ---- (3) a 403 on the target's first write forces exactly one extra
  // contextinfo (forced refresh), the retry succeeds, and the source's
  // already-cached digest is not re-fetched ---------------------------------
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });

    // Establish the source's digest first, on this same fresh page/cache.
    await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const source = createSpRestClient();
      await source.connectWeb(`${origin}/sites/lsrc`);
      const sourceWrite = createSpWriteClient({ client: source });
      await sourceWrite.postJson('_test/write-src', { x: 1 });
    }, origin);

    state.flags.failNext_ldst = true;
    const outcome = await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const target = createSpRestClient();
      await target.connectWeb(`${origin}/sites/ldst`);
      const targetWrite = createSpWriteClient({ client: target });
      await targetWrite.postJson('_test/write-tgt', { x: 1 });
      return 'resolved';
    }, origin);

    await check('live: the target’s first write gets a 403, sp-write.js forces exactly one extra /sites/ldst contextinfo, the retry succeeds, and the source’s digest is not re-fetched', () => {
      const ldstAttempts = state.writes.filter((w) => w.site === 'ldst' && w.path === '_test/write-tgt');
      return outcome === 'resolved'
        && ldstAttempts.length === 2   // the 403'd attempt + the retry, both recorded
        && ldstAttempts[0].digest === 'DIGEST-LDST' && ldstAttempts[1].digest === 'DIGEST-LDST'
        && (state.digestCounts.ldst || 0) === 2
        && (state.digestCounts.lsrc || 0) === 1;
    });

    await page.close();
  }

  // ---- (4) request-shape checks: createPage, savePage, moveFileByPath -----
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });

    await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const { createSpPages } = await import('/src/workbench/sp-pages.js');
      const client = createSpRestClient();
      await client.connectWeb(`${origin}/sites/ldst`);
      const write = createSpWriteClient({ client });
      const pages = createSpPages({ client, write });
      await pages.createPage({ pageLayoutType: 'Article', promotedState: 0 });
      await pages.savePage(42, { Title: 'X' });
      await pages.moveFileByPath(
        `${origin}/sites/ldst/SitePages/A.aspx`,
        `${origin}/sites/ldst/SitePages/B.aspx`,
        { overwrite: false },
      );
    }, origin);

    await check('live: createPage sends a Content-Type carrying odata=verbose and a body whose __metadata.type is SP.Publishing.SitePage', () => {
      const w = state.writes.find((x) => x.path === 'sitepages/pages');
      if (!w) return false;
      let body;
      try { body = JSON.parse(w.body); } catch { return false; }
      return w.contentType.includes('odata=verbose') && body.__metadata?.type === 'SP.Publishing.SitePage';
    });

    await check('live: savePage sends If-Match: *', () => {
      const w = state.writes.find((x) => x.path.endsWith('/savepage'));
      return Boolean(w) && w.ifMatch === '*';
    });

    await check('live: moveFileByPath posts SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=false with absolute DecodedUrl values', () => {
      const w = state.writes.find((x) => x.path.startsWith('sp.movecopyutil.movefilebypath'));
      if (!w) return false;
      if (!w.url.endsWith('MoveFileByPath(overwrite=@a1)?@a1=false')) return false;
      let body;
      try { body = JSON.parse(w.body); } catch { return false; }
      return body.srcPath?.DecodedUrl === `${origin}/sites/ldst/SitePages/A.aspx`
        && body.destPath?.DecodedUrl === `${origin}/sites/ldst/SitePages/B.aspx`;
    });

    await page.close();
  }

  // ---- (5) a lost response (aborted POST) rejects with code 'network' and
  // is never retried ----------------------------------------------------------
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });

    state.flags.abortNext_ldst = true;
    const result = await page.evaluate(async (origin) => {
      const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
      const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
      const client = createSpRestClient();
      await client.connectWeb(`${origin}/sites/ldst`);
      const write = createSpWriteClient({ client });
      try {
        await write.postJson('_test/write-lost', { x: 1 });
        return { rejected: false };
      } catch (err) {
        return { rejected: true, code: err?.code };
      }
    }, origin);

    await check('live: a lost response (aborted POST) makes sp-write.js reject with code \'network\', and the route saw exactly one attempt (no automatic retry)', () => {
      const attempts = state.writes.filter((w) => w.path === '_test/write-lost');
      return result.rejected === true && result.code === 'network' && attempts.length === 1;
    });

    await page.close();
  }

  // ---- (6)/(7): drive the real dialog against a stub source client + the
  // real live createSpRestClient for the target. The stub answers exactly
  // the reads page-copy-dialog.js's phase-0 snapshot makes (design/
  // PAGE-COPY.md §2.2) for one JSON-canvas Article page, and is strict about
  // the projection: an item read without $expand=FieldValuesAsText, or a
  // fields read without EntityPropertyName in $select, 400s — the same
  // shape as the Workbench Items tab's own pinned $select=* stub
  // (tests/workbench.mjs), applied here to page-copy-dialog.js's own reads.
  const STUB_HELPER = `
    import { openPageCopyDialog } from '/src/workbench/page-copy-dialog.js';
    import { createSpRestClient } from '/src/workbench/sp-rest.js';

    const origin = location.origin;

    const PAGE_DTO = {
      Id: 7, Title: 'Stub Page', Description: '', TopicHeader: '', AuthorByline: [],
      PageLayoutType: 'Article', PromotedState: 0,
      CanvasContent1: '[]', LayoutWebpartsContent: '[]',
      BannerImageUrl: '', BannerThumbnailUrl: '', FileName: 'Stub-Page.aspx',
    };
    const ITEM_DTO = {
      Id: 7, FileRef: origin + '/sites/lsrc/SitePages/Stub-Page.aspx',
      FileDirRef: origin + '/sites/lsrc/SitePages', FileLeafRef: 'Stub-Page.aspx',
      CommentsDisabled: false, OData__UIVersionString: '1.0', Modified: '2026-09-23T00:00:00Z',
      FieldValuesAsText: { Title: 'Stub Page' },
    };
    const FIELDS = [
      { Id: 'f1', Title: 'Title', InternalName: 'Title', EntityPropertyName: 'Title',
        TypeAsString: 'Text', FieldTypeKind: 2, Required: false, Hidden: false,
        ReadOnlyField: false, Choices: null, FillInChoice: false },
    ];
    const WEB_IDENTITY = { Id: 'lsrc-web-id', Title: 'LSRC Web', Url: origin + '/sites/lsrc', ServerRelativeUrl: '/sites/lsrc' };
    const SITE_IDENTITY = { Id: 'lsrc-site-id', Url: origin + '/sites/lsrc', ServerRelativeUrl: '/sites/lsrc' };

    function makeStubClient() {
      return {
        webUrl: () => origin + '/sites/lsrc',
        context: () => ({ live: true }),
        async get(path, opts) {
          if (path === 'sitepages/pages(7)') return structuredClone(PAGE_DTO);
          if (path === 'web') return structuredClone(WEB_IDENTITY);
          if (path === 'site') return structuredClone(SITE_IDENTITY);
          if (path === "web/lists(guid'stub-list')/items(7)") {
            if (opts?.expand !== 'FieldValuesAsText') {
              const err = new Error("The field or property 'FieldValuesAsText' does not exist.");
              err.status = 400;
              throw err;
            }
            return structuredClone(ITEM_DTO);
          }
          throw new Error('stub: unhandled get ' + path);
        },
        async getAll(path, opts) {
          if (path === "web/lists(guid'stub-list')/fields") {
            if (!(Array.isArray(opts?.select) && opts.select.includes('EntityPropertyName'))) {
              const err = new Error("The field or property 'EntityPropertyName' does not exist.");
              err.status = 400;
              throw err;
            }
            return { items: structuredClone(FIELDS), partial: false };
          }
          throw new Error('stub: unhandled getAll ' + path);
        },
      };
    }

    window.__pc = {
      open() {
        const client = makeStubClient();
        const library = {
          listId: 'stub-list', title: 'Site Pages', baseTemplate: 119,
          hidden: false, rootPath: '/sites/lsrc/SitePages',
        };
        const p = openPageCopyDialog({
          client, createClient: () => createSpRestClient(), mockWriter: undefined,
          library, pageId: 7, pageName: 'Stub Page', statusOf: undefined, analyzers: null,
          runCopy: async () => ({
            outcome: 'done',
            journal: { pageId: null, currentPath: '', createdBy: 'unknown', assets: [], steps: [], retries: [], drift: [] },
            pageUrl: '',
          }),
          discardCopy: async () => ({ recycled: [], leftovers: [] }),
        });
        p.then((r) => { window.__pcOutcome = r; });
        return true;
      },
      // Fires both connects back-to-back via dispatchEvent rather than
      // Playwright's click(): connectBtn.disabled is set synchronously the
      // instant connect() starts (page-copy-dialog.js:434), so a real
      // actionability-gated click on the still-disabled button would just
      // wait out the slow connect instead of racing it.
      raceConnect(slowUrl, ldstUrl) {
        const other = document.querySelector('.wb-pc-dest-other');
        other.checked = true;
        other.dispatchEvent(new Event('change', { bubbles: true }));
        const urlInput = document.querySelector('.wb-pc-url');
        const connectBtn = document.querySelector('.wb-pc-connect');
        urlInput.value = slowUrl;
        urlInput.dispatchEvent(new Event('input', { bubbles: true }));
        connectBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        urlInput.value = ldstUrl;
        urlInput.dispatchEvent(new Event('input', { bubbles: true }));
        connectBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      },
      // Starts a slow connect, then edits the URL WITHOUT connecting again:
      // the slow answer arriving afterwards must not install /sites/slow as
      // the destination while the field names another site.
      slowConnectThenEdit(slowUrl, editedUrl) {
        const other = document.querySelector('.wb-pc-dest-other');
        other.checked = true;
        other.dispatchEvent(new Event('change', { bubbles: true }));
        const urlInput = document.querySelector('.wb-pc-url');
        const connectBtn = document.querySelector('.wb-pc-connect');
        urlInput.value = slowUrl;
        urlInput.dispatchEvent(new Event('input', { bubbles: true }));
        connectBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        urlInput.value = editedUrl;
        urlInput.dispatchEvent(new Event('input', { bubbles: true }));
      },
    };
  `;

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

  // ---- (6) the connect race: a slow /sites/slow connect started first must
  // not win over the immediately-following /sites/ldst connect -------------
  {
    const { page, state } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
    await page.addScriptTag({ type: 'module', content: STUB_HELPER });
    await page.waitForFunction(() => typeof window.__pc === 'object');

    await page.evaluate(() => window.__pc.open());
    await page.waitForSelector('.wb-pc-version');   // phase-0 read finished, form built

    await page.evaluate(({ slowUrl, ldstUrl }) => window.__pc.raceConnect(slowUrl, ldstUrl),
      { slowUrl: `${origin}/sites/slow`, ldstUrl: `${origin}/sites/ldst` });

    // The slow web's /_api/web answer is delayed 1500ms; give it (and the
    // fast ldst connect that should overtake it) time to fully settle.
    await page.waitForTimeout(2500);

    const settled = await page.evaluate(() => ({
      statusHidden: document.querySelector('.wb-pc-target-status')?.hidden !== false,
      checkDisabled: document.querySelector('.wb-pc-check')?.disabled,
    }));

    await check('live: the connect race settles on /sites/ldst (no error status, Check enabled) rather than being clobbered by the slower /sites/slow answer arriving later', () =>
      settled.statusHidden === true && settled.checkDisabled === false);

    const readsBeforeCheck = state.reads.length;
    await page.evaluate(() => document.querySelector('.wb-pc-check').click());
    await waitReportVisible(page);
    const probes = state.reads.slice(readsBeforeCheck);

    await check('live: the subsequent Check sends its probes only under /sites/ldst/_api/ (never re-touching /sites/slow or /sites/lsrc)', () =>
      probes.length > 0 && probes.every((r) => r.site === 'ldst'));

    await page.close();
  }

  // ---- (6b) editing the destination abandons an in-flight connect (xo
  // review): the late answer must not arm Check for a web the form no longer
  // names ----------------------------------------------------------------------
  {
    const { page } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
    await page.addScriptTag({ type: 'module', content: STUB_HELPER });
    await page.waitForFunction(() => typeof window.__pc === 'object');

    await page.evaluate(() => window.__pc.open());
    await page.waitForSelector('.wb-pc-version');

    await page.evaluate(({ slowUrl, ldstUrl }) => window.__pc.slowConnectThenEdit(slowUrl, ldstUrl),
      { slowUrl: `${origin}/sites/slow`, ldstUrl: `${origin}/sites/ldst` });
    // Outlast the slow connect entirely: it reads the slow /_api/web twice
    // (connectWeb, then webIdentity), 1500ms each, before it would land.
    await page.waitForTimeout(5000);

    const after = await page.evaluate(() => ({
      checkDisabled: document.querySelector('.wb-pc-check')?.disabled,
      connectDisabled: document.querySelector('.wb-pc-connect')?.disabled,
      status: document.querySelector('.wb-pc-target-status')?.textContent || '',
    }));

    await check('live: editing the URL while a slow connect is in flight abandons it — its late answer leaves Check disabled, Connect enabled, and the status asking to connect', () =>
      after.checkDisabled === true && after.connectDisabled === false
        && /Connect to check this site/.test(after.status));

    await page.close();
  }

  // ---- (7) projection strictness: the dialog's own snapshot reads must
  // carry the exact select/expand the stub enforces, or it 400s -------------
  {
    const { page } = await newLivePage(browser);
    await page.goto(WB_URL);
    await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
    await page.addScriptTag({ type: 'module', content: STUB_HELPER });
    await page.waitForFunction(() => typeof window.__pc === 'object');

    await page.evaluate(() => window.__pc.open());
    // Either the form loads (phase-0 reads passed the stub's projection
    // check) or a refusal banner appears (a wrong projection 400'd) — wait
    // for whichever settles first, then assert it was the form.
    await page.waitForFunction(
      () => document.querySelector('.wb-pc-version') || document.querySelector('.sp-metadata-dialog .sp-files-error'),
      null, { timeout: 10000 },
    );

    await check('live: the snapshot’s item read carries $expand=FieldValuesAsText and the fields read carries EntityPropertyName in $select — the dialog’s snapshot loads without an error', async () => {
      const versionPresent = await page.$('.wb-pc-version');
      // buildForm's OWN (Check-phase) error banner shares the sp-files-error
      // class and is always present-but-hidden once the form has built
      // successfully — presence alone can't distinguish it from
      // renderRefusal's always-visible phase-0 refusal banner, so this
      // checks visibility, not just presence.
      const errorVisible = await page.evaluate(() =>
        Array.from(document.querySelectorAll('.sp-metadata-dialog .sp-files-error'))
          .some((e) => e.hidden === false));
      return Boolean(versionPresent) && !errorVisible;
    });

    await page.close();
  }
}
