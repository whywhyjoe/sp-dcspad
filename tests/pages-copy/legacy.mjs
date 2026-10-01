// SP Workbench — page copy: the legacy-HTML same-web path (design/PAGE-COPY.md
// §1 table, decision 2, §4.2, §7, §11 Q5).
//
// A page whose CanvasContent1 is not JSON (the pre-modern HTML canvas format)
// can only be duplicated on its own web, and only through
// SP.MoveCopyUtil.CopyFileByPath — never the sitepages/pages API. Cross-site
// is refused outright (page-copy.js sourceEligibility). page-copy-run.js's
// 'reset' step replaces the API-path's name/assets/content/move sequence: the
// file copy already carries everything, so reset only neutralises a promoted
// source (PromotedState cleared before anyone can checkout/save) and resets
// the Title.
//
// Fixture: mock-pagecopy.js page 8 on /sites/pagesrc, 'Legacy-News.aspx',
// CanvasContent1 = the HTML-format string `<div><div data-sp-canvascontrol>
// </div></div>`. House style: tests/pages-copy/runner.mjs (injected
// window.__pc helper, module shape); this file's helper is the same recipe,
// scoped to page 8's copyFile path instead of page 7's sitepages-API path.

const HELPER_LINES = [
  "import { createSpRestClient } from '/src/workbench/sp-rest.js';",
  "import { createSpWriteClient } from '/src/workbench/sp-write.js';",
  "import { mockResolver, mockWriter } from '/src/workbench/mock-data.js';",
  "import { createSpPages } from '/src/workbench/sp-pages.js';",
  "import * as pageCopy from '/src/workbench/page-copy.js';",
  "import { runCopy, discardCopy } from '/src/workbench/page-copy-run.js';",
  "import { resetPageCopyMock } from '/src/workbench/mock-pagecopy.js';",
  "",
  "const FIELD_SELECT = ['Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'Required', 'Hidden', 'ReadOnlyField', 'Group', 'DefaultValue', 'Choices', 'Description', 'FillInChoice'];",
  "const LEGACY_PAGE_ID = 8;",
  "",
  "async function connectPair(webPath) {",
  "  const client = createSpRestClient({ mockResolver });",
  "  await client.connectWeb(webPath);",
  "  const write = createSpWriteClient({ client, mockWriter });",
  "  const pages = createSpPages({ client, write });",
  "  return { client, write, pages };",
  "}",
  "",
  "async function readSnapshot(pair, pageId) {",
  "  const page = await pair.pages.getPage(pageId);",
  "  const library = await pair.pages.sitePagesLibrary();",
  "  const item = await pair.client.get(`web/lists(guid'${library.id}')/items(${pageId})`, { expand: 'FieldValuesAsText' });",
  "  const fieldsRes = await pair.client.getAll(`web/lists(guid'${library.id}')/fields`, { select: FIELD_SELECT });",
  "  const web = await pair.pages.webIdentity();",
  "  return pageCopy.snapshotFromReads({",
  "    page, item, fields: fieldsRes.items,",
  "    library: { id: library.id, baseTemplate: 119, hidden: false, rootPath: library.rootPath },",
  "    web,",
  "  });",
  "}",
  "",
  "async function buildTarget(pair, snapshot, options) {",
  "  const web = await pair.pages.webIdentity();",
  "  const library = await pair.pages.sitePagesLibrary();",
  "  const fieldsRes = await pair.client.getAll(`web/lists(guid'${library.id}')/fields`, { select: FIELD_SELECT });",
  "  const folder = String(options.folder || '').replace(/^\\/+|\\/+$/g, '');",
  "  const libraryRoot = library.rootPath.replace(/\\/+$/, '');",
  "  const finalDir = folder ? `${libraryRoot}/${folder}` : libraryRoot;",
  "  const stem = pageCopy.fileStem(snapshot.fileName) || 'Page';",
  "  const candidates = [`${stem}.aspx`, `${stem}-copy.aspx`, `${stem}-copy-2.aspx`];",
  "  const takenFinal = new Set();",
  "  for (const name of candidates) {",
  "    if (await pair.pages.exists(`${finalDir}/${name}`)) takenFinal.add(name.toLowerCase());",
  "  }",
  "  return {",
  "    webUrl: web.webUrl, webServerRelativeUrl: web.webServerRelativeUrl,",
  "    webId: web.webId, siteId: web.siteId,",
  "    library: { id: library.id, rootPath: library.rootPath },",
  "    fields: fieldsRes.items,",
  "    takenFinal, takenRoot: new Set(),",
  "    assetFolderExists: false,",
  "  };",
  "}",
  "",
  "function slimPlan(plan) {",
  "  return {",
  "    engine: plan.engine, createPath: plan.createPath, sameWeb: plan.sameWeb, sameSite: plan.sameSite,",
  "    title: plan.title, publish: plan.publish, promoteAsNews: plan.promoteAsNews,",
  "    source: plan.source, target: plan.target, pageLayoutType: plan.pageLayoutType,",
  "    blockers: plan.blockers, warnings: plan.warnings,",
  "  };",
  "}",
  "",
  "window.__pcLegacy = {",
  "  async analyze() {",
  "    resetPageCopyMock();",
  "    const src = await connectPair('/sites/pagesrc');",
  "    const snapshot = await readSnapshot(src, LEGACY_PAGE_ID);",
  "",
  "    const sameWebPair = await connectPair('/sites/pagesrc');",
  "    const sameWebTarget = await buildTarget(sameWebPair, snapshot, {});",
  "    const sameWebPlan = pageCopy.analyzeCopy({ snapshot, target: sameWebTarget, options: {} });",
  "",
  "    const crossWebPair = await connectPair('/sites/pagedst');",
  "    const crossWebTarget = await buildTarget(crossWebPair, snapshot, {});",
  "    const crossWebPlan = pageCopy.analyzeCopy({ snapshot, target: crossWebTarget, options: {} });",
  "",
  "    return { sameWeb: slimPlan(sameWebPlan), crossWeb: slimPlan(crossWebPlan) };",
  "  },",
  "",
  "  async run(opts) {",
  "    opts = opts || {};",
  "    const sourceWeb = opts.sourceWeb || '/sites/pagesrc';",
  "    const pageId = opts.pageId || LEGACY_PAGE_ID;",
  "    const targetWeb = opts.targetWeb || sourceWeb;",
  "    const options = opts.options || {};",
  "    if (opts.reset !== false) { resetPageCopyMock(); globalThis.__PAGECOPY_MOCK_FAIL__ = null; }",
  "",
  "    if (opts.promoted) {",
  "      const state = globalThis.__PAGECOPY_MOCK__.state[sourceWeb];",
  "      const fixture = state.pages.get(pageId);",
  "      fixture.dto.PromotedState = 2;",
  "      fixture.item.PromotedState = 2;",
  "    }",
  "",
  "    const src = await connectPair(sourceWeb);",
  "    const snapshot = await readSnapshot(src, pageId);",
  "    const dst = await connectPair(targetWeb);",
  "    const target = await buildTarget(dst, snapshot, options);",
  "    const plan = pageCopy.analyzeCopy({ snapshot, target, options });",
  "",
  "    if (opts.injectFailMatch) {",
  "      globalThis.__PAGECOPY_MOCK_FAIL__ = [{ match: opts.injectFailMatch.match, code: opts.injectFailMatch.code || 'write', status: opts.injectFailMatch.status || 500 }];",
  "    }",
  "",
  "    const stepLog = [];",
  "    const deps = {",
  "      source: { pages: src.pages },",
  "      target: { pages: dst.pages, write: dst.write },",
  "      rewrite: (transferResults) => pageCopy.rewriteContent(snapshot, plan, transferResults),",
  "      verify: (dto, saved) => pageCopy.compareReadBack(plan, saved, dto),",
  "    };",
  "    const result = await runCopy({ plan }, deps, {",
  "      onStep: (record) => stepLog.push({ name: record.name, status: record.status, detail: record.detail }),",
  "    });",
  "",
  "    let postDto = null;",
  "    if (result.journal.pageId) {",
  "      try { postDto = await dst.pages.getPage(result.journal.pageId); } catch (err) { /* left null — the page may not exist after a failed run */ }",
  "    }",
  "",
  "    let discard = null;",
  "    if (opts.thenDiscard) {",
  "      discard = await discardCopy(result.journal, { target: { pages: dst.pages } });",
  "    }",
  "",
  "    return {",
  "      plan: slimPlan(plan),",
  "      result, stepLog,",
  "      writes: (globalThis.__DCSPAD_WB_WRITES__ || []).map((w) => ({ url: w.url, body: w.body })),",
  "      postDto, discard,",
  "    };",
  "  },",
  "};",
];

export async function run({ browser, check, WB_URL }) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
  // addScriptTag's promise resolves once the <script type=module> element's
  // own load event fires, which in practice can land before the module body
  // has finished executing (its dependency graph is still being evaluated) —
  // waiting for window.__pcLegacy directly is what actually guarantees the
  // helper is ready.
  await page.addScriptTag({ type: 'module', content: HELPER_LINES.join('\n') });
  await page.waitForFunction(() => typeof window.__pcLegacy === 'object');

  // ---- eligibility: same-web plans copyFile clean; cross-web is blocked ----

  const analysis = await page.evaluate(() => window.__pcLegacy.analyze());

  await check('legacy: same-web plans engine "copyFile" with no blockers; cross-web to /sites/pagedst blocks with a reason naming the legacy HTML canvas format', () =>
    analysis.sameWeb.engine === 'copyFile'
    && Array.isArray(analysis.sameWeb.blockers) && analysis.sameWeb.blockers.length === 0
    && analysis.crossWeb.engine === 'copyFile'
    && Array.isArray(analysis.crossWeb.blockers)
    && analysis.crossWeb.blockers.some((b) => /legacy html canvas format/i.test(b)));

  // ---- same-web run: exactly one CopyFileByPath, correct request shape ----

  const dup = await page.evaluate(() => window.__pcLegacy.run({}));

  await check('legacy: a same-web run posts exactly one SP.MoveCopyUtil.CopyFileByPath with absolute srcPath/destPath, ResetAuthorAndCreatedOnCopy true and overwrite false, and never sitepages/pages create, MoveFileByPath or an asset upload', () => {
    const copyWrites = dup.writes.filter((w) => /sp\.movecopyutil\.copyfilebypath/i.test(w.url));
    const usedCreate = dup.writes.some((w) => /\/sitepages\/pages(\?|$)/i.test(new URL(w.url).pathname));
    const usedMove = dup.writes.some((w) => /sp\.movecopyutil\.movefilebypath/i.test(w.url));
    const usedUpload = dup.writes.some((w) => /addusingpath/i.test(w.url));
    if (copyWrites.length !== 1) return false;
    const overwriteParam = /[?&]@a1=(true|false)/i.exec(copyWrites[0].url);
    let body = {};
    try { body = JSON.parse(copyWrites[0].body); } catch { /* leave {} — assertion below fails */ }
    return /^https?:\/\//.test(body.srcPath?.DecodedUrl || '')
      && /^https?:\/\//.test(body.destPath?.DecodedUrl || '')
      && body.options?.ResetAuthorAndCreatedOnCopy === true
      && Boolean(overwriteParam) && overwriteParam[1].toLowerCase() === 'false'
      && !usedCreate && !usedMove && !usedUpload;
  });

  await check('legacy: the run outcome is done or done-with-warnings and the copy lands as Legacy-News-copy.aspx', () =>
    (dup.result.outcome === 'done' || dup.result.outcome === 'done-with-warnings')
    && dup.plan.target.fileName === 'Legacy-News-copy.aspx'
    && dup.plan.target.finalPath === '/sites/pagesrc/SitePages/Legacy-News-copy.aspx');

  // ---- promoted source: reset clears PromotedState before checkout/save ----

  const promoted = await page.evaluate(() => window.__pcLegacy.run({ promoted: true }));

  await check('legacy: when the source is promoted, the reset step\'s ValidateUpdateListItem (PromotedState "0") is posted before checkoutpage and savepage, and the copy reads back PromotedState 0', () => {
    const shapeOf = (u) => {
      const p = new URL(u).pathname.toLowerCase();
      if (/validateupdatelistitem$/.test(p)) return 'metadata';
      if (/checkoutpage$/.test(p)) return 'checkout';
      if (/savepage$/.test(p)) return 'save';
      if (/copyfilebypath/i.test(u)) return 'copy';
      return `other:${p}`;
    };
    const shapes = promoted.writes.map((w) => shapeOf(w.url));
    const metaIdx = shapes.indexOf('metadata');
    const checkoutIdx = shapes.indexOf('checkout');
    const saveIdx = shapes.indexOf('save');
    let metaBody = {};
    try { metaBody = JSON.parse(promoted.writes[metaIdx]?.body || '{}'); } catch { /* leave {} */ }
    const promotedFieldValue = (metaBody.formValues || []).find((f) => f.FieldName === 'PromotedState');
    return metaIdx !== -1 && checkoutIdx !== -1 && saveIdx !== -1
      && metaIdx < checkoutIdx && metaIdx < saveIdx
      && Boolean(promotedFieldValue) && promotedFieldValue.FieldValue === '0'
      && Number((promoted.postDto || {}).PromotedState) === 0;
  });

  // ---- content is carried byte-for-byte by the file copy itself -----------

  await check('legacy: the copy\'s CanvasContent1 equals the source\'s legacy HTML string', () =>
    typeof (dup.postDto || {}).CanvasContent1 === 'string'
    && dup.postDto.CanvasContent1 === '<div><div data-sp-canvascontrol></div></div>');

  // ---- publish only when ticked --------------------------------------------

  const unpublished = await page.evaluate(() => window.__pcLegacy.run({}));
  const published = await page.evaluate(() => window.__pcLegacy.run({ options: { publish: true } }));

  await check('legacy: publish only happens when ticked — no /publish write and a checkin instead when not ticked; a /publish write and no checkin when ticked', () => {
    const hasPublish = (r) => r.writes.some((w) => /\/publish$/i.test(new URL(w.url).pathname));
    const hasCheckin = (r) => r.writes.some((w) => /\/checkin\(/i.test(w.url));
    return !hasPublish(unpublished) && hasCheckin(unpublished)
      && hasPublish(published) && !hasCheckin(published);
  });

  // ---- a failure on the reset savepage fails the run and can be discarded --

  const failed = await page.evaluate(() => window.__pcLegacy.run({
    injectFailMatch: { match: '/savepage', code: 'write', status: 500 },
    thenDiscard: true,
  }));

  await check('legacy: a failure on the reset step\'s savepage yields outcome "failed" with journal.createdBy "this run", and discardCopy recycles the copied file', () => {
    const resetStep = failed.stepLog.find((s) => s.name === 'reset');
    return failed.result.outcome === 'failed'
      && Boolean(resetStep) && resetStep.status === 'failed'
      && failed.result.journal.createdBy === 'this run'
      && Array.isArray(failed.discard.recycled) && failed.discard.recycled.length === 1
      && failed.discard.recycled[0] === failed.result.journal.currentPath;
  });

  await page.close();
}
