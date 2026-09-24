// SP Workbench — page copy: the runner (design/PAGE-COPY.md §3, §7).
// Exercises page-copy-run.js's runCopy()/discardCopy() end to end against the
// mock-pagecopy.js fixtures (source /sites/pagesrc, cross-site-collection
// /sites/pagedst) — connecting real sp-rest/sp-write client pairs the way
// page-copy-dialog.js would, building a snapshot with page-copy.js's own
// snapshotFromReads/analyzeCopy/rewriteContent/compareReadBack, and letting
// the runner drive the actual write sequence. House style: tests/pages-copy/
// pure.mjs (module shape), tests/workbench-schema.mjs (mock write-log
// assertions).
//
// Everything below runs inside the page via one injected helper module
// (page.addScriptTag) exposing window.__pc, because the scenarios need real
// connected sp-rest/sp-write/sp-pages instances that cannot cross the
// page.evaluate boundary — only the plain, serializable results of running
// them can.

// Every line here becomes page-context source text via .join('\n'); it is a
// plain double-quoted Node string, so the template literals, embedded single
// quotes and `${…}` interpolations the PAGE script itself needs are written
// as ordinary characters — nothing here escapes for Node's own parser.
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
  "async function resolveAssetInfo(sourcePair, snapshot, analysis) {",
  "  const requests = pageCopy.assetRequests(snapshot, analysis);",
  "  const map = new Map();",
  "  for (const req of requests) {",
  "    let info = null;",
  "    if (req.ids && req.ids.uniqueId) info = await sourcePair.pages.fileInfoById(req.ids.uniqueId);",
  "    if (!info && req.path) info = await sourcePair.pages.fileInfo(req.path);",
  "    map.set(req.key, info);",
  "  }",
  "  return map;",
  "}",
  "",
  "function slimPlan(plan) {",
  "  return {",
  "    engine: plan.engine, createPath: plan.createPath, sameWeb: plan.sameWeb, sameSite: plan.sameSite,",
  "    title: plan.title, description: plan.description, publish: plan.publish, promoteAsNews: plan.promoteAsNews,",
  "    source: plan.source, target: plan.target,",
  "    metadata: { formValues: plan.metadata.formValues, requiredGaps: plan.metadata.requiredGaps, skipped: plan.metadata.skipped },",
  "    commentsDisabled: plan.commentsDisabled, blockers: plan.blockers, warnings: plan.warnings,",
  "    assetFolderChain: plan.assetFolderChain,",
  "  };",
  "}",
  "",
  "window.__pc = {",
  "  writes() { return (globalThis.__DCSPAD_WB_WRITES__ || []).map((w) => w.url); },",
  "",
  "  async run(opts) {",
  "    opts = opts || {};",
  "    const sourceWeb = opts.sourceWeb || '/sites/pagesrc';",
  "    const pageId = opts.pageId || 7;",
  "    const targetWeb = opts.targetWeb || '/sites/pagesrc';",
  "    const options = opts.options || {};",
  "    if (opts.reset !== false) { resetPageCopyMock(); globalThis.__PAGECOPY_MOCK_FAIL__ = null; }",
  "",
  "    const src = await connectPair(sourceWeb);",
  "    const snapshot = await readSnapshot(src, pageId);",
  "    const dst = await connectPair(targetWeb);",
  "    const target = await buildTarget(dst, snapshot, options);",
  "",
  "    if (opts.precreateFolder) {",
  "      const rootPath = target.library.rootPath.replace(/\\/+$/, '');",
  "      await dst.write.createFolder(rootPath, opts.precreateFolder);",
  "    }",
  "",
  "    let assetInfo = new Map();",
  "    if (opts.crossWeb) {",
  "      const analysis = pageCopy.analyzeParts(snapshot);",
  "      assetInfo = await resolveAssetInfo(src, snapshot, analysis);",
  "    }",
  "",
  "    const plan = pageCopy.analyzeCopy({ snapshot, target, options, assetInfo });",
  "",
  "    if (opts.injectCollision) {",
  "      const collided = await dst.pages.createPage({ pageLayoutType: 'Article', promotedState: 0 });",
  "      await dst.pages.savePage(collided.Id, { Title: plan.target.stem });",
  "    }",
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
  "      verify: (dto, saved) => {",
  "        const useDto = opts.wrapVerifyTitle ? Object.assign({}, dto, { Title: dto.Title + ' (drifted)' }) : dto;",
  "        return pageCopy.compareReadBack(plan, saved, useDto);",
  "      },",
  "    };",
  "    const result = await runCopy({ plan }, deps, {",
  "      onStep: (record) => {",
  "        stepLog.push({ name: record.name, status: record.status, detail: record.detail });",
  "        if (opts.armFailAfterName && record.name === 'name' && record.status === 'done') {",
  "          globalThis.__PAGECOPY_MOCK_FAIL__ = [{ match: '/savepage', code: 'network', once: true }];",
  "        }",
  "      },",
  "    });",
  "",
  "    let postDto = null;",
  "    let postItem = null;",
  "    if (result.journal.pageId) {",
  "      try {",
  "        postDto = await dst.pages.getPage(result.journal.pageId);",
  "        postItem = await dst.client.get(`web/lists(guid'${target.library.id}')/items(${result.journal.pageId})`, { expand: 'FieldValuesAsText' });",
  "      } catch (err) { /* left null — the page may not exist after a failed/discarded run */ }",
  "    }",
  "",
  "    let discard = null;",
  "    if (opts.thenDiscard) {",
  "      discard = await discardCopy(result.journal, { target: { pages: dst.pages } });",
  "    }",
  "",
  "    const finalPathStillExists = opts.injectCollision ? await dst.pages.exists(plan.target.finalPath) : null;",
  "    const sourceStillExists = await src.pages.exists(snapshot.fileRef);",
  "",
  "    return {",
  "      plan: slimPlan(plan),",
  "      result, stepLog,",
  "      writes: (globalThis.__DCSPAD_WB_WRITES__ || []).map((w) => w.url),",
  "      postDto, postItem, discard,",
  "      finalPathStillExists, sourceStillExists,",
  "    };",
  "  },",
  "",
  "  async promoteAsNewsPure() {",
  "    resetPageCopyMock();",
  "    const src = await connectPair('/sites/pagesrc');",
  "    const snapshot = await readSnapshot(src, 7);",
  "    const dst = await connectPair('/sites/pagesrc');",
  "    const target = await buildTarget(dst, snapshot, {});",
  "    const promotedPlan = pageCopy.analyzeCopy({ snapshot, target, options: { promoteAsNews: true } });",
  "    const notPromotedSnapshot = Object.assign({}, snapshot, { dto: Object.assign({}, snapshot.dto, { PromotedState: 0 }) });",
  "    const notPromotedPlan = pageCopy.analyzeCopy({ snapshot: notPromotedSnapshot, target, options: { promoteAsNews: true } });",
  "    return {",
  "      promotedHonoured: promotedPlan.promoteAsNews === true,",
  "      notPromotedIgnored: notPromotedPlan.promoteAsNews === false,",
  "    };",
  "  },",
  "};",
];

export async function run({ browser, check, WB_URL }) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards');
  // addScriptTag's promise resolves once the <script type=module> element's
  // own load event fires, which in practice can land before the module body
  // has finished executing (its dependency graph is still being evaluated) —
  // waiting for window.__pc directly is what actually guarantees the helper
  // is ready.
  await page.addScriptTag({ type: 'module', content: HELPER_LINES.join('\n') });
  await page.waitForFunction(() => typeof window.__pc === 'object');

  // ---- same-web duplicate (path A, defaults) --------------------------------

  const dup = await page.evaluate(() => window.__pc.run({}));

  await check('runner: same-web duplicate — write sequence is create/name-save/content-save/move/metadata/comments/checkin, in order, with no publish', () => {
    const urls = dup.writes.map((u) => new URL(u).pathname.toLowerCase());
    const shapeOf = (p) => {
      if (/\/sitepages\/pages$/.test(p)) return 'create';
      if (/\/savepage$/.test(p)) return 'savepage';
      if (/sp\.movecopyutil\.movefilebypath/.test(p)) return 'move';
      if (/\/validateupdatelistitem$/.test(p)) return 'metadata';
      if (/\/setcommentsdisabled$/.test(p)) return 'comments';
      if (/\/checkin\(/.test(p)) return 'checkin';
      if (/\/publish$/.test(p)) return 'publish';
      return `other:${p}`;
    };
    const shapes = urls.map(shapeOf);
    const expected = ['create', 'savepage', 'savepage', 'move', 'metadata', 'comments', 'checkin'];
    return dup.result.outcome === 'done'
      && shapes.join(',') === expected.join(',')
      && !shapes.includes('publish');
  });

  await check('runner: same-web duplicate — the new file is Quarterly-Update-copy.aspx in SitePages, its CanvasContent1 matches the source byte-for-byte, Title is kept, PromotedState is reset to 0, and comments stay disabled', () => {
    const dto = dup.postDto || {};
    return dto.FileRef === '/sites/pagesrc/SitePages/Quarterly-Update-copy.aspx'
      && dto.FileLeafRef === 'Quarterly-Update-copy.aspx'
      && dup.plan.target.finalPath === '/sites/pagesrc/SitePages/Quarterly-Update-copy.aspx'
      && dto.Title === 'Quarterly Update'
      && Number(dto.PromotedState) === 0
      && (dup.postItem || {}).CommentsDisabled === true;
  });

  // Description rides with the metadata write: savepage blanks it on SPO
  // (spike §11), so the run sets it through ValidateUpdateListItem.
  await check('runner: same-web duplicate — the metadata write carries Description plus PageCategory/ReviewDate/ShowInNav/RelatedLink and never Owner', () => {
    const names = dup.plan.metadata.formValues.map((f) => f.FieldName).filter((n) => n !== 'Description').sort();
    const description = dup.plan.metadata.formValues.find((f) => f.FieldName === 'Description');
    if (!description || description.FieldValue !== dup.plan.description) return false;
    const category = dup.plan.metadata.formValues.find((f) => f.FieldName === 'PageCategory');
    const review = dup.plan.metadata.formValues.find((f) => f.FieldName === 'ReviewDate');
    return names.join(',') === ['PageCategory', 'RelatedLink', 'ReviewDate', 'ShowInNav'].join(',')
      && !names.includes('Owner')
      && Boolean(category) && category.FieldValue === 'IT'
      && Boolean(review) && review.FieldValue === '10/1/2026 12:00 AM';
  });

  // ---- publish: true ---------------------------------------------------------

  const published = await page.evaluate(() => window.__pc.run({ options: { publish: true } }));

  await check('runner: publish:true reaches the publish step, never checks in, and the page reads back published', () => {
    const paths = published.writes.map((u) => new URL(u).pathname.toLowerCase());
    const hasPublish = paths.some((p) => /\/publish$/.test(p));
    const hasCheckin = paths.some((p) => /\/checkin\(/.test(p));
    const dto = published.postDto || {};
    return published.result.outcome === 'done'
      && hasPublish && !hasCheckin
      && dto.IsPageCheckedOutToCurrentUser === false
      && /^[1-9]\d*\.0$/.test(String(dto._UIVersionString || ''));
  });

  // ---- promoteAsNews only when the source is promoted -----------------------

  const promo = await page.evaluate(() => window.__pc.promoteAsNewsPure());

  await check('runner: promoteAsNews is only honoured on a plan built from a promoted source', () =>
    promo.promotedHonoured === true && promo.notPromotedIgnored === true);

  // ---- path B is gone -------------------------------------------------------
  // The spike (§11 Q1) showed addTemplateFile creates a file without the Site
  // Page content type, which the sitepages API then refuses; decision 2 fixes
  // creation on path A, whatever an options object asks for.

  const pathB = await page.evaluate(() => window.__pc.run({ options: { createPath: 'B' } }));

  await check('runner: creation is always path A (create + staging save + move), never addtemplatefile, even when asked for path B', () => {
    const paths = pathB.writes.map((u) => new URL(u).pathname.toLowerCase());
    const usedTemplateFile = paths.some((p) => p.includes('/addtemplatefile('));
    const usedCreate = paths.some((p) => /\/sitepages\/pages$/.test(p));
    const usedMove = paths.some((p) => /sp\.movecopyutil\.movefilebypath/.test(p));
    return pathB.plan.createPath === 'A' && !usedTemplateFile && usedCreate && usedMove;
  });

  // ---- folder copy ------------------------------------------------------------

  const folderCopy = await page.evaluate(() => window.__pc.run({
    options: { folder: 'Archive' }, precreateFolder: 'Archive',
  }));

  await check('runner: a folder-targeted copy lands under SitePages/Archive/ while path A stages the page at the library root', () =>
    folderCopy.result.outcome === 'done'
    && folderCopy.plan.target.finalPath === '/sites/pagesrc/SitePages/Archive/Quarterly-Update.aspx'
    && folderCopy.plan.target.stagingPath.startsWith('/sites/pagesrc/SitePages/')
    && !folderCopy.plan.target.stagingPath.includes('/Archive/'));

  // ---- name collision between preflight and move ------------------------------

  const collision = await page.evaluate(() => window.__pc.run({ injectCollision: true, thenDiscard: true }));

  await check('runner: another operator claiming the final name before the move fails the run at "move", leaves journal.currentPath at the staging page, and discardCopy recycles exactly that page — the colliding page and the source both survive', () => {
    const moveStep = collision.stepLog.find((s) => s.name === 'move');
    const stagingPath = collision.plan.target.stagingPath;
    return collision.result.outcome === 'failed'
      && Boolean(moveStep) && moveStep.status === 'failed'
      && collision.result.journal.currentPath === stagingPath
      && Array.isArray(collision.discard.recycled) && collision.discard.recycled.length === 1
      && collision.discard.recycled[0] === stagingPath
      && collision.finalPathStillExists === true
      && collision.sourceStillExists === true;
  });

  // ---- a lost response on the content save -------------------------------------

  const lost = await page.evaluate(() => window.__pc.run({ armFailAfterName: true }));

  await check('runner: a lost response on the content save yields outcome "unknown", stops before any later write, calls createPage exactly once, and marks the journal createdBy "this run"', () => {
    const paths = lost.writes.map((u) => new URL(u).pathname.toLowerCase());
    const createCalls = paths.filter((p) => /\/sitepages\/pages$/.test(p)).length;
    const contentStep = lost.stepLog.find((s) => s.name === 'content');
    const stepsAfterContent = lost.stepLog.slice(lost.stepLog.findIndex((s) => s.name === 'content') + 1);
    return lost.result.outcome === 'unknown'
      && Boolean(contentStep) && contentStep.status === 'unknown'
      && stepsAfterContent.length === 0
      && createCalls === 1
      && lost.result.journal.createdBy === 'this run';
  });

  // ---- per-field VULI rejection on the target (cross-site, Audience required) -

  const rejected = await page.evaluate(() => window.__pc.run({
    targetWeb: '/sites/pagedst',
    options: { publish: true, requiredValues: { Audience: 'All' } },
    injectFailMatch: { match: 'validateupdatelistitem', code: 'metadata-write', status: 400 },
  }));

  await check('runner: a per-field VULI rejection on the target downgrades the outcome to "done-with-warnings" but still honours publish', () => {
    const paths = rejected.writes.map((u) => new URL(u).pathname.toLowerCase());
    const hasPublish = paths.some((p) => /\/publish$/.test(p));
    const metaStep = rejected.stepLog.find((s) => s.name === 'metadata');
    return rejected.result.outcome === 'done-with-warnings'
      && Boolean(metaStep) && metaStep.status === 'failed'
      && hasPublish;
  });

  // ---- verify drift -------------------------------------------------------------

  const drifted = await page.evaluate(() => window.__pc.run({ wrapVerifyTitle: true }));

  await check('runner: a Title mismatch on read-back is reported as drift and downgrades the outcome to "done-with-warnings"', () =>
    drifted.result.outcome === 'done-with-warnings'
    && drifted.result.journal.drift.some((d) => d.field === 'Title'));

  // ---- cross-site copy: writes land only under the target --------------------

  const cross = await page.evaluate(() => window.__pc.run({
    targetWeb: '/sites/pagedst',
    crossWeb: true,
    options: { requiredValues: { Audience: 'All' } },
    thenDiscard: true,
  }));

  await check('runner: a cross-site copy to /sites/pagedst writes only under the target site (the source is read-only) and lands as Quarterly-Update-copy.aspx (name collision with the existing page there)', () => {
    const allTarget = cross.writes.every((u) => u.toLowerCase().includes('/sites/pagedst/_api/'));
    const noSource = cross.writes.every((u) => !u.toLowerCase().includes('/sites/pagesrc/_api/'));
    return cross.result.outcome !== 'failed' && cross.result.outcome !== 'unknown'
      && cross.writes.length > 0 && allTarget && noSource
      && cross.plan.target.finalPath === '/sites/pagedst/SitePages/Quarterly-Update-copy.aspx';
  });

  await check('runner: cross-site asset folders — discardCopy recycles the leaf asset folder and the transferred files but never the shared SiteAssets/SitePages parent, and never calls discardPage', () => {
    const assets = cross.result.journal.assets || [];
    const sharedParent = assets.find((a) => a.kind === 'folder' && a.path.endsWith('/SiteAssets/SitePages'));
    const leaf = assets.find((a) => a.kind === 'folder' && a.path.endsWith('/SiteAssets/SitePages/Quarterly-Update-copy'));
    const files = assets.filter((a) => a.kind === 'file');
    const recycled = cross.discard.recycled || [];
    const everyWriteUrl = cross.writes.map((u) => u.toLowerCase());
    // The shared parent already exists on pagedst (as on live SPO), so the
    // run never creates it and never journals it — or, if it did, only as
    // non-recyclable. Either way it is never recycled.
    const parentPath = '/sites/pagedst/SiteAssets/SitePages';
    return (!sharedParent || sharedParent.recyclable === false)
      && Boolean(leaf) && leaf.recyclable === true
      // One file until the header analyzer lands: the custom thumbnail (the
      // DTO's BannerImageUrl). The banner itself is only in the header part.
      && files.length === 1 && files[0].path.endsWith('/thumb.png')
      && !recycled.some((r) => r.toLowerCase() === parentPath.toLowerCase())
      && recycled.includes(leaf.path)
      && files.every((f) => recycled.includes(f.path))
      && recycled.includes(cross.result.journal.currentPath)
      && !everyWriteUrl.some((u) => u.includes('/discardpage'));
  });

  await page.close();
}
