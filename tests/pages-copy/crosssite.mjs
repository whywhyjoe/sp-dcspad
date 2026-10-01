// SP Workbench — page copy: cross-site copies in mock mode (design/PAGE-COPY.md
// §5, §8, §11). Exercises the real cross-site-collection path (/sites/pagesrc
// -> /sites/pagedst) and the same-site-collection subweb path (/sites/pagesrc
// -> /sites/pagesrc/team) end to end: analyzeCopy -> runCopy -> the saved
// mock DTO, always passing `analyzers: ANALYZERS` so the shipped analyzers
// (text, quick links, news, list/library) actually patch content — house
// style and helper approach borrowed from tests/pages-copy/runner.mjs.
//
// Two web parts referenced by the fixture page (Call to action, Video) have
// no analyzer yet (page-copy-analyzers.js only registers text/quick-links/
// news/list-library); two checks below (#3 dedup naming, #6 oversized-asset
// retention) need an *asset*-classified reference that isn't naturally
// produced by the shipped registry, so they supplement it locally with a
// tiny stand-in analyzer for exactly that one web part id, scoped to that
// single run — this is test-only code, nothing in src/ changes.

const HELPER_LINES = [
  "import { createSpRestClient } from '/src/workbench/sp-rest.js';",
  "import { createSpWriteClient } from '/src/workbench/sp-write.js';",
  "import { mockResolver, mockWriter } from '/src/workbench/mock-data.js';",
  "import { createSpPages } from '/src/workbench/sp-pages.js';",
  "import * as pageCopy from '/src/workbench/page-copy.js';",
  "import { ANALYZERS } from '/src/workbench/page-copy-analyzers.js';",
  "import { runCopy, discardCopy } from '/src/workbench/page-copy-run.js';",
  "import { resetPageCopyMock } from '/src/workbench/mock-pagecopy.js';",
  "",
  "const FIELD_SELECT = ['Id', 'Title', 'InternalName', 'TypeAsString', 'FieldTypeKind', 'Required', 'Hidden', 'ReadOnlyField', 'Group', 'DefaultValue', 'Choices', 'Description', 'FillInChoice'];",
  "",
  "// A stand-in analyzer for the Call to action background image — CTA has no",
  "// real analyzer yet, but the fixture already carries a real",
  "// imageSources.imageSource + customMetadata.imageSource shape for it",
  "// (mock-pagecopy.js control d15e...005), so this exercises the real",
  "// dedup-naming path (page-copy.js's uniqueName()) against two genuinely",
  "// different source files that share a basename (icons/logo.png vs",
  "// img/logo.png).",
  "const CTA_ID = 'df8e44e7-edd5-46d5-90da-aca1539313b8';",
  "const ctaAnalyzer = {",
  "  id: CTA_ID,",
  "  refs(instance) {",
  "    const spc = instance?.webPartData?.serverProcessedContent || {};",
  "    const src = spc.imageSources && spc.imageSources.imageSource;",
  "    if (!src) return [];",
  "    const meta = (spc.customMetadata && spc.customMetadata.imageSource) || {};",
  "    return [{ path: ['webPartData', 'serverProcessedContent', 'imageSources', 'imageSource'], value: src, class: 'asset', asset: { path: src, ids: meta } }];",
  "  },",
  "  patch(instance, ctx) {",
  "    const spc = instance && instance.webPartData && instance.webPartData.serverProcessedContent;",
  "    const src = spc && spc.imageSources && spc.imageSources.imageSource;",
  "    if (!src) return 0;",
  "    const meta = (spc.customMetadata && spc.customMetadata.imageSource) || {};",
  "    const mapped = ctx.mapAsset({ path: src, ids: meta });",
  "    if (!mapped) return 0;",
  "    spc.imageSources.imageSource = mapped.url || mapped.path;",
  "    return 1;",
  "  },",
  "};",
  "",
  "// A stand-in analyzer for the Video part's background — Video has no real",
  "// analyzer yet either; this exercises the oversized-asset retain path",
  "// (page-copy.js's MAX_ASSET_BYTES check) against the real 60 MB fixture",
  "// file (mock-pagecopy.js 'huge.mp4').",
  "const VIDEO_ID = '275c0095-a77e-4f6d-a2a0-6a7626911518';",
  "const videoAnalyzer = {",
  "  id: VIDEO_ID,",
  "  refs(instance) {",
  "    const src = instance?.webPartData?.serverProcessedContent?.links?.videoSource;",
  "    if (!src) return [];",
  "    return [{ path: ['webPartData', 'serverProcessedContent', 'links', 'videoSource'], value: src, class: 'asset', asset: { path: src, ids: {} } }];",
  "  },",
  "  patch() { return 0; },",
  "};",
  "",
  "function analyzersFor(flag) {",
  "  const extra = flag === 'cta' ? ctaAnalyzer : flag === 'video' ? videoAnalyzer : null;",
  "  if (!extra) return ANALYZERS;",
  "  return {",
  "    header: ANALYZERS.header,",
  "    text: ANALYZERS.text,",
  "    forId: (id) => (String(id || '').toLowerCase() === extra.id.toLowerCase() ? extra : ANALYZERS.forId(id)),",
  "  };",
  "}",
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
  "  const wpItems = await pair.pages.clientSideWebParts();",
  "  const webParts = new Set(wpItems.map((w) => pageCopy.normalizeGuid(w.Id)));",
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
  "    webParts,",
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
  "    engine: plan.engine, sameWeb: plan.sameWeb, sameSite: plan.sameSite,",
  "    title: plan.title, description: plan.description, rewriteLinks: plan.rewriteLinks,",
  "    source: plan.source, target: plan.target,",
  "    assets: plan.assets.map((a) => ({ key: a.key, name: a.name, sourcePath: a.sourcePath, destName: a.destName, destPath: a.destPath, retainReason: a.retainReason })),",
  "    metadata: { formValues: plan.metadata.formValues, requiredGaps: plan.metadata.requiredGaps },",
  "    warnings: plan.warnings, blockers: plan.blockers,",
  "    parts: plan.analysis.parts.map((p) => ({ index: p.index, label: p.label, webPartId: p.webPartId, kind: p.kind, available: p.available })),",
  "  };",
  "}",
  "",
  "window.__cs = {",
  "  async run(opts) {",
  "    opts = opts || {};",
  "    const sourceWeb = opts.sourceWeb || '/sites/pagesrc';",
  "    const pageId = opts.pageId || 7;",
  "    const targetWeb = opts.targetWeb || '/sites/pagedst';",
  "    const userOptions = opts.options || {};",
  "    if (opts.reset !== false) { resetPageCopyMock(); globalThis.__PAGECOPY_MOCK_FAIL__ = null; }",
  "",
  "    const analyzers = analyzersFor(opts.extraAnalyzer);",
  "",
  "    const src = await connectPair(sourceWeb);",
  "    const snapshot = await readSnapshot(src, pageId);",
  "    const dst = await connectPair(targetWeb);",
  "    const target = await buildTarget(dst, snapshot, userOptions);",
  "",
  "    const analysisForAssets = pageCopy.analyzeParts(snapshot, { analyzers, targetWebParts: target.webParts });",
  "    const assetInfo = await resolveAssetInfo(src, snapshot, analysisForAssets);",
  "",
  "    const requiredValues = userOptions.requiredValues !== undefined ? userOptions.requiredValues",
  "      : (targetWeb === '/sites/pagedst' ? { Audience: 'All' } : {});",
  "    const options = Object.assign({}, userOptions, { requiredValues });",
  "",
  "    const plan = pageCopy.analyzeCopy({ snapshot, target, options, analyzers, assetInfo });",
  "",
  "    if (opts.injectFailMatch) {",
  "      globalThis.__PAGECOPY_MOCK_FAIL__ = [{ match: opts.injectFailMatch.match, code: opts.injectFailMatch.code || 'write', status: opts.injectFailMatch.status || 500 }];",
  "    }",
  "",
  "    const stepLog = [];",
  "    const deps = {",
  "      source: { pages: src.pages },",
  "      target: { pages: dst.pages, write: dst.write },",
  "      rewrite: (transferResults) => pageCopy.rewriteContent(snapshot, plan, transferResults, analyzers),",
  "      verify: (dto, saved) => pageCopy.compareReadBack(plan, saved, dto),",
  "    };",
  "    const result = await runCopy({ plan }, deps, {",
  "      onStep: (record) => stepLog.push({ name: record.name, status: record.status, detail: record.detail }),",
  "    });",
  "",
  "    let postDto = null;",
  "    if (result.journal.pageId) {",
  "      try { postDto = await dst.pages.getPage(result.journal.pageId); } catch (err) { /* left null — page may not exist after a failed run */ }",
  "    }",
  "",
  "    let discard = null;",
  "    if (opts.thenDiscard) {",
  "      discard = await discardCopy(result.journal, { target: { pages: dst.pages } });",
  "    }",
  "",
  "    return {",
  "      plan: slimPlan(plan),",
  "      targetIdentity: { siteId: target.siteId, webId: target.webId, webServerRelativeUrl: target.webServerRelativeUrl },",
  "      result, stepLog,",
  "      writes: (globalThis.__DCSPAD_WB_WRITES__ || []).map((w) => w.url),",
  "      postDto, discard,",
  "    };",
  "  },",
  "};",
];

export async function run({ browser, check, WB_URL }) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards', { timeout: 60000 });
  await page.addScriptTag({ type: 'module', content: HELPER_LINES.join('\n') });
  await page.waitForFunction(() => typeof window.__cs === 'object');

  // ---- #1 writes land only under the target; outcome is done or done-with-warnings --

  const main = await page.evaluate(() => window.__cs.run({}));

  await check('crosssite: a pagesrc -> pagedst run writes only under /sites/pagedst/_api/ and finishes done or done-with-warnings', () => {
    const allTarget = main.writes.length > 0 && main.writes.every((u) => u.toLowerCase().includes('/sites/pagedst/_api/'));
    const noSource = main.writes.every((u) => !u.toLowerCase().includes('/sites/pagesrc/_api/'));
    return allTarget && noSource
      && (main.result.outcome === 'done' || main.result.outcome === 'done-with-warnings');
  });

  // ---- #2 assets land in the right destination folder, one entry per file -----------
  // The header analyzer (page-copy-analyzers/header.js) is now registered, so the
  // title-area banner (banner.jpg) transfers alongside the custom thumbnail (via the
  // DTO's own BannerImageUrl, handled outside any analyzer — §11), the Text part's
  // image (chart.png) and the Quick links thumbnail (logo.png): four distinct source
  // files, each appearing exactly once in the plan (page-copy.js's assetRequests()
  // dedupes by source identity — a repeated reference to the same file, as banner.jpg
  // and thumb.png each are here, still yields exactly one map entry).

  await check('crosssite: assets land under SiteAssets/SitePages/Quarterly-Update-copy/ on pagedst — the banner, the custom thumbnail, the Text part\'s image and the Quick links thumbnail, one each, none retained', () => {
    const destFolder = '/sites/pagedst/SiteAssets/SitePages/Quarterly-Update-copy';
    const transferred = main.plan.assets.filter((a) => a.destName);
    const byName = (n) => transferred.filter((a) => a.destName === n);
    const names = ['banner.jpg', 'thumb.png', 'chart.png', 'logo.png'];
    // Other parts' files (Image, Hero, File viewer…) may transfer too now
    // that their analyzers exist; every transfer lands in the one folder.
    return transferred.length >= 4
      && names.every((n) => byName(n).length === 1 && byName(n)[0].destPath === `${destFolder}/${n}`)
      && transferred.every((a) => a.destPath.startsWith(`${destFolder}/`))
      && main.plan.assets.every((a) => !a.retainReason);
  });

  // ---- #3 two different source files sharing a basename get distinct names --------

  const ctaRun = await page.evaluate(() => window.__cs.run({ extraAnalyzer: 'cta' }));

  await check('crosssite: two source files sharing a basename (icons/logo.png from Quick links, img/logo.png from Call to action) get distinct destination names logo.png / logo-2.png', () => {
    const logos = ctaRun.plan.assets.filter((a) => a.destName && /^logo(-2)?\.png$/.test(a.destName));
    const plain = logos.find((a) => a.sourcePath.endsWith('/icons/logo.png'));
    const numbered = logos.find((a) => a.sourcePath.endsWith('/img/logo.png'));
    return logos.length === 2
      && Boolean(plain) && plain.destName === 'logo.png'
      && Boolean(numbered) && numbered.destName === 'logo-2.png'
      && plain.destPath !== numbered.destPath;
  });

  // ---- #6 an oversized asset is retained, not transferred, and the run still succeeds --

  const oversized = await page.evaluate(() => window.__cs.run({ extraAnalyzer: 'video' }));

  await check('crosssite: the 60 MB video asset is retained (not transferred) with a warning naming the 50 MB limit, and the run still finishes done or done-with-warnings', () => {
    const huge = oversized.plan.assets.find((a) => a.name === 'huge.mp4');
    const warned = oversized.plan.warnings.some((w) => w.includes('huge.mp4') && w.includes('50 MB'));
    return Boolean(huge) && huge.destName === '' && huge.retainReason.includes('50 MB')
      && warned
      && (oversized.result.outcome === 'done' || oversized.result.outcome === 'done-with-warnings')
      && !oversized.writes.some((u) => u.toLowerCase().includes('huge.mp4'));
  });

  // ---- #4 the saved canvas/layout carry no source ids for transferred, analyzer- --
  // ---- patched files, and untouched parts keep their source references ----------

  await check('crosssite: the saved layout (title-area banner) and canvas (Quick links thumbnail, Text image) carry destination paths/ids, never the source ones', () => {
    const canvas = JSON.parse(main.postDto.CanvasContent1);
    const layout = JSON.parse(main.postDto.LayoutWebpartsContent);
    const header = layout.find((l) => l.id === 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788');
    const headerImage = header?.serverProcessedContent?.imageSources?.imageSource || '';
    const headerIds = header?.serverProcessedContent?.customMetadata?.imageSource || {};
    const quickLinks = canvas.find((c) => c.webPartId === 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd');
    const qlImage = quickLinks?.webPartData?.serverProcessedContent?.imageSources?.['items[0].image.url'] || '';
    const qlUniqueId = quickLinks?.webPartData?.properties?.items?.[0]?.uniqueId || '';
    const text = canvas.find((c) => c.controlType === 4);
    const textHtml = text?.innerHTML || '';

    return headerImage.toLowerCase().includes('/sites/pagedst/') && !headerImage.toLowerCase().includes('/sites/pagesrc/')
      && headerIds.uniqueId !== 'd15f0000-0000-4000-8000-000000000001'
      && qlImage.toLowerCase().includes('/sites/pagedst/') && !qlImage.toLowerCase().includes('/icons/logo.png')
      && qlUniqueId !== 'd15f0000-0000-4000-8000-000000000005'
      && textHtml.toLowerCase().includes('/sites/pagedst/') && !textHtml.includes('/sites/pagesrc/SiteAssets/img/chart.png');
  });

  await check('crosssite: an untouched part (List) still names the source list — a "data → warn" reference the copy never patches', () => {
    const canvas = JSON.parse(main.postDto.CanvasContent1);
    const list = canvas.find((c) => c.webPartId === 'f92bf067-bc19-489e-a556-7fe95f508720'
      && c.webPartData?.properties?.isDocumentLibrary !== true);
    return Boolean(list) && list.webPartData.properties.selectedListId === 'd1560000-0000-4000-8000-000000000001';
  });

  // ---- #5 the saved BannerImageUrl is the destination thumbnail, never afdcache or source --

  await check('crosssite: the saved BannerImageUrl is the destination thumbnail — never the afdcache CDN URL, never the source thumbnail', () => {
    const banner = String(main.postDto.BannerImageUrl || '').toLowerCase();
    return banner.includes('/sites/pagedst/') && banner.endsWith('/thumb.png')
      && !banner.includes('/sites/pagesrc/') && !banner.includes('afdcache');
  });

  // ---- #7 Quick chart is unavailable on pagedst — warned, but kept in the canvas ----

  await check('crosssite: Quick chart (absent from pagedst\'s GetClientSideWebParts) is reported unavailable and kept in the saved canvas', () => {
    const part = main.plan.parts.find((p) => p.webPartId === '91a50c94-865f-4f5c-8b4e-e49659e69772');
    const warned = main.plan.warnings.some((w) => w.startsWith('Quick chart') && w.includes('not available'));
    const canvas = JSON.parse(main.postDto.CanvasContent1);
    const stillThere = canvas.some((c) => c.webPartId === '91a50c94-865f-4f5c-8b4e-e49659e69772');
    return Boolean(part) && part.available === false && warned && stillThere;
  });

  // ---- #8 the site-scoped SPFx id and the unknown id are kept and reported unavailable --

  await check('crosssite: the site-scoped SPFx id and the unknown id are both reported unavailable and both kept in the saved canvas', () => {
    const spfxId = 'd15c0000-0000-4000-8000-000000000001';
    const unknownId = 'd15c0000-0000-4000-8000-0000000000ff';
    const spfxPart = main.plan.parts.find((p) => p.webPartId === spfxId);
    const unknownPart = main.plan.parts.find((p) => p.webPartId === unknownId);
    const warnedSpfx = main.plan.warnings.some((w) => w.includes('Site-scoped SPFx') && w.includes('not available'));
    const warnedUnknown = main.plan.warnings.some((w) => w.includes('Unknown part') && w.includes('not available'));
    const canvas = JSON.parse(main.postDto.CanvasContent1);
    return Boolean(spfxPart) && spfxPart.available === false && warnedSpfx && canvas.some((c) => c.webPartId === spfxId)
      && Boolean(unknownPart) && unknownPart.available === false && warnedUnknown && canvas.some((c) => c.webPartId === unknownId);
  });

  // ---- #9 link rewrite is opt-in; path-boundary matching never touches a look-alike web --

  await check('crosssite: rewriteLinks off (default) — the Text part\'s link into the source web is unchanged in the saved canvas', () => {
    const canvas = JSON.parse(main.postDto.CanvasContent1);
    const text = canvas.find((c) => c.controlType === 4);
    return String(text?.innerHTML || '').includes('href="/sites/pagesrc/SitePages/Policies.aspx"');
  });

  const rewritten = await page.evaluate(() => window.__cs.run({ options: { rewriteLinks: true } }));

  await check('crosssite: rewriteLinks:true — the Text part\'s link points at /sites/pagedst/… while the look-alike /sites/pagesrc-other/… link is left alone', () => {
    const canvas = JSON.parse(rewritten.postDto.CanvasContent1);
    const text = canvas.find((c) => c.controlType === 4);
    const html = String(text?.innerHTML || '');
    return html.includes('href="/sites/pagedst/SitePages/Policies.aspx"')
      && !html.includes('/sites/pagesrc/SitePages/Policies.aspx')
      && html.includes('href="/sites/pagesrc-other/SitePages/X.aspx"');
  });

  // ---- #10 a same-site-collection subweb copy is sameSite:true, sameWeb:false --------

  const team = await page.evaluate(() => window.__cs.run({ targetWeb: '/sites/pagesrc/team' }));

  await check('crosssite: a /sites/pagesrc -> /sites/pagesrc/team run is sameSite:true, sameWeb:false and transfers assets into the team web\'s own SiteAssets/SitePages/…', () => {
    const transferred = team.plan.assets.filter((a) => a.destName);
    const destFolder = '/sites/pagesrc/team/SiteAssets/SitePages/Quarterly-Update';
    return team.plan.sameSite === true && team.plan.sameWeb === false
      && transferred.length > 0
      && transferred.every((a) => a.destPath.toLowerCase().startsWith(destFolder.toLowerCase()))
      && (team.result.outcome === 'done' || team.result.outcome === 'done-with-warnings');
  });

  // ---- #11 an upload failure mid-run: outcome failed, page already created, discard -----
  // ---- recycles the page and the uploaded files but no asset folder -------------------------

  const failed = await page.evaluate(() => window.__cs.run({
    injectFailMatch: { match: 'thumb.png', code: 'write', status: 500 },
    thenDiscard: true,
  }));

  await check('crosssite: an upload failure mid-run (thumb.png) yields outcome "failed" with the page already created, and discardCopy recycles the page and the confirmed uploaded files', () => {
    const assetsStep = failed.stepLog.find((s) => s.name === 'assets');
    const fileAssets = (failed.result.journal.assets || []).filter((a) => a.kind === 'file');
    const recycled = failed.discard?.recycled || [];
    return failed.result.outcome === 'failed'
      && failed.result.journal.createdBy === 'this run'
      && Boolean(failed.result.journal.pageId)
      && Boolean(assetsStep) && assetsStep.status === 'failed'
      && fileAssets.length > 0 && fileAssets.every((a) => recycled.includes(a.path))
      && recycled.includes(failed.result.journal.currentPath);
  });

  await check('crosssite: discardCopy recycles no asset folder — the shared SiteAssets/SitePages parent is untouched and the leaf folder it created is left in place and reported', () => {
    const folderAssets = (failed.result.journal.assets || []).filter((a) => a.kind === 'folder');
    const sharedParent = folderAssets.find((a) => a.path.toLowerCase().endsWith('/siteassets/sitepages'));
    const leaf = folderAssets.find((a) => a.path.toLowerCase().includes('/siteassets/sitepages/') && a !== sharedParent);
    const recycled = failed.discard?.recycled || [];
    // An existing shared parent is never journaled (spike §11: the run probes
    // and skips it); if it were, it would be non-recyclable. Never recycled.
    return (!sharedParent || sharedParent.recyclable === false)
      && !recycled.some((r) => /\/siteassets\/sitepages$/i.test(r))
      && Boolean(leaf) && leaf.recyclable === true && !recycled.includes(leaf.path)
      && (failed.discard?.leftovers || []).some((l) => l.path === leaf.path);
  });

  await page.close();
}
