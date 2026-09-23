// SP Workbench — page copy: pure units (design/PAGE-COPY.md §8 "planner/
// rewriter units"). No page state, no mock webs — every check imports
// src/workbench/page-copy.js in the page context and exercises it against
// inline fixtures. House style: tests/workbench-schema.mjs.

export async function run({ browser, check, WB_URL }) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  await page.goto(WB_URL);
  await page.waitForSelector('.wb-home-cards');

  await check('pure: defaultFileName keeps the source name, then -copy, then -copy-2 when taken (same-folder duplicates are always suffixed)', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const free = m.defaultFileName('Report.aspx', new Set());
      const takenOnce = m.defaultFileName('Report.aspx', new Set(['report.aspx']));
      const takenTwice = m.defaultFileName('Report.aspx', new Set(['report.aspx', 'report-copy.aspx']));
      return free === 'Report.aspx' && takenOnce === 'Report-copy.aspx' && takenTwice === 'Report-copy-2.aspx';
    }));

  await check('pure: fileNameProblem and folderProblem refuse bad input and accept good input', () =>
    page.evaluate(async () => {
      const { fileNameProblem, folderProblem } = await import('/src/workbench/page-copy.js');
      const noExt = fileNameProblem('Report');
      const badChar = fileNameProblem('Rep:ort.aspx');
      const leadingDot = fileNameProblem('.foo.aspx');
      const tooLong = fileNameProblem(`${'a'.repeat(130)}.aspx`);
      const goodName = fileNameProblem('Report-Copy.aspx');
      const badFolder = folderProblem('Archive/../etc');
      const goodFolder = folderProblem('Archive/2026');
      const rootFolder = folderProblem('');
      return noExt !== '' && badChar !== '' && leadingDot !== '' && tooLong !== ''
        && goodName === '' && badFolder !== '' && goodFolder === '' && rootFolder === '';
    }));

  await check('pure: stagingStem builds a collision-safe root-level temporary stem', () =>
    page.evaluate(async () => {
      const { stagingStem } = await import('/src/workbench/page-copy.js');
      return stagingStem('Report', 'abc123') === 'Report~copy-abc123';
    }));

  await check('pure: sourceEligibility refuses a legacy-HTML page cross-web but allows it same-web with legacyHtml true', () =>
    page.evaluate(async () => {
      const { sourceEligibility } = await import('/src/workbench/page-copy.js');
      const base = { library: { baseTemplate: 119 }, canvasFormat: 'html' };
      const crossWeb = sourceEligibility(base, { sameWeb: false });
      const sameWeb = sourceEligibility(base, { sameWeb: true });
      return crossWeb.ok === false && crossWeb.legacyHtml === true
        && sameWeb.ok === true && sameWeb.legacyHtml === true;
    }));

  await check('pure: sourceEligibility refuses a RepostPage layout, a non-119 library and unparsable canvas JSON', () =>
    page.evaluate(async () => {
      const { sourceEligibility } = await import('/src/workbench/page-copy.js');
      const repost = sourceEligibility({ library: { baseTemplate: 119 }, canvasFormat: 'json', dto: { PageLayoutType: 'RepostPage' } });
      const lib850 = sourceEligibility({ library: { baseTemplate: 850 }, canvasFormat: 'json', dto: { PageLayoutType: 'Article' } });
      const invalidJson = sourceEligibility({ library: { baseTemplate: 119 }, canvasFormat: 'invalid', dto: { PageLayoutType: 'Article' } });
      return repost.ok === false && lib850.ok === false && invalidJson.ok === false;
    }));

  await check('pure: sourceEligibility accepts Article, Home and SingleWebPartAppPage layouts', () =>
    page.evaluate(async () => {
      const { sourceEligibility } = await import('/src/workbench/page-copy.js');
      const mk = (layout) => sourceEligibility({ library: { baseTemplate: 119 }, canvasFormat: 'json', dto: { PageLayoutType: layout } });
      const article = mk('Article');
      const home = mk('Home');
      const single = mk('SingleWebPartAppPage');
      return article.ok === true && !article.legacyHtml && home.ok === true && single.ok === true;
    }));

  await check('pure: targetEligibility requires an active Site Pages feature and a resolved library', () =>
    page.evaluate(async () => {
      const { targetEligibility } = await import('/src/workbench/page-copy.js');
      const ok = targetEligibility({ featureActive: true, library: { id: 'L' } });
      const noFeature = targetEligibility({ featureActive: false, library: { id: 'L' } });
      const noLibrary = targetEligibility({ featureActive: true, library: null });
      return ok.ok === true && noFeature.ok === false && noFeature.reason.includes('feature')
        && noLibrary.ok === false && noLibrary.reason.includes('Site Pages library');
    }));

  await check('pure: computeCarrySet never carries Title/Description/CommentsDisabled/PromotedState/BannerImageUrl, skips hidden/read-only/type-mismatched destination fields', () =>
    page.evaluate(async () => {
      const { computeCarrySet } = await import('/src/workbench/page-copy.js');
      const sourceFields = [
        { InternalName: 'Title', TypeAsString: 'Text' },
        { InternalName: 'Description', TypeAsString: 'Note' },
        { InternalName: 'CommentsDisabled', TypeAsString: 'Boolean' },
        { InternalName: 'PromotedState', TypeAsString: 'Number' },
        { InternalName: 'BannerImageUrl', TypeAsString: 'URL' },
        { InternalName: 'HiddenField', Title: 'Hidden Field', TypeAsString: 'Text' },
        { InternalName: 'ROField', Title: 'RO Field', TypeAsString: 'Text' },
        { InternalName: 'MismatchField', Title: 'Mismatch Field', TypeAsString: 'Text' },
      ];
      const targetFields = [
        { InternalName: 'Title', TypeAsString: 'Text' },
        { InternalName: 'Description', TypeAsString: 'Note' },
        { InternalName: 'CommentsDisabled', TypeAsString: 'Boolean' },
        { InternalName: 'PromotedState', TypeAsString: 'Number' },
        { InternalName: 'BannerImageUrl', TypeAsString: 'URL' },
        { InternalName: 'HiddenField', TypeAsString: 'Text', Hidden: true },
        { InternalName: 'ROField', TypeAsString: 'Text', ReadOnlyField: true },
        { InternalName: 'MismatchField', TypeAsString: 'Number' },
      ];
      const item = {
        Title: 'T', Description: 'D', CommentsDisabled: true, PromotedState: 2,
        BannerImageUrl: { Url: '/x', Description: '' },
        HiddenField: 'h', ROField: 'r', MismatchField: 'm',
      };
      const r = computeCarrySet({ sourceFields, targetFields, item, itemAsText: {} });
      const names = r.carried.map((c) => c.internalName);
      const noOwned = !names.includes('Title') && !names.includes('Description')
        && !names.includes('CommentsDisabled') && !names.includes('PromotedState') && !names.includes('BannerImageUrl');
      const hiddenSkip = r.skipped.find((s) => s.internalName === 'HiddenField');
      const roSkip = r.skipped.find((s) => s.internalName === 'ROField');
      const mismatchSkip = r.skipped.find((s) => s.internalName === 'MismatchField');
      return noOwned
        && Boolean(hiddenSkip) && hiddenSkip.reason.includes('read-only or hidden')
        && Boolean(roSkip) && roSkip.reason.includes('read-only or hidden')
        && Boolean(mismatchSkip) && mismatchSkip.reason.includes('Text → Number');
    }));

  await check('pure: computeCarrySet skips a Choice value missing on the destination, but a FillInChoice destination accepts it', () =>
    page.evaluate(async () => {
      const { computeCarrySet } = await import('/src/workbench/page-copy.js');
      const sourceFields = [{ InternalName: 'PageCategory', Title: 'Category', TypeAsString: 'Choice', Choices: ['Finance', 'IT'] }];
      const item = { PageCategory: 'IT' };
      const noFillIn = computeCarrySet({
        sourceFields, item, itemAsText: {},
        targetFields: [{ InternalName: 'PageCategory', TypeAsString: 'Choice', Choices: ['Finance'] }],
      });
      const withFillIn = computeCarrySet({
        sourceFields, item, itemAsText: {},
        targetFields: [{ InternalName: 'PageCategory', TypeAsString: 'Choice', Choices: ['Finance'], FillInChoice: true }],
      });
      const skip = noFillIn.skipped.find((s) => s.internalName === 'PageCategory');
      const carried = withFillIn.carried.find((c) => c.internalName === 'PageCategory');
      return Boolean(skip) && skip.reason.includes('does not exist on the destination')
        && !noFillIn.carried.some((c) => c.internalName === 'PageCategory')
        && Boolean(carried) && carried.value === 'IT';
    }));

  await check('pure: computeCarrySet lists User/Lookup fields as skipped only when the source has a value', () =>
    page.evaluate(async () => {
      const { computeCarrySet } = await import('/src/workbench/page-copy.js');
      const sourceFields = [
        { InternalName: 'Owner', TypeAsString: 'User' },
        { InternalName: 'RelatedTask', TypeAsString: 'Lookup' },
      ];
      const targetFields = [
        { InternalName: 'Owner', TypeAsString: 'User' },
        { InternalName: 'RelatedTask', TypeAsString: 'Lookup' },
      ];
      const withValue = computeCarrySet({ sourceFields, targetFields, itemAsText: {}, item: { Owner: { Id: 3 }, RelatedTask: { LookupId: 5 } } });
      const withoutValue = computeCarrySet({ sourceFields, targetFields, itemAsText: {}, item: {} });
      const ownerSkip = withValue.skipped.find((s) => s.internalName === 'Owner');
      const taskSkip = withValue.skipped.find((s) => s.internalName === 'RelatedTask');
      return Boolean(ownerSkip) && ownerSkip.reason.includes('not copied')
        && Boolean(taskSkip)
        && withoutValue.skipped.length === 0 && withoutValue.carried.length === 0;
    }));

  await check('pure: computeCarrySet carries DateTime using the FieldValuesAsText display text, never the raw ISO string', () =>
    page.evaluate(async () => {
      const { computeCarrySet } = await import('/src/workbench/page-copy.js');
      const sourceFields = [{ InternalName: 'ReviewDate', TypeAsString: 'DateTime' }];
      const targetFields = [{ InternalName: 'ReviewDate', TypeAsString: 'DateTime' }];
      const item = { ReviewDate: '2026-10-01T07:00:00Z' };
      const itemAsText = { ReviewDate: '10/1/2026 12:00 AM' };
      const r = computeCarrySet({ sourceFields, targetFields, item, itemAsText });
      const entry = r.carried.find((c) => c.internalName === 'ReviewDate');
      return Boolean(entry) && entry.value === '10/1/2026 12:00 AM' && !entry.value.includes('2026-10-01T07');
    }));

  await check('pure: computeCarrySet reports requiredGaps, then satisfies them from requiredValues', () =>
    page.evaluate(async () => {
      const { computeCarrySet } = await import('/src/workbench/page-copy.js');
      const targetFields = [{ InternalName: 'Audience', Title: 'Audience', TypeAsString: 'Text', Required: true }];
      const gap = computeCarrySet({ sourceFields: [], targetFields, item: {}, itemAsText: {} });
      const filled = computeCarrySet({ sourceFields: [], targetFields, item: {}, itemAsText: {}, requiredValues: { Audience: 'All' } });
      const gapEntry = gap.requiredGaps.find((g) => g.internalName === 'Audience');
      const carried = filled.carried.find((c) => c.internalName === 'Audience');
      return Boolean(gapEntry) && gapEntry.supportable === true
        && filled.requiredGaps.length === 0
        && Boolean(carried) && carried.value === 'All' && carried.supplied === true;
    }));

  await check('pure: underPath matches a path and its children but never a similarly-prefixed sibling, case-insensitively', () =>
    page.evaluate(async () => {
      const { underPath } = await import('/src/workbench/page-copy.js');
      return underPath('/sites/source/x', '/sites/source') === true
        && underPath('/sites/source-other/x', '/sites/source') === false
        && underPath('/sites/Source', '/sites/source/') === true
        && underPath('/sites/source?x=1', '/sites/source') === true
        && underPath('/sites/source#frag', '/sites/source') === true;
    }));

  await check('pure: rewriteLink handles a relative link, an absolute same-origin link, an other-origin link (null) and an outside-the-web link (null)', () =>
    page.evaluate(async () => {
      const { rewriteLink } = await import('/src/workbench/page-copy.js');
      const opts = { fromPath: '/sites/src', toPath: '/sites/dst', origin: 'https://t.sharepoint.com' };
      const relative = rewriteLink('/sites/src/SitePages/A.aspx', opts);
      const absolute = rewriteLink('https://t.sharepoint.com/sites/src/Shared%20Documents/a%20b.pdf?web=1', opts);
      const otherOrigin = rewriteLink('https://other.sharepoint.com/sites/src/x.aspx', opts);
      const outsideWeb = rewriteLink('/sites/src-other/A.aspx', opts);
      return relative === '/sites/dst/SitePages/A.aspx'
        && absolute === 'https://t.sharepoint.com/sites/dst/Shared%20Documents/a%20b.pdf?web=1'
        && otherOrigin === null && outsideWeb === null;
    }));

  await check('pure: inventoryReferences finds a rich-text link, a getpreview URL and a listId GUID, and never mutates the snapshot', () =>
    page.evaluate(async () => {
      const { inventoryReferences } = await import('/src/workbench/page-copy.js');
      const web = {
        webUrl: 'https://t.sharepoint.com/sites/pagesrc', webServerRelativeUrl: '/sites/pagesrc',
        siteId: 'aaaaaaaa-0000-4000-8000-000000000001', webId: 'bbbbbbbb-0000-4000-8000-000000000001',
      };
      const canvas = [
        { controlType: 4, id: 't1', innerHTML: '<p><a href="/sites/pagesrc/SitePages/Other.aspx">link</a></p>' },
        {
          controlType: 3, id: 'wp1', webPartId: 'dddddddd-0000-4000-8000-000000000001',
          webPartData: { properties: { listId: 'eeeeeeee-0000-4000-8000-000000000001' } },
        },
      ];
      const dto = {
        BannerImageUrl: 'https://t.sharepoint.com/sites/pagesrc/_layouts/15/getpreview.ashx'
          + '?guidSite=aaaaaaaa00004000800000000001&guidWeb=bbbbbbbb00004000800000000001&guidFile=ffffffff00004000800000000001',
      };
      const snapshot = { web, library: { id: 'cccccccc-0000-4000-8000-000000000001' }, canvas, layout: [], dto, customThumbnail: false };
      const before = JSON.stringify(snapshot);
      const refs = inventoryReferences(snapshot);
      const untouched = JSON.stringify(snapshot) === before;
      const hasLink = refs.some((r) => r.kind === 'path' && String(r.value).includes('Other.aspx'));
      const hasPreview = refs.some((r) => r.kind === 'preview');
      const hasGuid = refs.some((r) => r.kind === 'guid' && r.path.includes('listId'));
      return untouched && hasLink && hasPreview && hasGuid;
    }));

  await check('pure: analyzeCopy builds a same-web plan with the api engine, path A, and the staging/final paths in the library', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const srcFields = [{ InternalName: 'Title', TypeAsString: 'Text' }];
      const canvas = JSON.stringify([
        { controlType: 4, id: 't1', innerHTML: '<p>hi</p>' },
        { controlType: 0, pageSettingsSlice: { isDefaultThumbnail: true } },
      ]);
      const snap = m.snapshotFromReads({
        page: { Id: 7, Title: 'Q', CanvasContent1: canvas, LayoutWebpartsContent: '[]', PageLayoutType: 'Article', PromotedState: 2, FileName: 'Q.aspx', BannerImageUrl: '' },
        item: { Id: 7, FileRef: '/sites/src/SitePages/Q.aspx', FileLeafRef: 'Q.aspx', _UIVersionString: '2.0', Modified: 'x', CommentsDisabled: true },
        fields: srcFields, library: { id: 'L', baseTemplate: 119, rootPath: '/sites/src/SitePages' },
        web: { webId: 'aaaaaaaa-0000-4000-8000-000000000001', siteId: 'bbbbbbbb-0000-4000-8000-000000000001', webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' },
      });
      const target = {
        webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src',
        webId: 'aaaaaaaa-0000-4000-8000-000000000001', siteId: 'bbbbbbbb-0000-4000-8000-000000000001',
        library: { id: 'L', rootPath: '/sites/src/SitePages' }, fields: srcFields,
        takenFinal: new Set(), takenRoot: new Set(),
      };
      const noPromote = m.analyzeCopy({ snapshot: snap, target, options: { runId: 'abc123' } });
      const promote = m.analyzeCopy({ snapshot: snap, target, options: { runId: 'abc123', promoteAsNews: true } });
      return noPromote.engine === 'api' && noPromote.createPath === 'A' && noPromote.sameWeb === true
        && noPromote.blockers.length === 0
        && noPromote.target.finalPath === '/sites/src/SitePages/Q.aspx'
        && noPromote.target.stagingPath === '/sites/src/SitePages/Q~copy-abc123.aspx'
        && noPromote.promoteAsNews === false
        && promote.promoteAsNews === true;
    }));

  await check('pure: analyzeCopy treats a name already present in the final folder as a blocker', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const srcFields = [{ InternalName: 'Title', TypeAsString: 'Text' }];
      const canvas = JSON.stringify([{ controlType: 4, id: 't1', innerHTML: '<p>hi</p>' }]);
      const snap = m.snapshotFromReads({
        page: { Id: 7, Title: 'Q', CanvasContent1: canvas, LayoutWebpartsContent: '[]', PageLayoutType: 'Article', PromotedState: 0, FileName: 'Q.aspx', BannerImageUrl: '' },
        item: { Id: 7, FileRef: '/sites/src/SitePages/Q.aspx', FileLeafRef: 'Q.aspx', _UIVersionString: '1.0', Modified: 'x', CommentsDisabled: false },
        fields: srcFields, library: { id: 'L', baseTemplate: 119, rootPath: '/sites/src/SitePages' },
        web: { webId: 'aaaaaaaa-0000-4000-8000-000000000001', siteId: 'bbbbbbbb-0000-4000-8000-000000000001', webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' },
      });
      const target = {
        webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src',
        webId: 'aaaaaaaa-0000-4000-8000-000000000001', siteId: 'bbbbbbbb-0000-4000-8000-000000000001',
        library: { id: 'L', rootPath: '/sites/src/SitePages' }, fields: srcFields,
        takenFinal: new Set(['q.aspx']), takenRoot: new Set(),
      };
      const plan = m.analyzeCopy({ snapshot: snap, target, options: { runId: 'abc123', fileName: 'Q.aspx' } });
      return plan.blockers.some((b) => b.includes('already exists'));
    }));

  await check('pure: rewriteContent returns the raw CanvasContent1/LayoutWebpartsContent strings byte-identical for a verbatim same-web copy, and uses BannerThumbnailUrl as BannerImageUrl for a custom thumbnail', () =>
    page.evaluate(async () => {
      const { rewriteContent } = await import('/src/workbench/page-copy.js');
      const canvasRaw = JSON.stringify([{ controlType: 4, id: 't1', innerHTML: '<p>hi</p>' }]);
      const layoutRaw = '[]';
      const baseSnapshot = {
        dto: { Title: 'Q', BannerImageUrl: '', BannerThumbnailUrl: '', Description: '', TopicHeader: '', AuthorByline: [] },
        canvasRaw, layoutRaw, canvas: JSON.parse(canvasRaw), layout: [], customThumbnail: false,
        web: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' },
      };
      const basePlan = { sameWeb: true, dropped: [], title: 'Q', target: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' }, rewriteLinks: false };
      const saved = rewriteContent(baseSnapshot, basePlan, new Map());
      const verbatim = saved.CanvasContent1 === canvasRaw && saved.LayoutWebpartsContent === layoutRaw;

      const thumbSnapshot = {
        ...baseSnapshot,
        dto: { ...baseSnapshot.dto, BannerImageUrl: '/sites/src/SiteAssets/banner.jpg', BannerThumbnailUrl: '/sites/src/SiteAssets/thumb.png' },
        customThumbnail: true,
      };
      const savedThumb = rewriteContent(thumbSnapshot, basePlan, new Map());
      return verbatim && savedThumb.BannerImageUrl === '/sites/src/SiteAssets/thumb.png';
    }));

  await check('pure: rewriteContent with a changed title patches only the header part\'s properties.title, leaving the canvas and other layout parts untouched', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const headerPart = { id: m.PAGE_HEADER_ID, instanceId: 'h1', title: 'Header', properties: { title: 'Old Title', imageSourceType: 4 } };
      const otherPart = { id: 'other-part-id', instanceId: 'o1', properties: { foo: 'bar' } };
      const canvasRaw = JSON.stringify([{ controlType: 4, id: 't1', innerHTML: '<p>hi</p>' }]);
      const layoutRaw = JSON.stringify([headerPart, otherPart]);
      const snapshot = {
        dto: { Title: 'Old Title', BannerImageUrl: '', BannerThumbnailUrl: '', Description: '', TopicHeader: '', AuthorByline: [] },
        canvasRaw, layoutRaw, canvas: JSON.parse(canvasRaw), layout: JSON.parse(layoutRaw), customThumbnail: false,
        web: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' },
      };
      const plan = { sameWeb: true, dropped: [], title: 'New Title', target: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' }, rewriteLinks: false };
      const saved = m.rewriteContent(snapshot, plan, new Map());
      const newLayout = JSON.parse(saved.LayoutWebpartsContent);
      return saved.CanvasContent1 === canvasRaw
        && newLayout[0].properties.title === 'New Title'
        && newLayout[1].properties.foo === 'bar';
    }));

  await check('pure: dropping an instance id removes its control from the copy, and dependentConsumers lists the consumer of a dropped dynamic provider', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const providerInstanceId = 'aaaaaaaa-0000-4000-8000-000000000001';
      const consumerInstanceId = 'bbbbbbbb-0000-4000-8000-000000000002';
      const provider = {
        controlType: 3, id: providerInstanceId, webPartId: 'cccccccc-0000-4000-8000-000000000003',
        webPartData: { instanceId: providerInstanceId, title: 'List' },
      };
      const consumer = {
        controlType: 3, id: consumerInstanceId, webPartId: 'dddddddd-0000-4000-8000-000000000004',
        webPartData: { instanceId: consumerInstanceId, title: 'Chart', dynamicDataPaths: { series: `WebPart.${providerInstanceId}` } },
      };
      const snapshot = {
        dto: { Title: 'Q', BannerImageUrl: '', BannerThumbnailUrl: '', Description: '', TopicHeader: '', AuthorByline: [] },
        canvasRaw: JSON.stringify([provider, consumer]), layoutRaw: '[]',
        canvas: [provider, consumer], layout: [], customThumbnail: false,
        web: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src', siteId: 's', webId: 'w' },
        library: { id: 'L' },
      };
      const analysis = m.analyzeParts(snapshot);
      const consumers = m.dependentConsumers(analysis, [providerInstanceId]);
      const plan = { sameWeb: true, dropped: [providerInstanceId], title: 'Q', target: { webUrl: 'https://t.sharepoint.com/sites/src', webServerRelativeUrl: '/sites/src' }, rewriteLinks: false };
      const saved = m.rewriteContent(snapshot, plan, new Map());
      const remaining = JSON.parse(saved.CanvasContent1);
      return consumers.length === 1 && consumers[0].instanceId === consumerInstanceId && consumers[0].label === 'Chart'
        && remaining.length === 1 && remaining[0].id === consumerInstanceId;
    }));

  await check('pure: compareReadBack reports FileName/Title drift and a non-zero PromotedState when not promoting, and is empty for a faithful read-back', () =>
    page.evaluate(async () => {
      const { compareReadBack } = await import('/src/workbench/page-copy.js');
      const plan = { title: 'New Title', target: { fileName: 'Q-copy.aspx' }, pageLayoutType: 'Article', promoteAsNews: false };
      const saved = { Description: 'D', TopicHeader: 'T', CanvasContent1: JSON.stringify([{ controlType: 4, id: 't1' }]), BannerImageUrl: '/sites/x/banner.jpg' };
      const driftDto = {
        Title: 'Old Title', FileName: 'Different.aspx', PageLayoutType: 'Article',
        Description: 'D', TopicHeader: 'T', CanvasContent1: saved.CanvasContent1, BannerImageUrl: '/sites/x/banner.jpg', PromotedState: 2,
      };
      const faithfulDto = {
        Title: 'New Title', FileName: 'Q-copy.aspx', PageLayoutType: 'Article',
        Description: 'D', TopicHeader: 'T', CanvasContent1: saved.CanvasContent1, BannerImageUrl: '/sites/x/banner.jpg', PromotedState: 0,
      };
      const drift = compareReadBack(plan, saved, driftDto);
      const faithful = compareReadBack(plan, saved, faithfulDto);
      return drift.some((d) => d.field === 'Title') && drift.some((d) => d.field === 'FileName')
        && drift.some((d) => d.field === 'PromotedState' && d.actual === 2)
        && faithful.length === 0;
    }));

  await page.close();
}
