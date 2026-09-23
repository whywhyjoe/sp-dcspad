// Page header (title area) analyzer unit tests (design/PAGE-COPY.md §0
// "Banner metadata", §5.3, §5.4, §11, id cbe7b0a9-3504-44dd-a3a3-0e5cacd07788).
// House style: tests/pages-copy/analyzers/quick-links.mjs — imports
// src/workbench/page-copy-analyzers/header.js and page-copy.js's
// analyzerContext in the page context and exercises them against fixture
// instances. Two sources:
//   - banner-less editor-authored samples: tests/pages-copy/fixtures/
//     live-shapes.json → titleArea (both entries, imageSourceType 4, empty
//     imageSources — no banner set)
//   - a page with a custom banner: not in the repo fixture (live-shapes.json
//     only captured banner-less pages), so reproduced inline from the dev
//     tenant spike capture (design/PAGE-COPY.md §11 Q7,
//     layout:cbe7b0a9-3504-44dd-a3a3-0e5cacd07788 → the zz-pagecopy-q5.aspx
//     sample) — imageSourceType 2, customMetadata.imageSource + the mirrored
//     properties ids PnPjs's save() writes (PnP-pages 340-353). The dev
//     tenant's real ids are kept for authenticity; PAGECOPY_IDS
//     (src/workbench/mock-pagecopy.js) supplies the source/destination web
//     ids for the end-to-end check, matching the other analyzer suites'
//     house style.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapes = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes.json'), 'utf8'));

// The zz-pagecopy-q5.aspx sample, byte-for-byte from the dev tenant capture.
const BANNER_PART = {
  id: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
  instanceId: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
  title: 'Title area',
  description: 'Title Region Description',
  serverProcessedContent: {
    htmlStrings: {},
    searchablePlainTexts: {},
    imageSources: { imageSource: '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/banner.png' },
    links: {},
    customMetadata: {
      imageSource: {
        siteId: '93de26bf-15a3-4fac-8f3e-2b8007d813a9',
        webId: '773960b7-b16d-4f81-957a-02de394949b8',
        listId: '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45',
        uniqueId: '51d45a56-d7b3-4781-99f5-74a41dae0840',
      },
    },
  },
  dataVersion: '1.4',
  properties: {
    title: 'zz-pagecopy Q5',
    imageSourceType: 2,
    layoutType: 'FullWidthImage',
    textAlignment: 'Left',
    showTopicHeader: true,
    showPublishDate: true,
    topicHeader: 'Finance',
    enableGradientEffect: true,
    authors: [],
    authorByline: [],
    webId: '773960b7-b16d-4f81-957a-02de394949b8',
    siteId: '93de26bf-15a3-4fac-8f3e-2b8007d813a9',
    listId: '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45',
    uniqueId: '51d45a56-d7b3-4781-99f5-74a41dae0840',
  },
};

const BANNER_SOURCE = {
  webUrl: 'https://nervedotnet.sharepoint.com/sites/NewNerve',
  webServerRelativeUrl: '/sites/NewNerve',
  siteId: '93de26bf-15a3-4fac-8f3e-2b8007d813a9',
  webId: '773960b7-b16d-4f81-957a-02de394949b8',
};

export async function checks({ page, check }) {
  await check('analyzer header: refs on the banner part — one asset ref carrying the decoded path and the four customMetadata ids', () =>
    page.evaluate(async ({ part, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/header.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(part, ctx);
      const assetRefs = refs.filter((r) => r.class === 'asset');

      return assetRefs.length === 1 && refs.length === 1 // authorByline/authors are empty on this page
        && assetRefs[0].path.join('.') === 'serverProcessedContent.imageSources.imageSource'
        && assetRefs[0].asset.path === '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/banner.png'
        && assetRefs[0].asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9'
        && assetRefs[0].asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8'
        && assetRefs[0].asset.ids.listId === '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45'
        && assetRefs[0].asset.ids.uniqueId === '51d45a56-d7b3-4781-99f5-74a41dae0840';
    }, { part: BANNER_PART, source: BANNER_SOURCE }));

  await check('analyzer header: refs on banner-less editor-authored title-area parts (imageSourceType 4, empty imageSources) — no asset refs', () =>
    page.evaluate(async ({ parts, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/header.js')).default;
      const ctx = analyzerContext({ web: source });
      return parts.every((part) => analyzer.refs(part, ctx).filter((r) => r.class === 'asset').length === 0);
    }, { parts: liveShapes.titleArea, source: liveShapes.source }));

  await check('analyzer header: patch with no mapping changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ part, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/header.js')).default;
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => null, mapLink: () => null };
      const before = JSON.stringify(part);
      const copy = structuredClone(part);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { part: BANNER_PART, source: BANNER_SOURCE }));

  await check('analyzer header: patch with a mapAsset mapping rewrites exactly imageSources.imageSource, customMetadata.imageSource\'s four ids and properties\' four ids — nothing else', () =>
    page.evaluate(async ({ part, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/header.js')).default;
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/banner.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/banner.png',
      };
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(part);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(part);
      expected.serverProcessedContent.imageSources.imageSource = mapping.path;
      Object.assign(expected.serverProcessedContent.customMetadata.imageSource, mapping.ids);
      Object.assign(expected.properties, mapping.ids);

      return n === 9 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { part: BANNER_PART, source: BANNER_SOURCE }));

  await check('analyzer header: end to end — rewriteContent through createAnalyzers([header]) on a cross-web snapshot rewrites the header\'s banner from transferResults, and maps the DTO BannerImageUrl thumbnail separately', () =>
    page.evaluate(async () => {
      const m = await import('/src/workbench/page-copy.js');
      const { createAnalyzers } = await import('/src/workbench/page-copy-analyzers.js');
      const header = (await import('/src/workbench/page-copy-analyzers/header.js')).default;
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const ids = PAGECOPY_IDS;
      const analyzers = createAnalyzers([header]);

      const bannerUniqueId = 'd15f0000-0000-4000-8000-0000000000b1';
      const thumbUniqueId = 'd15f0000-0000-4000-8000-0000000000b2';
      const bannerSourcePath = '/sites/pagesrc/SiteAssets/zz-pagecopy-assets/banner.png';
      const thumbSourcePath = '/sites/pagesrc/SiteAssets/zz-pagecopy-assets/thumb.png';

      const headerPart = {
        id: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
        instanceId: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
        title: 'Title area',
        serverProcessedContent: {
          htmlStrings: {}, searchablePlainTexts: {},
          imageSources: { imageSource: bannerSourcePath },
          links: {},
          customMetadata: { imageSource: { siteId: ids.srcSite, webId: ids.srcWeb, listId: ids.srcSiteAssets, uniqueId: bannerUniqueId } },
        },
        dataVersion: '1.4',
        properties: {
          title: 'zz-pagecopy Q5', imageSourceType: 2, layoutType: 'FullWidthImage', textAlignment: 'Left',
          showTopicHeader: true, showPublishDate: true, topicHeader: 'Finance', authors: [], authorByline: [],
          siteId: ids.srcSite, webId: ids.srcWeb, listId: ids.srcSiteAssets, uniqueId: bannerUniqueId,
        },
      };

      const web = { webId: ids.srcWeb, siteId: ids.srcSite, webUrl: `${location.origin}/sites/pagesrc`, webServerRelativeUrl: '/sites/pagesrc' };
      const snapshot = m.snapshotFromReads({
        page: {
          Id: 42, Title: 'zz-pagecopy Q5', CanvasContent1: '[]', LayoutWebpartsContent: JSON.stringify([headerPart]),
          PageLayoutType: 'Article', PromotedState: 0, FileName: 'zz-pagecopy-q5.aspx',
          // A distinct file from the banner — §11: the thumbnail on a
          // custom-thumbnail page, mapped by mapBannerUrl(), never by this
          // analyzer.
          BannerImageUrl: thumbSourcePath,
        },
        item: { Id: 42, FileRef: '/sites/pagesrc/SitePages/zz-pagecopy-q5.aspx', FileLeafRef: 'zz-pagecopy-q5.aspx', _UIVersionString: '1.0', Modified: 'x', CommentsDisabled: false },
        fields: [], library: { id: ids.srcSitePages, baseTemplate: 119, hidden: false, rootPath: '/sites/pagesrc/SitePages' },
        web,
      });

      const target = {
        webUrl: `${location.origin}/sites/pagedst`, webServerRelativeUrl: '/sites/pagedst',
        webId: ids.dstWeb, siteId: ids.dstSite,
        library: { id: ids.dstSitePages, rootPath: '/sites/pagedst/SitePages' }, fields: [],
        takenFinal: new Set(), takenRoot: new Set(),
      };
      const plan = m.analyzeCopy({ snapshot, target, options: { runId: 'e2e001', fileName: 'zz-pagecopy-q5-copy.aspx' }, analyzers });

      const bannerMapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/banner.png',
        ids: { siteId: ids.dstSite, webId: ids.dstWeb, listId: ids.dstSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-0000000000c1' },
        url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/banner.png`,
      };
      const thumbMapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/thumb.png',
        ids: { siteId: ids.dstSite, webId: ids.dstWeb, listId: ids.dstSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-0000000000c2' },
        url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-q5-copy/thumb.png`,
      };
      const transferResults = new Map([
        [m.assetKey({ uniqueId: bannerUniqueId }), bannerMapping],
        [m.assetKey({ path: thumbSourcePath }), thumbMapping],
      ]);

      const saved = m.rewriteContent(snapshot, plan, transferResults, analyzers);
      const newLayout = JSON.parse(saved.LayoutWebpartsContent);
      const newHeader = newLayout[0];

      return plan.blockers.length === 0
        && newHeader.serverProcessedContent.imageSources.imageSource === bannerMapping.path
        && newHeader.serverProcessedContent.customMetadata.imageSource.siteId === bannerMapping.ids.siteId
        && newHeader.serverProcessedContent.customMetadata.imageSource.webId === bannerMapping.ids.webId
        && newHeader.serverProcessedContent.customMetadata.imageSource.listId === bannerMapping.ids.listId
        && newHeader.serverProcessedContent.customMetadata.imageSource.uniqueId === bannerMapping.ids.uniqueId
        && newHeader.properties.siteId === bannerMapping.ids.siteId
        && newHeader.properties.webId === bannerMapping.ids.webId
        && newHeader.properties.listId === bannerMapping.ids.listId
        && newHeader.properties.uniqueId === bannerMapping.ids.uniqueId
        // The DTO thumbnail is a different file, mapped separately from the
        // header's own banner, and to its own destination path.
        && saved.BannerImageUrl === thumbMapping.path
        && saved.BannerImageUrl !== bannerMapping.path;
    }));
}
