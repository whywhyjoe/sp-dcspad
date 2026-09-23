// Quick links analyzer unit tests (design/PAGE-COPY.md §5.1-5.2, id
// c70391ea-0b10-4ee9-b2b4-006d3fcad0cd). House style: tests/pages-copy/pure.mjs
// — imports src/workbench/page-copy-analyzers/quick-links.js and
// src/workbench/page-copy.js's analyzerContext in the page context and
// exercises them against fixture instances. Two sources:
//   - the real editor-authored capture (tests/pages-copy/fixtures/
//     live-shapes.json → quickLinks[0], both tiles pointing at the stock
//     go.microsoft.com fwlinks — no source-web link in the capture itself,
//     so a clone with one tile repointed into the source web exercises the
//     'link' class and the rewrite-links path)
//   - the mock fixture's custom-thumbnail Quick links control
//     (src/workbench/mock-pagecopy.js, search 'c70391ea') reproduced here
//     inline, its ids pulled from the module's own PAGECOPY_IDS so they
//     can never drift from the mock.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapes = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapes.quickLinks[0];
  const realSource = liveShapes.source;

  await check('analyzer quick links: refs on the real fixture — baseUrl is config, the stock go.microsoft.com links are ignored, and the part-level ids are config', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);

      const baseUrlRef = refs.find((r) => r.path.join('.') === 'webPartData.serverProcessedContent.links.baseUrl');
      const linkRefs = refs.filter((r) => r.class === 'link');
      const configIdRefs = refs.filter((r) => r.class === 'config' && r.path[1] === 'properties');
      const assetRefs = refs.filter((r) => r.class === 'asset');

      return Boolean(baseUrlRef) && baseUrlRef.class === 'config' && baseUrlRef.value === '/sites/NewNerve'
        && linkRefs.length === 0 // both tiles point at go.microsoft.com — external, not refs
        && configIdRefs.length === 2
        && configIdRefs.some((r) => r.path[2] === 'siteId') && configIdRefs.some((r) => r.path[2] === 'webId')
        && assetRefs.length === 0; // imageSources is empty on this capture
    }, { control: realControl, source: realSource }));

  await check('analyzer quick links: refs — a tile link repointed into the source web classifies as a link ref, not config or asset', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const ctx = analyzerContext({ web: source });
      const clone = structuredClone(control);
      clone.webPartData.serverProcessedContent.links['items[0].sourceItem.url'] = '/sites/NewNerve/SitePages/Team.aspx';
      const refs = analyzer.refs(clone, ctx);
      const linkRefs = refs.filter((r) => r.class === 'link');
      return linkRefs.length === 1
        && linkRefs[0].path.join('.') === 'webPartData.serverProcessedContent.links.items[0].sourceItem.url'
        && linkRefs[0].value === '/sites/NewNerve/SitePages/Team.aspx';
    }, { control: realControl, source: realSource }));

  await check('analyzer quick links: refs on the mock custom-thumbnail variant — the tile image is an asset ref carrying the item\'s four ids, the tile link is a link ref, and the part-level ids are config', () =>
    page.evaluate(async () => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const ids = PAGECOPY_IDS;
      const control = {
        controlType: 3, id: 'd15e0000-0000-4000-8000-000000000004', webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
        webPartData: {
          id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', title: 'Quick links',
          properties: {
            siteId: ids.srcSite, webId: ids.srcWeb,
            items: [{ siteId: ids.srcSite, webId: ids.srcWeb, listId: ids.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000005', thumbnailType: 3 }],
          },
          serverProcessedContent: {
            htmlStrings: {}, searchablePlainTexts: {},
            imageSources: { 'items[0].image.url': '/sites/pagesrc/SiteAssets/icons/logo.png' },
            links: { 'items[0].sourceItem.url': '/sites/pagesrc/SitePages/Policies.aspx' },
          },
        },
      };
      const source = { webUrl: `${location.origin}/sites/pagesrc`, webServerRelativeUrl: '/sites/pagesrc', siteId: ids.srcSite, webId: ids.srcWeb };
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);

      const asset = refs.find((r) => r.class === 'asset');
      const link = refs.find((r) => r.class === 'link');
      const configIds = refs.filter((r) => r.class === 'config' && r.path[1] === 'properties');

      return Boolean(asset)
        && asset.path.join('.') === 'webPartData.serverProcessedContent.imageSources.items[0].image.url'
        && asset.asset.path === '/sites/pagesrc/SiteAssets/icons/logo.png'
        && asset.asset.ids.siteId === ids.srcSite && asset.asset.ids.webId === ids.srcWeb
        && asset.asset.ids.listId === ids.srcSiteAssets && asset.asset.ids.uniqueId === 'd15f0000-0000-4000-8000-000000000005'
        && Boolean(link) && link.value === '/sites/pagesrc/SitePages/Policies.aspx'
        && configIds.length === 2;
    }));

  await check('analyzer quick links: patch with no mapping and rewriteLinks false changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async () => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const ids = PAGECOPY_IDS;
      const control = {
        controlType: 3, id: 'd15e0000-0000-4000-8000-000000000004', webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
        webPartData: {
          id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', title: 'Quick links',
          properties: {
            siteId: ids.srcSite, webId: ids.srcWeb,
            items: [{ siteId: ids.srcSite, webId: ids.srcWeb, listId: ids.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000005', thumbnailType: 3 }],
          },
          serverProcessedContent: {
            htmlStrings: {}, searchablePlainTexts: {},
            imageSources: { 'items[0].image.url': '/sites/pagesrc/SiteAssets/icons/logo.png' },
            links: { 'items[0].sourceItem.url': '/sites/pagesrc/SitePages/Policies.aspx' },
          },
        },
      };
      const source = { webUrl: `${location.origin}/sites/pagesrc`, webServerRelativeUrl: '/sites/pagesrc', siteId: ids.srcSite, webId: ids.srcWeb };
      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: false,
        mapAsset: () => null,
        mapLink: () => null,
      };
      const before = JSON.stringify(control);
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }));

  await check('analyzer quick links: patch with a mapAsset mapping rewrites the imageSources value and the item\'s four ids and nothing else', () =>
    page.evaluate(async () => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const ids = PAGECOPY_IDS;
      const control = {
        controlType: 3, id: 'd15e0000-0000-4000-8000-000000000004', webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
        webPartData: {
          id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd', title: 'Quick links',
          properties: {
            siteId: ids.srcSite, webId: ids.srcWeb,
            items: [{ siteId: ids.srcSite, webId: ids.srcWeb, listId: ids.srcSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-000000000005', thumbnailType: 3 }],
          },
          serverProcessedContent: {
            htmlStrings: {}, searchablePlainTexts: {},
            imageSources: { 'items[0].image.url': '/sites/pagesrc/SiteAssets/icons/logo.png' },
            links: { 'items[0].sourceItem.url': '/sites/pagesrc/SitePages/Policies.aspx' },
          },
        },
      };
      const source = { webUrl: `${location.origin}/sites/pagesrc`, webServerRelativeUrl: '/sites/pagesrc', siteId: ids.srcSite, webId: ids.srcWeb };
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/Quarterly-Update-copy/logo.png',
        ids: { siteId: ids.dstSite, webId: ids.dstWeb, listId: ids.dstSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-0000000000f1' },
        url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/Quarterly-Update-copy/logo.png`,
      };
      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: false,
        mapAsset: () => mapping,
        mapLink: () => null,
      };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.imageSources['items[0].image.url'] = mapping.path;
      Object.assign(expected.webPartData.properties.items[0], mapping.ids);

      return n === 5 && JSON.stringify(copy) === JSON.stringify(expected);
    }));

  await check('analyzer quick links: patch with rewriteLinks true and a mapLink mapping rewrites a source-web tile link and baseUrl, leaving the external tile link and part-level ids untouched', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/quick-links.js')).default;
      const clone = structuredClone(control);
      // items[1] stays the stock go.microsoft.com fwlink (external); repoint
      // items[0] into the source web so there is a link ref to rewrite.
      clone.webPartData.serverProcessedContent.links['items[0].sourceItem.url'] = '/sites/NewNerve/SitePages/Team.aspx';
      const externalBefore = clone.webPartData.serverProcessedContent.links['items[1].sourceItem.url'];
      const idsBefore = JSON.stringify(clone.webPartData.properties.siteId) + JSON.stringify(clone.webPartData.properties.webId);

      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: true,
        mapAsset: () => null,
        mapLink: (value) => (value === '/sites/NewNerve/SitePages/Team.aspx' ? '/sites/pagedst/SitePages/Team.aspx' : null),
      };
      const copy = structuredClone(clone);
      const n = analyzer.patch(copy, ctx);
      const links = copy.webPartData.serverProcessedContent.links;
      const idsAfter = JSON.stringify(copy.webPartData.properties.siteId) + JSON.stringify(copy.webPartData.properties.webId);

      return n === 2
        && links['items[0].sourceItem.url'] === '/sites/pagedst/SitePages/Team.aspx'
        && links.baseUrl === '/sites/pagedst'
        && links['items[1].sourceItem.url'] === externalBefore
        && idsAfter === idsBefore; // part-level config ids are never patched
    }, { control: realControl, source: realSource }));
}
