// Image web part analyzer unit tests (design/PAGE-COPY.md §5.2, id
// d1d91016-032f-456d-98a4-721247c305e8). House style: tests/pages-copy/
// analyzers/header.mjs, quick-links.mjs — imports src/workbench/
// page-copy-analyzers/image.js and page-copy.js's analyzerContext in the
// page context and exercises them against a fixture instance.
//
// Source: tests/pages-copy/fixtures/live-shapes-editor.json → image (the
// real editor-authored capture, §11 Q7) + source. linkUrl is empty on this
// capture, so only the asset ref shows up; the stock/CDN variant is a
// synthetic clone (no such capture exists — stock images are served from
// cdn.hubblecontent.osi.office.net, never the source web).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapes = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapes.image;
  const realSource = liveShapes.source;

  await check('analyzer image: refs on the real instance — one asset ref carrying the decoded path and the four customMetadata ids', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);
      const assetRefs = refs.filter((r) => r.class === 'asset');

      return assetRefs.length === 1 && refs.length === 1 // linkUrl is empty on this capture
        && assetRefs[0].path.join('.') === 'webPartData.serverProcessedContent.imageSources.imageSource'
        && assetRefs[0].asset.path === '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/logo.png'
        && assetRefs[0].asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9'
        && assetRefs[0].asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8'
        && assetRefs[0].asset.ids.listId === '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45'
        && assetRefs[0].asset.ids.uniqueId === 'efab1d13-d80d-4f4f-b22c-9a4ec4640c11';
    }, { control: realControl, source: realSource }));

  await check('analyzer image: refs on a stock/CDN variant (cdn.hubblecontent.osi.office.net, not the source web) — no asset ref', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image.js')).default;
      const clone = structuredClone(control);
      clone.webPartData.serverProcessedContent.imageSources.imageSource =
        'https://cdn.hubblecontent.osi.office.net/services/stockphotos/v1/w:900/id:abc123/image.jpg';
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(clone, ctx);
      return refs.filter((r) => r.class === 'asset').length === 0;
    }, { control: realControl, source: realSource }));

  await check('analyzer image: patch with no mapping changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image.js')).default;
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => null, mapLink: () => null };
      const before = JSON.stringify(control);
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { control: realControl, source: realSource }));

  await check('analyzer image: patch with a mapAsset mapping rewrites exactly imageSources.imageSource, customMetadata.imageSource\'s four ids and properties\' four ids — nothing else', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image.js')).default;
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/logo.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/logo.png',
      };
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.imageSources.imageSource = mapping.path;
      const customImageSource = expected.webPartData.serverProcessedContent.customMetadata.imageSource;
      customImageSource.siteid = mapping.ids.siteId;
      customImageSource.webid = mapping.ids.webId;
      customImageSource.listid = mapping.ids.listId;
      customImageSource.uniqueid = mapping.ids.uniqueId;
      Object.assign(expected.webPartData.properties, mapping.ids);

      return n === 9 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: realControl, source: realSource }));
}
