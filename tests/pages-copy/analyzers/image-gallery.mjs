// Image gallery analyzer unit tests (design/PAGE-COPY.md §5.1-5.2, id
// af8be689-990e-492a-81f7-ba3e4cd3ed9c). House style: quick-links.mjs /
// header.mjs — imports src/workbench/page-copy-analyzers/image-gallery.js
// and page-copy.js's analyzerContext in the page context and exercises them
// against the real editor-authored capture, tests/pages-copy/fixtures/
// live-shapes-editor.json → imageGallery (two per-image-list tiles, both
// images copied into SiteAssets/SitePages/zz-pagecopy-shapes/ on the source
// page, imageSourceType 1).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapesEditor = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapesEditor.imageGallery;
  const realSource = liveShapesEditor.source;

  await check('analyzer image gallery: refs on the real instance — two asset refs, each carrying its source path and four ids, and no data ref (imageSourceType is the per-image list)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image-gallery.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);
      const assetRefs = refs.filter((r) => r.class === 'asset');
      const dataRefs = refs.filter((r) => r.class === 'data');

      const first = assetRefs.find((r) => r.path.join('.') === 'webPartData.serverProcessedContent.imageSources.images[0].url');
      const second = assetRefs.find((r) => r.path.join('.') === 'webPartData.serverProcessedContent.imageSources.images[1].url');

      return assetRefs.length === 2 && dataRefs.length === 0
        && Boolean(first) && Boolean(second)
        && first.asset.path === '/sites/NewNerve/SiteAssets/SitePages/zz-pagecopy-shapes/copy-(1)-gallery2-425624.png'
        && first.asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9'
        && first.asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8'
        && first.asset.ids.listId === '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45'
        && first.asset.ids.uniqueId === 'c9825d1a-d155-4f4c-91ab-af3dd3bac447'
        && second.asset.path === '/sites/NewNerve/SiteAssets/SitePages/zz-pagecopy-shapes/copy-(1)-gallery1-977682.png'
        && second.asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9'
        && second.asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8'
        && second.asset.ids.listId === '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45'
        && second.asset.ids.uniqueId === 'dcd30560-ccdd-45da-8aa0-6ec359d83dcc';
    }, { control: realControl, source: realSource }));

  await check('analyzer image gallery: refs — a library/folder source mode (imageSourceType other than the per-image list) adds a data ref', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image-gallery.js')).default;
      const ctx = analyzerContext({ web: source });
      const clone = structuredClone(control);
      clone.webPartData.properties.imageSourceType = 0;
      const refs = analyzer.refs(clone, ctx);
      const dataRefs = refs.filter((r) => r.class === 'data');
      return dataRefs.length === 1
        && dataRefs[0].path.join('.') === 'webPartData.properties.imageSourceType'
        && dataRefs[0].value === 0;
    }, { control: realControl, source: realSource }));

  await check('analyzer image gallery: patch with no mappings changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image-gallery.js')).default;
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => null, mapLink: () => null };
      const before = JSON.stringify(control);
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { control: realControl, source: realSource }));

  await check('analyzer image gallery: patch with mappings for both images rewrites each image\'s path and its four ids (imageSources, customMetadata\'s lowercase keys, and properties.images[].{siteId,webId,listId,id}) and nothing else', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image-gallery.js')).default;
      const mapping0 = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery2.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery2.png',
      };
      const mapping1 = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery1.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a2',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery1.png',
      };
      const byUniqueId = new Map([
        ['c9825d1a-d155-4f4c-91ab-af3dd3bac447', mapping0],
        ['dcd30560-ccdd-45da-8aa0-6ec359d83dcc', mapping1],
      ]);
      const ctx = {
        ...analyzerContext({ web: source }),
        mapAsset: (asset) => byUniqueId.get(asset?.ids?.uniqueId) || null,
        mapLink: () => null,
      };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.imageSources['images[0].url'] = mapping0.path;
      expected.webPartData.serverProcessedContent.imageSources['images[1].url'] = mapping1.path;
      const cm0 = expected.webPartData.serverProcessedContent.customMetadata['images[0].url'];
      cm0.siteid = mapping0.ids.siteId; cm0.webid = mapping0.ids.webId; cm0.listid = mapping0.ids.listId; cm0.uniqueid = mapping0.ids.uniqueId;
      const cm1 = expected.webPartData.serverProcessedContent.customMetadata['images[1].url'];
      cm1.siteid = mapping1.ids.siteId; cm1.webid = mapping1.ids.webId; cm1.listid = mapping1.ids.listId; cm1.uniqueid = mapping1.ids.uniqueId;
      const item0 = expected.webPartData.properties.images[0];
      item0.siteId = mapping0.ids.siteId; item0.webId = mapping0.ids.webId; item0.listId = mapping0.ids.listId; item0.id = mapping0.ids.uniqueId;
      const item1 = expected.webPartData.properties.images[1];
      item1.siteId = mapping1.ids.siteId; item1.webId = mapping1.ids.webId; item1.listId = mapping1.ids.listId; item1.id = mapping1.ids.uniqueId;

      return n === 18 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: realControl, source: realSource }));

  await check('analyzer image gallery: patch with a mapping for only one image changes only that image, leaving the other untouched', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/image-gallery.js')).default;
      const mapping0 = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery2.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/gallery2.png',
      };
      const ctx = {
        ...analyzerContext({ web: source }),
        mapAsset: (asset) => (asset?.ids?.uniqueId === 'c9825d1a-d155-4f4c-91ab-af3dd3bac447' ? mapping0 : null),
        mapLink: () => null,
      };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const untouchedSecond = JSON.stringify(control.webPartData.serverProcessedContent.imageSources['images[1].url'])
        === JSON.stringify(copy.webPartData.serverProcessedContent.imageSources['images[1].url']);
      const untouchedSecondCm = JSON.stringify(control.webPartData.serverProcessedContent.customMetadata['images[1].url'])
        === JSON.stringify(copy.webPartData.serverProcessedContent.customMetadata['images[1].url']);
      const untouchedSecondItem = JSON.stringify(control.webPartData.properties.images[1])
        === JSON.stringify(copy.webPartData.properties.images[1]);

      return n === 9
        && copy.webPartData.serverProcessedContent.imageSources['images[0].url'] === mapping0.path
        && copy.webPartData.properties.images[0].id === mapping0.ids.uniqueId
        && untouchedSecond && untouchedSecondCm && untouchedSecondItem;
    }, { control: realControl, source: realSource }));
}
