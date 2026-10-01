// Call to action analyzer unit tests (design/PAGE-COPY.md §5.1-5.2, id
// df8e44e7-edd5-46d5-90da-aca1539313b8). House style: tests/pages-copy/
// pure.mjs — imports src/workbench/page-copy-analyzers/call-to-action.js and
// src/workbench/page-copy.js's analyzerContext in the page context and
// exercises them against the real editor-authored capture
// (tests/pages-copy/fixtures/live-shapes-editor.json → callToAction):
// background image chosen with "Don't copy" (asset, ids on
// properties.image.itemInfo), button labeled "Join" linked into the source
// web (link, patched only via ctx.mapLink).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapesEditor = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapesEditor.callToAction;
  const realSource = liveShapesEditor.source;

  await check('analyzer call to action: refs on the real fixture — the background image is an asset ref carrying the properties.image.itemInfo ids, the button link is a link ref, and nothing else is a reference', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/call-to-action.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);

      const asset = refs.find((r) => r.class === 'asset');
      const link = refs.find((r) => r.class === 'link');
      const itemInfo = control.webPartData.properties.image.itemInfo;

      return refs.length === 2
        && Boolean(asset)
        && asset.path.join('.') === 'webPartData.serverProcessedContent.imageSources.image.url'
        && asset.value === control.webPartData.serverProcessedContent.imageSources['image.url']
        && asset.asset.path === '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/banner.png'
        && asset.asset.ids.siteId === itemInfo.siteId && asset.asset.ids.webId === itemInfo.webId
        && asset.asset.ids.listId === itemInfo.listId && asset.asset.ids.uniqueId === itemInfo.uniqueId
        && Boolean(link)
        && link.path.join('.') === 'webPartData.serverProcessedContent.links.button.linkUrl'
        && link.value === control.webPartData.serverProcessedContent.links['button.linkUrl'];
    }, { control: realControl, source: realSource }));

  await check('analyzer call to action: patch with no mappings changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/call-to-action.js')).default;
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
    }, { control: realControl, source: realSource }));

  await check('analyzer call to action: patch with a mapAsset mapping rewrites only the background image path and its four ids', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/call-to-action.js')).default;
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/Test-copy/banner.png',
        ids: {
          siteId: '11111111-1111-4111-8111-111111111111',
          webId: '22222222-2222-4222-8222-222222222222',
          listId: '33333333-3333-4333-8333-333333333333',
          uniqueId: '44444444-4444-4444-8444-444444444444',
        },
        url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/Test-copy/banner.png`,
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
      expected.webPartData.serverProcessedContent.imageSources['image.url'] = mapping.path;
      Object.assign(expected.webPartData.properties.image.itemInfo, mapping.ids);

      return n === 5 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: realControl, source: realSource }));

  await check('analyzer call to action: patch with rewriteLinks true and a mapLink mapping rewrites only the button link, keeping its absolute form', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/call-to-action.js')).default;
      const clone = structuredClone(control);
      const absoluteLink = `${source.webUrl}/SitePages/Test.aspx`;
      clone.webPartData.serverProcessedContent.links['button.linkUrl'] = absoluteLink;
      const mappedLink = `${source.webUrl.replace('/sites/NewNerve', '/sites/pagedst')}/SitePages/Test.aspx`;

      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: true,
        mapAsset: () => null,
        mapLink: (value) => (value === absoluteLink ? mappedLink : null),
      };
      const copy = structuredClone(clone);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(clone);
      expected.webPartData.serverProcessedContent.links['button.linkUrl'] = mappedLink;

      return n === 1
        && JSON.stringify(copy) === JSON.stringify(expected)
        && /^https?:\/\//i.test(copy.webPartData.serverProcessedContent.links['button.linkUrl']);
    }, { control: realControl, source: realSource }));
}
