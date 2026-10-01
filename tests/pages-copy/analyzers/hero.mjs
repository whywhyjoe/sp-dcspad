// Hero analyzer unit tests (design/PAGE-COPY.md §5.1-5.2, id
// c4bd7b2f-7b6e-4599-8485-16504575f590). House style:
// tests/pages-copy/analyzers/quick-links.mjs — imports
// src/workbench/page-copy-analyzers/hero.js and page-copy.js's
// analyzerContext in the page context and exercises them against fixture
// instances from tests/pages-copy/fixtures/live-shapes-editor.json:
//   - heroConfigured: an editor-authored carousel with one slide — a
//     custom background image (asset) and a link into the source web —
//     plus an auto-selected previewImage taken from the linked item
//     (link-dependent config, never an asset).
//   - hero: the manifest-default instance the page is seeded with, no
//     `content` at all — must yield no refs and never crash.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapesEditor = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const heroConfigured = liveShapesEditor.heroConfigured;
  const heroDefault = liveShapesEditor.hero;
  const source = liveShapesEditor.source;

  await check('analyzer hero: refs on heroConfigured — the custom image is an asset ref carrying the slide\'s four ids, the slide link is a link ref, and the auto-selected preview image is link-dependent config', () =>
    page.evaluate(async ({ control, src }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = analyzerContext({ web: src });
      const refs = analyzer.refs(control, ctx);

      const asset = refs.find((r) => r.class === 'asset');
      const link = refs.find((r) => r.class === 'link');
      const preview = refs.find((r) => r.class === 'config'
        && r.path.join('.') === 'webPartData.serverProcessedContent.imageSources.content[0].previewImage.url');

      return Boolean(asset)
        && asset.path.join('.') === 'webPartData.serverProcessedContent.imageSources.content[0].image.url'
        && asset.asset.path === '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/hero.png'
        && asset.asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9'
        && asset.asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8'
        && asset.asset.ids.listId === '8af0bba1-d6fb-4db3-9ec3-e71be6a23a45'
        && asset.asset.ids.uniqueId === '612ddc56-e831-4e33-a69f-1e22f7463855'
        && Boolean(link)
        && link.path.join('.') === 'webPartData.serverProcessedContent.links.content[0].link'
        && link.value === '/sites/NewNerve/SitePages/Test.aspx'
        && Boolean(preview) // reported, so it isn't silently dropped
        && !preview.asset; // but never carries an asset identity to transfer
    }, { control: heroConfigured, src: source }));

  await check('analyzer hero: refs on the empty manifest-default instance — no refs, no crash', () =>
    page.evaluate(async ({ control, src }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = analyzerContext({ web: src });
      const refs = analyzer.refs(control, ctx);
      return Array.isArray(refs) && refs.length === 0;
    }, { control: heroDefault, src: source }));

  await check('analyzer hero: patch with no mapping and rewriteLinks false changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, src }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = {
        ...analyzerContext({ web: src }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: false,
        mapAsset: () => null,
        mapLink: () => null,
      };
      const before = JSON.stringify(control);
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { control: heroConfigured, src: source }));

  await check('analyzer hero: patch with a mapAsset mapping rewrites the tile image\'s path plus its customMetadata and content[].image ids, and nothing else (link/previewImage untouched)', () =>
    page.evaluate(async ({ control, src }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/hero.png',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-shapes-copy/hero.png`,
      };
      const ctx = { ...analyzerContext({ web: src }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.imageSources['content[0].image.url'] = mapping.path;
      const meta = expected.webPartData.serverProcessedContent.customMetadata['content[0].image.url'];
      meta.siteid = mapping.ids.siteId;
      meta.webid = mapping.ids.webId;
      meta.listid = mapping.ids.listId;
      meta.uniqueid = mapping.ids.uniqueId;
      const img = expected.webPartData.properties.content[0].image;
      img.siteId = mapping.ids.siteId;
      img.webId = mapping.ids.webId;
      img.listId = mapping.ids.listId;
      img.id = mapping.ids.uniqueId;
      // Derived copies of the location follow too (live shape): resolvedUrl
      // becomes the destination url, and imageUrl's embedded site/web/list/
      // item GUIDs are swapped for the destination's — nothing else in it.
      // Written out literally rather than recomputed, so the check does not
      // share the swap logic it is testing.
      const srcImg = control.webPartData.properties.content[0].image;
      img.resolvedUrl = mapping.url;
      img.imageUrl = 'https://nervedotnet.sharepoint.com/_api/v2.1/sites/nervedotnet.sharepoint.com,'
        + 'd1510000-0000-4000-8000-000000000003,d1520000-0000-4000-8000-000000000003'
        + '/lists/d1540000-0000-4000-8000-000000000003/items/d15f0000-0000-4000-8000-0000000000a1'
        + '/driveItem/thumbnails/0/c400x99999/content?prefer=noRedirect,extendCacheMaxAge&clientType=modernWebPart&format=webp';

      return n === 11 && img.imageUrl !== srcImg.imageUrl && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: heroConfigured, src: source }));

  await check('analyzer hero: patch with rewriteLinks true and a mapLink mapping rewrites only the slide link, leaving the image and previewImage untouched', () =>
    page.evaluate(async ({ control, src }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = {
        ...analyzerContext({ web: src }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: true,
        mapAsset: () => null,
        mapLink: (value) => (value === '/sites/NewNerve/SitePages/Test.aspx' ? '/sites/pagedst/SitePages/Test.aspx' : null),
      };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.links['content[0].link'] = '/sites/pagedst/SitePages/Test.aspx';

      return n === 1 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: heroConfigured, src: source }));

  // Guards on the derived fields (xo review): each is rewritten only when it
  // demonstrably names the transferred file, and the GUID swap touches whole
  // GUID tokens only.
  const derivedMapping = {
    path: '/sites/pagedst/SiteAssets/SitePages/x/hero.png',
    ids: {
      siteId: 'd1510000-0000-4000-8000-000000000003',
      webId: 'd1520000-0000-4000-8000-000000000003',
      listId: 'd1540000-0000-4000-8000-000000000003',
      uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
    },
    url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/x/hero.png',
  };

  await check('analyzer hero: a resolvedUrl naming a DIFFERENT file in the source web is left alone while the image source, ids and imageUrl still move', () =>
    page.evaluate(async ({ control, src, mapping }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = { ...analyzerContext({ web: src }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(control);
      const other = 'https://nervedotnet.sharepoint.com/sites/NewNerve/SiteAssets/zz-pagecopy-assets/other.png';
      copy.webPartData.properties.content[0].image.resolvedUrl = other;
      const n = analyzer.patch(copy, ctx);
      const img = copy.webPartData.properties.content[0].image;
      return n === 10 && img.resolvedUrl === other
        && copy.webPartData.serverProcessedContent.imageSources['content[0].image.url'] === mapping.path
        && img.imageUrl.includes('d15f0000-0000-4000-8000-0000000000a1');
    }, { control: heroConfigured, src: source, mapping: derivedMapping }));

  await check('analyzer hero: the imageUrl swap changes whole GUID tokens only — a longer hex run containing the item id (bare) is kept byte-for-byte', () =>
    page.evaluate(async ({ control, src, mapping }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = { ...analyzerContext({ web: src }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(control);
      const img = copy.webPartData.properties.content[0].image;
      const decoy = '0612ddc56e8314e33a69f1e22f74638550';   // contains the bare item id, is not it
      img.imageUrl = `${img.imageUrl}&cache=${decoy}`;
      analyzer.patch(copy, ctx);
      return img.imageUrl.endsWith(`&cache=${decoy}`)
        && img.imageUrl.includes('/items/d15f0000-0000-4000-8000-0000000000a1/')
        && !img.imageUrl.includes('612ddc56-e831-4e33-a69f-1e22f7463855');
    }, { control: heroConfigured, src: source, mapping: derivedMapping }));

  await check('analyzer hero: an imageUrl carrying the item id only in bare, upper-case form is still recognised and swapped to the bare destination id', () =>
    page.evaluate(async ({ control, src, mapping }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/hero.js')).default;
      const ctx = { ...analyzerContext({ web: src }), mapAsset: () => mapping, mapLink: () => null };
      const copy = structuredClone(control);
      const img = copy.webPartData.properties.content[0].image;
      img.imageUrl = 'https://nervedotnet.sharepoint.com/_api/v2.1/drives/x/items/612DDC56E8314E33A69F1E22F7463855/thumbnails/0/c400x99999/content';
      analyzer.patch(copy, ctx);
      return img.imageUrl === 'https://nervedotnet.sharepoint.com/_api/v2.1/drives/x/items/d15f00000000400080000000000000a1/thumbnails/0/c400x99999/content';
    }, { control: heroConfigured, src: source, mapping: derivedMapping }));
}
