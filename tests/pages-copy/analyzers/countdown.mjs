// Countdown timer analyzer unit tests (design/PAGE-COPY.md §5.1-5.2, id
// 62cac389-787f-495d-beca-e11786162ef4). House style:
// tests/pages-copy/analyzers/quick-links.mjs — imports
// src/workbench/page-copy-analyzers/countdown.js and page-copy.js's
// analyzerContext in the page context and exercises them against fixture
// instances. Two sources:
//   - the real editor-authored capture (tests/pages-copy/fixtures/
//     live-shapes-editor.json → countdown — background set, CTA off, so a
//     clone with buttonURL populated exercises the 'link' class and the
//     rewrite path)
//   - the mock ids (src/workbench/mock-pagecopy.js's PAGECOPY_IDS) for the
//     end-to-end mapAsset/mapLink checks, matching the other analyzer
//     suites' convention.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapes = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapes.countdown;
  const realSource = liveShapes.source;

  await check('analyzer countdown: refs on the real fixture — the background image is an asset ref carrying the four ids off properties.backgroundImage (id, not uniqueId), and the CTA-off buttonURL produces no link ref', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/countdown.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);

      const asset = refs.find((r) => r.class === 'asset');
      const linkRefs = refs.filter((r) => r.class === 'link');
      const bg = control.webPartData.properties.backgroundImage;

      return Boolean(asset)
        && asset.path.join('.') === 'webPartData.serverProcessedContent.imageSources.backgroundImage.url'
        && asset.asset.path === '/sites/NewNerve/SiteAssets/zz-pagecopy-assets/thumb.png'
        && asset.asset.ids.siteId === bg.siteId && asset.asset.ids.webId === bg.webId
        && asset.asset.ids.listId === bg.listId && asset.asset.ids.uniqueId === bg.id
        && linkRefs.length === 0; // buttonURL is empty on this capture — CTA off
    }, { control: realControl, source: realSource }));

  await check('analyzer countdown: refs — a synthetic variant with a call-to-action link into the source web classifies as a link ref', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/countdown.js')).default;
      const ctx = analyzerContext({ web: source });
      const clone = structuredClone(control);
      clone.webPartData.properties.showButton = true;
      clone.webPartData.properties.buttonURL = '/sites/NewNerve/SitePages/RSVP.aspx';
      const refs = analyzer.refs(clone, ctx);
      const linkRefs = refs.filter((r) => r.class === 'link');
      return linkRefs.length === 1
        && linkRefs[0].path.join('.') === 'webPartData.properties.buttonURL'
        && linkRefs[0].value === '/sites/NewNerve/SitePages/RSVP.aspx';
    }, { control: realControl, source: realSource }));

  await check('analyzer countdown: patch with no mapping and rewriteLinks false changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/countdown.js')).default;
      const clone = structuredClone(control);
      clone.webPartData.properties.showButton = true;
      clone.webPartData.properties.buttonURL = '/sites/NewNerve/SitePages/RSVP.aspx';
      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: false,
        mapAsset: () => null,
        mapLink: () => null,
      };
      const before = JSON.stringify(clone);
      const copy = structuredClone(clone);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { control: realControl, source: realSource }));

  await check('analyzer countdown: patch with a mapAsset mapping rewrites the background image path and the four ids on properties.backgroundImage (id key) and nothing else', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/countdown.js')).default;
      const ids = PAGECOPY_IDS;
      const ctx = {
        ...analyzerContext({ web: source }),
        target: { webUrl: `${location.origin}/sites/pagedst`, webPath: '/sites/pagedst' },
        rewriteLinks: false,
        mapAsset: () => ({
          path: '/sites/pagedst/SiteAssets/SitePages/Countdown-copy/thumb.png',
          ids: { siteId: ids.dstSite, webId: ids.dstWeb, listId: ids.dstSiteAssets, uniqueId: 'd15f0000-0000-4000-8000-0000000000f2' },
          url: `${location.origin}/sites/pagedst/SiteAssets/SitePages/Countdown-copy/thumb.png`,
        }),
        mapLink: () => null,
      };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.serverProcessedContent.imageSources['backgroundImage.url'] = '/sites/pagedst/SiteAssets/SitePages/Countdown-copy/thumb.png';
      Object.assign(expected.webPartData.properties.backgroundImage, {
        id: 'd15f0000-0000-4000-8000-0000000000f2', siteId: ids.dstSite, webId: ids.dstWeb, listId: ids.dstSiteAssets,
      });

      return n === 5 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: realControl, source: realSource }));
}
