// SP Workbench — page copy: the text analyzer (design/PAGE-COPY.md §5.1–5.2,
// controlType 4 rich text). Every check imports
// src/workbench/page-copy-analyzers/text.js in the page context and
// exercises it directly against inline fixtures — no mock webs, no dialog.
// House style: tests/pages-copy/pure.mjs (module shape).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = JSON.parse(
  readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes.json'), 'utf8'),
);

// A ctx shaped like analyzerContext(snapshot) in page-copy.js (§2.1): the
// analyzer never imports page-copy.js directly, so these checks build one by
// hand from the fixture's captured source web.
function baseCtx(overrides = {}) {
  return {
    origin: 'https://nervedotnet.sharepoint.com',
    sourceWebPath: '/sites/NewNerve',
    sourceIds: { siteId: FIXTURE.source.siteId, webId: FIXTURE.source.webId },
    ...overrides,
  };
}

export async function checks({ page, check }) {
  // ---- refs: real fixture instances (all links are external) ----------------

  await check('analyzer text: refs classifies source-web links and images, ignoring external ones', () =>
    page.evaluate(async ({ textFixtures, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/text.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      // Every captured Test.aspx text part links only to bing.com — external
      // to the source web, so refs() must report nothing for any of them.
      const externalOnly = textFixtures.every((instance) => {
        const refs = analyzer.refs(instance, liveCtx);
        return Array.isArray(refs) && refs.length === 0;
      });

      // A synthetic instance exercises positive classification: a relative
      // link into the source web, an absolute link into the source web, an
      // external link, a relative image into the source web, and a data-*
      // attribute into the source web.
      const synthetic = {
        controlType: 4,
        innerHTML: ''
          + '<p><a href="/sites/NewNerve/SitePages/Other.aspx">relative link</a></p>'
          + '<p><a href="https://nervedotnet.sharepoint.com/sites/NewNerve/SitePages/Other2.aspx">absolute link</a></p>'
          + '<p><a href="https://bing.com/">external</a></p>'
          + '<p><img src="/sites/NewNerve/SiteAssets/logo.png" data-foo="/sites/NewNerve/SiteAssets/other.png"></p>',
      };
      const refs = analyzer.refs(synthetic, liveCtx);
      const link1 = refs.find((r) => r.value === '/sites/NewNerve/SitePages/Other.aspx');
      const link2 = refs.find((r) => r.value === 'https://nervedotnet.sharepoint.com/sites/NewNerve/SitePages/Other2.aspx');
      const img = refs.find((r) => r.value === '/sites/NewNerve/SiteAssets/logo.png');
      const dataAttr = refs.find((r) => r.value === '/sites/NewNerve/SiteAssets/other.png');
      const externalRef = refs.find((r) => r.value === 'https://bing.com/');

      return externalOnly
        && refs.length === 4
        && Boolean(link1) && link1.class === 'link' && link1.path[0] === 'innerHTML'
        && Boolean(link2) && link2.class === 'link'
        && Boolean(img) && img.class === 'asset' && img.asset && img.asset.path === '/sites/NewNerve/SiteAssets/logo.png'
          && JSON.stringify(img.asset.ids) === '{}'
        && Boolean(dataAttr) && dataAttr.class === 'link'
        && !externalRef;
    }, { textFixtures: FIXTURE.text, ctx: baseCtx() }));

  // ---- path-boundary matching ------------------------------------------------

  await check('analyzer text: path-boundary — /sites/NewNerve-other/x is not a ref when the source is /sites/NewNerve', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/text.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };
      const instance = {
        controlType: 4,
        innerHTML: '<p><a href="/sites/NewNerve-other/x">not a boundary match</a></p>'
          + '<p><a href="/sites/NewNerve/x">a boundary match</a></p>',
      };
      const refs = analyzer.refs(instance, liveCtx);
      const bad = refs.find((r) => r.value === '/sites/NewNerve-other/x');
      const good = refs.find((r) => r.value === '/sites/NewNerve/x');
      return refs.length === 1 && !bad && Boolean(good);
    }, { ctx: baseCtx() }));

  // ---- patch: rewriteLinks false changes nothing -----------------------------

  await check('analyzer text: patch with rewriteLinks false changes nothing (innerHTML byte-identical)', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/text.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      // Only link references — no image, so this isolates the "rewriteLinks
      // false" case cleanly: per the contract, ctx.mapLink returns null
      // unless the operator opted into link rewriting (page-copy.js's
      // rewriteContent builds mapLink exactly this way), so patch() must
      // leave the markup byte-identical and report zero changes.
      const html = '<h4><a href="/sites/NewNerve/SitePages/Other.aspx">Kept text →</a></h4>'
        + '<p><a href="https://bing.com/">external stays too</a></p>';
      const instance = { controlType: 4, innerHTML: html };
      const patchCtx = {
        ...ctx, underPath,
        rewriteLinks: false,
        mapLink: () => null,
        mapAsset: () => null,
      };
      const changed = analyzer.patch(instance, patchCtx);
      return changed === 0 && instance.innerHTML === html;
    }, { ctx: baseCtx() }));

  // ---- patch: mapLink + mapAsset rewrite exactly the mapped attributes ------

  await check('analyzer text: patch with mapLink and mapAsset rewrites exactly those attributes and keeps all other markup', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/text.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const html = ''
        + '<h4><a href="/sites/NewNerve/SitePages/Other.aspx">Kept text →</a></h4>'
        + '<p><span class="fontSizeMediumPlus">Untouched paragraph.</span></p>'
        + '<p><img src="/sites/NewNerve/SiteAssets/logo.png" alt="logo"></p>'
        + '<p><a href="https://bing.com/">external stays</a></p>';
      const instance = { controlType: 4, innerHTML: html };
      const patchCtx = {
        ...ctx, underPath,
        rewriteLinks: true,
        mapLink: (value) => (value === '/sites/NewNerve/SitePages/Other.aspx'
          ? '/sites/NewSite/SitePages/Other.aspx' : null),
        mapAsset: (identity) => (identity.path === '/sites/NewNerve/SiteAssets/logo.png'
          ? { path: '/sites/NewSite/SiteAssets/logo.png', ids: {}, url: 'https://nervedotnet.sharepoint.com/sites/NewSite/SiteAssets/logo.png' }
          : null),
      };
      const changed = analyzer.patch(instance, patchCtx);
      const out = instance.innerHTML;
      return changed === 2
        && out.includes('href="/sites/NewSite/SitePages/Other.aspx"')
        && out.includes('src="/sites/NewSite/SiteAssets/logo.png"')
        && out.includes('Kept text')
        && out.includes('Untouched paragraph.')
        && out.includes('alt="logo"')
        && out.includes('href="https://bing.com/"')
        && !out.includes('/sites/NewNerve/SitePages/Other.aspx')
        && !out.includes('/sites/NewNerve/SiteAssets/logo.png');
    }, { ctx: baseCtx() }));

  // ---- patch: an unmapped image keeps its src --------------------------------

  await check('analyzer text: an unmapped image keeps its src', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/text.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const html = '<p><img src="/sites/NewNerve/SiteAssets/missing.png" alt="gone"></p>';
      const instance = { controlType: 4, innerHTML: html };
      const patchCtx = {
        ...ctx, underPath,
        rewriteLinks: true,
        mapLink: () => null,
        mapAsset: () => null, // not transferred — retained, per §5.3
      };
      const changed = analyzer.patch(instance, patchCtx);
      return changed === 0 && instance.innerHTML === html;
    }, { ctx: baseCtx() }));
}
