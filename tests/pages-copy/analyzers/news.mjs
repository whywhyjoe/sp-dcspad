// SP Workbench — page copy: the News analyzer (design/PAGE-COPY.md §5.1–5.2,
// §11 Q7). Every check imports
// src/workbench/page-copy-analyzers/news.js in the page context and
// exercises it directly against inline fixtures / the real captured
// instances — no mock webs, no dialog. House style: tests/pages-copy/
// analyzers/text.mjs (module shape), tests/pages-copy/pure.mjs.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = JSON.parse(
  readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes.json'), 'utf8'),
);

// A ctx shaped like analyzerContext(snapshot) in page-copy.js (§2.1): the
// analyzer never imports page-copy.js directly, so these checks build one
// by hand from the fixture's captured source web.
function baseCtx(overrides = {}) {
  return {
    origin: 'https://nervedotnet.sharepoint.com',
    sourceWebPath: '/sites/NewNerve',
    sourceIds: { siteId: FIXTURE.source.siteId, webId: FIXTURE.source.webId },
    ...overrides,
  };
}

export async function checks({ page, check }) {
  const newsFixtures = Object.values(FIXTURE.news);

  // ---- refs: both real fixture instances are "this site" (context-relative) --

  await check("analyzer news: refs classifies the real fixtures' site/web source as context-relative, nothing as asset", () =>
    page.evaluate(async ({ instances, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/news.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      return instances.every((instance) => {
        const refs = analyzer.refs(instance, liveCtx);
        if (refs.some((r) => r.class === 'asset')) return false;
        const siteRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.siteId');
        const webRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.webId');
        return refs.length === 2
          && Boolean(siteRef) && siteRef.class === 'data' && /context-relative/.test(siteRef.note)
          && siteRef.value === instance.webPartData.properties.siteId
          && Boolean(webRef) && webRef.class === 'data' && /context-relative/.test(webRef.note)
          && webRef.value === instance.webPartData.properties.webId;
      });
    }, { instances: newsFixtures, ctx: baseCtx() }));

  // ---- refs: a selected-sites source is fixed; links are walked, baseUrl isn't -

  await check('analyzer news: refs classifies a selected-sites source as fixed (not context-relative) and classifies source-web links, ignoring baseUrl and external ones', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/news.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      const instance = {
        controlType: 3,
        webPartId: '8c88f208-6c77-4bdb-86a0-0c47b4316588',
        webPartData: {
          id: '8c88f208-6c77-4bdb-86a0-0c47b4316588',
          properties: {
            newsDataSourceProp: 2,
            newsSiteList: [
              '11111111-1111-1111-1111-111111111111',
              '22222222-2222-2222-2222-222222222222',
            ],
          },
          serverProcessedContent: {
            links: {
              baseUrl: '/sites/NewNerve',
              'items[0].sourceItem.url': '/sites/NewNerve/SitePages/Other.aspx',
              'items[1].sourceItem.url': 'https://bing.com/external',
            },
          },
        },
      };
      const refs = analyzer.refs(instance, liveCtx);
      const hasAsset = refs.some((r) => r.class === 'asset');
      const site0 = refs.find((r) => r.value === '11111111-1111-1111-1111-111111111111');
      const site1 = refs.find((r) => r.value === '22222222-2222-2222-2222-222222222222');
      const baseUrlRef = refs.find((r) => r.path[r.path.length - 1] === 'baseUrl');
      const itemLink = refs.find((r) => r.value === '/sites/NewNerve/SitePages/Other.aspx');
      const external = refs.find((r) => r.value === 'https://bing.com/external');

      return !hasAsset
        && Boolean(site0) && site0.class === 'data' && !/context-relative/.test(site0.note)
        && Boolean(site1) && site1.class === 'data' && !/context-relative/.test(site1.note)
        && !baseUrlRef
        && Boolean(itemLink) && itemLink.class === 'link'
          && itemLink.path.join('.') === 'webPartData.serverProcessedContent.links.items[0].sourceItem.url'
        && !external;
    }, { ctx: baseCtx() }));

  // ---- patch: rewriteLinks false changes nothing -----------------------------

  await check('analyzer news: patch with rewriteLinks false changes nothing (deep-equal)', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/news.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const before = JSON.stringify(instance);
      const patchCtx = { ...ctx, underPath, rewriteLinks: false, mapLink: () => null, mapAsset: () => null };
      const changed = analyzer.patch(instance, patchCtx);
      return changed === 0 && JSON.stringify(instance) === before;
    }, { instance: JSON.parse(JSON.stringify(newsFixtures[0])), ctx: baseCtx() }));

  // ---- patch: mapLink rewrites only link refs, data refs stay put -----------

  await check('analyzer news: patch with mapLink rewrites only link refs, leaving data refs untouched', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/news.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      const instance = {
        controlType: 3,
        webPartId: 'a5df8fdf-b508-4b66-98a6-d83bc2597f63',
        webPartData: {
          id: 'a5df8fdf-b508-4b66-98a6-d83bc2597f63',
          properties: {
            newsDataSourceProp: 1,
            webId: ctx.sourceIds.webId,
            siteId: ctx.sourceIds.siteId,
            newsSiteList: [],
          },
          serverProcessedContent: {
            links: {
              baseUrl: '/sites/NewNerve',
              'items[0].sourceItem.url': '/sites/NewNerve/SitePages/Other.aspx',
            },
          },
        },
      };
      const originalProps = JSON.stringify(instance.webPartData.properties);

      const patchCtx = {
        ...liveCtx,
        rewriteLinks: true,
        mapLink: (value) => (value === '/sites/NewNerve/SitePages/Other.aspx'
          ? '/sites/NewSite/SitePages/Other.aspx' : null),
        mapAsset: () => null,
      };
      const changed = analyzer.patch(instance, patchCtx);

      return changed === 1
        && instance.webPartData.serverProcessedContent.links['items[0].sourceItem.url'] === '/sites/NewSite/SitePages/Other.aspx'
        && instance.webPartData.serverProcessedContent.links.baseUrl === '/sites/NewNerve'
        && JSON.stringify(instance.webPartData.properties) === originalProps;
    }, { ctx: baseCtx() }));
}
