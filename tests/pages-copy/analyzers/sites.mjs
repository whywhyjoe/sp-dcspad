// SP Workbench — page copy: the Sites analyzer (design/PAGE-COPY.md
// §5.1–5.2). Every check imports
// src/workbench/page-copy-analyzers/sites.js in the page context and
// exercises it directly against the real captured shape (fixtures/
// live-shapes-editor.json "sites" — two selected sites) and a frequent-sites
// variant — no mock webs, no dialog. House style: tests/pages-copy/
// analyzers/list-library.mjs (module shape, baseCtx).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = JSON.parse(
  readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'),
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
  // ---- refs: real fixture — two selected sites, one under the source web ----

  await check('analyzer sites: refs on the real fixture (selected sites) reports ids and urls as data for both sites — a subweb under the source path included, since re-basing it would aim the card at a site that does not exist — and flips the note when a site IS the source web', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/sites.js');
      const { underPath, normalizeGuid } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath, normalizeGuid };
      const refs = analyzer.refs(instance, liveCtx);

      const site0Ids = refs.filter((r) => r.path[3] === 0 && r.path.includes('ItemReference'));
      const site1Ids = refs.filter((r) => r.path[3] === 1 && r.path.includes('ItemReference'));
      const site0Url = refs.find((r) => r.path.at(-1) === 'sites[0].Url');
      const site1Url = refs.find((r) => r.path.at(-1) === 'sites[1].Url');

      const noAssets = !refs.some((r) => r.class === 'asset');
      // Site 0 (Group "SPUtils Test") carries SiteId+WebId+GroupId; site 1
      // (Site "TestSiteCollection") has no GroupId, so just SiteId+WebId.
      const site0Shape = site0Ids.length === 3 && site0Ids.every((r) => r.class === 'data' && /another site/i.test(r.note));
      const site1Shape = site1Ids.length === 2 && site1Ids.every((r) => r.class === 'data' && /another site/i.test(r.note));
      const urlShape = Boolean(site0Url) && site0Url.class === 'data' && /another site/i.test(site0Url.note)
        && Boolean(site1Url) && site1Url.class === 'data' && /another site/i.test(site1Url.note);

      // Now simulate that the second selected site (TestSiteCollection) IS
      // the source web: the id notes for that site must switch to the
      // self-reference phrasing, while the first site's notes are untouched.
      const selfCtx = {
        ...liveCtx,
        sourceIds: { ...ctx.sourceIds, webId: instance.webPartData.properties.sites[1].ItemReference.WebId },
      };
      const selfRefs = analyzer.refs(instance, selfCtx);
      const selfSite1Ids = selfRefs.filter((r) => r.path[3] === 1 && r.path.includes('ItemReference'));
      const selfSite0Ids = selfRefs.filter((r) => r.path[3] === 0 && r.path.includes('ItemReference'));
      const selfShape = selfSite1Ids.length === 2 && selfSite1Ids.every((r) => /points at the source site itself/i.test(r.note))
        && selfSite0Ids.length === 3 && selfSite0Ids.every((r) => /another site/i.test(r.note));

      return noAssets && site0Shape && site1Shape && urlShape && selfShape;
    }, { instance: FIXTURE.sites, ctx: baseCtx() }));

  // ---- refs: frequent sites for current user carries no references -----------

  await check('analyzer sites: refs on a frequent-sites variant (empty sites list) reports nothing', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/sites.js');
      const { underPath, normalizeGuid } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath, normalizeGuid };

      const instance = {
        controlType: 3,
        webPartId: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b',
        webPartData: {
          id: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b',
          properties: {
            sites: [],
            sourceType: 0,
            itemCount: 8,
            hideWebPartWhenEmpty: false,
            layoutId: 'FilmStrip',
            dataProviderId: 'SitesDataProvider',
            webId: ctx.sourceIds.webId,
            siteId: ctx.sourceIds.siteId,
          },
          serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: { baseUrl: '/sites/NewNerve' } },
        },
      };
      const refs = analyzer.refs(instance, liveCtx);
      return refs.length === 0;
    }, { ctx: baseCtx() }));

  // ---- patch: mapLink rewrites only the subweb url; data refs stay put -------

  await check('analyzer sites: patch never changes a selected site — even with rewriteLinks on and a mapLink that would re-base the subweb url', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/sites.js');
      const { underPath, normalizeGuid } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath, normalizeGuid };
      const before = JSON.parse(JSON.stringify(instance));

      const patchCtx = {
        ...liveCtx,
        rewriteLinks: true,
        mapLink: (value) => (value === '/sites/NewNerve/sputils-test' ? '/sites/OtherSite/sputils-test' : null),
        mapAsset: () => null,
      };
      const changed = analyzer.patch(instance, patchCtx);

      const urlChanged = instance.webPartData.serverProcessedContent.links['sites[0].Url'] === before.webPartData.serverProcessedContent.links['sites[0].Url'];
      const otherUrlSame = instance.webPartData.serverProcessedContent.links['sites[1].Url'] === before.webPartData.serverProcessedContent.links['sites[1].Url'];
      const propsSame = JSON.stringify(instance.webPartData.properties) === JSON.stringify(before.webPartData.properties);

      // No-op patch: nothing to map, so nothing changes and the instance is deep-equal.
      const noopCtx = { ...liveCtx, rewriteLinks: false, mapLink: () => null, mapAsset: () => null };
      const noopInstance = JSON.parse(JSON.stringify(before));
      const noopChanged = analyzer.patch(noopInstance, noopCtx);
      const noopSame = JSON.stringify(noopInstance) === JSON.stringify(before);

      return changed === 0 && urlChanged && otherUrlSame && propsSame && noopChanged === 0 && noopSame;
    }, { instance: FIXTURE.sites, ctx: baseCtx() }));
}
