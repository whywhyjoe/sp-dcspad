// SP Workbench — page copy: the Events analyzer (design/PAGE-COPY.md
// §5.1–5.2, §11 Q7). Every check imports
// src/workbench/page-copy-analyzers/events.js in the page context and
// exercises it directly against the real captured instance / a synthetic
// variant — no mock webs, no dialog. House style: tests/pages-copy/
// analyzers/news.mjs (module shape, baseCtx).

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const FIXTURE = JSON.parse(
  readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'),
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
  // ---- refs: the real fixture is "this site" (context-relative) ----------

  await check('analyzer events: refs on the real instance classifies the this-site source as context-relative, nothing as asset', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/events.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };
      const refs = analyzer.refs(instance, liveCtx);

      if (refs.some((r) => r.class === 'asset')) return false;
      const siteRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.siteId');
      const webRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.webId');
      const listRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.selectedListId');

      return refs.length === 3
        && Boolean(siteRef) && siteRef.class === 'data' && /context-relative/.test(siteRef.note)
          && siteRef.value === instance.webPartData.properties.siteId
        && Boolean(webRef) && webRef.class === 'data' && /context-relative/.test(webRef.note)
          && webRef.value === instance.webPartData.properties.webId
        && Boolean(listRef) && listRef.class === 'data' && /context-relative/.test(listRef.note)
          && listRef.value === instance.webPartData.properties.selectedListId;
    }, { instance: FIXTURE.events, ctx: baseCtx() }));

  // ---- refs: a synthetic selected-sites source is fixed; links walked ----

  await check('analyzer events: refs on a synthetic selected-sites variant classifies the source as fixed (not context-relative) and classifies source-web links, ignoring baseUrl and external ones', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/events.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      const instance = {
        controlType: 3,
        webPartId: '20745d7d-8581-4a6c-bf26-68279bc123fc',
        webPartData: {
          id: '20745d7d-8581-4a6c-bf26-68279bc123fc',
          properties: {
            dataSource: 7,
            selectedListId: '33333333-3333-3333-3333-333333333333',
            webId: ctx.sourceIds.webId,
            siteId: ctx.sourceIds.siteId,
            sites: [
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
      const siteRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.siteId');
      const webRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.webId');
      const listRef = refs.find((r) => r.path.join('.') === 'webPartData.properties.selectedListId');
      const baseUrlRef = refs.find((r) => r.path.at(-1) === 'baseUrl');
      const itemLink = refs.find((r) => r.value === '/sites/NewNerve/SitePages/Other.aspx');
      const external = refs.find((r) => r.value === 'https://bing.com/external');

      return !hasAsset
        && Boolean(site0) && site0.class === 'data' && !/context-relative/.test(site0.note)
        && Boolean(site1) && site1.class === 'data' && !/context-relative/.test(site1.note)
        && Boolean(siteRef) && !/context-relative/.test(siteRef.note)
        && Boolean(webRef) && !/context-relative/.test(webRef.note)
        && Boolean(listRef) && !/context-relative/.test(listRef.note)
        && !baseUrlRef
        && Boolean(itemLink) && itemLink.class === 'link'
          && itemLink.path.join('.') === 'webPartData.serverProcessedContent.links.items[0].sourceItem.url'
        && !external;
    }, { ctx: baseCtx() }));

  // ---- patch: rewriteLinks false changes nothing -------------------------

  await check('analyzer events: patch with rewriteLinks false changes nothing (deep-equal)', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/events.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const before = JSON.stringify(instance);
      const patchCtx = { ...ctx, underPath, rewriteLinks: false, mapLink: () => null, mapAsset: () => null };
      const changed = analyzer.patch(instance, patchCtx);
      return changed === 0 && JSON.stringify(instance) === before;
    }, { instance: JSON.parse(JSON.stringify(FIXTURE.events)), ctx: baseCtx() }));
}
