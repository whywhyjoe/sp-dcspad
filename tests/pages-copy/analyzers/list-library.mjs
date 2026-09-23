// SP Workbench — page copy: the List / Document library analyzer
// (design/PAGE-COPY.md §5.1–5.2, web part id
// f92bf067-bc19-489e-a556-7fe95f508720 — the same component serves both the
// List and the Document library controls). Every check imports
// src/workbench/page-copy-analyzers/list-library.js in the page context and
// exercises it directly against the real captured shape (fixtures/
// live-shapes.json "listOrLibrary") and the mock List control's shape
// (mock-pagecopy.js, PAGECOPY_IDS) — no mock webs, no dialog.
// House style: tests/pages-copy/analyzers/text.mjs (module shape, baseCtx).

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
  // ---- refs: real fixture — a Document library instance ----------------------

  await check('analyzer list/library: refs on the real fixture lists the list identity as data with an absolute-vs-relative note', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/list-library.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };
      const refs = analyzer.refs(instance, liveCtx);

      const listId = refs.find((r) => r.path.at(-1) === 'selectedListId');
      const viewId = refs.find((r) => r.path.at(-1) === 'selectedViewId');
      const absoluteUrl = refs.find((r) => r.path.at(-1) === 'selectedListUrl');
      const relativeUrl = refs.find((r) => r.path.at(-1) === 'webRelativeListUrl');

      const allData = refs.length > 0 && refs.every((r) => r.class === 'data');
      const namesTheList = [listId, viewId, absoluteUrl, relativeUrl].every((r) => r && /documents/i.test(r.note));

      return allData
        && Boolean(listId) && listId.value === instance.webPartData.properties.selectedListId
          && /identity/i.test(listId.note) && !/web-relative/i.test(listId.note)
        && Boolean(viewId) && viewId.value === instance.webPartData.properties.selectedViewId
          && /identity/i.test(viewId.note)
        && Boolean(absoluteUrl) && absoluteUrl.value === instance.webPartData.properties.selectedListUrl
          && /absolute/i.test(absoluteUrl.note) && !/web-relative/i.test(absoluteUrl.note)
          && /keeps pointing at the source/i.test(absoluteUrl.note)
        && Boolean(relativeUrl) && relativeUrl.value === instance.webPartData.properties.webRelativeListUrl
          && /web-relative/i.test(relativeUrl.note) && /destination/i.test(relativeUrl.note)
        && namesTheList;
    }, { instance: FIXTURE.listOrLibrary[0], ctx: baseCtx() }));

  // ---- refs: the mock List control -------------------------------------------

  await check('analyzer list/library: refs on the mock List control lists its identity and both URL shapes as data', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/list-library.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const liveCtx = { ...ctx, underPath };

      // Shape from mock-pagecopy.js (search f92bf067): no isDocumentLibrary
      // flag, so the analyzer falls back to the shared 'List / Document
      // library' label rather than guessing.
      const instance = {
        controlType: 3,
        id: 'd15e0000-0000-4000-8000-000000000006',
        webPartId: 'f92bf067-bc19-489e-a556-7fe95f508720',
        webPartData: {
          id: 'f92bf067-bc19-489e-a556-7fe95f508720',
          title: 'List',
          properties: {
            selectedListId: PAGECOPY_IDS.srcEvents,
            selectedListUrl: '/sites/pagesrc/Lists/Events',
            webRelativeListUrl: '/Lists/Events',
          },
          serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
        },
      };
      const refs = analyzer.refs(instance, liveCtx);
      const listId = refs.find((r) => r.path.at(-1) === 'selectedListId');
      const absoluteUrl = refs.find((r) => r.path.at(-1) === 'selectedListUrl');
      const relativeUrl = refs.find((r) => r.path.at(-1) === 'webRelativeListUrl');

      return refs.length === 3 && refs.every((r) => r.class === 'data')
        && Boolean(listId) && listId.value === PAGECOPY_IDS.srcEvents && /identity/i.test(listId.note)
        && Boolean(absoluteUrl) && absoluteUrl.value === '/sites/pagesrc/Lists/Events'
          && /absolute/i.test(absoluteUrl.note) && !/web-relative/i.test(absoluteUrl.note)
        && Boolean(relativeUrl) && relativeUrl.value === '/Lists/Events' && /web-relative/i.test(relativeUrl.note);
    }, { ctx: baseCtx() }));

  // ---- patch: never patches data -----------------------------------------------

  await check('analyzer list/library: patch returns 0 and leaves the instance deep-equal', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/list-library.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const before = JSON.stringify(instance);
      const patchCtx = {
        ...ctx, underPath,
        target: { webUrl: 'https://nervedotnet.sharepoint.com/sites/OtherSite', webPath: '/sites/OtherSite' },
        rewriteLinks: false,
        mapAsset: () => null,
        mapLink: () => null,
      };
      const changed = analyzer.patch(instance, patchCtx);
      return changed === 0 && JSON.stringify(instance) === before;
    }, { instance: FIXTURE.listOrLibrary[0], ctx: baseCtx() }));

  // ---- nothing classified as asset or link -------------------------------------

  await check('analyzer list/library: no reference on either fixture is ever classified as asset or link', () =>
    page.evaluate(async ({ realInstance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/list-library.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const { PAGECOPY_IDS } = await import('/src/workbench/mock-pagecopy.js');
      const liveCtx = { ...ctx, underPath };

      const mockInstance = {
        controlType: 3,
        id: 'd15e0000-0000-4000-8000-000000000006',
        webPartId: 'f92bf067-bc19-489e-a556-7fe95f508720',
        webPartData: {
          id: 'f92bf067-bc19-489e-a556-7fe95f508720',
          title: 'List',
          properties: {
            selectedListId: PAGECOPY_IDS.srcEvents,
            selectedListUrl: '/sites/pagesrc/Lists/Events',
            webRelativeListUrl: '/Lists/Events',
          },
          serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
        },
      };

      const realRefs = analyzer.refs(realInstance, liveCtx);
      const mockRefs = analyzer.refs(mockInstance, liveCtx);
      const none = (refs) => refs.length > 0 && !refs.some((r) => r.class === 'asset' || r.class === 'link');
      return none(realRefs) && none(mockRefs);
    }, { realInstance: FIXTURE.listOrLibrary[0], ctx: baseCtx() }));
}
