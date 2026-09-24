// File and media (File viewer) analyzer unit tests (design/PAGE-COPY.md
// §5.2-§5.3, web part id b7dd04e1-19ce-4b24-9132-b60a1c2b910d). House
// style: tests/pages-copy/analyzers/quick-links.mjs — imports
// src/workbench/page-copy-analyzers/file-viewer.js and page-copy.js's
// analyzerContext in the page context and exercises them against fixture
// instances.
//
// Source: the real editor-authored capture, tests/pages-copy/fixtures/
// live-shapes-editor.json → fileViewer (Shared Documents/zz-pagecopy-docs/
// zz-pagecopy-report.pdf). On that capture wopiurl is byte-identical to
// serverRelativeUrl; a Doc.aspx/WOPI `sourcedoc={guid}` form (§5.3) is not
// in any repo fixture, so one is reproduced inline from the same file's own
// ids for the dedicated sourcedoc check.

import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const liveShapesEditor = JSON.parse(readFileSync(join(__dirname, '..', 'fixtures', 'live-shapes-editor.json'), 'utf8'));

export async function checks({ page, check }) {
  const realControl = liveShapesEditor.fileViewer;
  const realSource = liveShapesEditor.source;

  await check('analyzer file viewer: refs on the real fixture — file, serverRelativeUrl and wopiurl are asset refs sharing one identity', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/file-viewer.js')).default;
      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(control, ctx);
      const assetRefs = refs.filter((r) => r.class === 'asset');
      const paths = assetRefs.map((r) => r.path.join('.'));

      return refs.length === 3 && assetRefs.length === 3
        && paths.includes('webPartData.properties.file')
        && paths.includes('webPartData.serverProcessedContent.links.serverRelativeUrl')
        && paths.includes('webPartData.serverProcessedContent.links.wopiurl')
        && assetRefs.every((r) => r.asset.path === '/sites/NewNerve/Shared Documents/zz-pagecopy-docs/zz-pagecopy-report.pdf')
        && assetRefs.every((r) => r.asset.ids.siteId === '93de26bf-15a3-4fac-8f3e-2b8007d813a9')
        && assetRefs.every((r) => r.asset.ids.webId === '773960b7-b16d-4f81-957a-02de394949b8')
        && assetRefs.every((r) => r.asset.ids.listId === 'acea6a08-5afa-4db6-bda8-c3dbb165760e')
        && assetRefs.every((r) => r.asset.ids.uniqueId === 'de981877-5316-45f1-949e-b2fb7a6eca59');
    }, { control: realControl, source: realSource }));

  await check('analyzer file viewer: patch with no mapping changes nothing (0 returned, deep-equal before/after)', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/file-viewer.js')).default;
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => null };
      const before = JSON.stringify(control);
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);
      return n === 0 && JSON.stringify(copy) === before;
    }, { control: realControl, source: realSource }));

  await check('analyzer file viewer: patch with a verified mapping rewrites properties.file, the four ids, links.serverRelativeUrl and links.wopiurl consistently — nothing else', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/file-viewer.js')).default;
      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-report-copy/zz-pagecopy-report.pdf',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-report-copy/zz-pagecopy-report.pdf',
      };
      const ctx = { ...analyzerContext({ web: source }), mapAsset: () => mapping };
      const copy = structuredClone(control);
      const n = analyzer.patch(copy, ctx);

      const expected = structuredClone(control);
      expected.webPartData.properties.file = mapping.url;
      Object.assign(expected.webPartData.properties, mapping.ids);
      expected.webPartData.serverProcessedContent.links.serverRelativeUrl = mapping.path;
      expected.webPartData.serverProcessedContent.links.wopiurl = mapping.path;

      return n === 7 && JSON.stringify(copy) === JSON.stringify(expected);
    }, { control: realControl, source: realSource }));

  await check('analyzer file viewer: a Doc.aspx/WOPI sourcedoc={guid} form of wopiurl has only its GUID token rewritten, the rest of the url untouched', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/file-viewer.js')).default;
      const clone = structuredClone(control);
      const docAspxUrl = '/sites/NewNerve/_layouts/15/Doc.aspx?sourcedoc=%7Bde981877-5316-45f1-949e-b2fb7a6eca59%7D&file=zz-pagecopy-report.pdf&action=default&mobileredirect=true';
      clone.webPartData.serverProcessedContent.links.wopiurl = docAspxUrl;

      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(clone, ctx);
      const wopiRef = refs.find((r) => r.path.join('.') === 'webPartData.serverProcessedContent.links.wopiurl');

      const mapping = {
        path: '/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-report-copy/zz-pagecopy-report.pdf',
        ids: {
          siteId: 'd1510000-0000-4000-8000-000000000003',
          webId: 'd1520000-0000-4000-8000-000000000003',
          listId: 'd1540000-0000-4000-8000-000000000003',
          uniqueId: 'd15f0000-0000-4000-8000-0000000000a1',
        },
        url: 'https://t.sharepoint.com/sites/pagedst/SiteAssets/SitePages/zz-pagecopy-report-copy/zz-pagecopy-report.pdf',
      };
      const patchCtx = { ...analyzerContext({ web: source }), mapAsset: () => mapping };
      const copy = structuredClone(clone);
      analyzer.patch(copy, patchCtx);
      const patchedWopi = copy.webPartData.serverProcessedContent.links.wopiurl;
      const expectedWopi = '/sites/NewNerve/_layouts/15/Doc.aspx?sourcedoc=%7Bd15f0000-0000-4000-8000-0000000000a1%7D&file=zz-pagecopy-report.pdf&action=default&mobileredirect=true';

      return Boolean(wopiRef) && patchedWopi === expectedWopi && patchedWopi !== mapping.path;
    }, { control: realControl, source: realSource }));

  await check('analyzer file viewer: a file outside the source web produces no asset ref', () =>
    page.evaluate(async ({ control, source }) => {
      const { analyzerContext } = await import('/src/workbench/page-copy.js');
      const analyzer = (await import('/src/workbench/page-copy-analyzers/file-viewer.js')).default;
      const clone = structuredClone(control);
      const props = clone.webPartData.properties;
      const links = clone.webPartData.serverProcessedContent.links;
      props.file = 'https://nervedotnet.sharepoint.com/sites/OtherSite/Shared%20Documents/other-report.pdf';
      links.serverRelativeUrl = '/sites/OtherSite/Shared%20Documents/other-report.pdf';
      links.wopiurl = '/sites/OtherSite/Shared%20Documents/other-report.pdf';
      props.siteId = 'aaaa0000-0000-4000-8000-000000000009';
      props.webId = 'bbbb0000-0000-4000-8000-000000000009';
      props.listId = 'cccc0000-0000-4000-8000-000000000009';
      props.uniqueId = 'dddd0000-0000-4000-8000-000000000009';

      const ctx = analyzerContext({ web: source });
      const refs = analyzer.refs(clone, ctx);
      return refs.filter((r) => r.class === 'asset').length === 0;
    }, { control: realControl, source: realSource }));
}
