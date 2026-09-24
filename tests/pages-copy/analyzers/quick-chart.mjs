// SP Workbench — page copy: the Quick chart analyzer (design/PAGE-COPY.md
// §5.1–5.2). Every check imports
// src/workbench/page-copy-analyzers/quick-chart.js in the page context and
// exercises it directly against the real captured shape (fixtures/
// live-shapes-editor.json "quickChart") and a manual-data variant — no mock
// webs, no dialog. House style: tests/pages-copy/analyzers/list-library.mjs
// (module shape, baseCtx).

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
  // ---- refs: real fixture — list mode ----------------------------------------

  await check('analyzer quick chart: refs on the real fixture (list mode) reports the list id and both column names as data, nothing as asset or link', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/quick-chart.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };
      const refs = analyzer.refs(instance, liveCtx);

      const listId = refs.find((r) => r.path.at(-1) === 'selectedListId');
      const labelField = refs.find((r) => r.path.at(-1) === 'selectedLabelFieldName');
      const valueField = refs.find((r) => r.path.at(-1) === 'selectedValueFieldName');

      const allData = refs.length > 0 && refs.every((r) => r.class === 'data');
      return allData && refs.length === 3
        && Boolean(listId) && listId.value === instance.webPartData.properties.selectedListId
        && Boolean(labelField) && labelField.value === 'Title'
        && Boolean(valueField) && valueField.value === 'Budget';
    }, { instance: FIXTURE.quickChart, ctx: baseCtx() }));

  // ---- refs: manual-data mode carries no references ---------------------------

  await check('analyzer quick chart: refs on a manual-data variant reports nothing', () =>
    page.evaluate(async ({ ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/quick-chart.js');
      const { underPath } = await import('/src/workbench/page-copy.js');
      const liveCtx = { ...ctx, underPath };

      const instance = {
        controlType: 3,
        webPartId: '91a50c94-865f-4f5c-8b4e-e49659e69772',
        webPartData: {
          id: '91a50c94-865f-4f5c-8b4e-e49659e69772',
          properties: {
            data: [
              { id: '7CFFD4B0-436E-430D-94C5-A4F9D22DB3FE', label: 'Q1', value: '10', valueNumber: 10 },
            ],
            type: 0,
            isInitialState: false,
            dataSourceType: 0,
            listItemOrderBy: 0,
            selectedListId: '',
            selectedLabelFieldName: '',
            selectedValueFieldName: '',
            xAxisLabel: '',
            yAxisLabel: '',
          },
        },
      };
      const refs = analyzer.refs(instance, liveCtx);
      return refs.length === 0;
    }, { ctx: baseCtx() }));

  // ---- patch: never patches data -----------------------------------------------

  await check('analyzer quick chart: patch returns 0 and leaves the instance deep-equal', () =>
    page.evaluate(async ({ instance, ctx }) => {
      const { default: analyzer } = await import('/src/workbench/page-copy-analyzers/quick-chart.js');
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
    }, { instance: FIXTURE.quickChart, ctx: baseCtx() }));
}
