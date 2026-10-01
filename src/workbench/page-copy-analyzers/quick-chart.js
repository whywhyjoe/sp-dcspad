// Page copy analyzer — Quick chart (design/PAGE-COPY.md §5.1–5.2).
//
// Web part id 91a50c94-865f-4f5c-8b4e-e49659e69772. Two data-source shapes
// share one properties bag (real fixture: tests/pages-copy/fixtures/
// live-shapes-editor.json "quickChart" — list mode, dataSourceType: 1,
// selectedListId pointing at the zz-schema-requests list, selectedValueFieldName
// "Budget", selectedLabelFieldName "Title"):
//
//   - list mode: `selectedListId` names the source list; `selectedLabelFieldName`
//     / `selectedValueFieldName` name its columns by internal name. All three
//     keep pointing at the source list wherever the copy lands — the §5.2
//     "data → warn" verdict.
//   - manual-data mode: chart values live inline in `properties.data`
//     (label/value pairs), and `selectedListId` is empty — no reference to
//     anything on the source site, so refs() reports nothing.
//
// Detection is presence-based (a non-empty `selectedListId`) rather than
// keyed off the numeric `dataSourceType`, so it does not depend on knowing
// every enum value the web part might send. Never patched — kept as-is with
// the warning refs() already produced, like every other "data" verdict
// (list-library.js, news.js are the same shape).

function isPlainObject(v) {
  return Boolean(v) && typeof v === 'object' && !Array.isArray(v);
}

function computeRefs(instance) {
  const props = instance?.webPartData?.properties;
  if (!isPlainObject(props)) return [];
  const listId = props.selectedListId;
  if (typeof listId !== 'string' || !listId.trim()) return []; // manual-data mode: no references

  const out = [
    {
      path: ['webPartData', 'properties', 'selectedListId'],
      value: listId,
      class: 'data',
      note: 'Quick chart data source — identity id (selectedListId); keeps pointing at the source list',
    },
  ];
  for (const key of ['selectedLabelFieldName', 'selectedValueFieldName']) {
    const value = props[key];
    if (typeof value === 'string' && value) {
      out.push({
        path: ['webPartData', 'properties', key],
        value,
        class: 'data',
        note: `Quick chart data source — column name (${key}); keeps pointing at the source list's own column`,
      });
    }
  }
  return out;
}

export default {
  id: '91a50c94-865f-4f5c-8b4e-e49659e69772',
  label: 'Quick chart',

  // eslint-disable-next-line no-unused-vars -- ctx is part of the analyzer contract
  refs(instance, ctx) {
    return computeRefs(instance);
  },

  // §5.2 "data → warn": the list data source is never patched — kept as-is
  // with the warning refs() already produced.
  // eslint-disable-next-line no-unused-vars -- (instance, ctx) is part of the analyzer contract
  patch(instance, ctx) {
    return 0;
  },
};
