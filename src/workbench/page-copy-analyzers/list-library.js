// Page copy analyzer — List / Document library (design/PAGE-COPY.md §5.1–5.2).
//
// Web part id f92bf067-bc19-489e-a556-7fe95f508720 serves BOTH the List and
// the Document library web parts; the same properties shape carries a list
// or a library identity depending on `isDocumentLibrary` (real fixture:
// tests/pages-copy/fixtures/live-shapes.json "listOrLibrary" — a Document
// library instance with `isDocumentLibrary: true`, `selectedListId`,
// `selectedListUrl`, `webRelativeListUrl`, `selectedViewId`). This is the
// §5.2 "data → warn" verdict: the list/library data source is kept as-is —
// it still points at the source list, or (for a web-relative config such as
// `webRelativeListUrl`) may re-resolve against the destination web instead —
// never patched. A Document library instance of this same id can ALSO be a
// dynamic-data provider for another part on the page; that connection is
// discovered generically by page-copy.js's own dynamicDataPaths handling
// (analyzeParts's `consider()`), not by this analyzer.

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function isPlainObject(v) {
  return Boolean(v) && typeof v === 'object' && !Array.isArray(v);
}

// Label from the instance's own properties when it says so (an explicit
// document-library flag); otherwise the shared fallback label, since this
// one web part id serves both controls (design/PAGE-COPY.md §5.1).
function kindLabel(instance) {
  const props = instance?.webPartData?.properties;
  if (isPlainObject(props)) {
    if (props.isDocumentLibrary === true) return 'Document library';
    if (props.isDocumentLibrary === false) return 'List';
  }
  return 'List / Document library';
}

// The list's own display name for the note, preferring the rendered title
// SharePoint already computed (serverProcessedContent.searchablePlainTexts.
// listTitle), then the last path segment of whichever URL property is
// present, then the generic kind label.
function listName(instance) {
  const props = instance?.webPartData?.properties || {};
  const spc = instance?.webPartData?.serverProcessedContent;
  const fromSpc = isPlainObject(spc) ? spc.searchablePlainTexts?.listTitle : undefined;
  if (typeof fromSpc === 'string' && fromSpc.trim()) return fromSpc.trim();
  const url = props.selectedListUrl || props.webRelativeListUrl || '';
  const seg = String(url).replace(/\/+$/, '').split('/').filter(Boolean).pop();
  return seg || kindLabel(instance);
}

// A `*Id`-named property (selectedListId, selectedViewId, listId, webId,
// siteId…) is a server-generated identity: it always names one specific
// object on the source, regardless of what URL got it there.
function isIdentityIdKey(key) {
  return /[a-z]id$/i.test(key) && !/^instanceid$/i.test(key);
}

// A `*ListUrl`-named property (selectedListUrl, webRelativeListUrl) or
// `siteUrl` names where the list/library lives. `webRelative…` is relative
// to whichever web renders the page — it may re-resolve on the destination;
// everything else is an absolute path/URL that keeps pointing at the source.
function isListUrlKey(key) {
  return /listurl$/i.test(key) || /^siteurl$/i.test(key);
}

function collectRefs(container, path, name, out) {
  if (!isPlainObject(container)) return;
  for (const [key, value] of Object.entries(container)) {
    if (value == null) continue;
    if (typeof value === 'string' && isIdentityIdKey(key)) {
      if (!GUID_RE.test(value)) continue;
      out.push({
        path: [...path, key],
        value,
        class: 'data',
        note: `${name} — identity id (${key}); always names this exact source list/view, wherever the copy lands`,
      });
      continue;
    }
    if (typeof value === 'string' && value !== '' && isListUrlKey(key)) {
      const relative = /webrelative/i.test(key);
      out.push({
        path: [...path, key],
        value,
        class: 'data',
        note: relative
          ? `${name} — web-relative (${key}); may re-resolve to a same-named list on the destination web, or break if there is none`
          : `${name} — absolute (${key}); keeps pointing at the source list`,
      });
      continue;
    }
    if (isPlainObject(value)) collectRefs(value, [...path, key], name, out);
  }
}

export default {
  id: 'f92bf067-bc19-489e-a556-7fe95f508720',
  label: 'List / Document library',

  // eslint-disable-next-line no-unused-vars -- ctx is part of the analyzer contract
  refs(instance, ctx) {
    const props = instance?.webPartData?.properties;
    if (!isPlainObject(props)) return [];
    const name = listName(instance);
    const out = [];
    collectRefs(props, ['webPartData', 'properties'], name, out);
    const spc = instance?.webPartData?.serverProcessedContent;
    if (isPlainObject(spc) && isPlainObject(spc.customMetadata)) {
      collectRefs(spc.customMetadata, ['webPartData', 'serverProcessedContent', 'customMetadata'], name, out);
    }
    return out;
  },

  // §5.2 "data → warn": the list/library data source is never patched —
  // kept as-is with the warning `refs()` already produced.
  // eslint-disable-next-line no-unused-vars -- (instance, ctx) is part of the analyzer contract
  patch(instance, ctx) {
    return 0;
  },
};
