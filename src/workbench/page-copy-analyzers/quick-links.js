// Quick links analyzer (design/PAGE-COPY.md §5.1-5.2, web part id c70391ea).
//
// Shapes observed (§11 Q7, the real editor-authored capture with a custom
// thumbnail — tests/pages-copy/fixtures/live-shapes-editor.json →
// quickLinksCustomThumb — and the mock fixture, src/workbench/
// mock-pagecopy.js 'c70391ea'): webPartData.properties.items[] holds each
// tile. The real capture's item carries { sourceItem: { guids: { siteId,
// webId, listId, uniqueId }, itemType, fileExtension, progId },
// thumbnailType, id, description, image: { guids: { siteId, webId, listId,
// uniqueId }, imageFit, minCanvasWidth }, altText,
// rawPreviewImageMinCanvasWidth } — the thumbnail's own identity lives
// under item.image.guids, never flat on the item; sourceItem.guids names
// the link target instead and is never touched here (§5.2 "link" only
// rewrites the URL, never the guids next to it — the same split the header
// and image analyzers keep between a part's own asset and a part's own
// link). The rendered thumbnail/icon lives in serverProcessedContent.
// imageSources, keyed 'items[N].image.url' — an asset (§5.2 "asset →
// copied"), whose identity mirrors onto serverProcessedContent.
// customMetadata[same key] (the real capture's key casing is lower-case:
// siteid/webid/listid/uniqueid, plus mincanvaswidth/fixedwidth that are
// never touched) — the same "primary customMetadata, properties/item
// fallback" split image.js uses, so itemIds() below resolves whichever
// location actually carries the ids rather than assuming one shape:
// customMetadata first, then item.image.guids, then a flat item.{siteId,
// webId,listId,uniqueId} (the hand-built mock's shape, kept working so the
// mock fixture never needs to model the real nesting). The link each tile
// points at lives in serverProcessedContent.links, keyed
// 'items[N].sourceItem.url' — a link into the source web when it resolves
// there (§5.2 "link"), never a ref when it points elsewhere (e.g. the stock
// go.microsoft.com fwlinks on out-of-the-box pages). links.baseUrl names the
// web the tiles' relative links are resolved against — config, never
// patched unless the operator opts into link rewriting (rewriteContent
// passes ctx.rewriteLinks only then). The part-level properties.siteId/webId
// name the source web the control was authored in — config, left alone
// (patching would invent a destination id nothing verified).

const ITEM_LINK_KEY = /^items\[(\d+)\]\.sourceItem\.url$/;

function underSource(value, ctx) {
  const raw = String(value ?? '');
  if (!raw) return false;
  if (raw.startsWith('/') && !raw.startsWith('//')) return ctx.underPath(raw, ctx.sourceWebPath);
  let url;
  try { url = new URL(raw); } catch { return false; }
  if (!ctx.origin || url.origin.toLowerCase() !== String(ctx.origin).toLowerCase()) return false;
  return ctx.underPath(decodePath(url.pathname), ctx.sourceWebPath);
}

function decodePath(p) { try { return decodeURIComponent(p); } catch { return p; } }

function sourceRelativePath(value) {
  const raw = String(value ?? '');
  if (raw.startsWith('/') && !raw.startsWith('//')) return decodePath(raw.split(/[?#]/)[0]);
  try { return decodeURIComponent(new URL(raw).pathname); } catch { return raw; }
}

function itemIndex(key) {
  const m = /^items\[(\d+)\]/.exec(String(key || ''));
  return m ? Number(m[1]) : null;
}

// Resolves the canonical id keys to whichever key names them in a
// case-insensitive object (customMetadata's real capture uses lower-case
// siteid/webid/listid/uniqueid; item.image.guids and the mock's flat item
// use camelCase) — never assumed, the same convention image.js's
// customImageIds() uses.
function idKeyMap(obj) {
  const keyMap = {};
  if (obj && typeof obj === 'object') {
    for (const actualKey of Object.keys(obj)) {
      const lower = actualKey.toLowerCase();
      if (lower === 'siteid') keyMap.siteId = actualKey;
      else if (lower === 'webid') keyMap.webId = actualKey;
      else if (lower === 'listid') keyMap.listId = actualKey;
      else if (lower === 'uniqueid') keyMap.uniqueId = actualKey;
    }
  }
  return keyMap;
}

function readIds(obj) {
  const keyMap = idKeyMap(obj);
  return {
    siteId: keyMap.siteId ? obj[keyMap.siteId] : undefined,
    webId: keyMap.webId ? obj[keyMap.webId] : undefined,
    listId: keyMap.listId ? obj[keyMap.listId] : undefined,
    uniqueId: keyMap.uniqueId ? obj[keyMap.uniqueId] : undefined,
  };
}

function hasAnyId(ids) {
  return Boolean(ids) && (ids.siteId !== undefined || ids.webId !== undefined || ids.listId !== undefined || ids.uniqueId !== undefined);
}

// The thumbnail's own identity, first location that actually carries it
// wins (never merged, so a stale mirror can't blend with a fresh one):
// serverProcessedContent.customMetadata[same imageSources key] is primary,
// item.image.guids is the real capture's own mirror, and a flat
// item.{siteId,webId,listId,uniqueId} is the hand-built mock's shape.
function itemIds(customEntry, item) {
  const fromCustom = readIds(customEntry);
  if (hasAnyId(fromCustom)) return fromCustom;
  const fromImageGuids = readIds(item?.image?.guids);
  if (hasAnyId(fromImageGuids)) return fromImageGuids;
  return readIds(item);
}

// Writes only the ids a mapping supplies into whichever keys `obj` actually
// has (never inventing a field) — returns how many values actually changed.
// Compares through normalizeGuid so a differently-formatted-but-equal guid
// (braces, case) is never counted as a change.
function writeIds(obj, ids, normalizeGuid) {
  const keyMap = idKeyMap(obj);
  let count = 0;
  for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
    const actualKey = keyMap[idKey];
    if (!actualKey) continue;
    const nextId = ids?.[idKey];
    if (nextId !== undefined && normalizeGuid(obj[actualKey]) !== normalizeGuid(nextId)) {
      obj[actualKey] = nextId;
      count += 1;
    }
  }
  return count;
}

export default {
  id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
  label: 'Quick links',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const items = Array.isArray(props.items) ? props.items : [];
    const imageSources = spc.imageSources || {};
    const links = spc.links || {};
    const refs = [];

    for (const [key, value] of Object.entries(imageSources)) {
      if (!underSource(value, ctx)) continue;
      const idx = itemIndex(key);
      const item = idx != null ? items[idx] : undefined;
      const customEntry = spc.customMetadata?.[key];
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', key],
        value,
        class: 'asset',
        asset: { path: sourceRelativePath(value), ids: itemIds(customEntry, item) },
      });
    }

    if (links.baseUrl && underSource(links.baseUrl, ctx)) {
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'links', 'baseUrl'],
        value: links.baseUrl,
        class: 'config',
        note: 'names the source web; the tiles\' relative links resolve against it',
      });
    }

    for (const [key, value] of Object.entries(links)) {
      if (key === 'baseUrl' || !ITEM_LINK_KEY.test(key)) continue;
      if (!underSource(value, ctx)) continue; // external links are not refs
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'links', key],
        value,
        class: 'link',
      });
    }

    for (const idKey of ['siteId', 'webId']) {
      const value = props[idKey];
      if (value && ctx.normalizeGuid(value) === ctx.sourceIds[idKey]) {
        refs.push({
          path: ['webPartData', 'properties', idKey],
          value,
          class: 'config',
          note: 'names the source site/web this control was authored in',
        });
      }
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    if (!props || !spc) return 0;
    const items = Array.isArray(props.items) ? props.items : [];
    const imageSources = spc.imageSources;
    const links = spc.links;
    let count = 0;

    if (imageSources) {
      for (const key of Object.keys(imageSources)) {
        const value = imageSources[key];
        if (!underSource(value, ctx)) continue;
        const idx = itemIndex(key);
        const item = idx != null ? items[idx] : undefined;
        const customEntry = spc.customMetadata?.[key];
        const mapping = ctx.mapAsset({ path: sourceRelativePath(value), ids: itemIds(customEntry, item) });
        if (!mapping) continue;
        const isAbsolute = /^https?:\/\//i.test(String(value));
        const next = isAbsolute && mapping.url ? mapping.url : mapping.path;
        if (next !== undefined && next !== value) {
          imageSources[key] = next;
          count += 1;
        }
        // Live shape (§11 Q7): the editor keeps the picked thumbnail's
        // absolute url in the part-level `imagePicker` too. It follows only
        // when it names this very file.
        if (typeof props.imagePicker === 'string' && underSource(props.imagePicker, ctx)
            && sourceRelativePath(props.imagePicker).toLowerCase() === sourceRelativePath(value).toLowerCase()) {
          const picked = /^https?:\/\//i.test(props.imagePicker) && mapping.url ? mapping.url : mapping.path;
          if (picked !== undefined && picked !== props.imagePicker) { props.imagePicker = picked; count += 1; }
        }
        // The thumbnail's identity mirrors across up to three places — only
        // whichever of them actually exist on this instance get written.
        if (mapping.ids) {
          count += writeIds(customEntry, mapping.ids, ctx.normalizeGuid);
          if (item) {
            count += writeIds(item.image?.guids, mapping.ids, ctx.normalizeGuid);
            count += writeIds(item, mapping.ids, ctx.normalizeGuid);
          }
        }
      }
    }

    if (links) {
      for (const key of Object.keys(links)) {
        if (key === 'baseUrl' || !ITEM_LINK_KEY.test(key)) continue;
        const value = links[key];
        if (!underSource(value, ctx)) continue;
        const mapped = ctx.mapLink(value);
        if (mapped != null && mapped !== value) {
          links[key] = mapped;
          count += 1;
        }
      }
      // baseUrl re-bases the tile links, so it follows them — but only when
      // it verifiably names the source web (the same test refs() uses).
      if (ctx.rewriteLinks && links.baseUrl && underSource(links.baseUrl, ctx)) {
        const nextBase = ctx.target?.webPath;
        if (nextBase !== undefined && links.baseUrl !== nextBase) {
          links.baseUrl = nextBase;
          count += 1;
        }
      }
    }

    // Part-level properties.siteId/webId are config — left unchanged; no
    // destination id has been verified for them.
    return count;
  },
};
