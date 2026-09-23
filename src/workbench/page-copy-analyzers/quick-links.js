// Quick links analyzer (design/PAGE-COPY.md §5.1-5.2, web part id c70391ea).
//
// Shapes observed (§11 Q7 + the mock fixture, src/workbench/mock-pagecopy.js
// 'c70391ea'): webPartData.properties.items[] holds each tile — an
// editor-authored page's items carry only { id, description, altText,
// thumbnailType, sourceItem: { itemType, fileExtension, progId } }; a
// custom-thumbnail item (thumbnailType 3) additionally carries the file's
// own { siteId, webId, listId, uniqueId } on the item itself. The rendered
// thumbnail/icon lives in serverProcessedContent.imageSources, keyed
// 'items[N].image.url' — an asset (§5.2 "asset → copied"). The link each
// tile points at lives in serverProcessedContent.links, keyed
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

function itemIds(item) {
  return {
    siteId: item?.siteId,
    webId: item?.webId,
    listId: item?.listId,
    uniqueId: item?.uniqueId,
  };
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
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', key],
        value,
        class: 'asset',
        asset: { path: sourceRelativePath(value), ids: itemIds(item) },
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
        const mapping = ctx.mapAsset({ path: sourceRelativePath(value), ids: itemIds(item) });
        if (!mapping) continue;
        const isAbsolute = /^https?:\/\//i.test(String(value));
        const next = isAbsolute && mapping.url ? mapping.url : mapping.path;
        if (next !== undefined && next !== value) {
          imageSources[key] = next;
          count += 1;
        }
        if (item && mapping.ids) {
          for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
            const nextId = mapping.ids[idKey];
            if (nextId !== undefined && ctx.normalizeGuid(item[idKey]) !== ctx.normalizeGuid(nextId)) {
              item[idKey] = nextId;
              count += 1;
            }
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
