// Hero analyzer (design/PAGE-COPY.md §5.1-5.2, web part id c4bd7b2f).
//
// Shape (editor-authored capture, tests/pages-copy/fixtures/
// live-shapes-editor.json → heroConfigured; the sibling `hero` key is the
// manifest-default instance the page is seeded with — no `content` at all —
// and must yield no refs without crashing):
// webPartData.properties.content[] holds each tile/slide — { id, type,
// color, image: { siteId, webId, listId, id, fileName, resolvedUrl,
// imageUrl, widthFactor, minCanvasWidth, imageSize }, description,
// showDescription, showTitle, alternateText, imageDisplayOption,
// isDefaultImage, showCallToAction, previewImage: { same shape as image } }.
// `image` is the tile's own custom background — an asset (§5.2
// "asset → copied"). `previewImage` is an auto-selected thumbnail taken
// from the item the tile links to (its ids name that linked item, not a
// file the tile owns — on this capture they equal the link's own page/list
// ids) — link-dependent config, kept, never transferred.
//
// The rendered/authoritative values live in serverProcessedContent, keyed
// 'content[N].image.url' / 'content[N].previewImage.url' (imageSources) and
// 'content[N].link' (links) — the same imageSources/links/customMetadata
// split quick-links.js and header.js use. customMetadata carries the
// image's four ids, but (unlike header's camelCase customMetadata.imageSource)
// this capture's keys are lower-cased: siteid/webid/listid/uniqueid — so ids
// are read tolerant of either casing, falling back to the tile's own
// image.{siteId,webId,listId,id} when customMetadata is missing a key.
// componentDependencies (heroLayoutComponentId/carouselLayoutComponentId)
// are SPFx component ids, not SharePoint site/web/list ids — never refs.

const IMAGE_KEY = /^content\[(\d+)\]\.image\.url$/;
const PREVIEW_KEY = /^content\[(\d+)\]\.previewImage\.url$/;
const LINK_KEY = /^content\[(\d+)\]\.link$/;

// customMetadata id bag → the four normalized ids, tolerant of the
// lower-cased key spelling this capture uses.
const CUSTOM_ID_ALIASES = {
  siteId: ['siteId', 'siteid'],
  webId: ['webId', 'webid'],
  listId: ['listId', 'listid'],
  uniqueId: ['uniqueId', 'uniqueid'],
};

// content[N].image's own id bag: siteId/webId/listId direct, uniqueId under `id`.
const IMAGE_PROP_ALIASES = {
  siteId: ['siteId'],
  webId: ['webId'],
  listId: ['listId'],
  uniqueId: ['id'],
};

function decodePath(p) {
  try { return decodeURIComponent(p); } catch { return p; }
}

function underSource(value, ctx) {
  const raw = String(value ?? '');
  if (!raw) return false;
  if (raw.startsWith('/') && !raw.startsWith('//')) return ctx.underPath(decodePath(raw), ctx.sourceWebPath);
  let url;
  try { url = new URL(raw); } catch { return false; }
  if (!ctx.origin || url.origin.toLowerCase() !== String(ctx.origin).toLowerCase()) return false;
  return ctx.underPath(decodePath(url.pathname), ctx.sourceWebPath);
}

function sourceRelativePath(value) {
  const raw = String(value ?? '');
  if (raw.startsWith('/') && !raw.startsWith('//')) return decodePath(raw.split(/[?#]/)[0]);
  try { return decodeURIComponent(new URL(raw).pathname); } catch { return raw; }
}

function tileIndex(key) {
  const m = /^content\[(\d+)\]/.exec(String(key || ''));
  return m ? Number(m[1]) : null;
}

function normalizeCustomIds(meta) {
  if (!meta || typeof meta !== 'object') return null;
  return {
    siteId: meta.siteId ?? meta.siteid,
    webId: meta.webId ?? meta.webid,
    listId: meta.listId ?? meta.listid,
    uniqueId: meta.uniqueId ?? meta.uniqueid,
  };
}

function imageIds(image) {
  if (!image || typeof image !== 'object') return null;
  return { siteId: image.siteId, webId: image.webId, listId: image.listId, uniqueId: image.id };
}

// Rewrites only the id keys already present in `bag`, from a verified
// mapping — never adds a key the shape didn't already carry.
function patchIdBag(bag, mapping, ctx, aliasMap) {
  if (!bag || typeof bag !== 'object') return 0;
  let count = 0;
  for (const [idKey, aliases] of Object.entries(aliasMap)) {
    const nextId = mapping.ids[idKey];
    if (nextId === undefined) continue;
    for (const alias of aliases) {
      if (!(alias in bag)) continue;
      if (ctx.normalizeGuid(bag[alias]) !== ctx.normalizeGuid(nextId)) {
        bag[alias] = nextId;
        count += 1;
      }
    }
  }
  return count;
}

// Swap every source GUID token (dashed or bare, any case) in a string for its
// mapped destination counterpart. Only whole GUID tokens named in `pairs`
// change, so a URL that embeds file identity (the Graph v2.1 thumbnail url
// the editor stores beside an image) follows the transfer; anything else in
// it is kept.
function swapGuids(value, pairs, ctx) {
  let out = String(value);
  for (const [from, to] of pairs) {
    const f = ctx.normalizeGuid(from);
    const t = ctx.normalizeGuid(to);
    if (!f || !t || f === t) continue;
    for (const [a, b] of [[f, t], [f.replace(/-/g, ''), t.replace(/-/g, '')]]) {
      out = out.replace(new RegExp(a, 'gi'), b);
    }
  }
  return out;
}

// The editor stores two derived copies of a tile image's location beside its
// ids: `resolvedUrl` (the file's absolute url) and `imageUrl` (a thumbnail
// url embedding site/web/list/item ids). Both follow a verified mapping; a
// value that does not name the source file is left alone.
function patchDerivedUrls(image, sourceIds, mapping, ctx) {
  if (!image || typeof image !== 'object') return 0;
  let n = 0;
  if (typeof image.resolvedUrl === 'string' && underSource(image.resolvedUrl, ctx)) {
    const next = /^https?:\/\//i.test(image.resolvedUrl) && mapping.url ? mapping.url : mapping.path;
    if (next && next !== image.resolvedUrl) { image.resolvedUrl = next; n += 1; }
  }
  if (typeof image.imageUrl === 'string' && sourceIds?.uniqueId
      && image.imageUrl.toLowerCase().includes(ctx.normalizeGuid(sourceIds.uniqueId))) {
    const pairs = ['siteId', 'webId', 'listId', 'uniqueId'].map((k) => [sourceIds[k], mapping.ids[k]]);
    const next = swapGuids(image.imageUrl, pairs, ctx);
    if (next !== image.imageUrl) { image.imageUrl = next; n += 1; }
  }
  return n;
}

export default {
  id: 'c4bd7b2f-7b6e-4599-8485-16504575f590',
  label: 'Hero',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const content = Array.isArray(props.content) ? props.content : [];
    const imageSources = spc.imageSources || {};
    const links = spc.links || {};
    const customMetadata = spc.customMetadata || {};
    const refs = [];

    for (const [key, value] of Object.entries(imageSources)) {
      if (!underSource(value, ctx)) continue;
      const idx = tileIndex(key);
      const item = idx != null ? content[idx] : undefined;

      if (IMAGE_KEY.test(key)) {
        const ids = normalizeCustomIds(customMetadata[key]) || imageIds(item?.image) || {};
        refs.push({
          path: ['webPartData', 'serverProcessedContent', 'imageSources', key],
          value,
          class: 'asset',
          asset: { path: sourceRelativePath(value), ids },
        });
      } else if (PREVIEW_KEY.test(key)) {
        // Not a file the tile owns — an auto-selected preview taken from the
        // linked item, so it rides on the tile's own link (§5.2 "link"),
        // kept and reported, never transferred as an asset.
        refs.push({
          path: ['webPartData', 'serverProcessedContent', 'imageSources', key],
          value,
          class: 'config',
          note: 'auto-selected preview image taken from the linked item, not a file the tile owns',
        });
      }
    }

    for (const [key, value] of Object.entries(links)) {
      if (!LINK_KEY.test(key)) continue;
      if (!underSource(value, ctx)) continue;
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'links', key],
        value,
        class: 'link',
      });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    if (!props || !spc) return 0;
    const content = Array.isArray(props.content) ? props.content : [];
    const imageSources = spc.imageSources;
    const links = spc.links;
    const customMetadata = spc.customMetadata;
    let count = 0;

    if (imageSources) {
      for (const key of Object.keys(imageSources)) {
        // previewImage.url is link-dependent config — never patched here.
        if (!IMAGE_KEY.test(key)) continue;
        const value = imageSources[key];
        if (!underSource(value, ctx)) continue;
        const idx = tileIndex(key);
        const item = idx != null ? content[idx] : undefined;
        const meta = customMetadata?.[key];
        const ids = normalizeCustomIds(meta) || imageIds(item?.image) || {};
        const mapping = ctx.mapAsset({ path: sourceRelativePath(value), ids });
        // Path and ids move together or not at all.
        if (!mapping || mapping.path === undefined || !mapping.ids) continue;

        const nextPath = /^https?:\/\//i.test(String(imageSources[key])) && mapping.url
          ? mapping.url : mapping.path;
        if (imageSources[key] !== nextPath) {
          imageSources[key] = nextPath;
          count += 1;
        }
        const sourceIds = { ...ids };
        count += patchIdBag(meta, mapping, ctx, CUSTOM_ID_ALIASES);
        count += patchIdBag(item?.image, mapping, ctx, IMAGE_PROP_ALIASES);
        count += patchDerivedUrls(item?.image, sourceIds, mapping, ctx);
      }
    }

    if (links) {
      for (const key of Object.keys(links)) {
        if (!LINK_KEY.test(key)) continue;
        const value = links[key];
        if (!underSource(value, ctx)) continue;
        const mapped = ctx.mapLink(value);
        if (mapped != null && mapped !== value) {
          links[key] = mapped;
          count += 1;
        }
      }
    }

    return count;
  },
};
