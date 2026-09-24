// Image gallery analyzer (design/PAGE-COPY.md §5.1-5.2, web part id
// af8be689-990e-492a-81f7-ba3e4cd3ed9c).
//
// Shape observed (real editor-authored capture,
// tests/pages-copy/fixtures/live-shapes-editor.json → imageGallery — the
// images were COPIED into SiteAssets/SitePages/zz-pagecopy-shapes/ on the
// source page): webPartData.properties.images[] holds each tile —
// { description, siteId, webId, listId, id (the file's UniqueId), isInList,
// altText, imageTileWidth }. The rendered thumbnail lives in
// serverProcessedContent.imageSources, keyed 'images[N].url' — an asset
// (§5.2 "asset → copied"). serverProcessedContent.customMetadata carries the
// same four ids again per image, keyed the same way, but with LOWERCASE
// field names (siteid/webid/listid/uniqueid/fixedwidth) — SharePoint's own
// casing on this web part, distinct from the header/quick-links camelCase
// customMetadata and never renamed here.
//
// Asset identity = the per-image `id` (the file's UniqueId, from
// properties.images[N]) + the server-relative path from imageSources — not
// customMetadata, which is a mirror kept in step by patch().
//
// properties.imageSourceType selects where the gallery's images come from:
// 1 is the per-image list captured above (each image individually chosen
// and carrying its own ids/path, handled entirely by the asset refs above).
// Any other value is a library/folder source mode — the gallery pulls
// whatever images live under a list/folder at render time, so there is
// nothing here for an analyzer to transfer or patch; it is `data → warn`
// (§5.2), kept pointing at the source and reported, exactly like
// list-library.js's list/library identity refs. No real capture of that
// mode exists yet, so this only fires when imageSourceType is present and
// not the per-image-list value — never invented from other properties.

const PER_IMAGE_LIST_SOURCE_TYPE = 1;

function decodePath(p) {
  try { return decodeURIComponent(p); } catch { return p; }
}

function splitPath(raw) {
  const m = /^([^?#]*)/.exec(raw);
  return m ? m[1] : raw;
}

// Server-relative under ctx.sourceWebPath, or absolute on ctx.origin with a
// path under it — otherwise not a reference at all.
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
  if (raw.startsWith('/') && !raw.startsWith('//')) return decodePath(splitPath(raw));
  try { return decodePath(splitPath(new URL(raw).pathname)); } catch { return raw; }
}

const IMAGE_URL_KEY = /^images\[(\d+)\]\.url$/;

function imageIndex(key) {
  const m = IMAGE_URL_KEY.exec(String(key || ''));
  return m ? Number(m[1]) : null;
}

function imageIds(item) {
  return {
    siteId: item?.siteId,
    webId: item?.webId,
    listId: item?.listId,
    uniqueId: item?.id,
  };
}

// customMetadata's own field names for this web part, lowercase — never
// treated as the camelCase keys the header/quick-links analyzers use.
const CUSTOM_META_ID_KEYS = [
  ['siteId', 'siteid'],
  ['webId', 'webid'],
  ['listId', 'listid'],
  ['uniqueId', 'uniqueid'],
];

export default {
  id: 'af8be689-990e-492a-81f7-ba3e4cd3ed9c',
  label: 'Image gallery',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const images = Array.isArray(props.images) ? props.images : [];
    const imageSources = spc.imageSources || {};
    const refs = [];

    for (const [key, value] of Object.entries(imageSources)) {
      if (imageIndex(key) == null) continue;
      if (!underSource(value, ctx)) continue;
      const idx = imageIndex(key);
      const item = images[idx];
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', key],
        value,
        class: 'asset',
        asset: { path: sourceRelativePath(value), ids: imageIds(item) },
      });
    }

    const sourceType = props.imageSourceType;
    if (sourceType !== undefined && sourceType !== PER_IMAGE_LIST_SOURCE_TYPE) {
      refs.push({
        path: ['webPartData', 'properties', 'imageSourceType'],
        value: sourceType,
        class: 'data',
        note: 'library/folder source — images resolve at render time against whatever list/folder this names, not the per-image list',
      });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    const imageSources = spc?.imageSources;
    if (!props || !spc || !imageSources) return 0;
    const images = Array.isArray(props.images) ? props.images : [];
    const customMetadata = spc.customMetadata;
    let count = 0;

    for (const key of Object.keys(imageSources)) {
      const idx = imageIndex(key);
      if (idx == null) continue;
      const value = imageSources[key];
      if (!underSource(value, ctx)) continue;

      const item = images[idx];
      const mapping = ctx.mapAsset({ path: sourceRelativePath(value), ids: imageIds(item) });
      // Path and ids move together or not at all — half a mapping would
      // leave this image's path and its id quadruple naming different files.
      if (!mapping || mapping.path === undefined || !mapping.ids) continue;

      const nextPath = /^https?:\/\//i.test(String(imageSources[key])) && mapping.url
        ? mapping.url : mapping.path;
      if (imageSources[key] !== nextPath) {
        imageSources[key] = nextPath;
        count += 1;
      }

      const cm = customMetadata && typeof customMetadata === 'object' ? customMetadata[key] : undefined;
      if (cm && typeof cm === 'object') {
        for (const [idKey, cmKey] of CUSTOM_META_ID_KEYS) {
          if (!(cmKey in cm)) continue;
          const nextId = mapping.ids[idKey];
          if (nextId !== undefined && ctx.normalizeGuid(cm[cmKey]) !== ctx.normalizeGuid(nextId)) {
            cm[cmKey] = nextId;
            count += 1;
          }
        }
      }

      if (item) {
        for (const idKey of ['siteId', 'webId', 'listId']) {
          if (!(idKey in item)) continue;
          const nextId = mapping.ids[idKey];
          if (nextId !== undefined && ctx.normalizeGuid(item[idKey]) !== ctx.normalizeGuid(nextId)) {
            item[idKey] = nextId;
            count += 1;
          }
        }
        // item.id carries the uniqueId (§ above) — the file's UniqueId, not
        // a distinct "id" concept.
        if ('id' in item) {
          const nextId = mapping.ids.uniqueId;
          if (nextId !== undefined && ctx.normalizeGuid(item.id) !== ctx.normalizeGuid(nextId)) {
            item.id = nextId;
            count += 1;
          }
        }
      }
    }

    // properties.imageSourceType (library/folder source mode) is data —
    // left unchanged; layout/grid/carousel settings, descriptions and
    // altText are left completely untouched.
    return count;
  },
};
