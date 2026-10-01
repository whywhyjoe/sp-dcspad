// Page header (title area) analyzer — the layout part (design/PAGE-COPY.md
// §0 "Banner metadata", §5.3, §5.4, §11). Selected by page-copy-analyzers.js
// via the `header: true` flag, not by id lookup: analyzeParts() and
// rewriteContent() in page-copy.js call it directly for every
// LayoutWebpartsContent entry (there is exactly one on a modern page, id
// cbe7b0a9-3504-44dd-a3a3-0e5cacd07788) — `part` here is that RAW layout
// entry, never the parsed canvas control shape web-part analyzers see.
//
// Shape (live capture, tests/pages-copy/fixtures/live-shapes.json →
// titleArea, and the dev-tenant q7 spike capture, both banner-less and
// with a custom banner set):
//   { id, instanceId, title, description,
//     serverProcessedContent: { htmlStrings, searchablePlainTexts,
//       imageSources: { imageSource? }, links: {},
//       customMetadata?: { imageSource?: { siteId, webId, listId, uniqueId } } },
//     properties: { title, imageSourceType, layoutType, textAlignment,
//       showTopicHeader, showPublishDate, topicHeader, authorByline, authors,
//       // only present once a banner has been set (PnP-pages 340-353):
//       siteId?, webId?, listId?, uniqueId? } }
//
// No banner: imageSourceType 4, serverProcessedContent.imageSources empty —
// no asset ref at all. A banner: imageSourceType 2, imageSources.imageSource
// is the source-relative path PnPjs's save() wrote there (§0), customMetadata
// .imageSource carries the file's four ids (PnP-pages 344-349), and the same
// four ids are mirrored onto properties (PnP-pages 350-353) — customMetadata
// is the primary source, properties the fallback when it is missing.
//
// properties.authorByline / properties.authors name people, who resolve per
// tenant and are never rewritten — classified 'config' (kept) only when
// actually populated, so an empty byline on nearly every page does not
// clutter the report. properties.topicHeader is derived from the DTO's
// TopicHeader (§5.4: DTO authoritative, the header part's copy is derived) —
// deliberately no ref for it here.
//
// The DTO's BannerImageUrl (a thumbnail, possibly a different file entirely
// from this header's banner — §11) is mapped separately by
// page-copy.js's mapBannerUrl(); this analyzer only ever touches the header
// part's own banner.

function decodePath(p) {
  try { return decodeURIComponent(p); } catch { return p; }
}

function splitPath(raw) {
  const m = /^([^?#]*)/.exec(raw);
  return m ? m[1] : raw;
}

// Server-relative under ctx.sourceWebPath, or absolute on ctx.origin with a
// path under it — otherwise not a reference at all (an out-of-tenant or
// other-web banner is left alone, never patched).
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

// customMetadata.imageSource is the primary identity; properties.{siteId,
// webId,listId,uniqueId} (mirrored there by PnPjs's save(), PnP-pages
// 350-353) is the fallback when it is missing.
function bannerIds(spc, props) {
  const custom = spc?.customMetadata?.imageSource;
  if (custom && typeof custom === 'object') {
    return { siteId: custom.siteId, webId: custom.webId, listId: custom.listId, uniqueId: custom.uniqueId };
  }
  return { siteId: props?.siteId, webId: props?.webId, listId: props?.listId, uniqueId: props?.uniqueId };
}

export default {
  id: 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788',
  header: true,
  label: 'Title area',

  refs(part, ctx) {
    const props = part?.properties || {};
    const spc = part?.serverProcessedContent || {};
    const imageSources = spc.imageSources || {};
    const refs = [];

    const bannerValue = imageSources.imageSource;
    if (bannerValue && underSource(bannerValue, ctx)) {
      refs.push({
        path: ['serverProcessedContent', 'imageSources', 'imageSource'],
        value: bannerValue,
        class: 'asset',
        asset: { path: sourceRelativePath(bannerValue), ids: bannerIds(spc, props) },
      });
    }

    if (Array.isArray(props.authorByline) && props.authorByline.length) {
      refs.push({
        path: ['properties', 'authorByline'],
        value: props.authorByline,
        class: 'config',
        note: 'authors resolve per tenant',
      });
    }
    if (Array.isArray(props.authors) && props.authors.length) {
      refs.push({
        path: ['properties', 'authors'],
        value: props.authors,
        class: 'config',
        note: 'authors resolve per tenant',
      });
    }

    // properties.topicHeader is derived from the DTO's TopicHeader (§5.4) —
    // deliberately not a reference here.
    return refs;
  },

  patch(part, ctx) {
    const props = part?.properties;
    const spc = part?.serverProcessedContent;
    const imageSources = spc?.imageSources;
    if (!props || !spc || !imageSources) return 0;

    const bannerValue = imageSources.imageSource;
    if (!bannerValue || !underSource(bannerValue, ctx)) return 0;

    const mapping = ctx.mapAsset({ path: sourceRelativePath(bannerValue), ids: bannerIds(spc, props) });
    // Path and ids move together or not at all — half a mapping would leave
    // the banner path and its id quadruple naming different files.
    if (!mapping || mapping.path === undefined || !mapping.ids) return 0;

    let count = 0;

    if (mapping.path !== undefined && imageSources.imageSource !== mapping.path) {
      imageSources.imageSource = mapping.path;
      count += 1;
    }

    if (mapping.ids) {
      const customImageSource = spc.customMetadata?.imageSource;
      if (customImageSource && typeof customImageSource === 'object') {
        for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
          if (!(idKey in customImageSource)) continue;
          const nextId = mapping.ids[idKey];
          if (nextId !== undefined && customImageSource[idKey] !== nextId) {
            customImageSource[idKey] = nextId;
            count += 1;
          }
        }
      }
      for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
        if (!(idKey in props)) continue;
        const nextId = mapping.ids[idKey];
        if (nextId !== undefined && props[idKey] !== nextId) {
          props[idKey] = nextId;
          count += 1;
        }
      }
    }

    // Everything else — translateX/Y, altText, layoutType, textAlignment,
    // showTopicHeader/showPublishDate, enableGradientEffect, title... — is
    // left completely untouched.
    return count;
  },
};
