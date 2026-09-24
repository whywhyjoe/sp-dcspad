// Image web part analyzer (design/PAGE-COPY.md §5.2, id d1d91016-032f-456d-98a4-721247c305e8).
//
// Shape (live editor-authored capture, tests/pages-copy/fixtures/
// live-shapes-editor.json → image): webPartData.properties carries
// { imageSourceType, captionText, altText, linkUrl, overlayText, fileName,
// siteId, webId, listId, uniqueId, ... }; webPartData.serverProcessedContent
// .imageSources.imageSource is the rendered image's source-relative path —
// the asset ref (§5.2 "asset → copied"), whose identity mirrors onto BOTH
// serverProcessedContent.customMetadata.imageSource (the capture's real key
// casing is lower-case: siteid/webid/listid/uniqueid, plus width/height that
// are never touched) and properties.{siteId,webId,listId,uniqueId} — the
// same "primary customMetadata, properties fallback" split the header
// analyzer uses for its banner, so customImageIds() below resolves the
// actual key names present rather than assuming a case.
//
// properties.linkUrl, when set and resolving into the source web, is the
// tile's own outbound link (§5.2 "link") — kept by default, rewritten only
// when the operator opts into link rewriting (ctx.mapLink). An empty
// linkUrl (the common case — no link configured) yields no ref at all.
//
// Stock/CDN images (served from cdn.hubblecontent.osi.office.net, never the
// source web) fail the origin check in underSource() and are correctly
// reported as having no asset ref — nothing under the source web resolves,
// so there is nothing to transfer or patch.

function decodePath(p) { try { return decodeURIComponent(p); } catch { return p; } }

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

// Maps the canonical id names to whatever key actually names them in
// customMetadata.imageSource (the real capture uses siteid/webid/listid/
// uniqueid, all lower-case) — case-insensitive match, never assumed.
function customImageIds(customImageSource) {
  const keyMap = {};
  if (customImageSource && typeof customImageSource === 'object') {
    for (const actualKey of Object.keys(customImageSource)) {
      const lower = actualKey.toLowerCase();
      if (lower === 'siteid') keyMap.siteId = actualKey;
      else if (lower === 'webid') keyMap.webId = actualKey;
      else if (lower === 'listid') keyMap.listId = actualKey;
      else if (lower === 'uniqueid') keyMap.uniqueId = actualKey;
    }
  }
  return keyMap;
}

// customMetadata.imageSource is the primary identity source; properties.
// {siteId,webId,listId,uniqueId} is the fallback when it is missing.
function imageIds(customImageSource, props) {
  const keyMap = customImageIds(customImageSource);
  if (Object.keys(keyMap).length) {
    return {
      siteId: keyMap.siteId ? customImageSource[keyMap.siteId] : undefined,
      webId: keyMap.webId ? customImageSource[keyMap.webId] : undefined,
      listId: keyMap.listId ? customImageSource[keyMap.listId] : undefined,
      uniqueId: keyMap.uniqueId ? customImageSource[keyMap.uniqueId] : undefined,
    };
  }
  return { siteId: props?.siteId, webId: props?.webId, listId: props?.listId, uniqueId: props?.uniqueId };
}

export default {
  id: 'd1d91016-032f-456d-98a4-721247c305e8',
  label: 'Image',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const imageSources = spc.imageSources || {};
    const refs = [];

    const imageValue = imageSources.imageSource;
    if (imageValue && underSource(imageValue, ctx)) {
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', 'imageSource'],
        value: imageValue,
        class: 'asset',
        asset: { path: sourceRelativePath(imageValue), ids: imageIds(spc.customMetadata?.imageSource, props) },
      });
    }

    const linkValue = props.linkUrl;
    if (linkValue && underSource(linkValue, ctx)) {
      refs.push({
        path: ['webPartData', 'properties', 'linkUrl'],
        value: linkValue,
        class: 'link',
      });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    let count = 0;

    if (props && spc) {
      const imageSources = spc.imageSources;
      const imageValue = imageSources?.imageSource;
      if (imageValue && underSource(imageValue, ctx)) {
        const mapping = ctx.mapAsset({ path: sourceRelativePath(imageValue), ids: imageIds(spc.customMetadata?.imageSource, props) });
        // Path and ids move together or not at all — half a mapping would
        // leave the image path and its id quadruple naming different files.
        if (mapping && mapping.path !== undefined && mapping.ids) {
          const nextPath = /^https?:\/\//i.test(String(imageSources.imageSource)) && mapping.url
            ? mapping.url : mapping.path;
          if (imageSources.imageSource !== nextPath) {
            imageSources.imageSource = nextPath;
            count += 1;
          }

          const customImageSource = spc.customMetadata?.imageSource;
          if (customImageSource && typeof customImageSource === 'object') {
            const keyMap = customImageIds(customImageSource);
            for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
              const actualKey = keyMap[idKey];
              if (!actualKey) continue;
              const nextId = mapping.ids[idKey];
              if (nextId !== undefined && ctx.normalizeGuid(customImageSource[actualKey]) !== ctx.normalizeGuid(nextId)) {
                customImageSource[actualKey] = nextId;
                count += 1;
              }
            }
          }

          for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
            if (!(idKey in props)) continue;
            const nextId = mapping.ids[idKey];
            if (nextId !== undefined && ctx.normalizeGuid(props[idKey]) !== ctx.normalizeGuid(nextId)) {
              props[idKey] = nextId;
              count += 1;
            }
          }
        }
      }

      const linkValue = props.linkUrl;
      if (linkValue && underSource(linkValue, ctx)) {
        const mapped = ctx.mapLink(linkValue);
        if (mapped != null && mapped !== linkValue) {
          props.linkUrl = mapped;
          count += 1;
        }
      }
    }

    // Everything else — captionText, altText, overlayText, fileName,
    // overlayTextStyles, alignment, isStretchEnabled... — is left
    // completely untouched.
    return count;
  },
};
