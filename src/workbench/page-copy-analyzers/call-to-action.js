// Call to action analyzer (design/PAGE-COPY.md §5.1-5.2, web part id
// df8e44e7-edd5-46d5-90da-aca1539313b8).
//
// Shape (editor-authored live capture, tests/pages-copy/fixtures/
// live-shapes-editor.json → callToAction; background set with "Don't copy",
// button label "Join", button linked to a page in the source web):
//   webPartData: {
//     serverProcessedContent: {
//       searchablePlainTexts: { 'button.label': <text — not a reference> },
//       imageSources: { 'image.url': <background image path/URL> },
//       links: { 'button.linkUrl': <button destination path/URL> },
//     },
//     properties: {
//       image: { itemInfo: { siteId, webId, listId, uniqueId }, zoomRatio },
//       button: {}, overlayText: {}, alignment, minimumLayoutWidth,
//     },
//   }
//
// Background image (§5.2 "asset → copied"): the path/URL lives in
// serverProcessedContent.imageSources['image.url'], its four ids on
// properties.image.itemInfo — one object, no primary/fallback split (unlike
// the header banner, which has both a customMetadata copy and a properties
// mirror). Button link (§5.2 "link"): serverProcessedContent
// .links['button.linkUrl'], patched only via ctx.mapLink (never invented —
// the operator has to opt into link rewriting). button.label is plain text,
// not a URL or id, so it is never a reference.

function decodePath(p) {
  try { return decodeURIComponent(p); } catch { return p; }
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
  if (raw.startsWith('/') && !raw.startsWith('//')) return decodePath(raw.split(/[?#]/)[0]);
  try { return decodeURIComponent(new URL(raw).pathname); } catch { return raw; }
}

function imageIds(props) {
  const info = props?.image?.itemInfo;
  return { siteId: info?.siteId, webId: info?.webId, listId: info?.listId, uniqueId: info?.uniqueId };
}

export default {
  id: 'df8e44e7-edd5-46d5-90da-aca1539313b8',
  label: 'Call to action',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const imageSources = spc.imageSources || {};
    const links = spc.links || {};
    const refs = [];

    const bgValue = imageSources['image.url'];
    if (bgValue && underSource(bgValue, ctx)) {
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', 'image.url'],
        value: bgValue,
        class: 'asset',
        asset: { path: sourceRelativePath(bgValue), ids: imageIds(props) },
      });
    }

    const linkValue = links['button.linkUrl'];
    if (linkValue && underSource(linkValue, ctx)) {
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'links', 'button.linkUrl'],
        value: linkValue,
        class: 'link',
      });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    if (!props || !spc) return 0;
    const imageSources = spc.imageSources;
    const links = spc.links;
    let count = 0;

    if (imageSources) {
      const bgValue = imageSources['image.url'];
      if (bgValue && underSource(bgValue, ctx)) {
        const mapping = ctx.mapAsset({ path: sourceRelativePath(bgValue), ids: imageIds(props) });
        // Path and ids move together or not at all — half a mapping would
        // leave the background path and its id quadruple naming different
        // files.
        if (mapping && mapping.path !== undefined && mapping.ids) {
          const isAbsolute = /^https?:\/\//i.test(String(bgValue));
          const next = isAbsolute && mapping.url ? mapping.url : mapping.path;
          if (next !== undefined && next !== bgValue) {
            imageSources['image.url'] = next;
            count += 1;
          }
          const itemInfo = props.image?.itemInfo;
          if (itemInfo) {
            for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
              if (!(idKey in itemInfo)) continue;
              const nextId = mapping.ids[idKey];
              if (nextId !== undefined && ctx.normalizeGuid(itemInfo[idKey]) !== ctx.normalizeGuid(nextId)) {
                itemInfo[idKey] = nextId;
                count += 1;
              }
            }
          }
        }
      }
    }

    if (links) {
      const linkValue = links['button.linkUrl'];
      if (linkValue && underSource(linkValue, ctx)) {
        const mapped = ctx.mapLink(linkValue);
        if (mapped != null && mapped !== linkValue) {
          links['button.linkUrl'] = mapped;
          count += 1;
        }
      }
    }

    // Everything else — button.label, overlayText, alignment,
    // minimumLayoutWidth, zoomRatio... — is left completely untouched.
    return count;
  },
};
