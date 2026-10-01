// Countdown timer analyzer (design/PAGE-COPY.md §5.1-5.2, web part id
// 62cac389-787f-495d-beca-e11786162ef4).
//
// Shape (live capture, tests/pages-copy/fixtures/live-shapes-editor.json →
// countdown — editor-authored, background set, CTA off):
//   webPartData: {
//     properties: {
//       showButton, buttonText, buttonURL, countDate, countDirection,
//       dateDisplay, title, description,
//       backgroundImage: { focalPosition, id, listId, siteId, webId,
//         zoomRatio, widthFactor },
//       backgroundOverlay: { color, opacity, useLightText },
//     },
//     serverProcessedContent: {
//       htmlStrings: {}, searchablePlainTexts: {},
//       imageSources: { 'backgroundImage.url': <source-relative path> },
//       links: {},
//     },
//   }
//
// The background image is a rendered asset (§5.2 "asset → copied"): its
// path lives in serverProcessedContent.imageSources['backgroundImage.url'],
// its identity on properties.backgroundImage — note the file's unique id is
// keyed `id` there, not `uniqueId` (unlike the header/quick-links shapes),
// so it is read from and written back to that key specifically; siteId/
// webId/listId sit alongside it under the same object. Path and ids move
// together or not at all (mapAsset returning an incomplete mapping leaves
// the background entirely untouched).
//
// buttonURL is the call-to-action link (§5.2 "link"): a plain URL property,
// not routed through serverProcessedContent.links like Quick links/News.
// The live fixture has the CTA off (buttonURL empty) — the ref only fires
// when the property is populated and points into the source web, kept by
// default and rewritten only through ctx.mapLink (opt-in — decision 3).
// showButton is a display toggle, not consulted here: an authored-but-
// hidden CTA still names a real link worth tracking.

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

// properties.backgroundImage.id is the file's unique id — the fourth
// identity component is spelled differently here than on header/quick-links.
function backgroundIds(bg) {
  return {
    siteId: bg?.siteId,
    webId: bg?.webId,
    listId: bg?.listId,
    uniqueId: bg?.id,
  };
}

export default {
  id: '62cac389-787f-495d-beca-e11786162ef4',
  label: 'Countdown timer',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const spc = instance?.webPartData?.serverProcessedContent || {};
    const imageSources = spc.imageSources || {};
    const refs = [];

    const bgValue = imageSources['backgroundImage.url'];
    if (bgValue && underSource(bgValue, ctx)) {
      refs.push({
        path: ['webPartData', 'serverProcessedContent', 'imageSources', 'backgroundImage.url'],
        value: bgValue,
        class: 'asset',
        asset: { path: sourceRelativePath(bgValue), ids: backgroundIds(props.backgroundImage) },
      });
    }

    const buttonURL = props.buttonURL;
    if (typeof buttonURL === 'string' && buttonURL && underSource(buttonURL, ctx)) {
      refs.push({
        path: ['webPartData', 'properties', 'buttonURL'],
        value: buttonURL,
        class: 'link',
        note: 'call-to-action link into the source web',
      });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const spc = instance?.webPartData?.serverProcessedContent;
    const imageSources = spc?.imageSources;
    let count = 0;

    if (props && imageSources) {
      const bgValue = imageSources['backgroundImage.url'];
      if (bgValue && underSource(bgValue, ctx)) {
        const bg = props.backgroundImage;
        const mapping = ctx.mapAsset({ path: sourceRelativePath(bgValue), ids: backgroundIds(bg) });
        // Path and ids move together or not at all.
        if (mapping && mapping.path !== undefined && mapping.ids) {
          const isAbsolute = /^https?:\/\//i.test(String(bgValue));
          const next = isAbsolute && mapping.url ? mapping.url : mapping.path;
          if (next !== undefined && next !== bgValue) {
            imageSources['backgroundImage.url'] = next;
            count += 1;
          }
          if (bg && typeof bg === 'object') {
            if ('id' in bg && mapping.ids.uniqueId !== undefined && bg.id !== mapping.ids.uniqueId) {
              bg.id = mapping.ids.uniqueId;
              count += 1;
            }
            for (const idKey of ['siteId', 'webId', 'listId']) {
              if (!(idKey in bg)) continue;
              const nextId = mapping.ids[idKey];
              if (nextId !== undefined && bg[idKey] !== nextId) {
                bg[idKey] = nextId;
                count += 1;
              }
            }
          }
        }
      }
    }

    if (props && typeof props.buttonURL === 'string' && props.buttonURL && underSource(props.buttonURL, ctx)) {
      const mapped = ctx.mapLink ? ctx.mapLink(props.buttonURL) : null;
      if (typeof mapped === 'string' && mapped && mapped !== props.buttonURL) {
        props.buttonURL = mapped;
        count += 1;
      }
    }

    return count;
  },
};
