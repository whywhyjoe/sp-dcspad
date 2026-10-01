// Page copy analyzer — Highlighted content (design/PAGE-COPY.md §5.1–5.2,
// §11 Q7).
//
// Web part id daf0b71c-6de8-4ef7-b511-faae7c388708. The real fixture instance
// (tests/pages-copy/fixtures/live-shapes-editor.json "highlightedContent",
// captured from zz-pagecopy-shapes.aspx) carries `properties.query
// .contentLocation: 1` with an empty `properties.sites` array and the
// *authoring* web's own webId/siteId echoed into `properties.webId`/
// `properties.siteId`. That is the §5.2 "data → warn" verdict,
// context-relative flavor: the only source mode this shape has been
// observed in is "this site" (1), so the note says the stored ids are only
// what the source picker recorded when the part was configured — not
// necessarily what a copy on another site will actually query. A populated
// `properties.sites` array is the opposite shape (mirroring News's
// `newsSiteList` / Events' `sites`): those entries name the query directly,
// so they keep pointing at the source and are reported without the
// context-relative caveat, and the whole instance's other data refs follow
// suit (no other `contentLocation` value has been observed live, so an
// unknown value degrades to the same non-context-relative wording rather
// than guessing). `properties.query.filters[].value` is a free-text search
// filter, not an id — never a ref. Neither flavor is ever patched — kept
// as-is with the warning `refs()` already produced, like every other "data"
// verdict (list-library.js/news.js/events.js are the same shape).
//
// serverProcessedContent.links (see the quickLinks fixture for the general
// shape: a `baseUrl` plus per-item `items[N]....url` entries) is walked
// separately: any entry other than `baseUrl` — which is resolution metadata,
// not a link itself — whose value points under the source web is a 'link'
// reference, patchable only through ctx.mapLink. The real Highlighted
// content fixture only carries `links.baseUrl` (no linked item), so this
// only fires on a page built with one; the walk is still generic so it does
// not silently miss one.

const CONTENT_LOCATION_SOURCE = {
  1: { label: 'this site', contextRelative: true },
};

function sourceNote(mode, hasSelectedSites) {
  if (hasSelectedSites) {
    return 'Highlighted content source: selected sites — keeps pointing at the source web/site';
  }
  const known = CONTENT_LOCATION_SOURCE[mode];
  const label = known ? known.label : `source mode ${JSON.stringify(mode)}`;
  return known && known.contextRelative
    ? `Highlighted content source: ${label} — context-relative; may re-resolve against the destination web instead of the source`
    : `Highlighted content source: ${label} — keeps pointing at the source web/site`;
}

function isPlainObject(v) {
  return Boolean(v) && typeof v === 'object' && !Array.isArray(v);
}

// Relative path under ctx.sourceWebPath, or absolute on ctx.origin with a
// path under it — the same rule the generic scanner and the other
// analyzers use (page-copy.js's underPath / text.js's matchSourcePath).
function underSourceWeb(value, ctx) {
  const raw = String(value ?? '');
  if (!raw) return false;
  if (raw.startsWith('/') && !raw.startsWith('//')) return ctx.underPath(raw, ctx.sourceWebPath);
  let url;
  try { url = new URL(raw); } catch { return false; }
  if (!ctx.origin || url.origin.toLowerCase() !== String(ctx.origin).toLowerCase()) return false;
  let path = url.pathname;
  try { path = decodeURIComponent(path); } catch { /* keep as-is */ }
  return ctx.underPath(path, ctx.sourceWebPath);
}

function computeRefs(instance, ctx) {
  const props = instance?.webPartData?.properties;
  const out = [];

  if (isPlainObject(props)) {
    const sites = Array.isArray(props.sites) ? props.sites.filter((id) => typeof id === 'string' && id) : [];
    const contentLocation = isPlainObject(props.query) ? props.query.contentLocation : undefined;
    const note = sourceNote(contentLocation, sites.length > 0);

    if (typeof props.siteId === 'string' && props.siteId) {
      out.push({ path: ['webPartData', 'properties', 'siteId'], value: props.siteId, class: 'data', note });
    }
    if (typeof props.webId === 'string' && props.webId) {
      out.push({ path: ['webPartData', 'properties', 'webId'], value: props.webId, class: 'data', note });
    }
    sites.forEach((id, i) => {
      out.push({
        path: ['webPartData', 'properties', 'sites', i],
        value: id,
        class: 'data',
        note: 'Highlighted content source: selected sites — keeps pointing at the source web/site',
      });
    });
  }

  const links = instance?.webPartData?.serverProcessedContent?.links;
  if (isPlainObject(links)) {
    for (const [key, value] of Object.entries(links)) {
      if (key === 'baseUrl') continue; // resolution metadata, not a link itself
      if (typeof value !== 'string' || !underSourceWeb(value, ctx)) continue;
      out.push({
        path: ['webPartData', 'serverProcessedContent', 'links', key],
        value,
        class: 'link',
        note: 'link into the source web',
      });
    }
  }

  return out;
}

function setAtPath(root, path, value) {
  let node = root;
  for (let i = 0; i < path.length - 1; i += 1) node = node[path[i]];
  node[path[path.length - 1]] = value;
}

export default {
  id: 'daf0b71c-6de8-4ef7-b511-faae7c388708',
  label: 'Highlighted content',

  refs(instance, ctx) {
    return computeRefs(instance, ctx);
  },

  // §5.2 "data → warn": the site/web data refs are never patched — kept
  // as-is with the warning refs() already produced. Only 'link' refs are
  // ever authorized to change, and only via ctx.mapLink.
  patch(instance, ctx) {
    let changed = 0;
    for (const ref of computeRefs(instance, ctx)) {
      if (ref.class !== 'link') continue;
      const mapped = ctx.mapLink ? ctx.mapLink(ref.value) : null;
      if (typeof mapped !== 'string' || !mapped || mapped === ref.value) continue;
      setAtPath(instance, ref.path, mapped);
      changed += 1;
    }
    return changed;
  },
};
