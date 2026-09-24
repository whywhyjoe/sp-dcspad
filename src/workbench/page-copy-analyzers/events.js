// Page copy analyzer — Events (design/PAGE-COPY.md §5.1–5.2, §11 Q7).
//
// Web part id 20745d7d-8581-4a6c-bf26-68279bc123fc. The real fixture instance
// (tests/pages-copy/fixtures/live-shapes-editor.json "events", captured from
// zz-pagecopy-shapes.aspx) carries `properties.dataSource: 7` with an empty
// `properties.sites` array and the *authoring* web's own webId/siteId echoed
// into `properties.webId`/`properties.siteId`, plus `properties.selectedListId`
// naming the events list the picker resolved on the source at configuration
// time. That is the §5.2 "data → warn" verdict, context-relative flavor: the
// only source mode this shape has been observed in is "this site" (7), so the
// note says the stored ids/list are only what the picker recorded when the
// part was configured — not necessarily what a copy on another site will
// actually query. A populated `properties.sites` array is the opposite shape
// (mirroring News's `newsSiteList`): those entries name the query directly,
// so they keep pointing at the source and are reported without the
// context-relative caveat, and the whole instance's other data refs follow
// suit (no other `dataSource` value has been observed live, so an unknown
// value degrades to the same non-context-relative wording rather than
// guessing). Neither flavor is ever patched — kept as-is with the warning
// `refs()` already produced, like every other "data" verdict
// (list-library.js/news.js are the same shape).
//
// serverProcessedContent.links (see the quickLinks fixture for the general
// shape: a `baseUrl` plus per-item `items[N]....url` entries) is walked
// separately: any entry other than `baseUrl` — which is resolution metadata,
// not a link itself — whose value points under the source web is a 'link'
// reference, patchable only through ctx.mapLink. The real Events fixture
// only carries `links.baseUrl` (no linked item), so this only fires on a
// page built with one; the walk is still generic so it does not silently
// miss one.

const EVENTS_SOURCE = {
  7: { label: 'this site', contextRelative: true },
};

function sourceNote(mode, hasSelectedSites) {
  if (hasSelectedSites) {
    return 'Events source: selected sites — keeps pointing at the source web/site';
  }
  const known = EVENTS_SOURCE[mode];
  const label = known ? known.label : `source mode ${JSON.stringify(mode)}`;
  return known && known.contextRelative
    ? `Events source: ${label} — context-relative; may re-resolve against the destination web instead of the source`
    : `Events source: ${label} — keeps pointing at the source web/site`;
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
    const note = sourceNote(props.dataSource, sites.length > 0);

    if (typeof props.siteId === 'string' && props.siteId) {
      out.push({ path: ['webPartData', 'properties', 'siteId'], value: props.siteId, class: 'data', note });
    }
    if (typeof props.webId === 'string' && props.webId) {
      out.push({ path: ['webPartData', 'properties', 'webId'], value: props.webId, class: 'data', note });
    }
    if (typeof props.selectedListId === 'string' && props.selectedListId) {
      out.push({
        path: ['webPartData', 'properties', 'selectedListId'],
        value: props.selectedListId,
        class: 'data',
        note,
      });
    }
    sites.forEach((id, i) => {
      out.push({
        path: ['webPartData', 'properties', 'sites', i],
        value: id,
        class: 'data',
        note: 'Events source: selected sites — keeps pointing at the source web/site',
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
  id: '20745d7d-8581-4a6c-bf26-68279bc123fc',
  label: 'Events',

  refs(instance, ctx) {
    return computeRefs(instance, ctx);
  },

  // §5.2 "data → warn": the site/web/list data refs are never patched —
  // kept as-is with the warning refs() already produced. Only 'link' refs
  // are ever authorized to change, and only via ctx.mapLink.
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
