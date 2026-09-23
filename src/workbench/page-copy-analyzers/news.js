// Page copy analyzer — News (design/PAGE-COPY.md §5.1–5.2, §11 Q7).
//
// Two ids ship as "News" from GetClientSideWebParts: 8c88f208-6c77-4bdb-
// 86a0-0c47b4316588 and a5df8fdf-b508-4b66-98a6-d83bc2597f63. Both real
// fixture instances (tests/pages-copy/fixtures/live-shapes.json "news".0/.1,
// captured from CollabHome.aspx and Test.aspx) carry
// `properties.newsDataSourceProp: 1` ("this site") with the *authoring*
// web's own webId/siteId echoed into `properties.webId`/`properties.siteId`
// and an empty `properties.newsSiteList`. That is the §5.2 "data → warn"
// verdict, context-relative flavor: the web part queries "this site" at
// render time, so the stored ids are only what the source picker recorded
// when the part was configured — not what a copy on another site will
// actually query — and the note says so. `newsDataSourceProp: 2` ("selected
// sites") is the opposite: the ids in `newsSiteList` ARE the query, so they
// keep pointing at the source and are reported without the
// context-relative caveat. Neither flavor is ever patched — kept as-is
// with the warning `refs()` already produced, like every other "data"
// verdict (list-library.js is the same shape).
//
// serverProcessedContent.links (see the quickLinks fixture for the general
// shape: a `baseUrl` plus per-item `items[N]....url` entries) is walked
// separately: any entry other than `baseUrl` — which is resolution
// metadata, not a link itself — whose value points under the source web is
// a 'link' reference, patchable only through ctx.mapLink. Real News
// fixtures never carry item links (no items to link to), so this only
// fires on a page built with a linked item; the walk is still generic so
// it does not silently miss one.

const NEWS_SOURCE = {
  0: { label: 'all sites', contextRelative: false },
  1: { label: 'this site', contextRelative: true },
  2: { label: 'selected sites', contextRelative: false },
  3: { label: 'recommended sites', contextRelative: false },
};

function sourceNote(mode) {
  const known = NEWS_SOURCE[mode];
  const label = known ? known.label : `source mode ${JSON.stringify(mode)}`;
  return known && known.contextRelative
    ? `News source: ${label} — context-relative; may re-resolve against the destination web instead of the source`
    : `News source: ${label} — keeps pointing at the source web/site`;
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
    const note = sourceNote(props.newsDataSourceProp);
    if (typeof props.siteId === 'string' && props.siteId) {
      out.push({ path: ['webPartData', 'properties', 'siteId'], value: props.siteId, class: 'data', note });
    }
    if (typeof props.webId === 'string' && props.webId) {
      out.push({ path: ['webPartData', 'properties', 'webId'], value: props.webId, class: 'data', note });
    }
    if (Array.isArray(props.newsSiteList)) {
      props.newsSiteList.forEach((id, i) => {
        if (typeof id !== 'string' || !id) return;
        out.push({
          path: ['webPartData', 'properties', 'newsSiteList', i],
          value: id,
          class: 'data',
          note: 'News source: selected sites — keeps pointing at the source web/site',
        });
      });
    }
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
  id: ['8c88f208-6c77-4bdb-86a0-0c47b4316588', 'a5df8fdf-b508-4b66-98a6-d83bc2597f63'],
  label: 'News',

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
