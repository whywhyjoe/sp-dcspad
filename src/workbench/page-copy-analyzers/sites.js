// Page copy analyzer — Sites (design/PAGE-COPY.md §5.1–5.2).
//
// Web part id 7cba020c-5ccb-42e8-b6fc-75b3149aba7b. `properties.sites` is
// populated only when the web part has a fixed set to show — "selected
// sites" or hub scope (real fixture: tests/pages-copy/fixtures/
// live-shapes-editor.json "sites" — two selected sites, the sputils-test
// subweb and the TestSiteCollection site collection). "Frequent sites for
// current user" is resolved at render time and carries no stored site list,
// so an empty/absent `sites` array means nothing to report — the §5.2
// "data → warn" verdict only fires once there is a fixed reference to warn
// about.
//
// Each selected-site entry carries its own SiteId/WebId/GroupId
// (`ItemReference`) plus a matching `serverProcessedContent.links`
// `sites[N].Url` entry. Ids and url together name ONE site, which keeps
// existing wherever the copy lands, so both are 'data' and never patched —
// even a subweb whose url sits under the source web's path: re-basing it
// with the opt-in link rewrite would aim the card at a site that does not
// exist at the destination while its ids still named the original.
//
// A selected site that IS the source web itself gets a distinct note
// ("points at the source site itself") from every other selected site
// ("another site — unaffected by the copy"), so the report doesn't lump a
// self-reference in with an unrelated site.

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function isPlainObject(v) {
  return Boolean(v) && typeof v === 'object' && !Array.isArray(v);
}

function siteLabel(entry, i, spc) {
  const fromSpc = isPlainObject(spc?.searchablePlainTexts) ? spc.searchablePlainTexts[`sites[${i}].Title`] : undefined;
  if (typeof fromSpc === 'string' && fromSpc.trim()) return fromSpc.trim();
  return (typeof entry?.Acronym === 'string' && entry.Acronym) || `Site ${i + 1}`;
}

function computeRefs(instance, ctx) {
  const props = instance?.webPartData?.properties;
  const spc = instance?.webPartData?.serverProcessedContent;
  const sites = Array.isArray(props?.sites) ? props.sites : [];
  if (!sites.length) return []; // frequent sites for current user: user-relative, nothing to warn about

  const sourceWebId = ctx.normalizeGuid(ctx.sourceIds?.webId);
  const out = [];

  sites.forEach((entry, i) => {
    if (!isPlainObject(entry)) return;
    const ref = isPlainObject(entry.ItemReference) ? entry.ItemReference : {};
    const label = siteLabel(entry, i, spc);
    const isSourceWeb = Boolean(sourceWebId) && ctx.normalizeGuid(ref.WebId) === sourceWebId;
    const idNote = isSourceWeb
      ? `${label} — points at the source site itself`
      : `${label} — another site — unaffected by the copy`;

    for (const key of ['SiteId', 'WebId', 'GroupId']) {
      const value = ref[key];
      if (typeof value === 'string' && GUID_RE.test(value)) {
        out.push({ path: ['webPartData', 'properties', 'sites', i, 'ItemReference', key], value, class: 'data', note: idNote });
      }
    }

    const url = isPlainObject(spc?.links) ? spc.links[`sites[${i}].Url`] : undefined;
    if (typeof url === 'string' && url) {
      out.push({
        path: ['webPartData', 'serverProcessedContent', 'links', `sites[${i}].Url`],
        value: url,
        class: 'data',
        note: idNote,
      });
    }
  });

  return out;
}

export default {
  id: '7cba020c-5ccb-42e8-b6fc-75b3149aba7b',
  label: 'Sites',

  refs(instance, ctx) {
    return computeRefs(instance, ctx);
  },

  // §5.2 "data → warn": site references are kept as-is with the warning
  // refs() already produced — never patched.
  patch() {
    return 0;
  },
};
