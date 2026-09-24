// File and media (File viewer) analyzer (design/PAGE-COPY.md §5.2-§5.3, web
// part id b7dd04e1-19ce-4b24-9132-b60a1c2b910d).
//
// Shape (live capture, tests/pages-copy/fixtures/live-shapes-editor.json →
// fileViewer — an editor-authored page pointing at Shared Documents/
// zz-pagecopy-docs/zz-pagecopy-report.pdf):
//   webPartData.properties: { file (absolute URL on the source origin),
//     siteId, webId, listId, uniqueId — the file's own four ids, carried
//     flat on properties (no customMetadata wrapper here, unlike the header
//     and quick-links shapes), authorName, photoUrl (a tokened afdcache CDN
//     avatar — never a ref, same convention as the header part's
//     BannerThumbnailUrl), webAbsoluteUrl, modifiedAt, ... }
//   webPartData.serverProcessedContent.links: { serverRelativeUrl (the
//     decoded-once, percent-encoded server-relative path to the same file),
//     wopiurl }
// On this capture wopiurl is byte-identical to serverRelativeUrl, but per
// §5.3 a WOPI frame url can instead take a `_layouts/15/Doc.aspx?
// sourcedoc={guid}&...` shape, where the file's own uniqueId is the only
// thing naming it inside the string — this analyzer recognizes and rewrites
// either form.
//
// The document is the one non-image asset verdict in §5.2 ("File viewer
// document"): readFileBytes + upload, the analyzer only reports the
// reference and authorizes the patch. A dynamic-data connection some File
// viewer instances carry (webPartData.dynamicDataPaths, e.g. bound to a
// list item) rides on the web part's own envelope and is handled
// generically by page-copy.js's analyzeParts() — never duplicated here.

function decodePath(p) { try { return decodeURIComponent(p); } catch { return p; } }

function splitPath(raw) { const m = /^([^?#]*)/.exec(String(raw ?? '')); return m ? m[1] : raw; }

// Server-relative under ctx.sourceWebPath, or absolute on ctx.origin with a
// path under it — the same test header.js/quick-links.js use.
function underSource(value, ctx) {
  const raw = String(value ?? '');
  if (!raw) return false;
  if (raw.startsWith('/') && !raw.startsWith('//')) return ctx.underPath(decodePath(splitPath(raw)), ctx.sourceWebPath);
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

function fileIds(props) {
  return { siteId: props?.siteId, webId: props?.webId, listId: props?.listId, uniqueId: props?.uniqueId };
}

// The identity's own site/web, when known, must agree with the source web —
// otherwise a `sourcedoc` GUID match alone (no path to test) could wrongly
// claim a file that only happens to share this instance's declared ids.
function identityInSourceWeb(ids, ctx) {
  const siteId = ctx.sourceIds?.siteId;
  const webId = ctx.sourceIds?.webId;
  if (!siteId && !webId) return true;
  if (ids.siteId && siteId && ctx.normalizeGuid(ids.siteId) !== siteId) return false;
  if (ids.webId && webId && ctx.normalizeGuid(ids.webId) !== webId) return false;
  return true;
}

// A `sourcedoc=` query token naming a file by GUID — the identity a
// Doc.aspx/WOPI-frame form of the link carries instead of a path.
// { full, prefix, guid, suffix } or null.
function sourceDocToken(value) {
  const raw = String(value ?? '');
  const m = /sourcedoc=(%7b|\{)?([0-9a-f-]{32,36})(%7d|\})?/i.exec(raw);
  return m ? { full: m[0], prefix: m[1] || '', guid: m[2], suffix: m[3] || '' } : null;
}

export default {
  id: 'b7dd04e1-19ce-4b24-9132-b60a1c2b910d',
  label: 'File and media',

  refs(instance, ctx) {
    const props = instance?.webPartData?.properties || {};
    const links = instance?.webPartData?.serverProcessedContent?.links || {};
    const ids = fileIds(props);
    const refs = [];

    const fileUnderSource = Boolean(props.file) && underSource(props.file, ctx);
    const serverRelUnderSource = Boolean(links.serverRelativeUrl) && underSource(links.serverRelativeUrl, ctx);
    const wopiUnderSource = Boolean(links.wopiurl) && underSource(links.wopiurl, ctx);
    const wopiToken = links.wopiurl ? sourceDocToken(links.wopiurl) : null;
    const wopiMatchesFile = Boolean(wopiToken) && Boolean(ids.uniqueId)
      && ctx.normalizeGuid(wopiToken.guid) === ctx.normalizeGuid(ids.uniqueId) && identityInSourceWeb(ids, ctx);

    if (!fileUnderSource && !serverRelUnderSource && !wopiUnderSource && !wopiMatchesFile) return refs;

    // The canonical path: prefer the server-relative form, then the
    // absolute one; a `sourcedoc` wopiurl names a system page, not the
    // file's own path, so it never supplies the canonical path.
    const canonicalPath = (serverRelUnderSource && sourceRelativePath(links.serverRelativeUrl))
      || (fileUnderSource && sourceRelativePath(props.file))
      || (wopiUnderSource && !wopiToken && sourceRelativePath(links.wopiurl))
      || '';
    // One shared identity object — every ref below points at the same file,
    // so assetRequests() (page-copy.js) dedupes them by uniqueId/path.
    const asset = { path: canonicalPath, ids };

    if (fileUnderSource) {
      refs.push({ path: ['webPartData', 'properties', 'file'], value: props.file, class: 'asset', asset });
    }
    if (serverRelUnderSource) {
      refs.push({ path: ['webPartData', 'serverProcessedContent', 'links', 'serverRelativeUrl'], value: links.serverRelativeUrl, class: 'asset', asset });
    }
    if (wopiUnderSource || wopiMatchesFile) {
      refs.push({ path: ['webPartData', 'serverProcessedContent', 'links', 'wopiurl'], value: links.wopiurl, class: 'asset', asset });
    }

    return refs;
  },

  patch(instance, ctx) {
    const props = instance?.webPartData?.properties;
    const links = instance?.webPartData?.serverProcessedContent?.links;
    if (!props) return 0;
    const ids = fileIds(props);

    const fileUnderSource = Boolean(props.file) && underSource(props.file, ctx);
    const serverRelUnderSource = Boolean(links?.serverRelativeUrl) && underSource(links.serverRelativeUrl, ctx);
    const canonicalPath = (serverRelUnderSource && sourceRelativePath(links.serverRelativeUrl))
      || (fileUnderSource && sourceRelativePath(props.file))
      || '';
    if (!canonicalPath && !ids.uniqueId) return 0;

    const mapping = ctx.mapAsset({ path: canonicalPath, ids });
    if (!mapping) return 0;

    let count = 0;

    if (fileUnderSource) {
      const next = mapping.url !== undefined ? mapping.url : mapping.path;
      if (next !== undefined && props.file !== next) { props.file = next; count += 1; }
    }

    // Live shape (§11 Q7): `webAbsoluteUrl` is the web holding the document —
    // the destination web once the document has been transferred there.
    if (typeof props.webAbsoluteUrl === 'string' && ctx.target?.webUrl
        && underSource(props.webAbsoluteUrl, ctx)) {
      const nextWeb = String(ctx.target.webUrl).replace(/\/+$/, '');
      if (props.webAbsoluteUrl.replace(/\/+$/, '') !== nextWeb) { props.webAbsoluteUrl = nextWeb; count += 1; }
    }

    if (mapping.ids) {
      for (const idKey of ['siteId', 'webId', 'listId', 'uniqueId']) {
        if (!(idKey in props)) continue;
        const nextId = mapping.ids[idKey];
        if (nextId !== undefined && props[idKey] !== nextId) { props[idKey] = nextId; count += 1; }
      }
    }

    if (links) {
      if (serverRelUnderSource && mapping.path !== undefined && links.serverRelativeUrl !== mapping.path) {
        links.serverRelativeUrl = mapping.path;
        count += 1;
      }

      if (links.wopiurl) {
        const original = links.wopiurl;
        const token = sourceDocToken(original);
        const tokenMatches = Boolean(token) && Boolean(ids.uniqueId)
          && ctx.normalizeGuid(token.guid) === ctx.normalizeGuid(ids.uniqueId) && identityInSourceWeb(ids, ctx);
        let next = original;
        if (tokenMatches && mapping.ids?.uniqueId) {
          // Rewrite only the GUID token — the rest of the WOPI frame url
          // (host, _layouts path, other query params) is a system page, not
          // the file, and is left exactly as it was.
          const replacement = `sourcedoc=${token.prefix}${mapping.ids.uniqueId}${token.suffix}`;
          const at = original.indexOf(token.full);
          next = `${original.slice(0, at)}${replacement}${original.slice(at + token.full.length)}`;
        } else if (underSource(original, ctx) && mapping.path !== undefined) {
          const isAbsolute = /^https?:\/\//i.test(original);
          next = isAbsolute && mapping.url !== undefined ? mapping.url : mapping.path;
        }
        if (next !== original) { links.wopiurl = next; count += 1; }
      }
    }

    return count;
  },
};
