// Text analyzer — rich text (controlType 4) (design/PAGE-COPY.md §5.1–5.2).
//
// Rich text lives entirely in `innerHTML`; there is no serverProcessedContent
// or properties bag to read (PAGE-COPY.md §5.1). This analyzer parses that
// markup with DOMParser and classifies every `a[href]`, `img[src]` and
// `[data-*]` attribute that points into the source web:
//   - an `<img src>` into the source web is an asset reference;
//   - an `<a href>` into the source web is a link reference (kept by
//     default; rewritten only when the operator opts in — decision 3);
//   - a `data-*` attribute into the source web is reported as a link too
//     (nothing recognized patches it — see patch() below).
// Anything outside the source web (external hosts, other webs on the same
// tenant) is not a reference at all: "Links elsewhere are not refs."
//
// patch() never string-replaces: it re-parses innerHTML, mutates only the
// attributes an analyzer is authorized to change (img src via
// ctx.mapAsset, a href via ctx.mapLink), and re-serializes from
// `body.innerHTML` only when something actually changed.

function parseHtml(html) {
  if (typeof DOMParser !== 'function') return null;
  try {
    return new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  } catch {
    return null;
  }
}

function splitPath(raw) {
  const m = /^([^?#]*)/.exec(raw);
  return m ? m[1] : raw;
}

function decodePath(p) {
  try { return decodeURIComponent(p); } catch { return p; }
}

// Relative server path under ctx.sourceWebPath, or absolute on ctx.origin
// with a path under it → the decoded server-relative path; else null.
function matchSourcePath(value, ctx) {
  const raw = String(value ?? '');
  if (!raw) return null;
  if (raw.startsWith('/') && !raw.startsWith('//')) {
    return ctx.underPath(raw, ctx.sourceWebPath) ? decodePath(splitPath(raw)) : null;
  }
  let url;
  try { url = new URL(raw); } catch { return null; }
  if (!ctx.origin || url.origin.toLowerCase() !== String(ctx.origin).toLowerCase()) return null;
  const path = decodePath(url.pathname);
  return ctx.underPath(path, ctx.sourceWebPath) ? path : null;
}

// Every a[href], img[src] and [data-*] attribute in the markup, in document
// order, each tagged with its sequential index among this combined set
// (matches the generic scanner's attrValues() indexing in page-copy.js).
function collectAttrs(doc) {
  const out = [];
  for (const node of doc.body.querySelectorAll('*')) {
    const tag = node.tagName.toLowerCase();
    for (const attr of [...node.attributes]) {
      const name = attr.name;
      const isHref = name === 'href' && tag === 'a';
      const isSrc = name === 'src' && tag === 'img';
      const isData = name.startsWith('data-');
      if (!isHref && !isSrc && !isData) continue;
      out.push({ node, tag, name, isHref, isSrc, isData });
    }
  }
  return out;
}

export default {
  id: 'text',
  text: true,
  label: 'Text',

  refs(instance, ctx) {
    const html = String(instance?.innerHTML || '');
    if (!html) return [];
    const doc = parseHtml(html);
    if (!doc) return [];
    const out = [];
    collectAttrs(doc).forEach(({ node, name, isSrc, isData }, index) => {
      const value = node.getAttribute(name);
      const path = matchSourcePath(value, ctx);
      if (path == null) return;
      if (isSrc) {
        out.push({
          path: ['innerHTML', `@${name}[${index}]`],
          value,
          class: 'asset',
          asset: { path, ids: {} },
          note: 'image into the source web',
        });
      } else {
        out.push({
          path: ['innerHTML', `@${name}[${index}]`],
          value,
          class: 'link',
          note: isData ? 'data attribute into the source web' : 'link into the source web',
        });
      }
    });
    return out;
  },

  patch(instance, ctx) {
    const html = String(instance?.innerHTML || '');
    if (!html) return 0;
    const doc = parseHtml(html);
    if (!doc) return 0;
    let changed = 0;

    for (const node of doc.body.querySelectorAll('img[src]')) {
      const src = node.getAttribute('src');
      const path = matchSourcePath(src, ctx);
      if (path == null) continue;
      const mapped = ctx.mapAsset ? ctx.mapAsset({ path }) : null;
      if (!mapped) continue;
      const wasAbsolute = /^https?:\/\//i.test(src);
      const next = wasAbsolute ? mapped.url : mapped.path;
      if (!next) continue;
      node.setAttribute('src', next);
      changed += 1;
    }

    for (const node of doc.body.querySelectorAll('a[href]')) {
      const href = node.getAttribute('href');
      const path = matchSourcePath(href, ctx);
      if (path == null) continue;
      const next = ctx.mapLink ? ctx.mapLink(href) : null;
      if (typeof next !== 'string' || !next) continue;
      node.setAttribute('href', next);
      changed += 1;
    }

    if (changed) instance.innerHTML = doc.body.innerHTML;
    return changed;
  },
};
