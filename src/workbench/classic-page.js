// Classic-page model — pure functions, no DOM writes, unit-testable.
//
// The Pages view was written for modern pages: BaseTemplate 119, content in
// CanvasContent1. Classic sites store the same idea three other ways, and the
// two axes are independent:
//
//   library kind (BaseTemplate)   modern 119 | publishing 850 | generic
//   page kind    (per item)       canvas | publishing | wiki | empty
//
// They are independent because a single 119 library can hold both modern
// canvas pages and old wiki pages, so the page kind is decided per item from
// which content field actually carries something — never from the library.
//
// On a publishing page most real content does NOT live in an item field: it
// lives in Content Editor / Script Editor web parts, whose payloads come from
// the page's limited web-part manager. `classicContentParts()` merges the
// body field and those web parts into the SAME part shape `contentParts()`
// produces for canvas pages, so Extract and the content export consume one
// reading model regardless of page kind.

const SITE_PAGES_BASE_TEMPLATE = 119;
const PUBLISHING_PAGES_BASE_TEMPLATE = 850;

// Content fields, in the order they are probed. CanvasContent1 is handled by
// canvas.js; these two are the classic bodies.
export const PUBLISHING_BODY_FIELD = 'PublishingPageContent';
export const WIKI_BODY_FIELD = 'WikiField';

export function libraryKindOf(list) {
  if (!list) return null;
  if (list.BaseTemplate === SITE_PAGES_BASE_TEMPLATE) return 'modern';
  if (list.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE) return 'publishing';
  return 'generic';
}

export const libraryKindLabel = (kind) => ({
  modern: 'modern Site Pages library',
  publishing: 'classic publishing Pages library',
  generic: 'pages library',
}[kind] || 'pages library');

// Does this markup render anything a reader would see? Markup that only
// carries an image still counts (same rule as page-export's text parts).
export function htmlHasContent(html) {
  const raw = String(html ?? '').trim();
  if (!raw) return false;
  if (/<img\b/i.test(raw)) return true;
  const doc = new DOMParser().parseFromString(raw, 'text/html');
  return Boolean((doc.body?.textContent || '').trim());
}

// Which body a given item actually carries. Decided per item, never from the
// library: a 119 library can hold canvas pages and wiki pages side by side.
export function pageContentKindOf(item) {
  const it = item || {};
  if (String(it.CanvasContent1 ?? '').trim()) return 'canvas';
  if (htmlHasContent(it[PUBLISHING_BODY_FIELD])) return 'publishing';
  if (htmlHasContent(it[WIKI_BODY_FIELD])) return 'wiki';
  return 'empty';
}

export const pageContentKindLabel = (kind) => ({
  canvas: 'modern canvas page',
  publishing: 'classic publishing page',
  wiki: 'classic wiki page',
  webparts: 'classic web-part page',
  empty: 'page with no readable body',
}[kind] || 'page');

// CEWP content is sometimes stored wrapped; REST hands back the raw string,
// but a CDATA wrapper survives on some tenants' exports.
function unwrapCdata(value) {
  const raw = String(value ?? '');
  const m = /^\s*<!\[CDATA\[([\s\S]*)\]\]>\s*$/.exec(raw);
  return m ? m[1] : raw;
}

// Rows from getlimitedwebpartmanager(...)/webparts?$expand=WebPart/Properties.
// The REST shape carries no class name, so the kind is inferred from the
// property bag: Content/ContentLink means an HTML-bearing part (Content
// Editor or Script Editor — both surface the same way, and telling them
// apart adds nothing a reader wants).
export function normalizeWebPart(entry, index = 0) {
  const wp = entry?.WebPart || {};
  const props = wp.Properties || {};
  const content = unwrapCdata(props.Content).trim();
  const contentLink = String(props.ContentLink ?? '').trim();
  return {
    id: entry?.Id || wp.Id || `wp-${index}`,
    title: String(wp.Title ?? '').trim(),
    zoneIndex: Number.isFinite(wp.ZoneIndex) ? wp.ZoneIndex : index,
    order: index,
    hidden: Boolean(wp.Hidden),
    closed: Boolean(wp.IsClosed),
    hasHtml: Boolean(content) || Boolean(contentLink),
    content,
    contentLink,
    properties: props,
  };
}

// Document order across zones, stable within a zone.
export function classicWebParts(entries) {
  return (entries || [])
    .map((entry, i) => normalizeWebPart(entry, i))
    .sort((a, b) => (a.zoneIndex - b.zoneIndex) || (a.order - b.order));
}

// The readable part a single web part contributes, or null when it carries
// nothing a reader wants. Linked content lives in another file; the
// reference is the readable fact there — the pad does not follow it.
function webPartPart(wp) {
  const label = wp.title || 'Embedded content';
  if (htmlHasContent(wp.content)) {
    return { kind: 'text', label, html: wp.content, lines: [] };
  }
  if (wp.contentLink) {
    return {
      kind: 'webpart', label, html: '', lines: [`Content linked from ${wp.contentLink}`],
    };
  }
  return null;
}

// First guid found in a string, dashes or underscores (wpbox markers use
// `div_<guid>`; some export shapes use underscores throughout).
function guidOf(value) {
  const text = String(value ?? '').replace(/_/g, '-');
  const m = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.exec(text);
  return m ? m[0].toLowerCase() : '';
}

// Body HTML split at its embedded web-part placeholders. Wiki (and rich
// publishing) bodies place web parts INSIDE the field as `.ms-rte-wpbox`
// markers whose ids carry the definition's storage guid — "intro → web part
// → conclusion" must read in that order, not body-then-web-parts. Each
// placeholder is swapped for a marker element, the serialized HTML is split
// on the markers, and matched definitions (added to `used`) are emitted in
// their true positions. Bodies without placeholders come back as one part.
function bodyParts(bodyHtml, bodyLabel, webParts, used) {
  const raw = String(bodyHtml ?? '').trim();
  const parts = [];
  const pushText = (html) => {
    if (htmlHasContent(html)) parts.push({ kind: 'text', label: bodyLabel, html: html.trim(), lines: [] });
  };
  if (!raw) return parts;

  const doc = new DOMParser().parseFromString(raw, 'text/html');
  const boxes = [...doc.querySelectorAll('.ms-rte-wpbox')];
  if (!boxes.length) { pushText(raw); return parts; }

  const byGuid = new Map();
  for (const wp of webParts) {
    const guid = guidOf(wp.id);
    if (guid && !byGuid.has(guid)) byGuid.set(guid, wp);
  }
  boxes.forEach((box, i) => {
    const marker = doc.createElement('dcspad-wp');
    marker.setAttribute('data-i', String(i));
    box.replaceWith(marker);
  });
  const segments = doc.body.innerHTML.split(/<dcspad-wp data-i="(\d+)"><\/dcspad-wp>/);
  for (let s = 0; s < segments.length; s += 1) {
    if (s % 2 === 0) { pushText(segments[s]); continue; }
    // The detached box still carries its children; its markup holds the guid.
    const box = boxes[Number(segments[s])];
    const wp = byGuid.get(guidOf(box.outerHTML));
    if (wp) {
      used.add(wp);
      const part = webPartPart(wp);
      if (part) parts.push(part);
    }
    // An unmatched placeholder renders nothing itself; if its definition
    // exists it is appended by the caller so nothing is lost.
  }
  return parts;
}

// Ordered reading model of a classic page, in the shape contentParts()
// returns for canvas pages: [{ kind, label, html, lines }], empty parts
// dropped, repeated labels numbered. The body field's fragments and its
// embedded web parts come first in document order, then every remaining
// HTML-bearing web part in zone order.
export function classicContentParts({ item = {}, webParts = [], contentKind = null } = {}) {
  const kind = contentKind || pageContentKindOf(item);
  const bodyField = kind === 'wiki' ? WIKI_BODY_FIELD : PUBLISHING_BODY_FIELD;
  const bodyLabel = kind === 'wiki' ? 'Wiki content' : 'Page content';

  const used = new Set();
  const parts = bodyParts(item[bodyField], bodyLabel, webParts, used);

  for (const wp of webParts) {
    if (used.has(wp) || !wp.hasHtml) continue;
    const part = webPartPart(wp);
    if (part) parts.push(part);
  }

  const counts = new Map();
  for (const part of parts) counts.set(part.label, (counts.get(part.label) || 0) + 1);
  const seen = new Map();
  for (const part of parts) {
    if (counts.get(part.label) > 1) {
      const n = (seen.get(part.label) || 0) + 1;
      seen.set(part.label, n);
      part.label = `${part.label} ${n}`;
    }
  }
  return { parts, unreadable: 0 };
}
