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

// Ordered reading model of a classic page, in the shape contentParts()
// returns for canvas pages: [{ kind, label, html, lines }], empty parts
// dropped, repeated labels numbered. The body field comes first, then each
// HTML-bearing web part in document order.
export function classicContentParts({ item = {}, webParts = [], contentKind = null } = {}) {
  const kind = contentKind || pageContentKindOf(item);
  const parts = [];

  const bodyField = kind === 'wiki' ? WIKI_BODY_FIELD : PUBLISHING_BODY_FIELD;
  const body = String(item[bodyField] ?? '').trim();
  if (htmlHasContent(body)) {
    parts.push({
      kind: 'text',
      label: kind === 'wiki' ? 'Wiki content' : 'Page content',
      html: body,
      lines: [],
    });
  }

  for (const wp of webParts) {
    if (!wp.hasHtml) continue;
    const label = wp.title || 'Embedded content';
    if (htmlHasContent(wp.content)) {
      parts.push({ kind: 'text', label, html: wp.content, lines: [] });
    } else if (wp.contentLink) {
      // Linked content lives in another file; the reference is the readable
      // fact here — the pad does not follow it.
      parts.push({
        kind: 'webpart',
        label,
        html: '',
        lines: [`Content linked from ${wp.contentLink}`],
      });
    }
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
