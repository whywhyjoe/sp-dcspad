// Page export builders — pure functions, no DOM writes, unit-testable.
//
// Two artifacts per page:
//   Content (.md)  — for human reading and archiving. Human-oriented
//     metadata on top (title, description, created, location), the merged
//     content of every part in document order under a heading per part, then
//     a standardized metadata block at the bottom. Text parts are converted
//     to real markdown — their own headings, lists, links and tables survive
//     as markdown, not as the HTML a text web part happens to store; other
//     parts contribute whatever searchable text they carry. Parts with
//     nothing to read are skipped — this artifact is for reading, not for
//     inventory.
//   Raw (.json)    — the list item plus the normalized controls, for later
//     script analysis.
//
// `contentParts()` is the shared reading model behind both the content export
// and the Pages → Extract tab, so the two never drift. It deliberately drops
// the technical framing ("web part", ids, control types): those live on the
// Web parts, Structure and Raw tabs.

import { webPartName, textOfControl, sanitizeHtml } from './canvas.js';
import { htmlToMarkdown } from '../html-markdown.js';

const fmtDate = (v) => (v ? String(v).slice(0, 10) : '');

// 'Site | Library' at the library root, 'Site | Library | news/fr' below it.
export function pageLocation({ siteTitle, libraryTitle, fileDirRef, libraryRootPath }) {
  const parts = [siteTitle, libraryTitle].filter(Boolean);
  const dir = String(fileDirRef || '');
  const root = String(libraryRootPath || '').replace(/\/+$/, '');
  let folder = '';
  if (root && dir.toLowerCase().startsWith(root.toLowerCase())) {
    folder = dir.slice(root.length).replace(/^\/+/, '');
  }
  if (folder) parts.push(folder);
  return parts.join(' | ');
}

// A text part is empty when it renders nothing a reader would see; markup that
// only carries an image still counts as content.
function textPartIsEmpty(html) {
  if (!html) return true;
  if (/<img\b/i.test(html)) return false;
  return !textOfControl({ kind: 'text', innerHTML: html });
}

// Ordered reading model of a page's canvas:
//   parts:      [{ kind, label, html, lines }] in document order, empty parts
//               dropped, repeated labels numbered ('Text 1', 'Text 2').
//   unreadable: count of canvas entries the parser could not make sense of.
// `html` is only ever set for text parts; `lines` carries the readable text of
// every other part.
export function contentParts(controls) {
  const parts = [];
  let unreadable = 0;
  for (const control of controls || []) {
    if (control.kind === 'text') {
      const html = String(control.innerHTML || '').trim();
      if (!textPartIsEmpty(html)) parts.push({ kind: 'text', label: 'Text', html, lines: [] });
    } else if (control.kind === 'webpart') {
      const lines = Object.values(
        control.webPartData?.serverProcessedContent?.searchablePlainTexts || {},
      ).filter((v) => typeof v === 'string' && v.trim()).map((v) => v.trim());
      if (!lines.length) continue;   // nothing to read — skipped by design
      const title = String(control.webPartData?.title || '').trim();
      parts.push({
        kind: 'webpart', label: title || webPartName(control.webPartId), html: '', lines,
      });
    } else if (control.kind === 'unknown') {
      unreadable += 1;
    }
    // 'section' and 'pageSettings' entries carry no content.
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
  return { parts, unreadable };
}

// A text part's body as markdown.
//
// Sanitized first: Script Editor payloads reach this path on classic pages,
// and a permissive markdown renderer executes inline HTML. The exact
// unsanitized payload stays available in the raw JSON export.
//
// Some markup has no markdown equivalent at all (a bare video embed, a styled
// container with no text). Rather than drop content silently, a part that
// converts to nothing falls back to the sanitized HTML — the old behaviour,
// now only where markdown genuinely cannot carry the part.
function textPartMarkdown(html) {
  const safe = sanitizeHtml(html);
  return htmlToMarkdown(safe, 'pageContent') || safe;
}

function contentBlocks(controls, override) {
  const { parts, unreadable } = override
    ? { parts: override, unreadable: 0 }
    : contentParts(controls);
  const blocks = [];
  for (const part of parts) {
    blocks.push(`## ${part.label}`);
    blocks.push(part.kind === 'text'
      ? textPartMarkdown(part.html)
      : part.lines.map((t) => `- ${t}`).join('\n'));
  }
  if (unreadable) {
    blocks.push(`*[${unreadable} part${unreadable === 1 ? '' : 's'} could not be read — `
      + 'see the raw export.]*');
  }
  return blocks;
}

// Item fields worth carrying into the standardized metadata block; content
// blobs and odata noise are excluded.
const METADATA_SKIP = new Set([
  'CanvasContent1', 'LayoutWebpartsContent', 'FieldValuesAsText',
  'PublishingPageContent', 'WikiField',
  'Author', 'Editor',   // flattened into Created/Modified lines
]);

export function buildContentExport({
  item = {}, controls = [], parts = null, siteTitle = '', webUrl = '',
  libraryTitle = '', libraryRootPath = '',
}) {
  const title = item.Title || item.FileLeafRef || 'Untitled page';
  const author = item.Author?.Title || '';
  const editor = item.Editor?.Title || '';
  const location = pageLocation({
    siteTitle, libraryTitle, fileDirRef: item.FileDirRef, libraryRootPath,
  });
  let fullUrl = '';
  if (item.FileRef) {
    try { fullUrl = `${new URL(webUrl).origin}${encodeURI(item.FileRef)}`; }
    catch { fullUrl = item.FileRef; }
  }

  const top = [`# ${title}`, ''];
  if (item.Description) top.push(`> ${String(item.Description).replace(/\r?\n/g, ' ')}`, '');
  top.push(`Created ${fmtDate(item.Created)}${author ? ` by ${author}` : ''}  `);
  if (location) top.push(`Location: ${location}`);
  // The '---' below must stay a thematic break: without this blank line it
  // would make the last front-matter line a setext heading instead.
  top.push('');

  const meta = ['## Metadata', ''];
  const metaLine = (label, value) => {
    if (value !== '' && value !== null && value !== undefined) {
      meta.push(`- ${label}: ${value}`);
    }
  };
  metaLine('Title', item.Title);
  metaLine('Description', item.Description);
  metaLine('Created', item.Created
    ? `${item.Created}${author ? ` by ${author}` : ''}` : '');
  metaLine('Modified', item.Modified
    ? `${item.Modified}${editor ? ` by ${editor}` : ''}` : '');
  metaLine('Site', siteTitle);
  metaLine('Library', libraryTitle);
  metaLine('URL', fullUrl);
  for (const [key, value] of Object.entries(item)) {
    if (METADATA_SKIP.has(key) || key.startsWith('odata') || key.startsWith('__')) continue;
    if (['Title', 'Description', 'Created', 'Modified', 'FileRef'].includes(key)) continue;
    if (value === null || value === undefined || typeof value === 'object') continue;
    meta.push(`- ${key}: ${value}`);
  }

  return [
    ...top,
    '---',
    '',
    contentBlocks(controls, parts).join('\n\n'),
    '',
    '---',
    '',
    ...meta,
    '',
  ].join('\n');
}

export function buildRawExport({ item = {}, controls = [], webParts = [] }) {
  const payload = { item, controls };
  if (webParts.length) payload.webParts = webParts;
  return JSON.stringify(payload, null, 2);
}

const slug = (s) => String(s).toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '');

export function exportFileStem(item) {
  const name = String(item.FileLeafRef || item.Title || 'page').replace(/\.aspx$/i, '');
  return slug(name) || 'page';
}

// ---- bulk export (Pages grid → one zip of content markdown) ----

// The path a page takes inside the bundle: its folder relative to the library
// root, then the same '<stem>-content.md' the single-page export writes. Folder
// segments go through the same slug as the stem, so nothing reaches a zip entry
// that could not appear in a file name the pad already produces.
export function bundleEntryName(item, libraryRootPath) {
  const dir = String(item?.FileDirRef || '');
  const root = String(libraryRootPath || '').replace(/\/+$/, '');
  let folder = '';
  if (root && dir.toLowerCase().startsWith(root.toLowerCase())) {
    folder = dir.slice(root.length).replace(/^\/+/, '');
  }
  const segments = folder.split('/').map(slug).filter(Boolean);
  segments.push(`${exportFileStem(item || {})}-content.md`);
  return segments.join('/');
}

// Two pages in one folder can slug to the same name ('Résumé.aspx' and
// 'Resume.aspx' both give 'resume'). A zip with a duplicate entry silently
// loses one on extraction, so the later one is numbered instead.
export function dedupeEntryNames(names) {
  const seen = new Set();
  return (names || []).map((raw) => {
    const name = String(raw);
    if (!seen.has(name)) {
      seen.add(name);
      return name;
    }
    const dot = name.lastIndexOf('.');
    const stem = dot > 0 ? name.slice(0, dot) : name;
    const ext = dot > 0 ? name.slice(dot) : '';
    let n = 2;
    while (seen.has(`${stem}-${n}${ext}`)) n += 1;
    const unique = `${stem}-${n}${ext}`;
    seen.add(unique);
    return unique;
  });
}

// Rides along in the bundle when a page could not be read. Neutral register on
// purpose (see denied.js): a page the account cannot open is a fact about the
// site, and the reason is SharePoint's own sentence rather than our gloss.
export function buildExportReport({ total = 0, exported = 0, failures = [] }) {
  const lines = ['# Export report', '', `${exported} of ${total} pages exported.`, ''];
  if (failures.length) {
    lines.push('Not exported:', '');
    for (const failure of failures) {
      const reason = String(failure.reason || '').replace(/\s+/g, ' ').trim();
      lines.push(`- ${failure.name}${reason ? ` — ${reason}` : ''}`);
    }
    lines.push('');
  }
  return lines.join('\n');
}
