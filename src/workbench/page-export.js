// Page export builders — pure functions, no DOM writes, unit-testable.
//
// Two artifacts per page:
//   Content (.md)  — for human reading and archiving. Human-oriented
//     metadata on top (title, description, created, location), the merged
//     content of every part in document order under a heading per part, then
//     a standardized metadata block at the bottom. Text parts contribute
//     their HTML (readable raw AND rendered by md viewers); other parts
//     contribute whatever searchable text they carry. Parts with nothing to
//     read are skipped — this artifact is for reading, not for inventory.
//   Raw (.json)    — the list item plus the normalized controls, for later
//     script analysis.
//
// `contentParts()` is the shared reading model behind both the content export
// and the Pages → Extract tab, so the two never drift. It deliberately drops
// the technical framing ("web part", ids, control types): those live on the
// Web parts, Structure and Raw tabs.

import { webPartName, textOfControl } from './canvas.js';

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

function contentBlocks(controls) {
  const { parts, unreadable } = contentParts(controls);
  const blocks = [];
  for (const part of parts) {
    blocks.push(`## ${part.label}`);
    blocks.push(part.kind === 'text' ? part.html : part.lines.map((t) => `- ${t}`).join('\n'));
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
  'Author', 'Editor',   // flattened into Created/Modified lines
]);

export function buildContentExport({
  item = {}, controls = [], siteTitle = '', webUrl = '',
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
  if (location) top.push(`Location: ${location}`, '');

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
    contentBlocks(controls).join('\n\n'),
    '',
    '---',
    '',
    ...meta,
    '',
  ].join('\n');
}

export function buildRawExport({ item = {}, controls = [] }) {
  return JSON.stringify({ item, controls }, null, 2);
}

export function exportFileStem(item) {
  const name = String(item.FileLeafRef || item.Title || 'page').replace(/\.aspx$/i, '');
  return name.toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '') || 'page';
}
