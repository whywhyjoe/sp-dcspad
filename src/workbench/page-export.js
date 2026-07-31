// Page export builders — pure functions, no DOM writes, unit-testable.
//
// Two artifacts per page:
//   Content (.md)  — for human reading and archiving. Human-oriented
//     metadata on top (title, description, created, location), the merged
//     content of every canvas control in document order, then a
//     standardized metadata block at the bottom. Text web parts contribute
//     their HTML (readable raw AND rendered by md viewers); other web parts
//     contribute a labelled block with whatever searchable text they carry.
//   Raw (.json)    — the list item plus the normalized controls, for later
//     script analysis.

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

function contentBlocks(controls) {
  const blocks = [];
  for (const control of controls || []) {
    if (control.kind === 'text') {
      const html = String(control.innerHTML || '').trim();
      if (html) blocks.push(html);
    } else if (control.kind === 'webpart') {
      const name = webPartName(control.webPartId);
      const title = control.webPartData?.title;
      const label = title && title !== name
        ? `**[Web part: ${name} — “${title}”]**`
        : `**[Web part: ${name}]**`;
      const texts = Object.values(
        control.webPartData?.serverProcessedContent?.searchablePlainTexts || {},
      ).filter((v) => typeof v === 'string' && v.trim());
      blocks.push(texts.length
        ? `${label}\n\n${texts.map((t) => `- ${t}`).join('\n')}`
        : label);
    } else if (control.kind === 'unknown') {
      blocks.push('*[Unparsed canvas entry — see the raw export]*');
    }
    // 'section' and 'pageSettings' entries carry no content.
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
