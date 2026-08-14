// List-content markdown export — pure builders, no DOM writes, unit-testable.
//
// The export is a reading document, not a table: each item renders as its
// Title as a heading with one `Field: value` line per column beneath it.
// SharePoint control fields are ignored except the user-facing system set —
// ID, Title, Created, Created By, Modified, Modified By — so the document
// carries the dates, the ID, and all the content, and nothing else.
// Attachments render as absolute links to where the files live. Rich-text
// (Note) fields are converted to markdown so embedded links, emphasis, and
// lists survive; empty fields are skipped.
//
// Column rules (agreed with Joe, 2026-08-13):
//   no view — Title (heading), ID, content columns in server return order,
//             then Created / Created By / Modified / Modified By;
//   a view  — the view supplies the content-column set and order; the
//             system framing (heading, ID, dates) stays identical.

// Types that never carry item content of their own.
const EXCLUDED_TYPES = new Set(['Computed', 'Attachments']);

// Control fields that survive the Hidden/ReadOnly filter on real lists but
// are noise in a content export. Attachments is handled separately.
const EXCLUDED_INTERNAL = new Set([
  'ContentType', 'Attachments', 'ComplianceAssetId', 'AppAuthor', 'AppEditor',
  'Edit', 'DocIcon', 'ItemChildCount', 'FolderChildCount', '_ColorTag',
  '_UIVersionString', 'LinkTitle', 'LinkTitleNoMenu',
  'LinkFilename', 'LinkFilenameNoMenu',
]);

// System fields the export frames every item with; they are emitted in a
// fixed position, never as ordinary content columns.
const SYSTEM_INTERNAL = new Set(['ID', 'Id', 'Title', 'Created', 'Modified', 'Author', 'Editor']);

// Classic views name the title/filename columns by their linked variants.
const VIEW_FIELD_ALIAS = {
  LinkTitle: 'Title',
  LinkTitleNoMenu: 'Title',
  LinkFilename: 'FileLeafRef',
  LinkFilenameNoMenu: 'FileLeafRef',
};

// Read-only fields a view may still legitimately export as content — the
// filename of a library item is content even though SharePoint marks the
// field read-only. Deliberately tiny; grow only with a reason.
const VIEW_ALLOWED_READONLY = new Set(['FileLeafRef']);

// The one content-field predicate: visible, writable, not system framing,
// not control noise. Both column paths run through it so choosing a view
// can only ever *narrow* the exported set, never widen it.
function isContentField(f) {
  return Boolean(f)
    && !f.Hidden && !f.ReadOnlyField
    && !EXCLUDED_TYPES.has(f.TypeAsString)
    && !EXCLUDED_INTERNAL.has(f.InternalName)
    && !SYSTEM_INTERNAL.has(f.InternalName);
}

// Content columns for the no-view export: every content field, in the
// order the server returned the field collection.
export function contentFields(fields) {
  return (fields || []).filter(isContentField);
}

// Content columns for a view export: the view's field names (aliases
// resolved) mapped onto the list's field metadata, view order preserved.
// Only content fields (plus the narrow read-only allow-list) survive — a
// view naming a hidden or read-only control field must not reintroduce it.
export function viewColumnFields(fields, viewFieldNames) {
  const byName = new Map((fields || [])
    .filter((f) => isContentField(f)
      || (f && !f.Hidden && VIEW_ALLOWED_READONLY.has(f.InternalName)))
    .map((f) => [f.InternalName, f]));
  const out = [];
  for (const raw of viewFieldNames || []) {
    const name = VIEW_FIELD_ALIAS[raw] || String(raw);
    if (SYSTEM_INTERNAL.has(name) || EXCLUDED_INTERNAL.has(name)) continue;
    const field = byName.get(name);
    if (field && !out.includes(field)) out.push(field);
  }
  return out;
}

// FieldValuesAsText keys OData-encode underscores in internal names
// (My_Field arrives as My_x005f_Field); try both spellings.
const asTextKey = (name) => String(name).replaceAll('_', '_x005f_');

function textOf(item, internalName) {
  const fvt = item?.FieldValuesAsText;
  if (!fvt || typeof fvt !== 'object') return undefined;
  return fvt[internalName] ?? fvt[asTextKey(internalName)];
}

// Raw-value fallback for items fetched without FieldValuesAsText (mocks,
// hand-built descriptors): render scalars directly and pull the readable
// core out of the common object shapes.
function scalarText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  if (Array.isArray(v)) return v.map(scalarText).filter(Boolean).join(', ');
  if (typeof v === 'object') {
    if (Array.isArray(v.results)) return scalarText(v.results);
    if (v.Title) return String(v.Title);
    if (v.Url) {
      const desc = String(v.Description || '');
      return desc && desc !== v.Url ? `${desc} (${v.Url})` : String(v.Url);
    }
    return '';
  }
  return String(v);
}

export function itemTitle(item) {
  const title = textOf(item, 'Title') ?? scalarText(item?.Title);
  return String(title || '').trim();
}

// Created By / Modified By: display text of the Author/Editor people fields.
export function personText(item, internalName) {
  return String(textOf(item, internalName) ?? scalarText(item?.[internalName]) ?? '').trim();
}

// Single-line plain text for grid cells; filter/sort/CSV all run over this.
export function fieldText(item, field) {
  const name = field.InternalName;
  const raw = item?.[name];
  if (field.TypeAsString === 'Note') {
    const text = textOf(item, name);
    if (text !== undefined) return String(text).replace(/\s+/g, ' ').trim();
    return htmlToMarkdown(typeof raw === 'string' ? raw : '').replace(/\s+/g, ' ').trim();
  }
  if (field.TypeAsString === 'URL') {
    return scalarText(raw) || String(textOf(item, name) ?? '');
  }
  const text = textOf(item, name);
  return String(text !== undefined ? text : scalarText(raw)).trim();
}

// Markdown link with the characters that would break the syntax escaped:
// ']' in the label, parentheses in the target (encodeURIComponent leaves
// them alone, and SharePoint file names may contain them).
const mdLink = (label, url) =>
  `[${String(label).replace(/\]/g, '\\]')}](${String(url).replace(/\(/g, '%28').replace(/\)/g, '%29')})`;

// Decoded server-relative path → href path, one encode per segment.
// encodeURI would leave '#' to become a URL fragment — SharePoint allows
// '#' and '%' in file and folder names.
const encodeSpPath = (path) =>
  String(path).split('/').map(encodeURIComponent).join('/');

// Markdown value for the export document. Rich text keeps its structure;
// URL fields become links; everything else is the plain text value.
export function fieldMarkdown(item, field) {
  const name = field.InternalName;
  const raw = item?.[name];
  if (field.TypeAsString === 'Note') {
    if (typeof raw === 'string' && raw.includes('<')) return htmlToMarkdown(raw);
    return String(raw ?? textOf(item, name) ?? '').trim();
  }
  if (field.TypeAsString === 'URL') {
    if (typeof raw === 'string') return raw.trim();
    const url = String(raw?.Url ?? raw?.url ?? '').trim();
    if (url) {
      const desc = String(raw?.Description ?? raw?.description ?? '').trim();
      return desc && desc !== url ? mdLink(desc, url) : url;
    }
    return String(textOf(item, name) ?? scalarText(raw)).trim();
  }
  return fieldText(item, field);
}

// Absolute markdown links to an item's attachments (expects the
// AttachmentFiles expansion; missing/empty resolves to no links).
export function attachmentLinks(item, origin = '') {
  const value = item?.AttachmentFiles;
  const files = Array.isArray(value) ? value : value?.results || [];
  return files.map((f) => {
    const rel = String(f?.ServerRelativeUrl || '');
    if (!rel) return '';
    const name = String(f?.FileName || rel.split('/').pop() || rel);
    return mdLink(name, `${origin}${encodeSpPath(rel)}`);
  }).filter(Boolean);
}

// ---- HTML → markdown ------------------------------------------------------
// Light conversion tuned for SharePoint rich-text values: paragraphs, line
// breaks, bold/italic, links, images, lists, blockquotes, and pre survive;
// headings flatten to bold lines (real headings would fight the per-item
// `##` structure); nested lists flatten to their top level; tables render
// as pipe-joined rows without a separator line (readable, not re-parseable);
// everything else contributes its text.

const BLOCK_TAGS = new Set([
  'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table',
  'blockquote', 'pre', 'section', 'article', 'header', 'footer', 'hr',
]);

const collapse = (s) => s
  .replace(/[ \t]*\n[ \t]*/g, '\n')
  .replace(/[ \t]{2,}/g, ' ')
  .replace(/^[ \t]+|[ \t]+$/g, '')
  .replace(/^\n+|\n+$/g, '');

function inlineChildren(node) {
  let out = '';
  for (const child of node.childNodes) out += inlineNode(child);
  return out;
}

function inlineNode(node) {
  if (node.nodeType === 3) return String(node.nodeValue).replace(/\s+/g, ' ');
  if (node.nodeType !== 1) return '';
  const tag = node.tagName.toLowerCase();
  if (tag === 'br') return '\n';
  if (tag === 'script' || tag === 'style') return '';
  if (BLOCK_TAGS.has(tag)) {
    const block = blockNode(node);
    return block ? `\n${block}\n` : '';
  }
  const body = inlineChildren(node);
  const core = body.trim();
  if (tag === 'strong' || tag === 'b') return core ? `**${core}**` : '';
  if (tag === 'em' || tag === 'i') return core ? `*${core}*` : '';
  if (tag === 'a') {
    const href = String(node.getAttribute('href') || '');
    const label = core || href;
    return href && !/^javascript:/i.test(href) ? `[${label}](${href})` : label;
  }
  if (tag === 'img') {
    const src = String(node.getAttribute('src') || '');
    return src ? `![${node.getAttribute('alt') || ''}](${src})` : '';
  }
  return body;
}

function blockNode(node) {
  const tag = node.tagName.toLowerCase();
  if (tag === 'hr') return '---';
  if (tag === 'ul' || tag === 'ol') {
    const items = [...node.children].filter((c) => c.tagName?.toLowerCase() === 'li');
    return items.map((li, i) =>
      `${tag === 'ol' ? `${i + 1}.` : '-'} ${collapse(inlineChildren(li)).replace(/\n+/g, ' ')}`)
      .join('\n');
  }
  if (/^h[1-6]$/.test(tag)) {
    const core = collapse(inlineChildren(node)).replace(/\n+/g, ' ');
    return core ? `**${core}**` : '';
  }
  if (tag === 'blockquote') {
    return blockChildren(node).split('\n').map((l) => `> ${l}`).join('\n');
  }
  if (tag === 'pre') {
    return `\`\`\`\n${String(node.textContent).replace(/\s+$/, '')}\n\`\`\``;
  }
  if (tag === 'table') {
    return [...node.querySelectorAll('tr')].map((tr) =>
      [...tr.children].map((td) => collapse(inlineChildren(td)).replace(/\n+/g, ' ')).join(' | '))
      .join('\n');
  }
  return blockChildren(node);
}

// Mixed containers (body, div, p): block children emit paragraphs; runs of
// inline content between them group into paragraphs of their own.
function blockChildren(container) {
  const parts = [];
  let run = '';
  const flush = () => {
    const text = collapse(run);
    if (text) parts.push(text);
    run = '';
  };
  for (const child of container.childNodes) {
    const tag = child.nodeType === 1 ? child.tagName.toLowerCase() : '';
    if (BLOCK_TAGS.has(tag)) {
      flush();
      const block = blockNode(child);
      if (block) parts.push(block);
    } else {
      run += inlineNode(child);
    }
  }
  flush();
  return parts.join('\n\n');
}

let sharedParser = null;

export function htmlToMarkdown(html) {
  const source = String(html ?? '');
  if (!source.trim()) return '';
  if (typeof DOMParser === 'undefined') {
    return source.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  }
  sharedParser = sharedParser || new DOMParser();
  const doc = sharedParser.parseFromString(source, 'text/html');
  return blockChildren(doc.body).replace(/\n{3,}/g, '\n\n').trim();
}

// ---- document builder -----------------------------------------------------

export function buildItemsMarkdown({
  listTitle = 'List', webUrl = '', viewTitle = '',
  items = [], fields = [], viewFieldNames = null,
  filter = '', orderby = '',
} = {}) {
  const columns = viewFieldNames
    ? viewColumnFields(fields, viewFieldNames)
    : contentFields(fields);
  let origin = '';
  try { origin = new URL(webUrl).origin; } catch { /* relative attachment links */ }

  const lines = [`# ${listTitle}`, ''];
  const source = [
    webUrl,
    viewTitle ? `view “${viewTitle}”` : 'all columns',
    filter ? `filter: ${filter}` : '',
    orderby ? `order: ${orderby}` : '',
    `${items.length} item${items.length === 1 ? '' : 's'}`,
  ].filter(Boolean).join(' · ');
  lines.push(source, '');

  // Headings and labels must stay one line, whatever the value contains.
  const oneLine = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();

  for (const item of items) {
    const id = item?.ID ?? item?.Id;
    const title = oneLine(itemTitle(item));
    lines.push(`## ${title || (id !== undefined && id !== null ? `Item ${id}` : 'Item')}`, '');
    // Two trailing spaces = markdown hard break, so the field lines stack.
    // Multiline values render inside a blockquote: field content — even a
    // line starting with '##' — can never masquerade as document structure.
    const put = (label, value) => {
      const v = String(value ?? '').trim();
      if (!v) return;   // empty fields are skipped by design
      if (v.includes('\n')) {
        lines.push(`${oneLine(label)}:`, '',
          ...v.split('\n').map((l) => (l.trim() ? `> ${l}` : '>')), '');
      } else {
        lines.push(`${oneLine(label)}: ${v}  `);
      }
    };
    put('ID', id);
    for (const field of columns) put(field.Title || field.InternalName, fieldMarkdown(item, field));
    const attachments = attachmentLinks(item, origin);
    if (attachments.length) put('Attachments', attachments.join(', '));
    put('Created', textOf(item, 'Created') ?? item?.Created);
    put('Created By', personText(item, 'Author'));
    put('Modified', textOf(item, 'Modified') ?? item?.Modified);
    put('Modified By', personText(item, 'Editor'));
    lines.push('');
  }

  return `${lines.join('\n').replace(/\n{3,}/g, '\n\n').trimEnd()}\n`;
}
