// Page copy core — pure. Builds the source snapshot, decides eligibility and
// names, computes the metadata carry set, inventories references, plans the
// copy, and rewrites the content strings the run saves. No I/O, no DOM beyond
// an optional DOMParser for rich-text attributes (design/PAGE-COPY.md §2.1).
//
// The copy payload is the RAW CanvasContent1 array and raw
// LayoutWebpartsContent, never the parsed display model: canvas.js normalizes
// for reading and would drop properties a page needs. Content that nothing
// patches is sent back byte-for-byte; a string is re-serialized only when an
// analyzer actually changed a value in it.

import { EDITABLE_TYPES, toFormValue, fromItemValue } from './field-editor.js';
import { webPartName } from './canvas.js';

export const SITE_PAGES_TEMPLATE = 119;
export const COPYABLE_LAYOUTS = new Set(['Article', 'Home', 'SingleWebPartAppPage']);
export const MAX_ASSET_BYTES = 50 * 1024 * 1024;
export const PAGE_HEADER_ID = 'cbe7b0a9-3504-44dd-a3a3-0e5cacd07788';

// Decision 2 (§4.2), settled by the spike (§1, §11): the sitepages API,
// created by path A (create in the root under a staging name, then move).
// CopyFileByPath lands a copy of a promoted page still promoted and in the
// news query before any reset can run, and addTemplateFile (path B) makes a
// file without the Site Page content type that the sitepages API refuses.
// 'copyFile' survives only for legacy-HTML canvases on the same web.
export const DEFAULT_ENGINE = 'api';

// §5.5 step 2: fields the copy operation itself owns. The carry set never
// writes these, whatever their type says.
export const OPERATION_OWNED_FIELDS = new Set([
  'Title', 'Description', 'BannerImageUrl', 'BannerThumbnailUrl', 'PromotedState',
  'FirstPublishedDate', 'PageLayoutType', 'CanvasContent1', 'LayoutWebpartsContent',
  '_TopicHeader', '_AuthorByline', '_SPSitePageFlags', 'ContentTypeId', 'FileLeafRef',
  'FileRef', 'Author', 'Editor', 'Created', 'Modified', 'CommentsDisabled',
  // Not in the spec's list but equally operation-owned: the item's own
  // bookkeeping, which VULI refuses or silently ignores.
  'ContentType', 'ID', 'Id', '_UIVersionString', 'CheckoutUser', '_ModerationStatus',
  'OData__ModerationStatus', '_ModerationComments',
]);

// People, lookups and managed metadata have no string convention the editor
// can round-trip (field-editor.js), so the API engine documents them as not
// carried (§4.1) rather than guessing.
const UNCARRIED_TYPES = new Set([
  'User', 'UserMulti', 'Lookup', 'LookupMulti', 'TaxonomyFieldType', 'TaxonomyFieldTypeMulti',
]);

const GUID = /[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}/i;
const lower = (v) => String(v ?? '').toLowerCase();
const trimSlash = (v) => String(v ?? '').replace(/\/+$/, '');
export const normalizeGuid = (v) => {
  const m = GUID.exec(String(v ?? ''));
  if (!m) return '';
  const hex = m[0].replace(/-/g, '').toLowerCase();
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
};

function parseJsonArray(raw) {
  const text = String(raw ?? '').trim();
  if (!text) return { ok: true, value: [], empty: true };
  if (text.startsWith('<')) return { ok: false, legacy: true };
  try {
    const value = JSON.parse(text);
    return Array.isArray(value) ? { ok: true, value } : { ok: false };
  } catch {
    return { ok: false };
  }
}

// ---- snapshot (§2.2) --------------------------------------------------------

// page: the sitepages DTO (content of record). item: the list item read with
// $select=*&$expand=FieldValuesAsText. fields: the library's field schema.
// status: derivePageStatus() output. library: { id, title, baseTemplate,
// hidden, rootPath }. web: sp-pages webIdentity().
export function snapshotFromReads({ page = {}, item = {}, fields = [], status = null, library = {}, web = {} }) {
  const canvasRaw = page.CanvasContent1 ?? item.CanvasContent1 ?? '';
  const layoutRaw = page.LayoutWebpartsContent ?? item.LayoutWebpartsContent ?? '';
  const canvas = parseJsonArray(canvasRaw);
  const layout = parseJsonArray(layoutRaw);
  const pageSettings = canvas.ok
    ? canvas.value.find((c) => c && typeof c === 'object' && c.pageSettingsSlice)?.pageSettingsSlice || null
    : null;
  // On a custom-thumbnail page the DTO's BannerImageUrl IS the thumbnail
  // (the banner itself lives in the header part's imageSources), and
  // BannerThumbnailUrl is a tokened afdcache CDN URL — never copied (§11).
  const bannerThumb = String(page.BannerThumbnailUrl || '');
  return {
    web: { ...web },
    library: { ...library },
    pageId: page.Id ?? item.Id,
    // No OData etag under nometadata; the version label plus Modified
    // changes on every save, which is all the binding needs.
    etag: `${item.OData__UIVersionString ?? item._UIVersionString ?? ''}|${item.Modified ?? page.Modified ?? ''}`,
    dto: {
      Title: page.Title ?? item.Title ?? '',
      Description: page.Description ?? item.Description ?? '',
      TopicHeader: page.TopicHeader ?? '',
      AuthorByline: Array.isArray(page.AuthorByline) ? [...page.AuthorByline] : [],
      PageLayoutType: page.PageLayoutType ?? item.PageLayoutType ?? '',
      PromotedState: Number(page.PromotedState ?? item.PromotedState ?? 0),
      BannerImageUrl: typeof page.BannerImageUrl === 'string'
        ? page.BannerImageUrl : String(page.BannerImageUrl?.Url || item.BannerImageUrl?.Url || ''),
      BannerThumbnailUrl: bannerThumb,
      FileName: page.FileName || item.FileLeafRef || '',
    },
    canvasRaw: String(canvasRaw ?? ''),
    canvasFormat: canvas.ok ? 'json' : (canvas.legacy ? 'html' : 'invalid'),
    canvas: canvas.ok ? canvas.value : null,
    layoutRaw: String(layoutRaw ?? ''),
    layout: layout.ok ? layout.value : null,
    pageSettings,
    customThumbnail: Boolean(pageSettings && pageSettings.isDefaultThumbnail === false),
    item,
    itemAsText: item.FieldValuesAsText || {},
    fields: [...fields],
    status,
    fileRef: item.FileRef || '',
    fileDirRef: item.FileDirRef || '',
    fileName: item.FileLeafRef || page.FileName || '',
    // The DTO is authoritative: on SPO the list item's CommentsDisabled
    // column kept reading false after SetCommentsDisabled took effect (§11).
    commentsDisabled: typeof page.CommentsDisabled === 'boolean' ? page.CommentsDisabled
      : typeof item.CommentsDisabled === 'boolean' ? item.CommentsDisabled : null,
  };
}

// ---- eligibility (§2.4) -----------------------------------------------------

// { ok, reason, legacyHtml }. legacyHtml pages copy on the same web through
// CopyFileByPath only; a cross-site copy of one is refused.
export function sourceEligibility(snapshot, { sameWeb = true } = {}) {
  const lib = snapshot.library || {};
  if (Number(lib.baseTemplate) !== SITE_PAGES_TEMPLATE) {
    return { ok: false, reason: 'Only pages in a modern Site Pages library can be copied.' };
  }
  // pickPagesLibrary ranks hidden 119 libraries too; they are vestigial, not
  // copy sources (§2.4).
  if (lib.hidden) {
    return { ok: false, reason: 'This Site Pages library is hidden, so its pages are not copied.' };
  }
  if (snapshot.canvasFormat === 'html') {
    return sameWeb
      ? { ok: true, legacyHtml: true, reason: '' }
      : { ok: false, legacyHtml: true, reason: 'This page uses the legacy HTML canvas format, which can only be duplicated within its own site.' };
  }
  if (snapshot.canvasFormat !== 'json') {
    return { ok: false, reason: 'This page’s canvas content could not be read as JSON, so it cannot be copied safely.' };
  }
  const layout = snapshot.dto.PageLayoutType;
  if (!COPYABLE_LAYOUTS.has(layout)) {
    return { ok: false, reason: `Pages with the ${layout || 'unknown'} layout are not supported (Article, Home and single-part app pages only).` };
  }
  return { ok: true, legacyHtml: false, reason: '' };
}

// The destination must carry the modern Site Pages feature and a 119 library.
export function targetEligibility({ featureActive, library }) {
  if (!featureActive) return { ok: false, reason: 'The destination site does not have the modern Site Pages feature active.' };
  if (!library) return { ok: false, reason: 'The destination site has no Site Pages library.' };
  return { ok: true, reason: '' };
}

// ---- naming (§1 decision 1, §3) ---------------------------------------------

const BAD_NAME = /["*:<>?/\\|#%]/;

export function fileStem(fileName) {
  return String(fileName || '').replace(/\.aspx$/i, '');
}

// '' when valid, else the sentence to show.
export function fileNameProblem(fileName) {
  const name = String(fileName || '').trim();
  if (!/\.aspx$/i.test(name) || !fileStem(name)) return 'The file name must end in .aspx.';
  if (BAD_NAME.test(name)) return 'File names cannot contain " * : < > ? / \\ | # %.';
  if (name.startsWith('.') || /\.\.aspx$/i.test(name)) return 'File names cannot start with a dot or end with one before .aspx.';
  if (name.length > 128) return 'The file name is too long (128 characters at most).';
  return '';
}

export function folderProblem(folder) {
  const clean = String(folder || '').replace(/^\/+|\/+$/g, '');
  if (!clean) return '';
  for (const segment of clean.split('/')) {
    if (!segment.trim() || /["*:<>?\\|]/.test(segment) || segment.startsWith('.') || segment.endsWith('.')) {
      return `“${segment}” is not a valid folder name.`;
    }
  }
  return '';
}

// Keep the source name; `{stem}-copy`, then `-copy-2`… only when taken.
// `taken` holds lowercase file names already present where the page lands.
export function defaultFileName(sourceFileName, taken = new Set()) {
  const stem = fileStem(sourceFileName) || 'Page';
  const has = (n) => taken.has(lower(n));
  if (!has(`${stem}.aspx`)) return `${stem}.aspx`;
  if (!has(`${stem}-copy.aspx`)) return `${stem}-copy.aspx`;
  for (let i = 2; i < 1000; i += 1) {
    if (!has(`${stem}-copy-${i}.aspx`)) return `${stem}-copy-${i}.aspx`;
  }
  return `${stem}-copy-${Date.now()}.aspx`;
}

// Path A's temporary root-level stem. Collision-safe by construction; the
// run still reads the name back, because SharePoint derives it from a Title.
export function stagingStem(stem, shortId) {
  return `${stem}~copy-${shortId}`;
}

export function shortRunId(random = Math.random) {
  return Math.floor(random() * 0x100000000).toString(16).padStart(8, '0').slice(0, 6);
}

// ---- metadata carry set (§5.5) ----------------------------------------------

const choicesOf = (field) => {
  const v = field?.Choices;
  const arr = Array.isArray(v) ? v : v?.results;
  return Array.isArray(arr) ? arr : [];
};

function itemValueOf(item, field) {
  for (const key of [field.EntityPropertyName, field.InternalName, `OData_${field.InternalName}`]) {
    if (key && Object.prototype.hasOwnProperty.call(item || {}, key)) return item[key];
  }
  return undefined;
}

function textValueOf(itemAsText, field) {
  const name = String(field.InternalName || '');
  const encoded = name.replace(/_/g, '_x005f_');
  for (const key of [field.EntityPropertyName, name, encoded]) {
    if (key && Object.prototype.hasOwnProperty.call(itemAsText || {}, key)) return itemAsText[key];
  }
  return undefined;
}

const isEmptyUi = (v) => v === '' || v === null || v === undefined
  || (Array.isArray(v) && !v.length) || (typeof v === 'object' && !Array.isArray(v) && 'url' in v && !v.url);

// sourceFields/targetFields: library field schemas. item/itemAsText: the
// source list item. requiredValues: { InternalName: uiValue } supplied by the
// operator for required-field gaps. dateText(field, item, itemAsText) → the
// FieldValue string for a DateTime field; the default is the source's own
// display text, which ValidateUpdateListItem reads in the web's regional
// settings (ISO strings are refused on live SPO).
export function computeCarrySet({
  sourceFields = [], targetFields = [], item = {}, itemAsText = {}, requiredValues = {},
  dateText = null, enabled = true,
} = {}) {
  const carried = [];
  const skipped = [];
  const formValues = [];
  const byName = new Map(targetFields.map((f) => [String(f.InternalName), f]));
  const writable = (f) => f && !f.Hidden && !f.ReadOnlyField && !OPERATION_OWNED_FIELDS.has(String(f.InternalName));
  const carriedNames = new Set();

  if (enabled) {
    for (const src of sourceFields) {
      const name = String(src.InternalName || '');
      if (!writable(src)) continue;
      const tgt = byName.get(name);
      const label = src.Title || name;
      if (!tgt) {
        // Worth saying only when there was something to lose.
        const had = itemValueOf(item, src);
        if (had !== null && had !== undefined && had !== '' && !(Array.isArray(had) && !had.length)) {
          skipped.push({ internalName: name, title: label, reason: 'the destination has no such column' });
        }
        continue;
      }
      const type = String(src.TypeAsString || '');
      if (!writable(tgt)) { skipped.push({ internalName: name, title: label, reason: 'read-only or hidden on the destination' }); continue; }
      if (String(tgt.TypeAsString || '') !== type) {
        skipped.push({ internalName: name, title: label, reason: `type differs (${type} → ${tgt.TypeAsString})` });
        continue;
      }
      const raw = itemValueOf(item, src);
      if (UNCARRIED_TYPES.has(type)) {
        if (raw !== null && raw !== undefined && raw !== '') {
          skipped.push({ internalName: name, title: label, reason: 'people, lookup and managed metadata columns are not copied' });
        }
        continue;
      }
      if (!EDITABLE_TYPES.has(type)) continue;
      const ui = fromItemValue(src, raw);
      if (isEmptyUi(ui)) continue;
      if (type === 'Choice' || type === 'MultiChoice') {
        const allowed = new Set(choicesOf(tgt));
        const values = Array.isArray(ui) ? ui : [ui];
        const missing = values.filter((v) => !allowed.has(v));
        if (missing.length && !tgt.FillInChoice) {
          skipped.push({ internalName: name, title: label, reason: `choice ${missing.map((v) => `“${v}”`).join(', ')} does not exist on the destination` });
          continue;
        }
      }
      let value;
      if (type === 'DateTime') {
        value = dateText ? dateText(src, item, itemAsText) : textValueOf(itemAsText, src);
        // Never fall back to toFormValue's ISO string — live SPO refuses it.
        if (value === undefined || value === null || value === '') {
          skipped.push({ internalName: name, title: label, reason: 'the date could not be read in the site’s own format' });
          continue;
        }
      } else {
        value = toFormValue(tgt, ui);
      }
      if (value === '') continue;
      formValues.push({ FieldName: name, FieldValue: String(value) });
      carried.push({ internalName: name, title: label, value: String(value) });
      carriedNames.add(name);
    }
  }

  // §5.5 step 5: required destination fields the carry leaves empty. Title
  // is written by the save itself, so it is never a gap.
  const requiredGaps = [];
  for (const tgt of targetFields) {
    const name = String(tgt.InternalName || '');
    if (!tgt.Required || !writable(tgt) || carriedNames.has(name)) continue;
    if (!EDITABLE_TYPES.has(String(tgt.TypeAsString || ''))) {
      requiredGaps.push({ internalName: name, title: tgt.Title || name, type: tgt.TypeAsString, supportable: false });
      continue;
    }
    const supplied = requiredValues[name];
    if (!isEmptyUi(supplied) && supplied !== undefined) {
      const value = toFormValue(tgt, supplied);
      if (value !== '') {
        formValues.push({ FieldName: name, FieldValue: value });
        carried.push({ internalName: name, title: tgt.Title || name, value, supplied: true });
        continue;
      }
    }
    requiredGaps.push({ internalName: name, title: tgt.Title || name, type: tgt.TypeAsString, supportable: true });
  }
  return { formValues, carried, skipped, requiredGaps };
}

// ---- link matching (§5.2 "link") --------------------------------------------

// Path-boundary prefix test: '/sites/a' matches '/sites/a', '/sites/a/x',
// '/sites/a?q', '/sites/a#h' — never '/sites/a-other'. Case-insensitive, as
// SharePoint paths are.
export function underPath(path, base) {
  const p = lower(path);
  const b = lower(trimSlash(base));
  if (!b) return p.startsWith('/');
  if (!p.startsWith(b)) return false;
  const next = p.charAt(b.length);
  return next === '' || next === '/' || next === '?' || next === '#';
}

// A link into the source web, relative or absolute on the source origin →
// the same link into the destination web; anything else → null.
export function rewriteLink(value, { fromPath, toPath, origin }) {
  const raw = String(value ?? '');
  if (!raw) return null;
  if (raw.startsWith('/') && !raw.startsWith('//')) {
    return underPath(raw, fromPath) ? `${trimSlash(toPath)}${raw.slice(trimSlash(fromPath).length)}` : null;
  }
  let url;
  try { url = new URL(raw); } catch { return null; }
  if (!origin || url.origin.toLowerCase() !== lower(origin)) return null;
  const path = decodeSafe(url.pathname);
  if (!underPath(path, fromPath)) return null;
  const rest = path.slice(trimSlash(fromPath).length);
  return `${url.origin}${encodePath(`${trimSlash(toPath)}${rest}`)}${url.search}${url.hash}`;
}

function decodeSafe(s) { try { return decodeURIComponent(s); } catch { return s; } }
function encodePath(p) { return p.split('/').map((seg) => encodeURIComponent(seg)).join('/'); }

// ---- reference inventory (§5.1 generic scanner — discovers, never rewrites) --

const ID_KEY = /^(site|web|list|unique|file|doc|library|item|term|group|view)(id|guid)$|^(siteid|webid|listid|uniqueid)$/i;

function attrValues(html) {
  const out = [];
  if (!html || !/[<]/.test(html)) return out;
  if (typeof DOMParser === 'function') {
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    for (const node of doc.body.querySelectorAll('*')) {
      for (const attr of node.attributes) {
        if (attr.name === 'href' || attr.name === 'src' || attr.name.startsWith('data-')) {
          out.push({ attr: attr.name, tag: node.tagName.toLowerCase(), value: attr.value });
        }
      }
    }
    return out;
  }
  const re = /\s(href|src|data-[\w-]+)\s*=\s*("([^"]*)"|'([^']*)')/gi;
  let m;
  while ((m = re.exec(html))) out.push({ attr: m[1].toLowerCase(), tag: '', value: m[3] ?? m[4] ?? '' });
  return out;
}

function classifyString(value, key, ctx) {
  const s = String(value);
  const hits = [];
  if (/getpreview\.ashx/i.test(s) && /guidFile=/i.test(s)) {
    hits.push('preview');
  } else if (s.startsWith('/') && !s.startsWith('//')) {
    if (underPath(s, ctx.webPath)) hits.push('path');
  } else if (/^https?:\/\//i.test(s)) {
    try {
      const url = new URL(s);
      if (url.origin.toLowerCase() === lower(ctx.origin)) {
        hits.push(underPath(decodeSafe(url.pathname), ctx.webPath) ? 'url' : 'host-url');
      }
    } catch { /* not a URL */ }
  }
  if (GUID.test(s) && !hits.length) {
    const g = normalizeGuid(s);
    const known = ctx.ids.has(g);
    if (known || (ID_KEY.test(String(key || '')) && s.length <= 40)) hits.push('guid');
  }
  return hits;
}

function walk(value, path, visit) {
  if (typeof value === 'string') { visit(value, path); return; }
  if (Array.isArray(value)) { value.forEach((v, i) => walk(v, [...path, i], visit)); return; }
  if (value && typeof value === 'object') {
    for (const [k, v] of Object.entries(value)) walk(v, [...path, k], visit);
  }
}

// Instance identity for the report: web parts by instance id, text parts too.
export function describeControl(control, index) {
  const webPartId = normalizeGuid(control?.webPartId || control?.webPartData?.id || '');
  const instanceId = String(control?.id || control?.webPartData?.instanceId || control?.instanceId || '');
  const kind = control?.pageSettingsSlice ? 'pageSettings'
    : control?.controlType === 4 ? 'text'
      : control?.controlType === 3 ? 'webpart' : 'other';
  const label = kind === 'text' ? 'Text'
    : kind === 'webpart' ? (control?.webPartData?.title || webPartName(webPartId) || webPartId)
      : kind;
  return { index, instanceId, webPartId, kind, label };
}

// Candidates only: { where, index, path, value, kind, attr?, control }.
// kind: 'path' | 'url' | 'host-url' | 'preview' | 'guid'.
export function inventoryReferences(snapshot) {
  const web = snapshot.web || {};
  let origin = '';
  try { origin = new URL(web.webUrl).origin; } catch { /* keep '' */ }
  const ids = new Set([web.siteId, web.webId, snapshot.library?.id].map(normalizeGuid).filter(Boolean));
  const ctx = { origin, webPath: web.webServerRelativeUrl || '', ids };
  const out = [];
  const scan = (where, root, control, basePath) => {
    walk(root, basePath, (value, path) => {
      const key = path[path.length - 1];
      for (const kind of classifyString(value, key, ctx)) out.push({ where, path, value, kind, control });
      if (key === 'innerHTML' || /</.test(value)) {
        attrValues(value).forEach((a, i) => {
          for (const kind of classifyString(a.value, a.attr, ctx)) {
            out.push({ where, path: [...path, `@${a.attr}[${i}]`], value: a.value, kind, attr: a.attr, control });
          }
        });
      }
    });
  };
  (snapshot.canvas || []).forEach((control, index) => {
    scan('canvas', control, describeControl(control, index), [index]);
  });
  (snapshot.layout || []).forEach((part, index) => {
    scan('layout', part, { index, instanceId: String(part?.instanceId || part?.id || ''), webPartId: normalizeGuid(part?.id), kind: 'layout', label: part?.title || 'Title area' }, [index]);
  });
  for (const key of ['BannerImageUrl']) {
    const value = snapshot.dto?.[key];
    if (!value) continue;
    for (const kind of classifyString(value, key, ctx)) {
      out.push({ where: 'dto', path: [key], value, kind, control: { index: -1, instanceId: '', webPartId: '', kind: 'dto', label: key } });
    }
  }
  return out;
}

// ---- preview URLs and asset identities (§5.3) -------------------------------

export function previewIds(url) {
  let parsed;
  try { parsed = new URL(String(url), 'https://placeholder.invalid'); } catch { return null; }
  if (!/getpreview\.ashx$/i.test(parsed.pathname)) return null;
  const q = (k) => normalizeGuid(parsed.searchParams.get(k) || parsed.searchParams.get(k.toLowerCase()) || '');
  const ids = { siteId: q('guidSite'), webId: q('guidWeb'), uniqueId: q('guidFile') };
  return ids.uniqueId ? ids : null;
}

// Identity key for dedupe: the file's UniqueId when known, else its path.
export function assetKey({ uniqueId, path }) {
  return normalizeGuid(uniqueId) || lower(path);
}

// Assets the analyzers asked for, plus the DTO banner/thumbnail, deduped.
// Returns [{ key, path, ids, refs: [{ where, index, path }] }] for the
// preflight to resolve through the source's fileInfo/fileInfoById.
export function assetRequests(snapshot, analysis) {
  const byKey = new Map();
  const add = (identity, ref) => {
    const key = assetKey({ uniqueId: identity.ids?.uniqueId, path: identity.path });
    if (!key) return;
    if (!byKey.has(key)) byKey.set(key, { key, path: identity.path || '', ids: { ...(identity.ids || {}) }, refs: [] });
    const entry = byKey.get(key);
    if (!entry.path && identity.path) entry.path = identity.path;
    entry.refs.push(ref);
  };
  for (const part of analysis.parts) {
    for (const ref of part.refs) {
      if (ref.class === 'asset' && ref.asset) add(ref.asset, { where: part.where, index: part.index, path: ref.path });
    }
  }
  // The DTO's BannerImageUrl: a getpreview.ashx URL (resolved by its GUIDs)
  // or a plain file URL — the custom thumbnail, when the page has one.
  const banner = snapshot.dto.BannerImageUrl;
  if (banner) {
    const ids = previewIds(banner);
    let path = '';
    if (!ids) {
      try {
        const url = new URL(banner, 'https://placeholder.invalid');
        const sameHost = url.origin === 'https://placeholder.invalid' || url.origin.toLowerCase() === lower(originOf(snapshot.web.webUrl));
        if (sameHost) path = decodeSafe(url.pathname);
      } catch { /* not a URL */ }
    }
    if (ids || path) add({ ids: ids || {}, path }, { where: 'dto', index: -1, path: ['BannerImageUrl'] });
  }
  return [...byKey.values()];
}

// ---- analysis (analyzers classify and authorize patches, §5.1–5.2) ----------

// analyzers: { forId(webPartId) → analyzer, header, text }, where an
// analyzer is { id, refs(instance, ctx) → [ref], patch(instance, ctx) }.
// A ref: { path (inside the instance), value, class: 'asset'|'data'|'config'
// |'link'|'dynamic'|'unverified', asset?: { path, ids }, note }.
export function analyzeParts(snapshot, { analyzers = null, targetWebParts = null } = {}) {
  const ctx = analyzerContext(snapshot);
  const findings = inventoryReferences(snapshot);
  const parts = [];
  const consider = (where, instance, index) => {
    const control = where === 'layout'
      ? { index, instanceId: String(instance?.instanceId || instance?.id || ''), webPartId: normalizeGuid(instance?.id), kind: 'layout', label: instance?.title || 'Title area' }
      : describeControl(instance, index);
    if (control.kind === 'pageSettings' || control.kind === 'other') return;
    const analyzer = analyzers
      ? (where === 'layout' ? analyzers.header : control.kind === 'text' ? analyzers.text : analyzers.forId(control.webPartId))
      : null;
    const scanned = findings.filter((f) => f.where === where && f.control.index === index);
    let refs;
    if (analyzer) {
      try { refs = analyzer.refs(instance, ctx) || []; }
      catch (err) { refs = [{ path: [], value: '', class: 'unverified', note: `analyzer failed: ${err?.message || err}` }]; }
    } else {
      refs = scanned.map((f) => ({ path: f.path.slice(1), value: f.value, class: 'unverified', note: f.kind }));
    }
    const available = control.kind !== 'webpart' || !targetWebParts || targetWebParts.has(control.webPartId);
    // Dynamic data rides on the web part's own envelope, whoever analyzes it.
    const dyn = instance?.webPartData?.dynamicDataPaths || instance?.dynamicDataPaths;
    const providers = [];
    if (dyn && typeof dyn === 'object') {
      for (const [k, v] of Object.entries(dyn)) {
        const m = /WebPart\.([0-9a-f-]{36})/i.exec(String(v));
        if (m) providers.push(m[1].toLowerCase());
        refs.push({ path: ['webPartData', 'dynamicDataPaths', k], value: String(v), class: 'dynamic', note: 'connected to another part on this page' });
      }
    }
    parts.push({
      where, index, ...control, analyzer: analyzer?.id || 'default', refs, available,
      providers, unverified: scanned.filter((f) => !refs.some((r) => sameRefPath(r.path, f.path.slice(1)))),
    });
  };
  (snapshot.canvas || []).forEach((c, i) => consider('canvas', c, i));
  (snapshot.layout || []).forEach((p, i) => consider('layout', p, i));
  return { parts, findings };
}

function sameRefPath(a, b) {
  return a.length === b.length && a.every((v, i) => String(v) === String(b[i]));
}

export function analyzerContext(snapshot) {
  const web = snapshot.web || {};
  let origin = '';
  try { origin = new URL(web.webUrl).origin; } catch { /* keep '' */ }
  return {
    origin,
    sourceWebPath: web.webServerRelativeUrl || '',
    sourceIds: { siteId: normalizeGuid(web.siteId), webId: normalizeGuid(web.webId) },
    normalizeGuid,
    underPath,
  };
}

// Consumers that lose their provider when the operator drops it (§5.2
// "dynamic"): every part whose dynamicDataPaths names a dropped instance.
export function dependentConsumers(analysis, droppedIds) {
  const dropped = new Set([...droppedIds].map(lower));
  return analysis.parts.filter((p) => p.providers.some((id) => dropped.has(id)) && !dropped.has(lower(p.instanceId)));
}

// ---- plan (§3) ---------------------------------------------------------------

// target: { webUrl, webServerRelativeUrl, siteId, webId, library: { id,
// rootPath }, siteAssetsRoot, fields, webParts: Set|null, takenFinal: Set,
// takenRoot: Set, assetFolderExists: bool }.
// options: { fileName, title, folder, publish, promoteAsNews, carryMetadata,
// rewriteLinks, engine, dropped, requiredValues, dateText,
// checkInDraft, runId }.
// assetInfo: Map(key → fileInfo | null) resolved by the preflight.
export function analyzeCopy({ snapshot, target, options = {}, analyzers = null, assetInfo = new Map() }) {
  const sameWeb = normalizeGuid(snapshot.web.webId) === normalizeGuid(target.webId)
    && Boolean(normalizeGuid(target.webId));
  const sameSite = normalizeGuid(snapshot.web.siteId) === normalizeGuid(target.siteId)
    && Boolean(normalizeGuid(target.siteId));
  const blockers = [];
  const warnings = [];

  const eligibility = sourceEligibility(snapshot, { sameWeb });
  if (!eligibility.ok) blockers.push(eligibility.reason);

  const legacy = Boolean(eligibility.legacyHtml);
  const engine = legacy ? 'copyFile' : (options.engine || DEFAULT_ENGINE);
  const createPath = engine === 'api' ? 'A' : null;
  if (engine === 'copyFile' && !sameWeb) blockers.push('Whole-file copy only runs within one site.');

  const folder = String(options.folder || '').replace(/^\/+|\/+$/g, '');
  const folderIssue = folderProblem(folder);
  if (folderIssue) blockers.push(folderIssue);
  const fileName = String(options.fileName || '').trim()
    || defaultFileName(snapshot.fileName, target.takenFinal || new Set());
  const nameIssue = fileNameProblem(fileName);
  if (nameIssue) blockers.push(nameIssue);
  if ((target.takenFinal || new Set()).has(lower(fileName))) {
    blockers.push(`${fileName} already exists in the destination folder.`);
  }
  const stem = fileStem(fileName);
  const runId = options.runId || shortRunId();
  const stage = stagingStem(stem, runId);
  if (createPath === 'A' && (target.takenRoot || new Set()).has(lower(`${stage}.aspx`))) {
    blockers.push('The staging name is taken in the library root — check again to pick another.');
  }
  const libraryRoot = trimSlash(target.library?.rootPath || '');
  const finalDir = folder ? `${libraryRoot}/${folder}` : libraryRoot;
  let origin = '';
  try { origin = new URL(target.webUrl).origin; } catch { /* keep '' */ }
  const abs = (sr) => `${origin}${sr}`;

  const title = String(options.title ?? snapshot.dto.Title ?? '').trim() || stem;

  // Parts, assets and links — only a cross-web copy moves anything.
  const analysis = analyzeParts(snapshot, { analyzers: sameWeb ? null : analyzers, targetWebParts: sameWeb ? null : target.webParts });
  // Within one web every reference still resolves, so nothing is suspicious:
  // the report keeps only the connections a per-part drop would break.
  if (sameWeb) {
    for (const part of analysis.parts) {
      part.refs = part.refs.filter((r) => r.class === 'dynamic');
      part.unverified = [];
    }
  }
  const dropped = new Set([...(options.dropped || [])].map(lower));
  const assets = [];
  const webSR = trimSlash(target.webServerRelativeUrl || '');
  const siteAssetsRoot = trimSlash(target.siteAssetsRoot || `${webSR}/SiteAssets`);
  const assetFolder = `${siteAssetsRoot}/SitePages/${stem}`;
  if (!sameWeb) {
    const usedNames = new Set();
    for (const req of assetRequests(snapshot, analysis)) {
      const info = assetInfo.get(req.key) ?? null;
      const name = info?.Name || String(req.path || '').split('/').pop() || 'file';
      let retainReason = '';
      if (!info) retainReason = 'the file could not be found in the source site';
      else if (Number(info.Length) > MAX_ASSET_BYTES) retainReason = 'larger than the 50 MB copy limit';
      const destName = retainReason ? '' : uniqueName(name, usedNames);
      assets.push({
        key: req.key,
        name,
        sourcePath: info?.ServerRelativeUrl || req.path || '',
        sourceIds: {
          siteId: normalizeGuid(info?.SiteId || req.ids.siteId),
          webId: normalizeGuid(info?.WebId || req.ids.webId),
          listId: normalizeGuid(info?.ListId || req.ids.listId),
          uniqueId: normalizeGuid(info?.UniqueId || req.ids.uniqueId),
        },
        length: Number(info?.Length) || 0,
        destName,
        destPath: destName ? `${assetFolder}/${destName}` : '',
        method: 'bytes',
        retainReason,
        refs: req.refs,
      });
      if (retainReason) warnings.push(`${name} stays linked to the source site (${retainReason}).`);
    }
    for (const part of analysis.parts) {
      if (!part.available && !dropped.has(lower(part.instanceId))) {
        warnings.push(`${part.label} is not available on the destination site; it is copied but may not render.`);
      }
    }
  }
  for (const consumer of dependentConsumers(analysis, dropped)) {
    warnings.push(`${consumer.label} is connected to a part you are dropping and will lose its data source.`);
  }

  const metadata = computeCarrySet({
    sourceFields: snapshot.fields,
    targetFields: target.fields || [],
    item: snapshot.item,
    itemAsText: snapshot.itemAsText,
    requiredValues: options.requiredValues || {},
    dateText: options.dateText || null,
    enabled: options.carryMetadata !== false,
  });
  // Description is operation-owned (never in the carry set), but savepage
  // cannot set it on SPO — so the run writes it with the metadata (§11).
  const description = String(snapshot.dto.Description ?? '');
  if (description) {
    metadata.formValues.unshift({ FieldName: 'Description', FieldValue: description });
  }
  for (const gap of metadata.requiredGaps) {
    blockers.push(gap.supportable
      ? `${gap.title} is required on the destination — enter a value to copy.`
      : `${gap.title} is required on the destination and cannot be filled from here.`);
  }

  const plan = {
    engine,
    createPath,
    sameWeb,
    sameSite,
    source: {
      webUrl: snapshot.web.webUrl,
      pageId: snapshot.pageId,
      listId: snapshot.library.id,
      fileRef: snapshot.fileRef,
      absoluteUrl: snapshot.fileRef ? `${originOf(snapshot.web.webUrl)}${snapshot.fileRef}` : '',
      fileName: snapshot.fileName,
      etag: snapshot.etag,
    },
    target: {
      webUrl: target.webUrl,
      webServerRelativeUrl: webSR,
      libraryId: target.library?.id || '',
      libraryRoot,
      folder,
      fileName,
      stem,
      finalPath: `${finalDir}/${fileName}`,
      finalAbsolute: abs(`${finalDir}/${fileName}`),
      stagingStem: stage,
      stagingPath: `${libraryRoot}/${stage}.aspx`,
      stagingAbsolute: abs(`${libraryRoot}/${stage}.aspx`),
    },
    title,
    description,
    pageLayoutType: snapshot.dto.PageLayoutType || 'Article',
    publish: Boolean(options.publish),
    promoteAsNews: Boolean(options.promoteAsNews) && snapshot.dto.PromotedState > 0,
    checkInDraft: options.checkInDraft !== false,
    assetFolder,
    assetFolderChain: [
      { path: `${siteAssetsRoot}/SitePages`, recyclable: false },
      { path: assetFolder, recyclable: !target.assetFolderExists },
    ],
    assets,
    metadata,
    commentsDisabled: snapshot.commentsDisabled,
    rewriteLinks: Boolean(options.rewriteLinks) && !sameWeb,
    dropped: [...dropped],
    analysis,
    warnings,
    blockers,
    legacyHtml: legacy,
  };
  return plan;
}

function originOf(url) {
  try { return new URL(url).origin; } catch { return ''; }
}

function uniqueName(name, used) {
  const dot = name.lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  let candidate = name;
  for (let i = 2; used.has(lower(candidate)); i += 1) candidate = `${base}-${i}${ext}`;
  used.add(lower(candidate));
  return candidate;
}

// ---- rewrite (§5.3 ordering: runs after the assets transferred) ----------------

// transferResults: Map(asset key → { path, ids: { siteId, webId, listId,
// uniqueId }, url }). Returns the savepage fields. A same-web copy returns the
// raw strings untouched.
export function rewriteContent(snapshot, plan, transferResults = new Map(), analyzers = null) {
  const dropped = new Set(plan.dropped || []);
  const titleChanged = plan.title !== snapshot.dto.Title;
  const sameWebVerbatim = plan.sameWeb && !dropped.size;

  let canvasOut = snapshot.canvasRaw;
  let layoutOut = snapshot.layoutRaw;
  const bannerOut = { image: snapshot.dto.BannerImageUrl };

  if (!sameWebVerbatim || titleChanged) {
    const ctx = {
      ...analyzerContext(snapshot),
      target: {
        webUrl: plan.target.webUrl,
        webPath: plan.target.webServerRelativeUrl,
      },
      rewriteLinks: plan.rewriteLinks,
      mapAsset: (identity) => {
        if (!identity) return null;
        const key = assetKey({ uniqueId: identity.ids?.uniqueId ?? identity.uniqueId, path: identity.path });
        const result = transferResults.get(key) || null;
        // Requests are deduped by UniqueId, so two parts naming different
        // files but carrying the same (stale) id share one key, and only one
        // file is transferred. A mapping authorizes a patch only when the
        // path the part names IS that file; otherwise the part keeps its
        // source reference.
        if (result && identity.path && result.sourcePath
            && lower(decodeSafe(identity.path)) !== lower(decodeSafe(result.sourcePath))) return null;
        return result;
      },
      mapLink: (value) => (plan.rewriteLinks
        ? rewriteLink(value, { fromPath: snapshot.web.webServerRelativeUrl, toPath: plan.target.webServerRelativeUrl, origin: originOf(snapshot.web.webUrl) })
        : null),
    };
    if (snapshot.canvas) {
      let changed = dropped.size > 0;
      const controls = [];
      snapshot.canvas.forEach((control, index) => {
        const desc = describeControl(control, index);
        if (desc.instanceId && dropped.has(lower(desc.instanceId))) return;
        if (plan.sameWeb || !analyzers || desc.kind === 'pageSettings' || desc.kind === 'other') { controls.push(control); return; }
        const analyzer = desc.kind === 'text' ? analyzers.text : analyzers.forId(desc.webPartId);
        if (!analyzer?.patch) { controls.push(control); return; }
        const copy = structuredClone(control);
        const n = analyzer.patch(copy, ctx) || 0;
        if (n) changed = true;
        controls.push(n ? copy : control);
      });
      if (changed) canvasOut = JSON.stringify(controls);
    }
    if (snapshot.layout) {
      let changed = false;
      const parts = snapshot.layout.map((part) => {
        let copy = part;
        if (!plan.sameWeb && analyzers?.header?.patch && normalizeGuid(part?.id) === PAGE_HEADER_ID) {
          copy = structuredClone(part);
          if (analyzers.header.patch(copy, ctx)) changed = true; else copy = part;
        }
        if (titleChanged && normalizeGuid(part?.id) === PAGE_HEADER_ID && copy?.properties
            && typeof copy.properties.title === 'string') {
          if (copy === part) copy = structuredClone(part);
          copy.properties.title = plan.title;
          changed = true;
        }
        return copy;
      });
      if (changed) layoutOut = JSON.stringify(parts);
    }
    if (!plan.sameWeb) bannerOut.image = mapBannerUrl(snapshot.dto.BannerImageUrl, ctx, plan);
  }

  // PnPjs's rule (PnP-pages 371–375): with a custom thumbnail the
  // BannerImageUrl sent is the thumbnail's URL — which is exactly what the
  // DTO's BannerImageUrl already holds on such a page (§11), so it is sent
  // as read (mapped to the transferred copy across webs). Description is
  // NOT sent: savepage blanks it on SPO; the run writes it through
  // ValidateUpdateListItem after the save (§11).
  const bannerForSave = bannerOut.image;
  const fields = {
    Title: plan.title,
    CanvasContent1: canvasOut,
    LayoutWebpartsContent: layoutOut,
    TopicHeader: snapshot.dto.TopicHeader,
    AuthorByline: snapshot.dto.AuthorByline,
  };
  if (bannerForSave) fields.BannerImageUrl = bannerForSave;
  return fields;
}

// A banner/thumbnail URL on the source: a getpreview.ashx URL is rebuilt
// against the destination web from the transferred file's ids; a plain path
// is swapped for the transferred path. Unmapped URLs stay as they were (the
// plan already warned they point at the source).
function mapBannerUrl(url, ctx, plan) {
  if (!url) return url;
  const ids = previewIds(url);
  if (ids) {
    const mapped = ctx.mapAsset({ ids });
    if (!mapped?.ids?.uniqueId) return url;
    const hex = (g) => normalizeGuid(g).replace(/-/g, '');
    return `${trimSlash(plan.target.webUrl)}/_layouts/15/getpreview.ashx?guidSite=${hex(mapped.ids.siteId)}`
      + `&guidWeb=${hex(mapped.ids.webId)}&guidFile=${hex(mapped.ids.uniqueId)}`;
  }
  let path = '';
  try { path = decodeSafe(new URL(url, 'https://placeholder.invalid').pathname); } catch { return url; }
  const mapped = ctx.mapAsset({ path });
  if (!mapped?.path) return url;
  return /^https?:/i.test(url) ? `${originOf(plan.target.webUrl)}${encodePath(mapped.path)}` : mapped.path;
}

// ---- read-back verification (§3 [verify]) ---------------------------------------

const webPartSequence = (raw) => {
  const parsed = parseJsonArray(raw);
  if (!parsed.ok) return null;
  return parsed.value
    .filter((c) => c && typeof c === 'object' && !c.pageSettingsSlice)
    .map((c) => (c.controlType === 4 ? 'text' : normalizeGuid(c.webPartId || c.webPartData?.id) || 'other'));
};

// saved: the fields the content save sent. dto: GET pages(id) after the run.
// Returns drift [{ field, expected, actual }] — empty when the copy matches.
export function compareReadBack(plan, saved, dto = {}) {
  const drift = [];
  const same = (field, expected, actual) => {
    if (String(expected ?? '') !== String(actual ?? '')) drift.push({ field, expected, actual });
  };
  same('Title', plan.title, dto.Title);
  same('FileName', plan.target.fileName.toLowerCase(), String(dto.FileName || '').toLowerCase());
  same('PageLayoutType', plan.pageLayoutType, dto.PageLayoutType);
  if (plan.description !== undefined && plan.engine !== 'copyFile') same('Description', plan.description, dto.Description ?? '');
  if (plan.commentsDisabled === true && dto.CommentsDisabled === false) {
    drift.push({ field: 'CommentsDisabled', expected: true, actual: false });
  }
  if (saved) {
    same('TopicHeader', saved.TopicHeader, dto.TopicHeader);
    const expectedParts = webPartSequence(saved.CanvasContent1);
    const actualParts = webPartSequence(dto.CanvasContent1);
    if (expectedParts && (!actualParts || expectedParts.join(',') !== actualParts.join(','))) {
      drift.push({ field: 'CanvasContent1', expected: `${expectedParts.length} parts`, actual: actualParts ? `${actualParts.length} parts (${actualParts.join(', ')})` : 'unreadable' });
    }
    if (saved.BannerImageUrl && !dto.BannerImageUrl) {
      drift.push({ field: 'BannerImageUrl', expected: 'a banner', actual: '(none)' });
    }
  }
  const promoted = Number(dto.PromotedState ?? 0);
  if (!plan.promoteAsNews && promoted > 0) drift.push({ field: 'PromotedState', expected: 0, actual: promoted });
  return drift;
}
