// List item DATA — pure core, no I/O. Stage 1b-a.
//
// Port of the console-first `exportListData`/`importListData` pair in
// utilities/dcspad-sp-utilities.js (SPUtils, lines ~2425-3146) into the
// Workbench's own ES modules, the same way list-schema.js ports
// getListSchema/createListFromSchema. The JSON document is the contract:
// `kind: dcspad-sputils-list-data` (list-schema.js's DATA_KIND), every v1 key
// name and type kept byte-for-byte, new keys additive, `version` moves to 2
// while a `version: 1` document (a SPUtils export) is still accepted on read.
//
// This module only builds and reads the document and the per-type value
// strings — it never talks to the network. list-data-capture.js (I/O) calls
// buildDataDoc(); the future import/apply module (stage 1b-b, not built
// here) will call writableFields/toImportFormValues/resolveLookupByShown
// Value/partitionPasses against a live target probe. Everything here is
// unit-tested directly (tests/workbench-schema.mjs).
//
// Two deliberate SMALLER-than-SPUtils behaviours, called out because SPUtils
// is unclear or does something only I/O can do:
//   1. toWebDateString() here takes a flat `offsetMinutes` instead of
//      SPUtils' per-UTC-day server lookup (calibrateDateFormat/webLocalDate,
//      utilities/dcspad-sp-utilities.js:2125-2184). The DST-exact, per-instant
//      offset needs a network round trip; that belongs to the future apply
//      module, which can supply the right offset per item/day the way
//      SPUtils' day cache does. Pure code here just applies whatever offset
//      (default 0 = UTC) it's given.
//   2. folderOrder() breaks ties by original array position ("stable") where
//      SPUtils breaks ties by source Id (dcspad-sputils-list-data.js:2975).
//      Both produce a valid parents-first order; Id-ordering only matters
//      when two folders share a path depth, which SPUtils itself doesn't
//      guarantee is meaningful either.

import { SpFileError } from '../sp-odata.js';
import {
  DATA_KIND, LOOKUP_TYPES, USER_TYPES, TAXONOMY_TYPES, NEVER_WRITE, NEVER_WRITE_TYPES,
} from './list-schema.js';

export { DATA_KIND };
export const DATA_VERSION = 2;

function cleanGuid(v) {
  const s = String(v ?? '').replace(/[{}]/g, '').trim();
  return s ? s.toLowerCase() : null;
}

function toIdList(v) {
  return (v == null ? [] : Array.isArray(v) ? v : [v]).filter((x) => x != null);
}

// SPUtils v1 docs carry people only inside each item's/folder's `_resolved`
// object — there is no top-level `users` array in that format. When `users`
// is missing or empty (a v1 document, or a v2 one that simply forgot the
// key), derive the user set by scanning every row's `_resolved`: Author,
// Editor, and every field the doc's own `fields` map marks as a User/
// UserMulti type. Deduped by Id.
function deriveUsersFromResolved(rows, fields) {
  const byId = new Map();
  for (const row of rows || []) {
    const resolved = row?._resolved;
    if (!resolved) continue;
    for (const [key, val] of Object.entries(resolved)) {
      const isUserField = key === 'Author' || key === 'Editor' || USER_TYPES.has(fields?.[key]?.type);
      if (!isUserField) continue;
      const arr = Array.isArray(val) ? val : [val];
      for (const u of arr) {
        if (!u || u.Id == null || byId.has(u.Id)) continue;
        byId.set(u.Id, { Id: u.Id, Email: u.Email || '', LoginName: u.LoginName || '', Title: u.Title || '' });
      }
    }
  }
  return [...byId.values()];
}

// ---- document shape ---------------------------------------------------------

// buildDataDoc keeps SPUtils' v1 `items` array shape exactly: one array,
// folders included inline (a folder row carries `_folder: true` and
// `_folderPath`) — that's the only key a v1 reader (SPUtils importListData)
// ever looks at. `folders` is repeated as its own v2 top-level array purely
// for the Workbench's convenience (folderOrder(), a folder count without
// filtering `items`); it is additive and a v1 reader ignores it.
export function buildDataDoc({
  source = {}, fields = {}, items = [], folders = [], users = [], warnings = [],
  generatorBuild = 'dev',
} = {}) {
  return {
    kind: DATA_KIND,
    version: DATA_VERSION,
    exported: new Date().toISOString(),
    generator: { tool: 'dcspad-workbench', build: generatorBuild },
    source,
    fields,
    items: [...folders, ...items],
    warnings,
    // v2 additions.
    folders,
    users,
  };
}

// Accepts a v1 (SPUtils) or v2 (Workbench) document and returns a v2-shaped
// copy: `items` holds only non-folder rows, `folders` the folder rows —
// reconstructed from `items` when the source document (v1) has no separate
// `folders` key, mirroring importListData's own split (dcspad-sp-utilities.js:2976).
// Throws SpFileError('bad-data') when `kind` doesn't match.
export function normalizeDataDoc(doc) {
  if (!doc || typeof doc !== 'object' || doc.kind !== DATA_KIND) {
    throw new SpFileError(
      `Not a list data document (expected kind ${DATA_KIND}).`,
      { code: 'bad-data' },
    );
  }
  const sourceVersion = Number(doc.version) || 1;
  const allItems = Array.isArray(doc.items) ? doc.items : [];
  const folders = Array.isArray(doc.folders) ? doc.folders : allItems.filter((i) => i?._folder);
  const items = allItems.filter((i) => !i?._folder);
  const fields = doc.fields || {};
  const users = (Array.isArray(doc.users) && doc.users.length)
    ? doc.users
    : deriveUsersFromResolved([...folders, ...items], fields);
  return {
    kind: doc.kind,
    version: DATA_VERSION,
    exported: doc.exported || '',
    generator: doc.generator || { tool: 'unknown', build: '' },
    source: doc.source || {},
    fields,
    items,
    folders,
    users,
    warnings: doc.warnings || [],
    _sourceVersion: sourceVersion,
  };
}

// Parents-first, by folder-path depth; stable within a depth (original
// position wins the tie — see the file-header note on why this differs from
// SPUtils' Id tiebreak).
export function folderOrder(folders) {
  const depthOf = (f) => {
    const path = f?._folderPath || f?.folderPath || '';
    return path ? path.split('/').length : 0;
  };
  return (folders || [])
    .map((f, i) => [f, i])
    .sort((a, b) => depthOf(a[0]) - depthOf(b[0]) || a[1] - b[1])
    .map(([f]) => f);
}

// ---- writable-field selection (importListData, dcspad-sp-utilities.js:2838-2851) --

// `schemaFields` is the DATA_KIND `fields` map (internalName -> {type,
// custom, readOnly, lookupList, lookupListId, lookupField, isSelfLookup,
// allowMultipleValues} — exactly buildDataDoc's `fields` key).
// `targetFields` is the TARGET list's field rows; either raw SharePoint rows
// (InternalName/TypeAsString/ReadOnlyField/LookupList, as a live probe would
// return) or list-schema-capture's normalized shape (internalName/type/
// readOnly/lookupListId) — both are accepted so this can be driven directly
// by list-schema-capture's probeTarget() output once the apply module lands.
// Writable = present on the target, not read-only, not system-owned or a
// SharePoint-computed type, and either custom on the source or the Title
// column — SPUtils' exact rule.
export function writableFields(schemaFields, targetFields) {
  const byName = new Map();
  for (const f of targetFields || []) {
    const name = f.InternalName ?? f.internalName;
    if (name) byName.set(name, f);
  }
  const out = [];
  for (const [name, meta] of Object.entries(schemaFields || {})) {
    const tf = byName.get(name);
    if (!tf) continue;   // missing on target
    const readOnly = tf.ReadOnlyField ?? tf.readOnly ?? false;
    const typeAsString = tf.TypeAsString ?? tf.type ?? '';
    if (readOnly) continue;
    if (NEVER_WRITE.has(name) || NEVER_WRITE_TYPES.has(typeAsString)) continue;
    if (!(meta?.custom || name === 'Title')) continue;
    const lookupListId = cleanGuid(tf.LookupList ?? tf.lookupListId);
    const sameLookupList = LOOKUP_TYPES.has(typeAsString) && !!meta?.lookupListId
      && lookupListId === cleanGuid(meta.lookupListId);
    out.push({ name, tf, meta, sameLookupList, typeAsString });
  }
  return out;
}

// ---- per-type FieldValue conventions (toFormValue, dcspad-sp-utilities.js:2935-2952) --

function pad2(n) { return String(n).padStart(2, '0'); }

// `dateFormat`: { order: 'mdy'|'dmy'|'ymd', sep, timeSep, offsetMinutes }.
// See the file-header note: offsetMinutes replaces SPUtils' live per-day
// server lookup with a flat number a caller supplies.
export function toWebDateString(value, dateFormat, { dateOnly = false } = {}) {
  const f = dateFormat || { order: 'mdy', sep: '/', timeSep: ':', offsetMinutes: 0 };
  const utc = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(utc.getTime())) return String(value);
  const local = new Date(utc.getTime() + (Number(f.offsetMinutes) || 0) * 60000);
  const parts = {
    y: String(local.getUTCFullYear()), m: pad2(local.getUTCMonth() + 1), d: pad2(local.getUTCDate()),
  };
  const dateText = (f.order || 'mdy').split('').map((k) => parts[k]).join(f.sep ?? '/');
  if (dateOnly) return dateText;
  return `${dateText} ${pad2(local.getUTCHours())}${f.timeSep ?? ':'}${pad2(local.getUTCMinutes())}`;
}

// DST boundaries differ between one date field and another on the very same
// item (a Created stamp and a custom DueOn value can straddle a transition
// differently) — the executor precomputes a per-field-name offset (see its
// dateFormatForFields()) rather than one offset for the whole item, and this
// picks the right one out of that map. A field with no per-field entry falls
// back to the flat offsetMinutes the caller supplied (mock mode, or a caller
// that never bothered with per-field offsets).
function fieldDateFormat(dateFormat, name) {
  if (dateFormat?.perField?.has(name)) {
    return { ...dateFormat, offsetMinutes: dateFormat.perField.get(name) };
  }
  return dateFormat;
}

// SPUtils userValue(): resolved {Email,LoginName} users -> the
// '[{"Key":"login"}]' JSON string ValidateUpdateListItem expects.
// `userIds`: Map(lowercased email-or-login -> target login name), built by
// the (future) apply module the way SPUtils' loginFor()/ensureUser() does —
// pure code here just looks the key up.
function userFormValue(users, userIds) {
  const keys = [];
  for (const u of users || []) {
    const key = String(u?.Email || u?.LoginName || '').toLowerCase();
    const login = key && userIds ? userIds.get(key) : null;
    if (login) keys.push({ Key: login });
  }
  return keys.length ? JSON.stringify(keys) : '';
}

// SPUtils lookupValue()'s ambiguity rule (dcspad-sp-utilities.js:2907-2929): a
// shown value matching exactly one candidate on the target resolves
// directly; a value matching more than one is only accepted when the column
// points at the very same source list AND one candidate carries the
// original source id (a same-site copy); anything else is left unresolved,
// with a reason instead of a guess.
export function resolveLookupByShownValue(shown, candidates, { sourceId = null, sameLookupList = false } = {}) {
  if (shown == null || shown === '') {
    return { id: null, error: `source id ${sourceId ?? '?'} had no shown value on the source; reference left empty` };
  }
  const list = candidates || [];
  if (list.length === 1) return { id: list[0], error: '' };
  if (list.length > 1 && sameLookupList && sourceId != null && list.includes(sourceId)) {
    return { id: sourceId, error: '' };
  }
  return {
    id: null,
    error: list.length
      ? `"${shown}" matches ${list.length} items on the target lookup list; reference left empty`
      : `"${shown}" (source id ${sourceId ?? '?'}) was not found on the target lookup list; reference left empty`,
  };
}

// `resolved` = item._resolved[fieldName]: [{Id: sourceId, value: shownText}].
// `lookupIdsByValue` = Map(shownValue -> [targetIds]), the target-side index
// SPUtils builds once per writable lookup column before importing any item.
function lookupFormValue(resolved, multi, lookupIdsByValue, sameLookupList, errs, fieldName) {
  const ids = [];
  for (const r of resolved || []) {
    const { id, error } = resolveLookupByShownValue(
      r?.value, lookupIdsByValue?.get(String(r?.value ?? '')) || [],
      { sourceId: r?.Id, sameLookupList },
    );
    if (id != null) ids.push(id);
    else errs.push({ field: fieldName, message: error });
  }
  if (!ids.length) return '';
  return multi ? `${ids.join(';#')};#` : String(ids[0]);
}

// A source id -> target id map (idMap, from a previous create pass) resolves
// a self-lookup directly — no shown-value ambiguity to weigh, since the
// reference and the target are the very same list.
function selfLookupFormValue(resolved, idMap, multi, errs, fieldName) {
  const ids = [];
  for (const r of resolved || []) {
    const mapped = idMap?.[r?.Id];
    if (mapped != null) ids.push(mapped);
    else errs.push({ field: fieldName, message: `references source item ${r?.Id}, which was not imported; reference left empty` });
  }
  if (!ids.length) return '';
  return multi ? `${ids.join(';#')};#` : String(ids[0]);
}

function taxonomyFormValue(raw) {
  const terms = (raw == null ? [] : Array.isArray(raw) ? raw : [raw]).filter((t) => t && t.Label);
  return terms.map((t) => `${t.Label}|${cleanGuid(t.TermGuid)};`).join('');
}

// Builds ValidateUpdateListItem formValues for one item over a writable-field
// list (writableFields() output, or its pass1/pass2 split from
// partitionPasses()). `lookupIds`: Map(fieldName -> Map(shownValue ->
// [targetIds])); `userIds`: Map(email/login -> target login); `idMap`:
// Map/object(sourceId -> targetId), required only when a field in `fields`
// is a self-lookup (its resolution goes through idMap, not lookupIds).
// Returns { values: [{FieldName, FieldValue}], errors: [{field, message}] } —
// SPUtils never lets one bad reference abort the whole item; it reports the
// field and writes an empty value instead (dcspad-sp-utilities.js:2935-2952).
export function toImportFormValues(item, fields, { dateFormat, userIds, lookupIds, idMap } = {}) {
  const values = [];
  const errors = [];
  for (const w of fields || []) {
    const raw = item?.[w.name];
    const type = w.typeAsString ?? w.tf?.TypeAsString ?? w.tf?.type ?? '';
    let value = '';
    switch (type) {
      case 'Boolean':
        value = raw == null ? '' : (raw ? '1' : '0');
        break;
      case 'MultiChoice': {
        const arr = raw == null ? [] : Array.isArray(raw) ? raw : [raw];
        value = arr.length ? `;#${arr.join(';#')};#` : '';
        break;
      }
      case 'URL':
        value = raw && raw.Url ? `${raw.Url}, ${raw.Description || raw.Url}` : '';
        break;
      case 'DateTime':
        value = raw ? toWebDateString(raw, fieldDateFormat(dateFormat, w.name), { dateOnly: (w.tf?.DisplayFormat ?? w.tf?.displayFormat) === 0 }) : '';
        break;
      case 'User':
      case 'UserMulti':
        value = userFormValue(item?._resolved?.[w.name], userIds);
        break;
      case 'Lookup':
      case 'LookupMulti': {
        const multi = type === 'LookupMulti';
        value = (w.meta?.isSelfLookup && idMap)
          ? selfLookupFormValue(item?._resolved?.[w.name], idMap, multi, errors, w.name)
          : lookupFormValue(item?._resolved?.[w.name], multi, lookupIds?.get(w.name), w.sameLookupList, errors, w.name);
        break;
      }
      case 'TaxonomyFieldType':
      case 'TaxonomyFieldTypeMulti':
        value = taxonomyFormValue(raw);
        break;
      default:
        value = raw == null ? '' : (typeof raw === 'object' ? JSON.stringify(raw) : String(raw));
    }
    values.push({ FieldName: w.name, FieldValue: value });
  }
  return { values, errors };
}

// SPUtils authorshipValues() (dcspad-sp-utilities.js:2955-2964): the four
// fields a preserve-authorship import restores after the item exists.
export function authorshipFormValues(item, { dateFormat, userIds } = {}) {
  const values = [];
  const author = userFormValue(item?._resolved?.Author, userIds);
  const editor = userFormValue(item?._resolved?.Editor, userIds);
  if (author) values.push({ FieldName: 'Author', FieldValue: author });
  if (editor) values.push({ FieldName: 'Editor', FieldValue: editor });
  if (item?.Created) values.push({ FieldName: 'Created', FieldValue: toWebDateString(item.Created, fieldDateFormat(dateFormat, 'Created')) });
  if (item?.Modified) values.push({ FieldName: 'Modified', FieldValue: toWebDateString(item.Modified, fieldDateFormat(dateFormat, 'Modified')) });
  return values;
}

// Splits a writable-field list (writableFields() output) and the item set
// into the three passes importListData applies in order (dcspad-sp-
// utilities.js:3034-3140): plain values on create (pass1, every item), self-
// referencing lookups once every item has a new id (pass2, only fields
// flagged isSelfLookup and only items that actually reference something),
// and authorship restoration last (pass3, only items that carry resolved
// Author/Editor or a Created/Modified stamp to restore). Pure and
// I/O-free: it decides WHAT belongs in each pass, not how a later apply
// module executes them (idMap resolution, network calls).
export function partitionPasses(items, fields) {
  const selfLookupFields = (fields || []).filter((w) => LOOKUP_TYPES.has(w.typeAsString) && w.meta?.isSelfLookup);
  const selfLookupNames = new Set(selfLookupFields.map((w) => w.name));
  const pass1Fields = (fields || []).filter((w) => !selfLookupNames.has(w.name));

  const hasAuthorship = (item) => Boolean(
    toIdList(item?._resolved?.Author).length || toIdList(item?._resolved?.Editor).length
    || item?.Created || item?.Modified,
  );
  const hasSelfLookupRef = (item) => selfLookupFields.some((w) => toIdList(item?._resolved?.[w.name]).length);

  return {
    pass1: { items: items || [], fields: pass1Fields },
    pass2: selfLookupFields.length
      ? { items: (items || []).filter(hasSelfLookupRef), fields: selfLookupFields }
      : { items: [], fields: [] },
    pass3: { items: (items || []).filter(hasAuthorship), fields: ['Author', 'Editor', 'Created', 'Modified'] },
  };
}

// Re-exported so a caller only needs one import for the "does this field
// carry user/lookup/taxonomy values" questions this module already answers.
export { LOOKUP_TYPES, USER_TYPES, TAXONOMY_TYPES };
