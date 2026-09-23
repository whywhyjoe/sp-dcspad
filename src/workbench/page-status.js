// Page status: publish state, check-out holder and broken inheritance for the
// Pages view — the grid's "Scan Page Status" columns, the drilldown's status
// chips and Permissions tab, and the Metadata tab's fixed row layout. Pure: no
// DOM, no I/O, so every rule here is unit-tested (tests/workbench.mjs).
//
// "Published" means the page has EVER had a published major version, even if
// a newer draft sits on top of it (Joe, 2026-09-23). It is read from the
// file's major version, never from _ModerationStatus alone: that field reads 0
// ("Approved") on every item of a library without content approval, drafts
// included, so on its own it would call every draft published.

// The status read, as a query ladder (see views/pages.js queryLadder). Rung 0
// is the full shape; each rung below gives up something SharePoint may refuse
// on schema grounds rather than dead-ending the scan.
//
// `hasModeration` — the library carries _ModerationStatus (content approval is
// or was on). Naming it on a library without the field is a 400, so it is
// only asked for when the fields probe found it.
export function pageStatusShapes({ hasModeration = false } = {}) {
  const moderation = hasModeration ? ['OData__ModerationStatus'] : [];
  const file = ['File/MajorVersion', 'File/MinorVersion', 'File/CheckOutType'];
  return [
    {
      options: {
        select: ['Id', 'HasUniqueRoleAssignments', ...file, 'CheckoutUser/Title', ...moderation],
        expand: ['File', 'CheckoutUser'],
      },
    },
    // Without the people projection: the check-out holder's name is lost, the
    // fact that the page is checked out is not.
    {
      options: { select: ['Id', 'HasUniqueRoleAssignments', ...file, ...moderation], expand: ['File'] },
      lost: 'the checked-out-to names',
    },
    // Without the File expand: the version label still says published or not.
    {
      options: { select: ['Id', 'HasUniqueRoleAssignments', 'OData__UIVersionString', ...moderation] },
      lost: 'the check-out state',
    },
    {
      options: { select: ['Id', 'HasUniqueRoleAssignments'] },
      lost: 'the publish and check-out state',
    },
  ];
}

// "3.2" -> { major: 3, minor: 2 }; anything else -> null.
export function parseVersionLabel(label) {
  const m = /^(\d+)\.(\d+)$/.exec(String(label ?? '').trim());
  return m ? { major: Number(m[1]), minor: Number(m[2]) } : null;
}

const numberOrNull = (v) => (typeof v === 'number' && Number.isFinite(v) ? v : null);

// _ModerationStatus: 0 Approved, 1 Rejected, 2 Pending, 3 Draft, 4 Scheduled
// (MS-WSSFO). A major version in 1, 2 or 4 has not gone live — a scheduled
// one is waiting for its publishing start date.
const NOT_LIVE = new Set([1, 2, 4]);

// Where a moderation value can sit: top level under its OData name (items,
// and versions when the select names it), the x005f-escaped form some version
// payloads use, the bare internal name, or a FieldValues dictionary.
function moderationOf(entity) {
  if (!entity || typeof entity !== 'object') return null;
  for (const key of ['OData__ModerationStatus', 'OData__x005f_ModerationStatus', '_ModerationStatus']) {
    const n = numberOrNull(entity[key]);
    if (n !== null) return n;
  }
  const fv = entity.FieldValues;
  return fv && typeof fv === 'object' ? numberOrNull(fv._ModerationStatus ?? fv.OData__ModerationStatus) : null;
}

// One item (from pageStatusShapes, or any item that happens to carry the same
// fields) -> the status the view shows. Each member is null when this read
// could not say — a chip or cell for it is then left out rather than guessed.
//   published          true | false | null
//   checkedOut         true | false | null
//   checkedOutTo       the holder's display name ('' when unknown)
//   brokenInheritance  true | false | null
export function derivePageStatus(item = {}) {
  const file = item.File && typeof item.File === 'object' ? item.File : null;
  const label = parseVersionLabel(item.OData__UIVersionString ?? item._UIVersionString);
  const major = numberOrNull(file?.MajorVersion) ?? label?.major ?? null;
  const minor = numberOrNull(file?.MinorVersion) ?? label?.minor ?? null;
  const moderation = moderationOf(item);

  let published = null;
  if (major !== null) {
    published = major > 0;
    // Content approval / scheduling: submitting 0.x makes it 1.0 at once,
    // Pending (2) until approved, Rejected (1) if turned down, or Scheduled
    // (4) until its start date. A 1.0 in any of those has never actually
    // been published. (A later x.0 in one of them still has an earlier live
    // major under it.)
    if (major === 1 && minor === 0 && NOT_LIVE.has(moderation)) published = false;
  }

  const holder = item.CheckoutUser && typeof item.CheckoutUser === 'object'
    ? String(item.CheckoutUser.Title || '') : '';
  // SP.CheckOutType: 0 online, 1 offline, 2 none.
  const checkOutType = numberOrNull(file?.CheckOutType);
  let checkedOut = null;
  if (holder) checkedOut = true;
  else if (checkOutType !== null) checkedOut = checkOutType !== 2;
  else if ('CheckoutUser' in item || 'CheckoutUserId' in item) {
    checkedOut = Boolean(item.CheckoutUserId);
  }

  const unique = item.HasUniqueRoleAssignments;
  return {
    published,
    checkedOut,
    checkedOutTo: holder,
    brokenInheritance: typeof unique === 'boolean' ? unique : null,
  };
}

export const publishLabel = (published) =>
  (published === true ? 'Published' : published === false ? 'Unpublished' : '');
export const inheritanceLabel = (broken) =>
  (broken === true ? 'Broken inheritance' : broken === false ? 'Inherited' : '');
// The grid's marker column: only the exception is worth a word.
export const inheritanceMarker = (broken) => (broken === true ? 'Broken' : '');
export const checkedOutLabel = (status) =>
  (status?.checkedOut ? (status.checkedOutTo || 'Checked out') : '');

// The item versions read for "last published". Same ladder idea: a select
// SharePoint refuses falls back to the bare collection (heavier — each
// version carries its field values — but still answerable).
//
// Versions name the moderation field differently from items. Live SPO (dev
// tenant, 2026-09-23, with and without content approval): selecting
// `OData__ModerationStatus` on items(id)/versions answers 200 but OMITS the
// value; selecting `OData__x005f_ModerationStatus` returns it under that key,
// which is also the key the bare collection uses. Items keep the plain name.
export const VERSION_MODERATION_FIELD = 'OData__x005f_ModerationStatus';
export function versionShapes({ hasModeration = false } = {}) {
  const base = ['VersionId', 'VersionLabel', 'IsCurrentVersion', 'Created'];
  return [
    { options: { select: hasModeration ? [...base, VERSION_MODERATION_FIELD] : base } },
    { options: {} },
  ];
}

// Whether a missing moderation value must make "last published" unknown.
// The _ModerationStatus FIELD exists on every pages library, approval or not
// (hidden), so the field probe alone says nothing about approval. A version
// read that comes back without moderation (the select naming the wrong key
// did exactly that on live SPO and blanked Publish Date on every library)
// may only withhold the date where the list really runs content approval —
// the list's own EnableModeration; unknown (null) keeps the conservative
// reading.
export function moderationApplies(hasModerationField, enableModeration) {
  return Boolean(hasModerationField) && enableModeration !== false;
}

// The date the page's CURRENT published version was published: the newest
// major (x.0) version, skipping one still pending, rejected or scheduled when
// the library runs approval. '' when the page was never published.
//
// `hasModeration` — the library carries _ModerationStatus. Then a candidate
// version whose moderation could NOT be read (the versions read fell back to
// a shape without it) makes the answer unknown (null) rather than letting a
// pending or scheduled version pass as live.
export function lastPublishedFrom(versions, { hasModeration = false } = {}) {
  const majors = [];
  for (const v of versions || []) {
    const label = parseVersionLabel(v?.VersionLabel);
    if (!label || label.major < 1 || label.minor !== 0) continue;
    majors.push({ label, v, moderation: moderationOf(v) });
  }
  majors.sort((a, b) => b.label.major - a.label.major);
  for (const m of majors) {
    if (m.moderation === null && hasModeration) return null;
    if (NOT_LIVE.has(m.moderation)) continue;
    return String(m.v.Created || m.v.Modified || '');
  }
  return '';
}

// ---- Metadata tab row layout ------------------------------------------------

// Joe's spec (2026-09-23): only these rows, in this order; a field the library
// does not carry is skipped. `field` is the internal name, `aliases` the
// display names to fall back on (a site-column internal name can differ from
// the name people know it by — Item Type is FolderType, a column whose use
// changed after it was created). Entries with `status` are synthesized from
// the page-status read rather than a list field.
export const METADATA_SPEC = [
  { status: 'id', label: 'ID' },
  { field: 'FileLeafRef', label: 'Name' },
  { field: 'Title', label: 'Title' },
  { status: 'published', label: 'Publish Status' },
  { status: 'lastPublished', label: 'Publish Date' },
  { field: 'FirstPublishedDate', label: 'First Published Date', aliases: ['First Published Date'] },
  { status: 'inheritance', label: 'Permissions' },
  { field: 'CheckoutUser', label: 'Checked Out To', aliases: ['Checked Out To'], status: 'checkedOutTo' },
  { field: 'PromotedState', label: 'Promoted State', display: 'promoted' },
  { field: 'bmocContentCategory', label: 'Content Category', aliases: ['Content Category'] },
  { field: 'Modified', label: 'Modified' },
  { field: 'Editor', label: 'Modified By' },
  { field: 'Created', label: 'Created' },
  { field: 'Author', label: 'Created By' },
  { field: 'FolderType', label: 'Item Type', aliases: ['Item Type'] },
  { field: 'Contact', label: 'Contact' },
  { field: 'Pillar', label: 'Pillar' },
  { field: 'Org', label: 'Org' },
  { field: 'ComplianceAssetId', label: 'Compliance Asset ID', aliases: ['Compliance Asset Id', 'Compliance Asset ID'] },
  { field: 'WikiField', label: 'Wiki Content', aliases: ['Wiki Content'] },
];

// PromotedState: 2 is a promoted news post; anything else reads "False".
export const promotedStateText = (v) => (Number(v) === 2 ? 'Promoted' : 'False');

const lower = (s) => String(s ?? '').toLowerCase();

// fields: the list's field entities. status: { id, published, lastPublished,
// brokenInheritance, checkedOut, checkedOutTo } — any member may be null
// (unknown), which skips a status-only row rather than showing a guess.
// Returns entries for createFieldEditorForm's `layout`:
//   { field, label, display? }   a real list field (display: override text)
//   { internal, label, text }    a synthesized read-only row
export function resolveMetadataLayout(fields, status = {}, spec = METADATA_SPEC) {
  const byInternal = new Map();
  const byTitle = new Map();
  for (const f of fields || []) {
    if (f?.InternalName) byInternal.set(lower(f.InternalName), f);
    // Display-name aliases only match visible fields: SharePoint's hidden
    // built-ins reuse friendly titles (FSObjType is titled "Item Type"), and
    // an alias is for the name people know a column by, never plumbing.
    if (f?.Title && !f.Hidden && !byTitle.has(lower(f.Title))) byTitle.set(lower(f.Title), f);
  }
  const used = new Set();
  const out = [];
  for (const entry of spec) {
    let field = entry.field ? byInternal.get(lower(entry.field)) : null;
    if (!field && entry.aliases) {
      for (const alias of entry.aliases) {
        field = byTitle.get(lower(alias));
        if (field) break;
      }
    }
    if (field && !used.has(field.InternalName)) {
      used.add(field.InternalName);
      const resolved = { field, label: entry.label };
      if (entry.display === 'promoted') resolved.display = promotedStateText;
      out.push(resolved);
      continue;
    }
    if (!entry.status) continue;
    const text = statusText(entry.status, status);
    if (text === null) continue;
    out.push({ internal: `__status_${entry.status}`, label: entry.label, text });
  }
  return out;
}

function statusText(kind, status) {
  switch (kind) {
    case 'id':
      return status.id === undefined || status.id === null ? null : String(status.id);
    case 'published':
      return status.published === null || status.published === undefined
        ? null : publishLabel(status.published);
    case 'lastPublished':
      // Blank if none — but only once the read has actually answered.
      return status.lastPublished === null || status.lastPublished === undefined
        ? null : String(status.lastPublished).slice(0, 10);
    case 'inheritance':
      return status.brokenInheritance === null || status.brokenInheritance === undefined
        ? null : inheritanceLabel(status.brokenInheritance);
    case 'checkedOutTo':
      return status.checkedOut === null || status.checkedOut === undefined
        ? null : (status.checkedOut ? (status.checkedOutTo || 'Checked out') : '');
    default:
      return null;
  }
}
