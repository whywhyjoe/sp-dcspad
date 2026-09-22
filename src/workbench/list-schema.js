// List schema — pure core, no I/O.
//
// This is the port of the console-first `getListSchema`/`createListFromSchema`
// pair in utilities/dcspad-sp-utilities.js (SPUtils) into the Workbench's own
// ES modules. The JSON document is the contract between the two: every v1 key
// name and type from SPUtils is kept byte-for-byte here, new keys are
// additive, and `version` moves to 2 while a `version: 1` document is still
// accepted on read (see normalizeSchemaDoc). A SPUtils export imports in the
// Workbench; a Workbench export imports in SPUtils, which checks only `kind`.
//
// Nothing here talks to the network. list-schema-capture.js reads via a
// client and calls the normalize* functions below; the future
// list-schema-apply.js (stage 1b/2 slice) executes the Plan this module
// builds. DOMParser/XMLSerializer are browser globals and fine here — the
// tests run in-page, same as canvas.js.

import { SpFileError } from '../sp-odata.js';

export const SCHEMA_KIND = 'dcspad-sputils-list-schema';
export const DATA_KIND = 'dcspad-sputils-list-data';
export const SCHEMA_VERSION = 2;

export const LOOKUP_TYPES = new Set(['Lookup', 'LookupMulti']);
export const USER_TYPES = new Set(['User', 'UserMulti']);
export const TAXONOMY_TYPES = new Set(['TaxonomyFieldType', 'TaxonomyFieldTypeMulti']);
// Never written on import (stage 1b); SharePoint owns them.
export const NEVER_WRITE = new Set([
  'ID', 'Id', 'Attachments', 'ContentType', 'ContentTypeId', 'Author', 'Editor',
  'Created', 'Modified', 'GUID', 'FileRef', 'FileDirRef', 'FileLeafRef', 'UniqueId',
  '_UIVersionString', 'Order',
]);
export const NEVER_WRITE_TYPES = new Set([
  'Computed', 'Counter', 'Attachments', 'File', 'ContentTypeId', 'Calculated', 'Guid',
]);

// Types the Workbench declines to re-index automatically after create — an
// under-approximation is safer than an over-approximation here: a skipped
// Indexed merge is a warning, a rejected one is a failed write.
const NOT_INDEXABLE_TYPES = new Set(['Note', 'Computed', 'Attachments', 'Calculated']);

function cleanGuid(v) {
  const s = String(v ?? '').replace(/[{}]/g, '').trim();
  return s ? s.toLowerCase() : null;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// A lookupMap value (or key) that is a GUID, braced or not, binds directly to
// that list id instead of being resolved by title.
function isGuidLike(v) {
  const g = cleanGuid(v);
  return !!g && GUID_RE.test(g);
}

// Case-insensitive lookupMap key lookup — SharePoint titles collide
// case-insensitively, so a mapping authored against the tenant's own casing
// should still match a differently-cased source title.
export function lookupMapGet(lookupMap, title) {
  if (!title || !lookupMap) return undefined;
  if (Object.prototype.hasOwnProperty.call(lookupMap, title)) return lookupMap[title];
  const wanted = String(title).toLowerCase();
  const key = Object.keys(lookupMap).find((k) => k.toLowerCase() === wanted);
  return key !== undefined ? lookupMap[key] : undefined;
}

// ---- classification --------------------------------------------------------

// A "custom" field is one the list author added: not inherited from the base
// type and deletable. Operates on the RAW (SharePoint-cased) field row, same
// as the SPUtils source this mirrors.
export function isCustomField(f) {
  return !f?.FromBaseType && f?.CanBeDeleted === true && !f?.Hidden
    && f?.TypeAsString !== 'Computed';
}

export function isLibrary(list) {
  return Number(list?.baseType ?? list?.BaseType) === 1;
}

// Dependency tier for field creation: plain → lookup → dependent lookup
// (needs its primary's new id) → calculated. Stable within a tier (the
// caller's original order is preserved).
export function fieldTier(f) {
  if (f?.type === 'Calculated') return 3;
  if (f?.isDependentLookup) return 2;
  if (LOOKUP_TYPES.has(f?.type)) return 1;
  return 0;
}

export function orderFields(fields) {
  return (fields || [])
    .map((f, i) => [f, i])
    .sort((a, b) => fieldTier(a[0]) - fieldTier(b[0]) || a[1] - b[1])
    .map(([f]) => f);
}

export function indexableType(type) {
  return !NOT_INDEXABLE_TYPES.has(type);
}

// List CT id → the site (parent) content type id it was attached from: a
// list-scoped id is the parent id plus "00" plus a 32-hex-char suffix; a
// genuinely site-level id (e.g. "0x0101") has no such suffix and is its own
// parent.
export function parentContentTypeId(listCtId) {
  const id = String(listCtId || '');
  const m = /^(.*?)00[0-9a-f]{32}$/i.exec(id);
  return m ? m[1] : id;
}

// ---- normalizers: raw nometadata rows → doc shapes -------------------------

const LIST_DEFAULTS = {
  title: '', description: '', baseTemplate: 100, enableVersioning: false,
  majorVersionLimit: null, enableMinorVersions: false, majorWithMinorVersionsLimit: null,
  draftVersionVisibility: 0, forceCheckout: false, hidden: false, contentTypesEnabled: false,
  enableAttachments: true, enableFolderCreation: false, enableModeration: false,
  validationFormula: '', validationMessage: '', onQuickLaunch: false, noCrawl: false,
  disableGridEditing: false, ordered: false, readSecurity: null, writeSecurity: null,
  listExperienceOptions: null, enableRequestSignOff: false, library: null, contentTypeOrder: [],
};

export function normalizeList(raw = {}) {
  return {
    ...LIST_DEFAULTS,
    title: raw.Title, description: raw.Description || '', baseTemplate: raw.BaseTemplate,
    enableVersioning: !!raw.EnableVersioning, majorVersionLimit: raw.MajorVersionLimit ?? null,
    enableMinorVersions: !!raw.EnableMinorVersions,
    majorWithMinorVersionsLimit: raw.MajorWithMinorVersionsLimit ?? null,
    draftVersionVisibility: raw.DraftVersionVisibility ?? 0,
    forceCheckout: !!raw.ForceCheckout, hidden: !!raw.Hidden,
    contentTypesEnabled: !!raw.ContentTypesEnabled,
    enableAttachments: raw.EnableAttachments !== false,
    enableFolderCreation: !!raw.EnableFolderCreation, enableModeration: !!raw.EnableModeration,
    validationFormula: raw.ValidationFormula || '', validationMessage: raw.ValidationMessage || '',
    onQuickLaunch: !!raw.OnQuickLaunch, noCrawl: !!raw.NoCrawl,
    disableGridEditing: !!raw.DisableGridEditing, ordered: !!raw.Ordered,
    readSecurity: raw.ReadSecurity ?? null, writeSecurity: raw.WriteSecurity ?? null,
    listExperienceOptions: raw.ListExperienceOptions ?? null,
    enableRequestSignOff: !!raw.EnableRequestSignOff,
    library: raw.BaseType === 1
      ? { documentTemplateUrl: raw.DocumentTemplateUrl || '', irmEnabled: !!raw.IrmEnabled }
      : null,
    contentTypeOrder: [],
  };
}

const FIELD_DEFAULTS = {
  fieldTypeKind: null, group: '', hidden: false, sealed: false, canBeDeleted: true,
  formula: '', outputType: '', displayFormat: null, richText: false,
  validationFormula: '', validationMessage: '', jsLink: '', baseTweak: null,
};

// A base-type field's Title/Required/Description are the only per-list
// tweaks SharePoint allows on Title and _ExtendedDescription — everything
// else about a non-custom field is fixed by the base type and not captured.
function baseTweakFor(raw, custom) {
  if (custom) return null;
  if (raw.InternalName !== 'Title' && raw.InternalName !== '_ExtendedDescription') return null;
  return { title: raw.Title, required: !!raw.Required, description: raw.Description || '' };
}

// `lookupTitleById` is a synchronous map/function — any async list-title
// resolution happens in list-schema-capture.js before this is called.
export function normalizeField(raw, { listId = '', lookupTitleById = () => null } = {}) {
  const type = raw.TypeAsString || '';
  const isLookup = LOOKUP_TYPES.has(type);
  const custom = isCustomField(raw);
  const lookupListId = isLookup && custom ? cleanGuid(raw.LookupList) : null;
  return {
    ...FIELD_DEFAULTS,
    // v1 keys, verbatim from SPUtils getListSchema.
    id: cleanGuid(raw.Id),
    internalName: raw.InternalName,
    staticName: raw.StaticName,
    displayName: raw.Title,
    type,
    required: !!raw.Required,
    readOnly: !!raw.ReadOnlyField,
    fromBaseType: !!raw.FromBaseType,
    custom,
    description: raw.Description || '',
    defaultValue: raw.DefaultValue ?? null,
    choices: Array.isArray(raw.Choices) ? raw.Choices : (raw.Choices?.results || []),
    maxLength: raw.MaxLength ?? null,
    indexed: !!raw.Indexed,
    enforceUniqueValues: !!raw.EnforceUniqueValues,
    allowMultipleValues: !!raw.AllowMultipleValues || type.endsWith('Multi'),
    customFormatter: raw.CustomFormatter || '',
    lookupListId,
    lookupList: lookupListId ? lookupTitleById(lookupListId) : null,
    lookupField: isLookup ? (raw.LookupField || 'Title') : null,
    isSelfLookup: !!lookupListId && lookupListId === cleanGuid(listId),
    isDependentLookup: !!raw.IsDependentLookup,
    primaryFieldId: raw.IsDependentLookup ? cleanGuid(raw.PrimaryFieldId) : null,
    schemaXml: raw.SchemaXml,
    // v2 additions.
    fieldTypeKind: raw.FieldTypeKind ?? null,
    group: raw.Group || '',
    hidden: !!raw.Hidden,
    sealed: !!raw.Sealed,
    canBeDeleted: raw.CanBeDeleted !== false,
    formula: raw.Formula || '',
    outputType: raw.OutputType || '',
    displayFormat: raw.DisplayFormat ?? null,
    richText: !!raw.RichText,
    validationFormula: raw.ValidationFormula || '',
    validationMessage: raw.ValidationMessage || '',
    jsLink: raw.JSLink || '',
    baseTweak: baseTweakFor(raw, custom),
  };
}

const VIEW_DEFAULTS = {
  viewTypeKind: null, scope: null, aggregations: '', aggregationsStatus: '',
  tabularView: true, mobileView: false, mobileDefaultView: false, viewData: '',
  viewJoins: '', readOnlyView: false, includeRootFolder: false, serverRelativeUrl: '',
};

export function normalizeView(raw = {}, viewFieldNames = []) {
  return {
    ...VIEW_DEFAULTS,
    id: raw.Id,
    title: raw.Title,
    defaultView: !!raw.DefaultView,
    hidden: !!raw.Hidden,
    viewType: raw.ViewType || 'HTML',
    viewQuery: raw.ViewQuery || '',
    rowLimit: raw.RowLimit ?? 30,
    paged: raw.Paged !== false,
    customFormatter: raw.CustomFormatter || '',
    jsLink: raw.JSLink || '',
    fields: viewFieldNames || [],
    viewTypeKind: raw.ViewTypeKind ?? null,
    scope: raw.Scope ?? null,
    aggregations: raw.Aggregations || '',
    aggregationsStatus: raw.AggregationsStatus || '',
    tabularView: raw.TabularView !== false,
    mobileView: !!raw.MobileView,
    mobileDefaultView: !!raw.MobileDefaultView,
    viewData: raw.ViewData || '',
    viewJoins: raw.ViewJoins || '',
    readOnlyView: !!raw.ReadOnlyView,
    includeRootFolder: !!raw.IncludeRootFolder,
    serverRelativeUrl: raw.ServerRelativeUrl || '',
  };
}

const CT_DEFAULTS = { parentId: '', sealed: false, documentTemplate: '', fieldLinks: [] };

export function normalizeContentType(raw = {}, fieldLinks = []) {
  const id = raw.StringId || raw.Id?.StringValue || String(raw.Id || '');
  return {
    ...CT_DEFAULTS,
    name: raw.Name,
    id,
    description: raw.Description || '',
    group: raw.Group || '',
    hidden: !!raw.Hidden,
    readOnly: !!raw.ReadOnly,
    parentId: parentContentTypeId(id),
    sealed: !!raw.Sealed,
    documentTemplate: raw.DocumentTemplate || '',
    fieldLinks: fieldLinks || [],
  };
}

export function buildSchemaDoc({
  source = {}, list = {}, fields = [], views = [], contentTypes = [], warnings = [],
  generatorBuild = 'dev',
} = {}) {
  return {
    kind: SCHEMA_KIND,
    version: SCHEMA_VERSION,
    exported: new Date().toISOString(),
    generator: { tool: 'dcspad-workbench', build: generatorBuild },
    source, list, fields, views, contentTypes, warnings,
  };
}

// Accepts a v1 (SPUtils) or v2 (Workbench) document and returns a v2-shaped
// copy with every missing v2 key defaulted. Throws SpFileError('bad-schema')
// when `kind` doesn't match — the one thing both producers agree to check.
export function normalizeSchemaDoc(doc) {
  if (!doc || typeof doc !== 'object' || doc.kind !== SCHEMA_KIND) {
    throw new SpFileError(
      `Not a list schema document (expected kind ${SCHEMA_KIND}).`,
      { code: 'bad-schema' },
    );
  }
  const sourceVersion = Number(doc.version) || 1;
  const list = { ...LIST_DEFAULTS, ...(doc.list || {}) };
  const fields = (doc.fields || []).map((f) => ({ ...FIELD_DEFAULTS, ...f }));
  const views = (doc.views || []).map((v) => ({ ...VIEW_DEFAULTS, ...v }));
  const contentTypes = (doc.contentTypes || []).map((c) => ({
    ...CT_DEFAULTS, ...c, parentId: c.parentId || parentContentTypeId(c.id),
  }));
  return {
    kind: doc.kind,
    version: SCHEMA_VERSION,
    exported: doc.exported || '',
    generator: doc.generator || { tool: 'unknown', build: '' },
    source: doc.source || {},
    list, fields, views, contentTypes,
    warnings: doc.warnings || [],
    _sourceVersion: sourceVersion,
  };
}

const BASE_TEMPLATE_LABELS = { 100: 'Generic list', 101: 'Document library' };

// Chip-ready summary for the Schema tab head. Sentence-case phrases, no
// signal colour — see design/INFO-CHIP.md.
export function schemaSummary(doc) {
  const d = normalizeSchemaDoc(doc);
  const fields = d.fields || [];
  const custom = fields.filter((f) => f.custom).length;
  const views = d.views || [];
  return {
    kind: BASE_TEMPLATE_LABELS[d.list.baseTemplate] || `List (template ${d.list.baseTemplate})`,
    fieldsText: `${fields.length} field${fields.length === 1 ? '' : 's'} · ${custom} custom`,
    viewsText: `${views.length} view${views.length === 1 ? '' : 's'}`,
    contentTypesText: d.list.contentTypesEnabled ? 'content types on' : 'content types off',
    versioningText: d.list.enableVersioning ? 'versioning on' : 'versioning off',
    isLibrary: d.list.baseTemplate === 101 || d.source?.baseType === 1,
  };
}

// ---- SchemaXml scrubbing ----------------------------------------------------

// Attributes tying SchemaXml to its source list/web. Indexed and
// EnforceUniqueValues are stripped too, even though SPUtils leaves them in —
// the Workbench applies them as a separate MERGE step after the field is
// created (some tenants reject them inline on createfieldasxml), so the
// create XML should never carry them.
const STRIP_ATTRS = [
  'ID', 'SourceID', 'ColName', 'RowOrdinal', 'Version', 'WebId', 'List', 'Sealed',
  'Customization', 'Indexed', 'EnforceUniqueValues',
];

// Two corrections beyond the SPUtils basis (see the plan's "corrections"
// section — recorded as SPUtils follow-ups, not changed there):
//   1. `List` is only ever a lookup-list binding; a User field's `List`
//      attribute is the literal string "UserInfo", not a GUID, and must
//      survive the strip instead of being dropped with the rest.
//   2. A Calculated field's <FieldRefs><FieldRef ID="…"/></FieldRefs> carries
//      the SOURCE field's GUID, meaningless on the target; the target
//      resolves each FieldRef by Name; the ID must go.
export function scrubSchemaXml(xml, { lookupListId = null, primaryFieldId = null, fieldType = '' } = {}) {
  let doc;
  try {
    doc = new DOMParser().parseFromString(String(xml || ''), 'text/xml');
  } catch (cause) {
    throw new SpFileError('SchemaXml could not be parsed.', { code: 'bad-xml', cause });
  }
  const root = doc.documentElement;
  if (!root || root.nodeName === 'parsererror' || doc.getElementsByTagName('parsererror').length) {
    throw new SpFileError('SchemaXml could not be parsed.', { code: 'bad-xml' });
  }
  for (const attr of STRIP_ATTRS) root.removeAttribute(attr);
  if (LOOKUP_TYPES.has(fieldType) && lookupListId) {
    root.setAttribute('List', `{${cleanGuid(lookupListId)}}`);
  } else if (USER_TYPES.has(fieldType)) {
    root.setAttribute('List', 'UserInfo');
  }
  if (LOOKUP_TYPES.has(fieldType) && primaryFieldId) {
    root.setAttribute('FieldRef', `{${cleanGuid(primaryFieldId)}}`);
  }
  const refs = root.getElementsByTagName('FieldRef');
  for (let i = 0; i < refs.length; i++) refs[i].removeAttribute('ID');
  return new XMLSerializer().serializeToString(root);
}

// A minimal single-line-of-text field, for the missingLookup:'text' policy —
// the column exists (nothing downstream breaks), but a later re-run cannot
// upgrade it back to a lookup without deleting and recreating it.
export function textFallbackXml(field) {
  const name = String(field?.internalName || '').replaceAll('"', '&quot;');
  const display = String(field?.displayName || field?.internalName || '').replaceAll('"', '&quot;');
  const required = field?.required ? ' Required="TRUE"' : '';
  return `<Field Type="Text" Name="${name}" DisplayName="${display}"${required} />`;
}

export function xmlHasAttr(xml, name) {
  try {
    const doc = new DOMParser().parseFromString(String(xml || ''), 'text/xml');
    return doc.documentElement?.hasAttribute(name) ?? false;
  } catch {
    return false;
  }
}

// ---- apply planning ---------------------------------------------------------
//
// buildApplyPlan is pure given a probe snapshot (list-schema-capture.js's
// probeTarget): it never itself reads or writes SharePoint. The executor
// (list-schema-apply.js, slice 2) walks the resulting Plan and turns each
// Step into a write; buildApplyPlan and the dry-run preview both call this
// same function so the plan a user reviews is the plan that runs.

// `final` marks a refusal the planner made from the probe (a type clash, a
// taken display name, a content type the target lacks): re-running the same
// write cannot succeed, so retryPlan never re-enters it.
function step(id, kind, label, {
  dependsOn = [], payload = {}, refs = {}, optional = false, status = 'planned', error = '',
  final = false,
} = {}) {
  return { id, kind, label, dependsOn, payload, refs, optional, status, error, final, result: null };
}

// Site (parent) content-type ids that every list already carries — never
// something to "attach".
export function isBuiltinParent(parentId) {
  return parentId === '0x01' || parentId === '0x0120' || parentId === '0x0101';
}

// Settings MERGE payloads. Group A is sent as one MERGE (the executor falls
// back to one property at a time if SharePoint rejects the group); group B
// only exists under its flags, the same conditions SPUtils applies. Null and
// undefined values are dropped — a v1 document carries no ReadSecurity etc.,
// and MERGE-ing an explicit null into SP.List is a 400, not a no-op.
// Adopting an existing list reconciles only what an import depends on
// (attachments, folders) and leaves every other setting alone — SPUtils'
// rule, and the one the consent sentence promises.
function settingsPayload(list, { reconcile = false } = {}) {
  if (reconcile) {
    const groupA = {};
    if (list.enableAttachments && !isLibrary(list)) groupA.EnableAttachments = true;
    if (list.enableFolderCreation) groupA.EnableFolderCreation = true;
    return { groupA, groupB: {} };
  }
  const groupA = {
    EnableVersioning: list.enableVersioning,
    EnableAttachments: list.enableAttachments,
    EnableFolderCreation: list.enableFolderCreation,
    EnableModeration: list.enableModeration,
    Hidden: list.hidden,
    OnQuickLaunch: list.onQuickLaunch,
    NoCrawl: list.noCrawl,
    DisableGridEditing: list.disableGridEditing,
    Ordered: list.ordered,
    ReadSecurity: list.readSecurity,
    WriteSecurity: list.writeSecurity,
    ListExperienceOptions: list.listExperienceOptions,
    EnableRequestSignOff: list.enableRequestSignOff,
    ForceCheckout: list.forceCheckout,
  };
  if (list.description) groupA.Description = list.description;
  const groupB = {};
  if (list.enableVersioning && list.majorVersionLimit) groupB.MajorVersionLimit = list.majorVersionLimit;
  if (list.enableVersioning && list.enableMinorVersions) {
    groupB.EnableMinorVersions = true;
    if (list.majorWithMinorVersionsLimit) groupB.MajorWithMinorVersionsLimit = list.majorWithMinorVersionsLimit;
  }
  if (list.enableModeration || list.enableMinorVersions) groupB.DraftVersionVisibility = list.draftVersionVisibility;
  // Libraries 400 on EnableAttachments (stage 2 carries the rest of the
  // library-only set).
  if (isLibrary(list) || Number(list.baseTemplate) === 101) delete groupA.EnableAttachments;
  const defined = (obj) => Object.fromEntries(Object.entries(obj).filter(([, v]) => v != null));
  return { groupA: defined(groupA), groupB: defined(groupB) };
}

export function defaultTargetTitle(doc, targetLists = []) {
  const source = doc?.list?.title || doc?.source?.listTitle || 'List';
  const taken = new Set((targetLists || []).map((l) => String(l.title || '').toLowerCase()));
  if (!taken.has(source.toLowerCase())) return source;
  // First free suffix: "X Copy", then "X Copy 2", "X Copy 3"…
  for (let n = 1; ; n++) {
    const candidate = n === 1 ? `${source} Copy` : `${source} Copy ${n}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}

// Post-create merges (Indexed / EnforceUniqueValues / CustomFormatter) — a
// field carries these regardless of whether it was just created or already
// existed with a matching type on resume, so both paths share this. See the
// scrubSchemaXml comment for why they don't ride the create XML.
function fieldMergePayload(f, warnings) {
  const merges = {};
  if (f.indexed && indexableType(f.type)) merges.Indexed = true;
  if (f.enforceUniqueValues && indexableType(f.type)) merges.EnforceUniqueValues = true;
  if (f.customFormatter && !xmlHasAttr(f.schemaXml, 'CustomFormatter')) {
    merges.CustomFormatter = f.customFormatter;
  }
  if (f.indexed && !indexableType(f.type)) {
    warnings.push(`Column ‘${f.internalName}’ (${f.type}) cannot be indexed — Indexed was not re-applied.`);
  }
  return merges;
}

// Stamped as its own step so the report and the dry-run preview show the
// merge as distinct work from the create/resume.
function pushMergeStep(steps, mergeStepIds, f, dependsOnId, merges) {
  if (!Object.keys(merges).length) return;
  const mergeId = `merge:${f.internalName}`;
  mergeStepIds.push(mergeId);
  steps.push(step(mergeId, 'field.merge', `Apply ${Object.keys(merges).join(', ')} on ‘${f.internalName}’`, {
    dependsOn: [dependsOnId],
    payload: { internalName: f.internalName, merges },
    optional: true,
  }));
}

// probe: { existingList, existingFields, existingViews, existingContentTypeIds,
//          availableContentTypes, targetLists }
export function buildApplyPlan(doc, options = {}, probe = {}) {
  const d = normalizeSchemaDoc(doc);
  const warnings = [];
  const title = String(options.title || '').trim() || defaultTargetTitle(d, probe.targetLists || []);
  const existing = probe.existingList || null;
  const steps = [];

  if (existing) {
    if ((options.existing || 'fail') === 'fail') {
      throw new SpFileError(
        `A list named ‘${title}’ already exists on the target.`,
        { code: 'exists' },
      );
    }
    steps.push(step('list', 'list.adopt', `Add to existing list ‘${title}’`, {
      payload: { listId: existing.id, title },
      refs: { listId: { self: true } },
    }));
  } else {
    steps.push(step('list', 'list.create', `Create list ‘${title}’ (${BASE_TEMPLATE_LABELS[d.list.baseTemplate] || d.list.baseTemplate})`, {
      payload: {
        title,
        description: options.description ?? d.list.description ?? '',
        baseTemplate: d.list.baseTemplate || 100,
        contentTypesEnabled: !!d.list.contentTypesEnabled,
        urlName: options.urlName || '',
      },
    }));
  }

  steps.push(step('settings', 'list.settings', existing ? 'Reconcile list settings' : 'Apply list settings', {
    dependsOn: ['list'],
    payload: settingsPayload(d.list, { reconcile: Boolean(existing) }),
    refs: { listId: { self: true } },
  }));

  // ---- content types --------------------------------------------------
  const ctStepIds = [];
  if (d.list.contentTypesEnabled) {
    const already = new Set(probe.existingContentTypeIds || []);
    const available = new Set((probe.availableContentTypes || []).map((c) => c.id));
    const seen = new Set();
    for (const ct of d.contentTypes) {
      const parentId = ct.parentId || parentContentTypeId(ct.id);
      if (isBuiltinParent(parentId) || already.has(parentId) || seen.has(parentId)) continue;
      seen.add(parentId);
      const id = `ct:${parentId}`;
      ctStepIds.push(id);
      const isAvailable = available.has(parentId);
      steps.push(step(id, 'ct.attach', `Attach content type ‘${ct.name}’`, {
        dependsOn: ['list'],
        payload: { contentTypeId: parentId, name: ct.name },
        refs: { listId: { self: true } },
        optional: true,
        status: isAvailable ? 'planned' : 'failed',
        final: !isAvailable,
        error: isAvailable ? '' : `‘${ct.name}’ is not available on the target web — publish the content type there first.`,
      }));
    }
  }

  // ---- fields (tier order: plain → lookup → dependent lookup → calculated)
  const custom = orderFields(d.fields.filter((f) => f.custom));
  const existingFields = probe.existingFields || [];
  const createdInternalNames = new Set();
  const fieldStepIds = [];
  const mergeStepIds = [];
  const missingLookupPolicy = options.missingLookup === 'text' ? 'text' : 'skip';
  const lookupMap = options.lookupMap || {};
  // A freshly created list already carries SharePoint's base columns
  // (Title, LinkTitle, ID, Created, Modified, Author, Editor, Attachments,
  // DocIcon…) — the executor re-checks each against a live probe later, but
  // the plan itself should count every non-custom source field as available.
  // Columns a fresh list already has: only ones inherited from the list's
  // base type. A non-custom column that came from a site column or content
  // type is not recreated, so it must not be assumed present in a view.
  const baseFieldNames = new Set(d.fields.filter((f) => !f.custom && f.fromBaseType).map((f) => f.internalName));
  // options is the single source of truth the executor and both script
  // emitters read: SP.AddFieldOptions.AddFieldInternalNameHint (8), plus
  // AddToDefaultContentType (4) whenever the list enables content types and
  // at least one is actually being attached.
  const fieldOptions = (d.list.contentTypesEnabled && ctStepIds.length > 0) ? (8 | 4) : 8;

  for (const f of custom) {
    const id = `field:${f.internalName}`;
    fieldStepIds.push(id);
    const present = existingFields.find((ef) => ef.internalName === f.internalName);
    if (present) {
      const same = present.typeAsString === f.type;
      steps.push(step(id, 'field.create', `Column ‘${f.displayName}’ (${f.internalName})`, {
        dependsOn: ['list', ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: same ? 'skipped' : 'failed',
        final: !same,
        error: same ? '' : `exists on the target as ${present.typeAsString}, source is ${f.type} — values will not import.`,
      }));
      if (same) {
        createdInternalNames.add(f.internalName);
        pushMergeStep(steps, mergeStepIds, f, id, fieldMergePayload(f, warnings));
      }
      continue;
    }
    const titleTaken = existingFields.some((ef) =>
      ef.internalName !== f.internalName
      && String(ef.title || '').toLowerCase() === String(f.displayName || '').toLowerCase());
    if (titleTaken) {
      steps.push(step(id, 'field.create', `Column ‘${f.displayName}’ (${f.internalName})`, {
        dependsOn: ['list', ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: 'failed',
        final: true,
        error: `a different column already uses the display name ‘${f.displayName}’.`,
      }));
      continue;
    }
    if (TAXONOMY_TYPES.has(f.type)) {
      // Skipped, not failed: nothing about a re-run could create it (it needs
      // a term-set binding), and a retry must never post its XML unbound.
      steps.push(step(id, 'field.create', `Skip column ‘${f.displayName}’ (${f.internalName}) — managed metadata is not recreated`, {
        dependsOn: ['list', ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: 'skipped',
        error: '',
        final: true,
      }));
      steps[steps.length - 1].skipReason = 'managed-metadata';
      continue;
    }

    const refs = {};
    let asText = false;
    let blocked = false;
    let blockReason = '';

    // Dependent lookups: the primary prerequisite is evaluated BEFORE the
    // missing-target policy below, so a dependent whose primary was skipped
    // (missing lookup target) or failed is 'blocked' — it never gets a
    // chance to independently resolve (and possibly skip on) its own
    // lookup target.
    if (f.isDependentLookup) {
      const primaryField = d.fields.find((pf) => pf.id === f.primaryFieldId);
      if (primaryField && createdInternalNames.has(primaryField.internalName)) {
        refs.primaryFieldId = { field: primaryField.internalName };
      } else {
        blocked = true;
        blockReason = 'its primary lookup column was not created on the target.';
      }
    }

    if (!blocked && LOOKUP_TYPES.has(f.type)) {
      if (f.isSelfLookup) {
        refs.lookupListId = { self: true };
      } else {
        // A mapping may be keyed by the source list's title or, when capture
        // could not read that title, by its source GUID.
        const mappedValue = lookupMapGet(lookupMap, f.lookupList)
          ?? lookupMapGet(lookupMap, f.lookupListId)
          ?? f.lookupList;
        if (isGuidLike(mappedValue)) {
          refs.lookupListId = { id: cleanGuid(mappedValue) };
        } else {
          const wanted = String(mappedValue || '').toLowerCase();
          const target = (probe.targetLists || []).find((l) => String(l.title || '').toLowerCase() === wanted);
          if (!target) {
            if (missingLookupPolicy === 'text') {
              asText = true;
              warnings.push(`Lookup ‘${f.internalName}’ target ‘${mappedValue}’ is missing on the target — created as a single line of text (policy: text; a re-run cannot upgrade it).`);
            } else {
              steps.push(step(id, 'field.create', `Skip column ‘${f.displayName}’ (${f.internalName}) — lookup target ‘${mappedValue}’ is missing (policy: skip)`, {
                dependsOn: ['list', ...ctStepIds],
                payload: { field: f, options: fieldOptions },
                status: 'skipped',
                error: '',
              }));
              steps[steps.length - 1].skipReason = 'lookup-target-missing';
              continue;
            }
          } else {
            refs.lookupListId = { list: target.title };
          }
        }
      }
    }

    steps.push(step(id, 'field.create', `${blocked ? 'Blocked' : 'Add'} ${f.type} column ‘${f.displayName}’ (${f.internalName})${asText ? ' as text' : ''}`, {
      dependsOn: ['list', ...ctStepIds],
      payload: { field: f, asText, options: fieldOptions },
      refs,
      status: blocked ? 'blocked' : 'planned',
      error: blockReason,
    }));
    if (!blocked) createdInternalNames.add(f.internalName);

    if (!blocked && !asText) {
      pushMergeStep(steps, mergeStepIds, f, id, fieldMergePayload(f, warnings));
    }
  }

  // ---- Title column fix-up (display name / required) -------------------
  const titleField = d.fields.find((f) => f.internalName === 'Title');
  if (titleField && (titleField.displayName !== 'Title' || titleField.required === false)) {
    steps.push(step('title', 'field.base', `Set the Title column’s display name and required flag`, {
      dependsOn: ['list'],
      payload: { displayName: titleField.displayName, required: titleField.required },
      optional: true,
    }));
  }

  // ---- views -------------------------------------------------------------
  for (const v of d.views.filter((view) => !view.hidden)) {
    const id = `view:${v.title}`;
    // Title match first; the source's default view also matches the target's
    // own default view, which carries a localized title ("Alle Elemente").
    const targetViews = probe.existingViews || [];
    const existingView = targetViews.find((ev) => String(ev.title).toLowerCase() === String(v.title).toLowerCase())
      || (v.defaultView ? targetViews.find((ev) => ev.defaultView) : null)
      || null;
    const wanted = v.fields.filter((name) => createdInternalNames.has(name)
      || existingFields.some((ef) => ef.internalName === name)
      || baseFieldNames.has(name));
    const missing = v.fields.filter((name) => !wanted.includes(name));
    if (missing.length) {
      warnings.push(`View ‘${v.title}’: columns not on the target were left out: ${missing.join(', ')}.`);
    }
    steps.push(step(id, 'view.upsert', `View ‘${v.title}’ (${wanted.length} column${wanted.length === 1 ? '' : 's'})`, {
      // Only the list: a view that names a column which failed to create
      // simply leaves it out (warned above). Depending on every field step
      // let one managed-metadata column block every view on the list.
      dependsOn: ['list'],
      payload: {
        title: v.title, viewId: existingView?.id || null, fields: wanted,
        viewQuery: v.viewQuery, rowLimit: v.rowLimit, paged: v.paged,
        defaultView: v.defaultView, customFormatter: v.customFormatter, jsLink: v.jsLink,
      },
    }));
  }

  // ---- list-level validation formula (last: it can reference any column) -
  if (d.list.validationFormula) {
    steps.push(step('validation', 'list.validation', 'Apply the list validation formula', {
      // Ordered last so every column the formula names exists; a formula
      // over a column that failed is refused by SharePoint on its own merits.
      dependsOn: ['list'],
      payload: { validationFormula: d.list.validationFormula, validationMessage: d.list.validationMessage },
      optional: true,
    }));
  }

  return { title, targetWebUrl: options.targetWebUrl || '', existingListId: existing?.id || null, steps, warnings };
}

export function newReport(plan) {
  return {
    dryRun: false,
    title: plan.title,
    listId: plan.existingListId || null,
    created: false,
    targetWebUrl: plan.targetWebUrl || '',
    rootFolder: '',
    adopted: Boolean(plan.existingListId),
    aborted: '',
    settings: { applied: [], failed: [] },
    validation: { applied: false, error: '' },
    contentTypes: { attached: 0, skipped: 0, failed: 0 },
    fields: { added: 0, skipped: 0, failed: [] },
    fieldMerges: { applied: 0, failed: 0 },
    views: { added: 0, updated: 0, failed: [] },
    fieldIdMap: {},
    warnings: [...(plan.warnings || [])],
    steps: (plan.steps || []).map((s) => ({ ...s })),
  };
}

// Merge a plan's steps into (or over) an existing report — the shape the
// executor and the dry-run preview share.
export function reportFromPlan(plan, report = {}) {
  return { ...newReport(plan), ...report, steps: (plan.steps || []).map((s) => ({ ...s })) };
}

// A plan carrying only the steps a previous run left 'failed' — the retry
// button's input. Bound to the list the earlier run created (or adopted).
export function retryPlan(report) {
  const steps = (report?.steps || [])
    .filter((s) => s.status === 'failed' && !s.final)
    .map((s) => ({ ...s, status: 'planned', error: '' }));
  return {
    title: report?.title || '',
    targetWebUrl: report?.targetWebUrl || '',
    existingListId: report?.listId || null,
    steps,
    warnings: [],
  };
}

export function buildApplyReport({ report, doc, targetWebUrl }) {
  const title = report?.title || doc?.list?.title || 'List';
  const dest = targetWebUrl || report?.targetWebUrl || 'the target site';
  const lines = ['# List schema report', ''];
  lines.push(report?.created ? `Created ‘${title}’ on ${dest}.` : `Updated ‘${title}’ on ${dest}.`, '');
  lines.push('## Fields', '',
    `- Added: ${report?.fields?.added ?? 0}`,
    `- Skipped: ${report?.fields?.skipped ?? 0}`,
    `- Failed: ${(report?.fields?.failed || []).length}`, '');
  lines.push('## Views', '',
    `- Added: ${report?.views?.added ?? 0}`,
    `- Updated: ${report?.views?.updated ?? 0}`,
    `- Failed: ${(report?.views?.failed || []).length}`, '');
  if (report?.contentTypes) {
    lines.push('## Content types', '',
      `- Attached: ${report.contentTypes.attached ?? 0}`,
      `- Skipped: ${report.contentTypes.skipped ?? 0}`,
      `- Failed: ${report.contentTypes.failed ?? 0}`, '');
  }
  const warnings = report?.warnings || [];
  if (warnings.length) {
    lines.push('## Warnings', '');
    for (const w of warnings) lines.push(`- ${w}`);
    lines.push('');
  }
  const failedSteps = (report?.steps || []).filter((s) => s.status === 'failed');
  if (failedSteps.length) {
    lines.push('## Failed steps', '');
    for (const s of failedSteps) lines.push(`- ${s.label}: ${s.error || 'failed'}`);
    lines.push('');
  }
  return lines.join('\n');
}
