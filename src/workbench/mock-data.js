// Mock /_api fixtures for off-SharePoint development and tests.
//
// Shapes mirror real nometadata responses closely enough that the sp-shapes
// inspector fingerprints them (SP.List needs BaseTemplate+EntityTypeName,
// SP.Field needs InternalName+TypeAsString, principals need
// LoginName+PrincipalType). Keep additions realistic — the mock is the local
// contract for every view.

import { defaultMockWriter } from './sp-write.js';

const WEB_URL = location.origin;

const LISTS = [
  list('Documents', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01', 101, 1, 42, false, '/Shared Documents'),
  list('Site Pages', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02', 119, 1, 17, false, '/SitePages'),
  list('Projects', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03', 100, 0, 128, false, '/Lists/Projects'),
  list('Tasks', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e04', 171, 0, 260, false, '/Lists/Tasks'),
  list('Site Assets', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e05', 101, 1, 96, false, '/SiteAssets'),
  list('User Information List', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e06', 112, 0, 57, true, '/_catalogs/users'),
  list('Master Page Gallery', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e07', 116, 1, 12, true, '/_catalogs/masterpage'),
  list('DevPadData', '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e08', 100, 0, 3, true, '/Lists/DevPadData'),
];

function list(title, id, template, baseType, itemCount, hidden, url) {
  return {
    Id: id,
    Title: title,
    BaseTemplate: template,
    BaseType: baseType,
    ItemCount: itemCount,
    Hidden: hidden,
    Created: '2025-11-02T15:04:00Z',
    LastItemModifiedDate: '2026-07-20T09:30:00Z',
    EntityTypeName: title.replaceAll(' ', '_x0020_'),
    Description: hidden ? '' : `${title} for the mock web.`,
    DefaultViewUrl: `${url}/Forms/AllItems.aspx`,
    RootFolder: { ServerRelativeUrl: url },
  };
}

let fieldSeq = 0;

const FIELDS = {
  '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03': [
    field('Title', 'Title', 'Text', 2, { Required: true }),
    field('Project Status', 'ProjectStatus', 'Choice', 6, {
      Choices: ['Planned', 'Active', 'Blocked', 'Done'], DefaultValue: 'Planned',
    }),
    field('Due Date', 'DueDate', 'DateTime', 4),
    field('Owner', 'ProjectOwner', 'User', 20),
    field('Budget', 'Budget', 'Currency', 10),
    field('Details', 'ProjectDetails', 'Note', 3, { RichText: true }),
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true, Hidden: false }),
    field('Content Type', 'ContentType', 'Computed', 12, { Hidden: true, ReadOnlyField: true }),
  ],
  // Documents library: the full editor-type spread for the Files browser,
  // including a read-only User field displayed via FieldValuesAsText.
  '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01': [
    field('Title', 'Title', 'Text', 2),
    field('Document category', 'DocCategory', 'Choice', 6, {
      Choices: ['Contract', 'Report', 'Misc'], DefaultValue: 'Misc',
    }),
    field('Confidential', 'Confidential', 'Boolean', 8),
    field('Published date', 'PublishedDate', 'DateTime', 4),
    field('Source link', 'SourceLink', 'URL', 11),
    field('DocVersion', 'DocVersion', 'Text', 2),
    field('Author', 'Author', 'User', 20, { ReadOnlyField: true }),
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true }),
  ],
  // Site Pages: one field per editor type the metadata form supports, plus
  // the content fields the editor must refuse to touch.
  '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02': [
    field('Title', 'Title', 'Text', 2, { Required: true }),
    field('Description', 'Description', 'Note', 3),
    field('Page category', 'PageCategory', 'Choice', 6, {
      Choices: ['Announcement', 'How-to', 'Reference'], DefaultValue: 'Reference',
    }),
    field('Review date', 'ReviewDate', 'DateTime', 4),
    field('Show in navigation', 'ShowInNav', 'Boolean', 8),
    field('Related link', 'RelatedLink', 'URL', 11),
    field('Promoted state', 'PromotedState', 'Number', 9, { ReadOnlyField: true }),
    field('Editor', 'Editor', 'User', 20, { ReadOnlyField: true }),
    field('Canvas content', 'CanvasContent1', 'Note', 3),
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true }),
  ],
};

const DEFAULT_FIELDS = [
  field('Title', 'Title', 'Text', 2, { Required: true }),
  field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true }),
  field('Modified', 'Modified', 'DateTime', 4, { ReadOnlyField: true }),
  field('Created', 'Created', 'DateTime', 4, { ReadOnlyField: true }),
];

function field(title, internal, type, kind, extra = {}) {
  fieldSeq++;
  return {
    Id: `af0e2c1d-2222-4444-8888-${String(fieldSeq).padStart(12, '0')}`,
    Title: title,
    InternalName: internal,
    TypeAsString: type,
    FieldTypeKind: kind,
    Required: false,
    Hidden: false,
    ReadOnlyField: false,
    Group: 'Custom Columns',
    DefaultValue: null,
    Description: '',
    EnforceUniqueValues: false,
    Indexed: false,
    SchemaXml: `<Field Name="${internal}" Type="${type}" DisplayName="${title}"/>`,
    ...extra,
  };
}

// Richer field fixture for the List Schema feature: a realistic SchemaXml
// (ID SourceID ColName RowOrdinal Version — the attributes scrubSchemaXml
// exists to strip) plus the raw-row keys captureListSchema reads that the
// plain field() helper above doesn't carry (StaticName, FromBaseType,
// CanBeDeleted, Sealed, CustomFormatter, LookupList, IsDependentLookup,
// PrimaryFieldId…). `xmlAttrs` adds/overrides SchemaXml attributes (List,
// ShowField, FieldRef, Indexed…); `xmlInner` supplies child elements
// (Formula/FieldRefs for Calculated columns).
function schemaField(title, internal, type, kind, { xmlAttrs = {}, xmlInner = '', ...extra } = {}) {
  fieldSeq++;
  const id = `bb1e2c1d-4444-5555-9999-${String(fieldSeq).padStart(12, '0')}`;
  const row = {
    Id: id, Title: title, InternalName: internal, StaticName: internal,
    TypeAsString: type, FieldTypeKind: kind, Required: false, Hidden: false,
    ReadOnlyField: false, FromBaseType: false, CanBeDeleted: true, Sealed: false,
    Group: 'Custom Columns', DefaultValue: null, Description: '',
    EnforceUniqueValues: false, Indexed: false, CustomFormatter: '',
    LookupList: null, LookupField: null, IsDependentLookup: false, PrimaryFieldId: null,
    ...extra,
  };
  const attrs = {
    ID: `{${id}}`, SourceID: '{deadbeef-0000-4000-8000-000000000000}', ColName: `tp_${internal}`,
    RowOrdinal: '0', Version: '1', Name: internal, StaticName: internal, DisplayName: title,
    Type: type,
    ...(row.Required ? { Required: 'TRUE' } : {}),
    ...(row.Indexed ? { Indexed: 'TRUE' } : {}),
    ...(row.EnforceUniqueValues ? { EnforceUniqueValues: 'TRUE' } : {}),
    ...xmlAttrs,
  };
  const attrText = Object.entries(attrs).map(([k, v]) => `${k}="${v}"`).join(' ');
  row.SchemaXml = xmlInner ? `<Field ${attrText}>${xmlInner}</Field>` : `<Field ${attrText} />`;
  return row;
}

// ---- List Schema feature: two dedicated webs ------------------------------
// Served for /sites/schema (capture source) and /sites/target (apply/probe
// target). Own webs, like /sites/classic and /sites/both above, so nothing
// on the default web's row counts moves. See design/list-schema plan,
// "Mock fixtures".

const SCHEMA_REQUESTS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a01';
const SCHEMA_CLIENTS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a02';
const SCHEMA_REGIONS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a03';
const SCHEMA_DOCUMENTS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a04';

const fRequestsTitle = schemaField('Title', 'Title', 'Text', 2, { FromBaseType: true, CanBeDeleted: false, Required: true });
const fRequestsId = schemaField('ID', 'ID', 'Counter', 5, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true });
const fStatus = schemaField('Status', 'Status', 'Choice', 6, { Choices: ['New', 'Active', 'Closed'], DefaultValue: 'New' });
const fNotes = schemaField('Notes', 'RequestNotes', 'Note', 3, { xmlAttrs: { AppendOnly: 'TRUE' } });
const fBudget = schemaField('Budget', 'Budget', 'Number', 9, {});
const fDue = schemaField('Due', 'RequestDue', 'DateTime', 4, {});
const fApproved = schemaField('Approved', 'Approved', 'Boolean', 8, {});
const fReference = schemaField('Reference', 'Reference', 'URL', 11, {});
const fOwner = schemaField('Owner', 'RequestOwner', 'User', 20, { xmlAttrs: { List: 'UserInfo', ShowField: 'Name' } });
const fClient = schemaField('Client', 'Client', 'Lookup', 7, {
  LookupList: SCHEMA_CLIENTS_ID, LookupField: 'Title',
  xmlAttrs: { List: `{${SCHEMA_CLIENTS_ID}}`, ShowField: 'Title' },
});
const fClientCode = schemaField('Client code', 'ClientCode', 'Lookup', 7, {
  LookupList: SCHEMA_CLIENTS_ID, LookupField: 'ClientCode', IsDependentLookup: true, PrimaryFieldId: fClient.Id,
  xmlAttrs: { List: `{${SCHEMA_CLIENTS_ID}}`, ShowField: 'ClientCode', FieldRef: `{${fClient.Id}}` },
});
const fParent = schemaField('Parent request', 'ParentRequest', 'Lookup', 7, {
  LookupList: SCHEMA_REQUESTS_ID, LookupField: 'Title',
  xmlAttrs: { List: `{${SCHEMA_REQUESTS_ID}}`, ShowField: 'Title' },
});
const fRegion = schemaField('Region', 'Region', 'Lookup', 7, {
  LookupList: SCHEMA_REGIONS_ID, LookupField: 'Title',
  xmlAttrs: { List: `{${SCHEMA_REGIONS_ID}}`, ShowField: 'Title' },
});
const fTotal = schemaField('Total', 'Total', 'Calculated', 12, {
  ReadOnlyField: true,
  xmlInner: `<Formula>=[Budget]*1</Formula><FieldRefs><FieldRef ID="{${fBudget.Id}}" Name="Budget"/></FieldRefs>`,
});
const fTracking = schemaField('Tracking code', 'TrackingCode', 'Text', 2, { Indexed: true, EnforceUniqueValues: true });
const fFormatted = schemaField('Formatted note', 'FormattedNote', 'Text', 2, { CustomFormatter: '{"schema":"https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json"}' });
const fCategory = schemaField('Category', 'RequestCategory', 'TaxonomyFieldType', 26, {});
const fContentType = schemaField('Content Type', 'ContentType', 'Computed', 12, { FromBaseType: true, CanBeDeleted: false, Hidden: true, ReadOnlyField: true });
const fAttachmentsF = schemaField('Attachments', 'Attachments', 'Attachments', 23, { FromBaseType: true, CanBeDeleted: false, Hidden: true });
const fAuthor = schemaField('Created By', 'Author', 'User', 20, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, xmlAttrs: { List: 'UserInfo' } });
const fLinkTitle = schemaField('Title', 'LinkTitle', 'Computed', 12, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true });

const SCHEMA_REQUESTS_FIELDS = [
  fRequestsTitle, fRequestsId, fStatus, fNotes, fBudget, fDue, fApproved, fReference, fOwner,
  fClient, fClientCode, fParent, fRegion, fTotal, fTracking, fFormatted, fCategory,
  fContentType, fAttachmentsF, fAuthor, fLinkTitle,
];

const SCHEMA_REQUESTS_LIST = {
  Id: SCHEMA_REQUESTS_ID, Title: 'Requests', BaseTemplate: 100, BaseType: 0, ItemCount: 12,
  Hidden: false, Created: '2025-01-05T00:00:00Z', LastItemModifiedDate: '2026-08-01T00:00:00Z',
  EntityTypeName: 'Requests', Description: 'Schema source list.',
  DefaultViewUrl: '/sites/schema/Lists/Requests/AllItems.aspx',
  RootFolder: { ServerRelativeUrl: '/sites/schema/Lists/Requests', Name: 'Requests' },
  ContentTypesEnabled: true, EnableVersioning: true, MajorVersionLimit: 50,
  EnableMinorVersions: false, ForceCheckout: false, EnableAttachments: true,
  EnableFolderCreation: false, EnableModeration: false, OnQuickLaunch: true,
  ValidationFormula: '=[Budget]>0', ValidationMessage: 'Budget must be positive.',
  NoCrawl: false, DisableGridEditing: false, Ordered: false,
  ReadSecurity: 1, WriteSecurity: 1, ListExperienceOptions: 0, EnableRequestSignOff: false,
};
const SCHEMA_CLIENTS_LIST = list('Clients', SCHEMA_CLIENTS_ID, 100, 0, 5, false, '/sites/schema/Lists/Clients');
const SCHEMA_REGIONS_LIST = list('Regions', SCHEMA_REGIONS_ID, 100, 0, 4, false, '/sites/schema/Lists/Regions');
const SCHEMA_DOCUMENTS_LIST = list('Documents', SCHEMA_DOCUMENTS_ID, 101, 1, 6, false, '/sites/schema/Documents');

const SCHEMA_LISTS = [SCHEMA_REQUESTS_LIST, SCHEMA_CLIENTS_LIST, SCHEMA_REGIONS_LIST, SCHEMA_DOCUMENTS_LIST];

const SCHEMA_FIELDS = {
  [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_FIELDS,
  [SCHEMA_CLIENTS_ID]: [
    schemaField('Title', 'Title', 'Text', 2, { FromBaseType: true, CanBeDeleted: false }),
    schemaField('Client code', 'ClientCode', 'Text', 2, {}),
  ],
  [SCHEMA_REGIONS_ID]: [schemaField('Title', 'Title', 'Text', 2, { FromBaseType: true, CanBeDeleted: false })],
  [SCHEMA_DOCUMENTS_ID]: [schemaField('Title', 'Title', 'Text', 2, {})],
};

const SCHEMA_REQUESTS_VIEWS = [
  {
    Id: 'cc1e2c1d-5555-6666-aaaa-000000000001', Title: 'All Items', DefaultView: true,
    PersonalView: false, Hidden: false, ServerRelativeUrl: '/sites/schema/Lists/Requests/AllItems.aspx',
    RowLimit: 30, Paged: true, ViewQuery: '<OrderBy><FieldRef Name="ID"/></OrderBy>',
    ViewFields: { Items: ['LinkTitle', 'Status', 'Budget'] },
  },
  {
    Id: 'cc1e2c1d-5555-6666-aaaa-000000000002', Title: 'Active only', DefaultView: false,
    PersonalView: false, Hidden: false, ServerRelativeUrl: '/sites/schema/Lists/Requests/Active.aspx',
    RowLimit: 100, Paged: true,
    ViewQuery: '<Where><Eq><FieldRef Name="Status"/><Value Type="Choice">Active</Value></Eq></Where>',
    ViewFields: { Items: ['LinkTitle', 'Status', 'Client'] },
  },
];

const REQUEST_PARENT_CT = '0x0100442912F2B6C7409A8FF25CE5504F1FD';
const REQUEST_LIST_CT = `${REQUEST_PARENT_CT}00${'A'.repeat(32)}`;
const SCHEMA_REQUESTS_CTS = [
  { Id: { StringValue: REQUEST_LIST_CT }, Name: 'Request', Group: 'Custom Content Types', Hidden: false, ReadOnly: false, Sealed: false, Description: 'A schema-source request.' },
  { Id: { StringValue: '0x01' }, Name: 'Item', Group: 'List Content Types', Hidden: false, ReadOnly: false, Sealed: false, Description: 'Create a new list item.' },
];

const VIEWS_BY_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_VIEWS };
const CONTENT_TYPES_BY_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_CTS };

// List Schema feature, stage 1b-a: a few Requests items exercising the
// shapes captureListData has to carry — a plain item, a self-lookup
// (ParentRequest) referencing another item, a User value (RequestOwner), an
// attachment (metadata only — never bytes), and a folder (FSObjType 1) with
// its own item row. `item()` (below, hoisted) supplies Id/ID/Title/Modified/
// Created; every internal name here matches SCHEMA_REQUESTS_FIELDS above.
const SCHEMA_REQUESTS_ITEMS = [
  item(1, 'Server upgrade', {
    Status: 'Active', Budget: 5000, RequestDue: '2026-09-01T00:00:00Z', Approved: true,
    RequestOwnerId: 11, FSObjType: 0,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/1_.000`,
    FieldValuesAsText: { Status: 'Active', RequestOwner: 'Mock Developer', Budget: '5000', Approved: 'Yes' },
  }),
  item(2, 'Server upgrade — phase 2', {
    Status: 'New', Budget: 2000, ParentRequestId: 1, Attachments: true, FSObjType: 0,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/2_.000`,
    AttachmentFiles: [{
      FileName: 'quote.pdf',
      ServerRelativeUrl: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/Attachments/2/quote.pdf`,
    }],
    FieldValuesAsText: { Status: 'New', ParentRequest: 'Server upgrade' },
  }),
  item(3, 'Archive', {
    FSObjType: 1,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/Archive`,
  }),
];
const ITEMS_BY_SCHEMA_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_ITEMS };

// Site users the schema web's captured User values resolve against (web/
// siteusers) — same shape as CURRENT_USER below (Id, Email, LoginName, Title).
const SCHEMA_SITE_USERS = [
  user(11, 'Mock Developer', 'dev@mock.local', true),
  user(14, 'Pat Example', 'pat@mock.local', false),
];

const TARGET_CLIENTS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b01';
const TARGET_TASKS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b02';
const TARGET_DOCUMENTS_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b03';
const TARGET_ARCHIVE_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b04';

// A reconcile target: has Title plus one of the source's custom fields
// already, under a different display name story (Status), so the "add
// missing fields and views" path has real work to do. No Regions list here
// on purpose — it exercises the missing-lookup-target skip policy.
const TARGET_LISTS = [
  list('Clients', TARGET_CLIENTS_ID, 100, 0, 3, false, '/sites/target/Lists/Clients'),
  list('Tasks', TARGET_TASKS_ID, 171, 0, 10, false, '/sites/target/Lists/Tasks'),
  list('Documents', TARGET_DOCUMENTS_ID, 101, 1, 4, false, '/sites/target/Documents'),
  list('Archive Requests', TARGET_ARCHIVE_ID, 100, 0, 2, false, '/sites/target/Lists/ArchiveRequests'),
];
const TARGET_FIELDS = {
  [TARGET_ARCHIVE_ID]: [
    schemaField('Title', 'Title', 'Text', 2, { FromBaseType: true, CanBeDeleted: false }),
    schemaField('Status', 'Status', 'Choice', 6, { Choices: ['New', 'Active', 'Closed'], DefaultValue: 'New' }),
  ],
  [TARGET_CLIENTS_ID]: [schemaField('Title', 'Title', 'Text', 2, { FromBaseType: true, CanBeDeleted: false })],
};
// Carries the Requests parent content type (so attach-CT can find it
// available) but deliberately no second source content type.
const TARGET_AVAILABLE_CTS = [
  { StringId: REQUEST_PARENT_CT, Name: 'Request', Group: 'Custom Content Types' },
  { StringId: '0x0101', Name: 'Document', Group: 'Document Content Types' },
];

// ---- List Schema feature: stateful mock writer -----------------------------
// A tiny in-memory SharePoint for the (slice 2) apply dialog: registers what
// gets POSTed under the posting web's base, so a probe that follows — "does
// this list exist now?", "open the new list" — resolves against what THIS
// session wrote instead of 404ing against the static fixtures above.
const writerState = {
  lists: new Map(),        // `${webBase}::${lower title}` -> list row
  fields: new Map(),       // listId -> Map(internalName -> field row)
  views: new Map(),        // listId -> Map(viewId -> { Id, Title, fields: [] })
  contentTypes: new Map(), // listId -> Set(parentId)
};

const webBaseOf = (url) => {
  const s = String(url);
  const i = s.indexOf('/_api/');
  return i === -1 ? '' : s.slice(0, i).replace(/\/+$/, '');
};

function xmlAttr(xml, name) {
  const m = new RegExp(`${name}="([^"]*)"`, 'i').exec(String(xml || ''));
  return m ? m[1] : '';
}

let mockWriteSeq = 0;
const nextMockId = (prefix) => `${prefix}${String(++mockWriteSeq).padStart(8, '0')}`;

// A list born via the stateful mock writer needs SharePoint's own base
// columns present immediately — the apply executor's post-create
// fields.probe (and any UI opening the new list right after) reads them
// unfiltered, same as a real tenant would answer.
function baseFieldRow(internalName, title, type, extra = {}) {
  return {
    Id: nextMockId('bf'), InternalName: internalName, Title: title, TypeAsString: type,
    FromBaseType: true, CanBeDeleted: false, Hidden: false, ReadOnlyField: false, ...extra,
  };
}
function baseFieldRows() {
  return [
    baseFieldRow('Title', 'Title', 'Text'),
    baseFieldRow('ID', 'ID', 'Counter', { ReadOnlyField: true }),
    baseFieldRow('LinkTitle', 'Title', 'Computed'),
    baseFieldRow('LinkTitleNoMenu', 'Title', 'Computed', { Hidden: true }),
    baseFieldRow('Attachments', 'Attachments', 'Attachments', { Hidden: true }),
    baseFieldRow('ContentType', 'Content Type', 'Computed', { Hidden: true }),
    baseFieldRow('Created', 'Created', 'DateTime', { ReadOnlyField: true }),
    baseFieldRow('Modified', 'Modified', 'DateTime', { ReadOnlyField: true }),
    baseFieldRow('Author', 'Created By', 'User', { ReadOnlyField: true }),
    baseFieldRow('Editor', 'Modified By', 'User', { ReadOnlyField: true }),
  ];
}

// mockWriter(url, body, contentType, headers) — the shape createSpWriteClient
// calls (sp-write.js). Falls back to sp-write's own defaultMockWriter for
// anything it doesn't recognize, so ValidateUpdateListItem/AddUsingPath/etc.
// used elsewhere in the Workbench keep working unmodified. Every handled
// branch records its own write exactly once via record() below — falling
// through to defaultMockWriter records once more, there — so a call is
// never double-counted in __DCSPAD_WB_WRITES__.
export function mockWriter(url, body, contentType, headers = {}) {
  const webBase = webBaseOf(url);
  const path = String(url).slice(String(url).indexOf('/_api/') + 6).toLowerCase();
  const method = headers?.['X-HTTP-Method'] || headers?.['x-http-method'] || '';
  const record = () => { (globalThis.__DCSPAD_WB_WRITES__ ||= []).push({ url, body, contentType, headers }); };

  if (/^web\/lists$/.test(path) && !method) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const id = nextMockId('cc00');
    let base = '';
    try { base = new URL(webBase).pathname.replace(/\/+$/, ''); } catch { /* keep '' */ }
    const rootUrl = `${base}/Lists/${String(data.Title || 'List').replace(/\s+/g, '')}`;
    const entry = {
      Id: id, Title: data.Title, BaseTemplate: data.BaseTemplate ?? 100,
      // Stage 1b-b's library gate (list-data-apply.js) reads BaseType off the
      // target list before importing items — a schema-created list is always
      // a generic list (BaseTemplate 100), never a library.
      BaseType: 0,
      ContentTypesEnabled: !!data.ContentTypesEnabled, RootFolder: { ServerRelativeUrl: rootUrl },
    };
    writerState.lists.set(`${webBase}::${String(data.Title || '').toLowerCase()}`, entry);
    const fields = new Map();
    for (const row of baseFieldRows()) fields.set(row.InternalName, row);
    writerState.fields.set(id, fields);
    const views = new Map();
    const viewId = nextMockId('vv00');
    views.set(viewId, { Id: viewId, Title: 'All Items', fields: ['LinkTitle'], defaultView: true });
    writerState.views.set(id, views);
    writerState.contentTypes.set(id, new Set());
    record();
    return { Id: id, Title: entry.Title, RootFolder: entry.RootFolder };
  }

  const listIdMatch = /lists\(guid'([0-9a-f-]+)'\)/i.exec(path);
  const listId = listIdMatch?.[1];

  // MERGE on the list itself, a field, or a view updates the registry so a
  // follow-up probe (or the executor's own re-read before the view steps)
  // sees what changed — scoped to what a follow-up read actually surfaces:
  // Title (list/view), Indexed/EnforceUniqueValues (field), DefaultView
  // (view). Settings/validation MERGEs still record but have no
  // registry-visible field to update.
  if (listId && method === 'MERGE') {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    if (new RegExp(`^web/lists\\(guid'${listId}'\\)$`).test(path)) {
      const entry = [...writerState.lists.values()].find((l) => l.Id === listId);
      if (entry && data.Title !== undefined) entry.Title = data.Title;
      record();
      return {};
    }
    const fieldMerge = new RegExp(`^web/lists\\(guid'${listId}'\\)/fields\\(guid'([0-9a-f-]+)'\\)$`).exec(path);
    if (fieldMerge) {
      const fields = writerState.fields.get(listId);
      const field = fields && [...fields.values()].find((f) => f.Id === fieldMerge[1]);
      if (field) {
        if (data.Indexed !== undefined) field.Indexed = !!data.Indexed;
        if (data.EnforceUniqueValues !== undefined) field.EnforceUniqueValues = !!data.EnforceUniqueValues;
        if (data.Title !== undefined) field.Title = data.Title;
      }
      record();
      return {};
    }
    const viewMergeMatch = new RegExp(`^web/lists\\(guid'${listId}'\\)/views\\(guid'([0-9a-f-]+)'\\)$`).exec(path);
    if (viewMergeMatch) {
      const view = writerState.views.get(listId)?.get(viewMergeMatch[1]);
      if (view) {
        if (data.Title !== undefined) view.Title = data.Title;
        if (data.DefaultView !== undefined) view.defaultView = !!data.DefaultView;
      }
      record();
      return {};
    }
    record();
    return {};
  }

  if (listId && path.includes('createfieldasxml')) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const xml = data?.parameters?.SchemaXml || '';
    const optionsBits = Number(data?.parameters?.Options) || 0;
    let internalName = xmlAttr(xml, 'Name') || xmlAttr(xml, 'StaticName');
    const displayName = xmlAttr(xml, 'DisplayName') || internalName;
    // Options bit 8 = AddFieldInternalNameHint. Without it, the mock mangles
    // the internal name from the display name — mirroring the real gap in
    // SPUtils' string-overload createFieldAsXml this port fixes by always
    // sending Options; a test pins the flag by checking the mock mangles it.
    if (!(optionsBits & 8)) internalName = displayName.replace(/[^A-Za-z0-9]+/g, '_x0020_');
    const type = xmlAttr(xml, 'Type');
    const fields = writerState.fields.get(listId) || new Map();
    const id = nextMockId('ff00');
    fields.set(internalName, { Id: id, InternalName: internalName, Title: displayName, TypeAsString: type });
    writerState.fields.set(listId, fields);
    record();
    return { Id: id, InternalName: internalName };
  }

  if (listId && path.includes('addavailablecontenttype')) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const cts = writerState.contentTypes.get(listId) || new Set();
    cts.add(data.contentTypeId);
    writerState.contentTypes.set(listId, cts);
    record();
    return {};
  }

  if (listId && /\/views$/.test(path) && !method) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const id = nextMockId('vv00');
    const views = writerState.views.get(listId) || new Map();
    views.set(id, { Id: id, Title: data.Title, fields: [], defaultView: !!data.DefaultView });
    writerState.views.set(listId, views);
    record();
    return { Id: id, Title: data.Title };
  }

  const viewMatch = /views\(guid'([0-9a-f-]+)'\)/i.exec(path);
  if (listId && viewMatch && path.includes('removeallviewfields')) {
    const view = writerState.views.get(listId)?.get(viewMatch[1]);
    if (view) view.fields = [];
    record();
    return {};
  }
  const addViewField = /addviewfield\('([^']*)'\)/i.exec(path);
  if (listId && viewMatch && addViewField) {
    const view = writerState.views.get(listId)?.get(viewMatch[1]);
    if (view) view.fields.push(decodeURIComponent(addViewField[1]));
    record();
    return {};
  }

  // Everything else (validation MERGE with no registry-visible effect, and
  // anything not touched by the schema feature) falls back to the pad's own
  // default mock behavior, which records the write itself.
  return defaultMockWriter(url, body, contentType, headers);
}

// List items, keyed by list id. Only the lists a Tier 2 view exercises need
// fixtures; unknown lists resolve to an empty collection.
const PROJECT_ITEMS = [
  item(1, 'Intranet refresh', {
    ProjectStatus: 'Active',
    DueDate: '2026-09-15T00:00:00Z',
    Budget: 12000,
    // Rich text + attachments: the item-export markdown path needs both.
    ProjectDetails: '<div><p>Kickoff <strong>done</strong>.</p><ul><li>Phase 1</li>'
      + '<li>Phase 2</li></ul><p>See the <a href="https://example.com/plan">plan</a>.</p></div>',
    AttachmentFiles: [
      { FileName: 'kickoff.pptx', ServerRelativeUrl: '/Lists/Projects/Attachments/1/kickoff.pptx' },
    ],
    FieldValuesAsText: { Author: 'Mock Developer', Editor: 'Pat Example', ProjectOwner: 'Mock Developer' },
  }),
  item(2, 'Records migration', { ProjectStatus: 'Planned', DueDate: '2026-11-01T00:00:00Z', Budget: 40000 }),
  item(3, 'Team site cleanup', { ProjectStatus: 'Done', DueDate: '2026-03-30T00:00:00Z', Budget: 1500 }),
  item(4, 'Permission audit', { ProjectStatus: 'Blocked', DueDate: '2026-08-05T00:00:00Z', Budget: 0 }),
  item(5, 'Search tuning', { ProjectStatus: 'Active', DueDate: '2026-10-20T00:00:00Z', Budget: 8000 }),
  item(6, 'Archive rollout', { ProjectStatus: 'Planned', DueDate: '2027-01-10T00:00:00Z', Budget: 22000 }),
];

function item(id, title, extra = {}) {
  return {
    Id: id,
    ID: id,
    Title: title,
    Modified: '2026-07-18T10:00:00Z',
    Created: '2026-05-02T09:00:00Z',
    ...extra,
  };
}

// Modern-page canvas fixtures for the Pages view. The JSON payload covers a
// full-width text section, a 6/6 two-column section (known + unknown web
// parts), a page-settings slice, and one malformed entry so the tolerant
// parser's degrade path stays exercised.
const HOME_CANVAS = JSON.stringify([
  {
    controlType: 4,
    id: 'a1000000-0000-4000-8000-000000000001',
    position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12, layoutIndex: 1 },
    emphasis: {},
    innerHTML: '<h2>Welcome</h2><p>Welcome to the mock intranet home page.</p>',
  },
  {
    controlType: 3,
    id: 'a1000000-0000-4000-8000-000000000002',
    position: { zoneIndex: 2, sectionIndex: 1, controlIndex: 1, sectionFactor: 6, layoutIndex: 1 },
    emphasis: { zoneEmphasis: 1 },
    webPartId: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
    webPartData: {
      id: 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd',
      title: 'Quick links',
      description: 'Mock quick links',
      properties: { items: [{ title: 'Docs' }, { title: 'Pad' }] },
      serverProcessedContent: {
        htmlStrings: {},
        searchablePlainTexts: { 'items[0].title': 'Docs', 'items[1].title': 'Pad' },
        imageSources: {},
        links: { baseUrl: '/SitePages' },
      },
    },
  },
  {
    controlType: 3,
    id: 'a1000000-0000-4000-8000-000000000003',
    position: { zoneIndex: 2, sectionIndex: 2, controlIndex: 1, sectionFactor: 6, layoutIndex: 1 },
    emphasis: { zoneEmphasis: 1 },
    webPartId: 'ffff0000-1111-2222-3333-444455556666',
    webPartData: {
      id: 'ffff0000-1111-2222-3333-444455556666',
      title: 'Mystery part',
      properties: {},
      serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
    },
  },
  { horrible: 'shape', with: ['no', 'controlType'] },
  { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } },
]);

// Legacy HTML storage format — attribute-encoded control JSON + nested RTE.
const LEGACY_CANVAS = '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0"'
  + ' data-sp-controldata="{&quot;controlType&quot;:4,&quot;id&quot;:&quot;b2000000-0000-4000-8000-000000000001&quot;,'
  + '&quot;position&quot;:{&quot;zoneIndex&quot;:1,&quot;sectionIndex&quot;:1,&quot;controlIndex&quot;:1,&quot;sectionFactor&quot;:12}}">'
  + '<div data-sp-rte=""><p>Legacy formatted news body.</p></div></div></div>';

const SITEPAGES_ITEMS = [
  {
    ...item(1, 'Home', {
      FileLeafRef: 'Home.aspx',
      FileRef: '/SitePages/Home.aspx',
      FileDirRef: '/SitePages',
      PromotedState: 0,
      UniqueId: 'ee000000-0000-4000-8000-000000000001',
      Author: { Title: 'Mock Developer' },
      Editor: { Title: 'Mock Developer' },
      CanvasContent1: HOME_CANVAS,
      LayoutWebpartsContent: null,
      Description: 'Mock landing page.',
      BannerImageUrl: null,
      PageCategory: 'Announcement',
      ReviewDate: '2026-08-01T00:00:00Z',
      ShowInNav: true,
      RelatedLink: { Url: 'https://example.com', Description: 'Example' },
      FieldValuesAsText: { Editor: 'Mock Developer', CanvasContent1: '(canvas markup)' },
    }),
  },
  {
    ...item(2, 'Release notes', {
      FileLeafRef: 'News-Update.aspx',
      FileRef: '/SitePages/News-Update.aspx',
      FileDirRef: '/SitePages',
      PromotedState: 2,
      UniqueId: 'ee000000-0000-4000-8000-000000000002',
      Author: { Title: 'Pat Example' },
      Editor: { Title: 'Pat Example' },
      CanvasContent1: LEGACY_CANVAS,
      FieldValuesAsText: { Editor: 'Pat Example' },
    }),
  },
  {
    ...item(3, 'Blank page', {
      FileLeafRef: 'Blank.aspx',
      FileRef: '/SitePages/Blank.aspx',
      FileDirRef: '/SitePages',
      PromotedState: 0,
      UniqueId: 'ee000000-0000-4000-8000-000000000003',
      Author: { Title: 'Mock Developer' },
      Editor: { Title: 'Mock Developer' },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: 'Mock Developer' },
    }),
  },
  // Pages in subfolders — the Pages view surfaces and sorts by folder.
  {
    ...item(4, 'Weekly roundup', {
      FileLeafRef: 'Weekly.aspx',
      FileRef: '/SitePages/news/Weekly.aspx',
      FileDirRef: '/SitePages/news',
      PromotedState: 2,
      UniqueId: 'ee000000-0000-4000-8000-000000000004',
      Author: { Title: 'Pat Example' },
      Editor: { Title: 'Pat Example' },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: 'Pat Example' },
    }),
  },
  {
    ...item(5, 'Résumé hebdo', {
      FileLeafRef: 'Hebdo.aspx',
      FileRef: '/SitePages/news/fr/Hebdo.aspx',
      FileDirRef: '/SitePages/news/fr',
      PromotedState: 2,
      UniqueId: 'ee000000-0000-4000-8000-000000000005',
      Author: { Title: 'Mock Developer' },
      Editor: { Title: 'Mock Developer' },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: 'Mock Developer' },
    }),
  },
];

const ITEMS = {
  '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03': PROJECT_ITEMS,
  '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02': SITEPAGES_ITEMS,
};

// Folder tree for the Files browser, keyed by lower-cased server-relative
// path. Deliberately includes binary types (.docx/.png/.zip) — the browser
// must list EVERYTHING, unlike the pad picker's code/text filter.
function mockFile(name, length, modified = '2026-07-10T09:00:00Z') {
  return {
    Name: name,
    ServerRelativeUrl: `__FOLDER__/${name}`,
    Length: length,
    TimeLastModified: modified,
    UIVersionLabel: '1.0',
    CheckOutType: 2,
  };
}

const MOCK_TREE = {
  '/shared documents': {
    folders: [
      { Name: 'Reports', ServerRelativeUrl: '/Shared Documents/Reports', ItemCount: 2, TimeLastModified: '2026-07-01T12:00:00Z' },
    ],
    files: [
      mockFile('proposal.docx', 48230),
      mockFile('logo.png', 15872),
      mockFile('archive.zip', 1048576),
      mockFile('notes.txt', 812),
      mockFile('widget.js', 2048),
      mockFile('data.csv', 5300),
    ],
  },
  '/shared documents/reports': {
    folders: [],
    files: [mockFile('q1-report.docx', 91000), mockFile('q2-report.docx', 87000)],
  },
};
for (const [folderPath, listing] of Object.entries(MOCK_TREE)) {
  for (const f of listing.files) {
    f.ServerRelativeUrl = f.ServerRelativeUrl.replace(
      '__FOLDER__',
      folderPath === '/shared documents' ? '/Shared Documents' : '/Shared Documents/Reports',
    );
  }
}

// File list-item metadata, keyed by lower-cased server-relative path.
const FILE_ITEMS = {
  '/shared documents/proposal.docx': {
    Id: 201,
    Title: 'Project proposal',
    DocCategory: 'Contract',
    Confidential: true,
    PublishedDate: '2026-06-01T00:00:00Z',
    SourceLink: { Url: 'https://example.com/spec', Description: 'Spec' },
    DocVersion: '1.4',
    FieldValuesAsText: { Author: 'Mock Developer' },
  },
};

const DOC_LIB_ID = '5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01';

const VIEWS = [
  { Id: 'bb0e2c1d-3333-4444-8888-000000000001', Title: 'All Items', DefaultView: true, PersonalView: false, Hidden: false, ServerRelativeUrl: '/Lists/Projects/AllItems.aspx', RowLimit: 30, Paged: true, ViewQuery: '<OrderBy><FieldRef Name="ID"/></OrderBy>' },
  { Id: 'bb0e2c1d-3333-4444-8888-000000000002', Title: 'Active only', DefaultView: false, PersonalView: false, Hidden: false, ServerRelativeUrl: '/Lists/Projects/Active.aspx', RowLimit: 100, Paged: true, ViewQuery: '<Where><Eq><FieldRef Name="ProjectStatus"/><Value Type="Choice">Active</Value></Eq></Where>' },
];

const CONTENT_TYPES = [
  { Id: { StringValue: '0x0100A1B2C3D4E5F601' }, Name: 'Item', Group: 'List Content Types', Hidden: false, ReadOnly: false, Sealed: false, Description: 'Create a new list item.' },
  { Id: { StringValue: '0x0120001122334455' }, Name: 'Folder', Group: '_Hidden', Hidden: true, ReadOnly: false, Sealed: true, Description: '' },
];

const GROUPS = [
  { Id: 3, Title: 'Mock Site Owners', Description: 'Full control of the mock site.', OwnerTitle: 'System Account', PrincipalType: 8, LoginName: 'Mock Site Owners', OnlyAllowMembersViewMembership: false },
  { Id: 5, Title: 'Mock Site Members', Description: 'Contribute to the mock site.', OwnerTitle: 'Mock Site Owners', PrincipalType: 8, LoginName: 'Mock Site Members', OnlyAllowMembersViewMembership: false },
  { Id: 7, Title: 'Mock Site Visitors', Description: 'Read-only visitors.', OwnerTitle: 'Mock Site Owners', PrincipalType: 8, LoginName: 'Mock Site Visitors', OnlyAllowMembersViewMembership: true },
];

const GROUP_USERS = {
  3: [user(11, 'Mock Developer', 'dev@mock.local', true)],
  5: [user(11, 'Mock Developer', 'dev@mock.local', true), user(14, 'Pat Example', 'pat@mock.local', false)],
  7: [user(19, 'Ronnie Reader', 'ronnie@mock.local', false)],
};

function user(id, title, email, admin) {
  return {
    Id: id,
    Title: title,
    LoginName: `i:0#.f|membership|${email}`,
    Email: email,
    IsSiteAdmin: admin,
    PrincipalType: 1,
  };
}

const ROLE_DEFINITIONS = [
  roleDef(1073741829, 'Full Control', 'Has full control.', 5, '2147483647', '4294967295'),
  roleDef(1073741827, 'Contribute', 'Can view, add, update, and delete list items and documents.', 3, '432', '1011028719'),
  roleDef(1073741826, 'Read', 'Can view pages and list items and download documents.', 2, '176', '138612833'),
  roleDef(1073741825, 'Limited Access', 'Can view specific lists when given access.', 1, '176', '138612801'),
];

function roleDef(id, name, description, kind, high, low) {
  return {
    Id: id, Name: name, Description: description, RoleTypeKind: kind,
    Hidden: kind === 1, BasePermissions: { High: high, Low: low },
  };
}

const ROLE_ASSIGNMENTS = [
  assignment(3, 'Mock Site Owners', 8, ['Full Control']),
  assignment(5, 'Mock Site Members', 8, ['Contribute']),
  assignment(7, 'Mock Site Visitors', 8, ['Read']),
];

function assignment(principalId, title, principalType, roleNames) {
  return {
    PrincipalId: principalId,
    Member: { Id: principalId, Title: title, LoginName: title, PrincipalType: principalType },
    RoleDefinitionBindings: roleNames.map((name) => ({
      Id: ROLE_DEFINITIONS.find((r) => r.Name === name)?.Id || 0, Name: name,
    })),
  };
}

const WEB = {
  Id: 'c0ffee00-1111-2222-3333-444455556666',
  Title: 'Mock Web',
  Description: 'Local workbench mock web.',
  Url: WEB_URL,
  ServerRelativeUrl: '/',
  WebTemplate: 'SITEPAGEPUBLISHING',
  Configuration: 0,
  Created: '2025-10-01T12:00:00Z',
  LastItemModifiedDate: '2026-07-25T08:00:00Z',
  Language: 1033,
  UIVersion: 15,
  QuickLaunchEnabled: true,
  MembersCanShare: true,
};

const SITE = {
  Id: 'deadbeef-7777-8888-9999-aaaabbbbcccc',
  Url: WEB_URL,
  ServerRelativeUrl: '/',
  ReadOnly: false,
  ShareByEmailEnabled: false,
};

const FEATURES = {
  site: [
    { DefinitionId: 'b50e3104-6812-424f-a011-cc90e6327318', DisplayName: 'BasicWebParts' },
    { DefinitionId: '8c6a6980-c3d9-440e-944c-77f93bc65a7e', DisplayName: '' },
  ],
  web: [
    { DefinitionId: '00bfea71-4ea5-48d4-a4ad-7ea5c011abe5', DisplayName: 'TeamCollab' },
    { DefinitionId: 'f151bb39-7c3b-414f-bb36-6bf18872052f', DisplayName: '' },
  ],
};

const SUBWEBS = [
  { Id: 'aaaa1111-0000-0000-0000-000000000001', Title: 'Archive', ServerRelativeUrl: '/archive', WebTemplate: 'STS', Created: '2025-12-01T00:00:00Z', Language: 1033 },
];

const ALL_PROPERTIES = {
  vti_x005f_defaultlanguage: 'en-us',
  vti_x005f_extenderversion: '16.0.0.26000',
  taxonomyhiddenlist: '{5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e99}',
  dcspad_x005f_deployfolder: '/Dev/tools/dcspad',
};

const REGIONAL_SETTINGS = {
  LocaleId: 1033,
  Time24: false,
  FirstDayOfWeek: 0,
  WorkDays: 62,
  AdjustHijriDays: 0,
  TimeZone: { Id: 10, Description: '(UTC-05:00) Eastern Time (US and Canada)' },
};

const CURRENT_USER = user(11, 'Mock Developer', 'dev@mock.local', true);


// ---- classic publishing web -----------------------------------------------
// Served for a web base ending in /sites/classic, so the classic Pages path
// can be exercised without disturbing the modern web's fixtures (and their
// row counts). A publishing library has NO PromotedState and NO
// CanvasContent1: the body lives in PublishingPageContent, and the rest of
// the content lives in Content Editor / Script Editor web parts.
const CLASSIC_LISTS = [
  list('Documents', '7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f01', 101, 1, 8, false, '/sites/classic/Documents'),
  list('Pages', '7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f02', 850, 1, 4, false, '/sites/classic/Pages'),
];

const CLASSIC_PAGE_ITEMS = [
  {
    Id: 1,
    Title: 'Benefits overview',
    FileLeafRef: 'Benefits.aspx',
    FileRef: '/sites/classic/Pages/Benefits.aspx',
    FileDirRef: '/sites/classic/Pages',
    UniqueId: 'ef000000-0000-4000-8000-000000000001',
    Created: '2019-04-02T10:00:00Z',
    Modified: '2026-06-11T14:20:00Z',
    Author: { Title: 'Pat Example' },
    Editor: { Title: 'Pat Example' },
    PublishingPageContent: '<h2>Benefits</h2><p>Open enrollment runs through March.</p>',
    FieldValuesAsText: { Editor: 'Pat Example' },
  },
  {
    // Body empty on purpose: everything readable is in the web parts.
    Id: 2,
    Title: 'Rates',
    FileLeafRef: 'Rates.aspx',
    FileRef: '/sites/classic/Pages/Rates.aspx',
    FileDirRef: '/sites/classic/Pages',
    UniqueId: 'ef000000-0000-4000-8000-000000000002',
    Created: '2018-09-14T08:30:00Z',
    Modified: '2026-05-02T09:05:00Z',
    Author: { Title: 'Mock Developer' },
    Editor: { Title: 'Mock Developer' },
    PublishingPageContent: '',
    FieldValuesAsText: { Editor: 'Mock Developer' },
  },
  {
    Id: 3,
    Title: 'Empty',
    FileLeafRef: 'Empty.aspx',
    FileRef: '/sites/classic/Pages/Empty.aspx',
    FileDirRef: '/sites/classic/Pages',
    UniqueId: 'ef000000-0000-4000-8000-000000000003',
    Created: '2020-01-05T11:00:00Z',
    Modified: '2026-01-05T11:00:00Z',
    Author: { Title: 'Mock Developer' },
    Editor: { Title: 'Mock Developer' },
    PublishingPageContent: null,
    FieldValuesAsText: { Editor: 'Mock Developer' },
  },
  {
    // Body with an EMBEDDED web part: rich bodies place web parts inside the
    // field as .ms-rte-wpbox markers, and reading order must interleave —
    // intro → web part → conclusion, never body-then-web-parts.
    Id: 4,
    Title: 'Newsletter',
    FileLeafRef: 'Newsletter.aspx',
    FileRef: '/sites/classic/Pages/Newsletter.aspx',
    FileDirRef: '/sites/classic/Pages',
    UniqueId: 'ef000000-0000-4000-8000-000000000004',
    Created: '2021-03-09T09:00:00Z',
    Modified: '2026-07-30T16:45:00Z',
    Author: { Title: 'Pat Example' },
    Editor: { Title: 'Pat Example' },
    PublishingPageContent: '<p>Intro paragraph.</p>'
      + '<div class="ms-rtestate-read ms-rte-wpbox">'
      + '<div id="div_c3000000-0000-4000-8000-000000000006"></div></div>'
      + '<p>Closing paragraph.</p>',
    FieldValuesAsText: { Editor: 'Pat Example' },
  },
];

// getlimitedwebpartmanager(scope=1)/webparts?$expand=WebPart/Properties,
// keyed by lower-cased FileRef. Zone indexes are deliberately out of array
// order so the document-order sort is exercised.
const CLASSIC_WEBPARTS = {
  '/sites/classic/pages/benefits.aspx': [
    {
      Id: 'c3000000-0000-4000-8000-000000000001',
      WebPart: {
        Title: 'Contact details',
        ZoneIndex: 2,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '<p>Call the benefits desk on x4120.</p>', ContentLink: '' },
      },
    },
    {
      Id: 'c3000000-0000-4000-8000-000000000002',
      WebPart: {
        Title: 'Eligibility',
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '<![CDATA[<p>All staff after 90 days.</p>]]>', ContentLink: '' },
      },
    },
  ],
  '/sites/classic/pages/rates.aspx': [
    {
      Id: 'c3000000-0000-4000-8000-000000000003',
      WebPart: {
        Title: 'Rate table',
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '', ContentLink: '/sites/classic/Style Library/rates.html' },
      },
    },
    {
      Id: 'c3000000-0000-4000-8000-000000000004',
      WebPart: {
        Title: 'Rate calculator',
        ZoneIndex: 2,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '<script>calcRates();</script><div id="calc">Rates</div>' },
      },
    },
    {
      // No Content and no ContentLink — an inventory row, not a reading part.
      Id: 'c3000000-0000-4000-8000-000000000005',
      WebPart: { Title: 'List view', ZoneIndex: 3, Hidden: true, IsClosed: false, Properties: { ListName: 'Rates' } },
    },
  ],
  '/sites/classic/pages/newsletter.aspx': [
    {
      // Matched by the wpbox marker in the Newsletter body — must land
      // BETWEEN the intro and closing paragraphs, not after the body.
      Id: 'c3000000-0000-4000-8000-000000000006',
      WebPart: {
        Title: 'Signup form',
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '<p>Subscribe at the front desk.</p>', ContentLink: '' },
      },
    },
  ],
};

const CLASSIC_ITEMS = {
  '7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f02': CLASSIC_PAGE_ITEMS,
};

// ---- web with BOTH pages libraries ----------------------------------------
// Served for a web base ending in /sites/both. This is the common upgraded
// shape, not an exotic one: a classic publishing site that has ever had a
// modern page added carries an 850 "Pages" library holding the real content
// AND a 119 "Site Pages" one that is often nearly empty. The ranking opens
// the 119 by default, so without a picker the view looks empty and the real
// library is unreachable — which is what the picker exists to fix. Its own
// web again, so the other fixtures' row counts are untouched.
const BOTH_SITE_PAGES_ID = '9c2d4e6f-1111-4222-8333-44445555a001';
const BOTH_PAGES_ID = '9c2d4e6f-1111-4222-8333-44445555a002';

const BOTH_LISTS = [
  list('Site Pages', BOTH_SITE_PAGES_ID, 119, 1, 1, false, '/sites/both/SitePages'),
  list('Pages', BOTH_PAGES_ID, 850, 1, 2, false, '/sites/both/Pages'),
];

// Real field sets, so the probed query plan has something to distinguish:
// the 119 library has PromotedState + CanvasContent1 (the modern marker
// pair), the 850 one has neither. Without these the probe returns an empty
// list for both and every library looks classic.
const BOTH_FIELDS = {
  [BOTH_SITE_PAGES_ID]: [
    field('Title', 'Title', 'Text', 2, { Required: true }),
    field('Promoted state', 'PromotedState', 'Number', 9, { ReadOnlyField: true }),
    field('Canvas content', 'CanvasContent1', 'Note', 3),
    field('Editor', 'Editor', 'User', 20, { ReadOnlyField: true }),
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true }),
  ],
  [BOTH_PAGES_ID]: [
    field('Title', 'Title', 'Text', 2, { Required: true }),
    field('Page content', 'PublishingPageContent', 'Note', 3),
    field('Editor', 'Editor', 'User', 20, { ReadOnlyField: true }),
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true }),
  ],
};

const BOTH_ITEMS = {
  [BOTH_SITE_PAGES_ID]: [
    {
      Id: 1,
      Title: 'Team news',
      FileLeafRef: 'TeamNews.aspx',
      FileRef: '/sites/both/SitePages/TeamNews.aspx',
      FileDirRef: '/sites/both/SitePages',
      PromotedState: 0,
      Modified: '2026-07-18T10:00:00Z',
      Editor: { Title: 'Mock Developer' },
      CanvasContent1: JSON.stringify([
        { controlType: 4, id: 'c1', innerHTML: '<p>The one modern page.</p>' },
      ]),
    },
  ],
  [BOTH_PAGES_ID]: [
    {
      Id: 1,
      Title: 'Policies',
      FileLeafRef: 'Policies.aspx',
      FileRef: '/sites/both/Pages/Policies.aspx',
      FileDirRef: '/sites/both/Pages',
      Modified: '2026-07-04T10:00:00Z',
      Editor: { Title: 'Pat Example' },
      PublishingPageContent: '<h2>Policies</h2><p>Where the real content is.</p>',
    },
    {
      Id: 2,
      Title: 'Handbook',
      FileLeafRef: 'Handbook.aspx',
      FileRef: '/sites/both/Pages/Handbook.aspx',
      FileDirRef: '/sites/both/Pages',
      Modified: '2026-07-05T10:00:00Z',
      Editor: { Title: 'Pat Example' },
      PublishingPageContent: '<p>Staff handbook.</p>',
    },
  ],
};

// ---- resolver -------------------------------------------------------------

const listIdOf = (url) => /lists\(guid'([0-9a-f-]+)'\)/i.exec(url)?.[1]?.toLowerCase();
const groupIdOf = (url) => /sitegroups\((\d+)\)/i.exec(url)?.[1];

// Returns response JSON for a mocked /_api URL, or null when unhandled.
// Collections use the nometadata { value: [...] } envelope.
export function mockResolver(rawUrl) {
  const url = String(rawUrl);
  const path = url.slice(url.indexOf('/_api/') + 6).toLowerCase();
  // Which mock web is being asked for. Everything outside /sites/classic is
  // the modern web, so existing fixtures and their row counts are untouched.
  const webBase = url.slice(0, url.indexOf('/_api/')).replace(/[/]+$/, '');
  const classic = /[/]sites[/]classic$/i.test(webBase);
  const both = /[/]sites[/]both$/i.test(webBase);
  const schema = /[/]sites[/]schema$/i.test(webBase);
  const target = /[/]sites[/]target$/i.test(webBase);
  const staticLists = both ? BOTH_LISTS : classic ? CLASSIC_LISTS : schema ? SCHEMA_LISTS : target ? TARGET_LISTS : LISTS;
  const itemsByList = both ? BOTH_ITEMS : classic ? CLASSIC_ITEMS : schema ? ITEMS_BY_SCHEMA_LIST : ITEMS;
  // A list the stateful mock writer created under this web (see mockWriter
  // above) is visible to reads the moment it's written — a probe right after
  // a mock `web/lists` POST must find it, not 404 against the static set.
  const dynamicLists = [...writerState.lists.entries()]
    .filter(([key]) => key.startsWith(`${webBase}::`))
    .map(([, entry]) => entry);
  const lists = [...staticLists, ...dynamicLists];

  // Classic pages carry their content in web parts, not item fields.
  const wpFile = /getfilebyserverrelativepath[(]decodedurl='([^']*)'[)][/]getlimitedwebpartmanager/
    .exec(path)?.[1];
  if (wpFile !== undefined && path.includes('/webparts')) {
    let decoded = wpFile;
    try { decoded = decodeURIComponent(wpFile); } catch { /* keep raw */ }
    return { value: CLASSIC_WEBPARTS[decoded] || [] };
  }

  if (/^web\/lists\(guid'/.test(path)) {
    const id = listIdOf(path);
    const found = lists.find((l) => l.Id.toLowerCase() === id);
    if (!found) return null;
    const itemId = /\/items\((\d+)\)/.exec(path)?.[1];
    if (itemId) {
      const single = (itemsByList[found.Id] || []).find((i) => i.Id === Number(itemId));
      return single ?? null;
    }
    // The mock ignores $filter/$select on items — live-stub tests assert the
    // real query URLs instead. A single-field $orderby IS honored: the Items
    // tab caps row counts, and live SharePoint orders before the cap applies.
    if (path.includes('/items')) {
      const rows = [...(itemsByList[found.Id] || [])];
      const order = /\$orderby=([a-z0-9_]+)(?:(?:%20| +)(asc|desc))?/.exec(path);
      if (order) {
        // The URL was lowercased for routing — recover the item key case.
        const key = Object.keys(rows[0] || {})
          .find((k) => k.toLowerCase() === order[1]) || order[1];
        const dir = order[2] === 'desc' ? -1 : 1;
        rows.sort((a, b) => (a[key] > b[key] ? dir : a[key] < b[key] ? -dir : 0));
      }
      return { value: rows };
    }
    if (path.includes('/fields')) {
      const dyn = writerState.fields.get(found.Id);
      if (dyn) return { value: [...dyn.values()] };
      const perWeb = both ? BOTH_FIELDS : schema ? SCHEMA_FIELDS : target ? TARGET_FIELDS : FIELDS;
      return { value: perWeb[found.Id] || DEFAULT_FIELDS };
    }
    if (/\/views\(guid'/.test(path) && path.includes('/viewfields')) {
      return { Items: ['LinkTitle', 'ProjectStatus', 'DueDate'] };
    }
    if (path.includes('/views')) {
      const dyn = writerState.views.get(found.Id);
      if (dyn) {
        return {
          value: [...dyn.values()].map((v) => ({
            Id: v.Id, Title: v.Title, DefaultView: !!v.defaultView, PersonalView: false, Hidden: false,
            RowLimit: 30, Paged: true, ViewQuery: '', ViewFields: { Items: v.fields },
          })),
        };
      }
      return { value: VIEWS_BY_LIST[found.Id] || VIEWS };
    }
    if (path.includes('/contenttypes')) {
      const dyn = writerState.contentTypes.get(found.Id);
      if (dyn) {
        return {
          value: [...dyn].map((id) => ({
            Id: { StringValue: id }, Name: id, Group: '', Hidden: false, ReadOnly: false,
            Sealed: false, Description: '',
          })),
        };
      }
      return { value: CONTENT_TYPES_BY_LIST[found.Id] || CONTENT_TYPES };
    }
    if (path.includes('/roleassignments')) return { value: ROLE_ASSIGNMENTS };
    return found;
  }
  if (path.startsWith('web/lists')) {
    if (path.includes('hasuniqueroleassignments')) {
      return {
        value: lists.map((l, i) => ({
          Id: l.Id, Title: l.Title, Hidden: l.Hidden, BaseTemplate: l.BaseTemplate,
          HasUniqueRoleAssignments: i === 2,
        })),
      };
    }
    return { value: lists };
  }

  if (/^web\/sitegroups\(\d+\)\/users/.test(path)) {
    const users = GROUP_USERS[groupIdOf(path)];
    return users ? { value: users } : { value: [] };
  }
  if (path.startsWith('web/sitegroups')) return { value: GROUPS };
  if (path.startsWith('web/roledefinitions')) return { value: ROLE_DEFINITIONS };
  if (path.startsWith('web/roleassignments')) return { value: ROLE_ASSIGNMENTS };

  // list-data-capture.js: bulk site-user read + the per-id fallback
  // (resolveUser's getbyid escape hatch) it uses when a referenced id isn't
  // in the bulk set.
  const siteUserId = /^web\/siteusers\/getbyid\((\d+)\)/.exec(path)?.[1];
  if (siteUserId !== undefined) {
    const found = SCHEMA_SITE_USERS.find((u) => u.Id === Number(siteUserId));
    return found || null;
  }
  if (path.startsWith('web/siteusers')) return { value: SCHEMA_SITE_USERS };

  // Files browser: ResourcePath folder/file endpoints. Paths arrive
  // percent-encoded inside the decodedUrl literal; decode before matching.
  const folderPathOf = /getfolderbyserverrelativepath\(decodedurl='([^']*)'\)/.exec(path)?.[1];
  if (folderPathOf !== undefined) {
    let decoded = folderPathOf;
    try { decoded = decodeURIComponent(folderPathOf); } catch { /* keep raw */ }
    const listing = MOCK_TREE[decoded];
    if (path.includes('/folders')) return { value: listing?.folders || [] };
    if (path.includes('/files')) return { value: listing?.files || [] };
    if (path.includes('parentlist')) {
      return { ListItemAllFields: { ParentList: { Id: DOC_LIB_ID } } };
    }
    return { Name: decoded.split('/').pop() || '', ServerRelativeUrl: decoded };
  }
  const filePathOf = /getfilebyserverrelativepath\(decodedurl='([^']*)'\)/.exec(path)?.[1];
  if (filePathOf !== undefined) {
    let decoded = filePathOf;
    try { decoded = decodeURIComponent(filePathOf); } catch { /* keep raw */ }
    if (path.includes('/listitemallfields')) {
      return FILE_ITEMS[decoded] || { Id: 0, Title: '' };
    }
    return null;
  }

  if (path.startsWith('web/allproperties')) return ALL_PROPERTIES;
  if (path.startsWith('web/regionalsettings')) return REGIONAL_SETTINGS;
  if (path.startsWith('web/currentuser')) return CURRENT_USER;
  if (path.startsWith('web/webs')) return { value: SUBWEBS };
  if (path.startsWith('web/features')) return { value: FEATURES.web };
  if (path.startsWith('site/features')) return { value: FEATURES.site };
  if (path.startsWith('site')) return SITE;
  // Ahead of the generic 'web' echo below (which would otherwise swallow
  // this and hand back a bare web entity) — the Schema tab's Copy-to dialog
  // (slice 2) probes this only on a content-types-enabled source.
  if (path.startsWith('web/availablecontenttypes')) {
    return { value: target ? TARGET_AVAILABLE_CTS : [] };
  }
  if (path.startsWith('web')) {
    // Echo the requested web base back so mock site-switching behaves like
    // the real thing (connectWeb canonicalizes on the returned Url).
    const base = url.slice(0, url.indexOf('/_api/')).replace(/\/+$/, '');
    let rel = '/';
    try { rel = decodeURIComponent(new URL(base).pathname) || '/'; } catch { /* keep '/' */ }
    return {
      ...WEB,
      Url: base || WEB.Url,
      ServerRelativeUrl: rel,
      Title: rel === '/' ? WEB.Title : `Mock Web (${rel})`,
    };
  }

  return null;
}
