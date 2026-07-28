// Mock /_api fixtures for off-SharePoint development and tests.
//
// Shapes mirror real nometadata responses closely enough that the sp-shapes
// inspector fingerprints them (SP.List needs BaseTemplate+EntityTypeName,
// SP.Field needs InternalName+TypeAsString, principals need
// LoginName+PrincipalType). Keep additions realistic — the mock is the local
// contract for every view.

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
    field('ID', 'ID', 'Counter', 5, { ReadOnlyField: true, Hidden: false }),
    field('Content Type', 'ContentType', 'Computed', 12, { Hidden: true, ReadOnlyField: true }),
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
  dcspad_x005f_deployfolder: '/SiteAssets/Code/dcspad-live',
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

// ---- resolver -------------------------------------------------------------

const listIdOf = (url) => /lists\(guid'([0-9a-f-]+)'\)/i.exec(url)?.[1]?.toLowerCase();
const groupIdOf = (url) => /sitegroups\((\d+)\)/i.exec(url)?.[1];

// Returns response JSON for a mocked /_api URL, or null when unhandled.
// Collections use the nometadata { value: [...] } envelope.
export function mockResolver(rawUrl) {
  const url = String(rawUrl);
  const path = url.slice(url.indexOf('/_api/') + 6).toLowerCase();

  if (/^web\/lists\(guid'/.test(path)) {
    const id = listIdOf(path);
    const found = LISTS.find((l) => l.Id.toLowerCase() === id);
    if (!found) return null;
    if (path.includes('/fields')) return { value: FIELDS[found.Id] || DEFAULT_FIELDS };
    if (/\/views\(guid'/.test(path) && path.includes('/viewfields')) {
      return { Items: ['LinkTitle', 'ProjectStatus', 'DueDate'] };
    }
    if (path.includes('/views')) return { value: VIEWS };
    if (path.includes('/contenttypes')) return { value: CONTENT_TYPES };
    if (path.includes('/roleassignments')) return { value: ROLE_ASSIGNMENTS };
    return found;
  }
  if (path.startsWith('web/lists')) {
    if (path.includes('hasuniqueroleassignments')) {
      return {
        value: LISTS.map((l, i) => ({
          Id: l.Id, Title: l.Title, Hidden: l.Hidden, BaseTemplate: l.BaseTemplate,
          HasUniqueRoleAssignments: i === 2,
        })),
      };
    }
    return { value: LISTS };
  }

  if (/^web\/sitegroups\(\d+\)\/users/.test(path)) {
    const users = GROUP_USERS[groupIdOf(path)];
    return users ? { value: users } : { value: [] };
  }
  if (path.startsWith('web/sitegroups')) return { value: GROUPS };
  if (path.startsWith('web/roledefinitions')) return { value: ROLE_DEFINITIONS };
  if (path.startsWith('web/roleassignments')) return { value: ROLE_ASSIGNMENTS };

  if (path.startsWith('web/allproperties')) return ALL_PROPERTIES;
  if (path.startsWith('web/regionalsettings')) return REGIONAL_SETTINGS;
  if (path.startsWith('web/currentuser')) return CURRENT_USER;
  if (path.startsWith('web/webs')) return { value: SUBWEBS };
  if (path.startsWith('web/features')) return { value: FEATURES.web };
  if (path.startsWith('site/features')) return { value: FEATURES.site };
  if (path.startsWith('site')) return SITE;
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
