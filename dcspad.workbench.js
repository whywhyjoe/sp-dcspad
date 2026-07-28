// ../src/bridge/sp-context.js
var MODERN_SITE_PAGES_FEATURE_ID = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";
var SERIALIZABLE_FIELDS = [
  "webAbsoluteUrl",
  "webServerRelativeUrl",
  "siteAbsoluteUrl",
  "siteServerRelativeUrl",
  "webTitle",
  "userId",
  "userLoginName",
  "userDisplayName",
  "currentLanguage",
  "currentCultureName",
  "layoutsUrl",
  "webUIVersion",
  "siteClientTag",
  "formDigestValue",
  "formDigestTimeoutSeconds"
];
var cached = null;
var isRecord = (value) => value !== null && typeof value === "object" && !Array.isArray(value);
function safeSameOriginUrl(value) {
  if (typeof value !== "string" || !value.trim()) return "";
  try {
    const url = new URL(value.trim(), location.href);
    if (url.origin !== location.origin) return "";
    return url.href.replace(/\/+$/, "");
  } catch {
    return "";
  }
}
function serverRelativeUrl(absoluteUrl) {
  try {
    return decodeURIComponent(new URL(absoluteUrl).pathname).replace(/\/+$/, "") || "/";
  } catch {
    return "/";
  }
}
function candidateWindows() {
  const candidates = [window];
  for (const key2 of ["parent", "top"]) {
    try {
      const candidate = window[key2];
      if (candidate && !candidates.includes(candidate)) {
        void candidate.location.href;
        candidates.push(candidate);
      }
    } catch {
    }
  }
  return candidates;
}
function hostContext(candidate) {
  try {
    const host = candidate.__DCSPAD_SP_CONTEXT__;
    if (!isRecord(host)) return null;
    const pageContext = isRecord(host.pageContext) ? host.pageContext : host;
    const webAbsoluteUrl = safeSameOriginUrl(
      host.webAbsoluteUrl || pageContext.webAbsoluteUrl
    );
    return webAbsoluteUrl ? { raw: host, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}
function globalContext(candidate) {
  try {
    const pageContext = candidate._spPageContextInfo;
    const webAbsoluteUrl = safeSameOriginUrl(pageContext?.webAbsoluteUrl);
    return webAbsoluteUrl ? { raw: pageContext, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}
function modernLegacyContext(candidate) {
  try {
    const pageContext = candidate.spModuleLoader?._bundledComponents?.[MODERN_SITE_PAGES_FEATURE_ID]?.PageManager?._instance?.pageContext?.legacyPageContext;
    const webAbsoluteUrl = safeSameOriginUrl(pageContext?.webAbsoluteUrl);
    return webAbsoluteUrl ? { raw: pageContext, pageContext, webAbsoluteUrl } : null;
  } catch {
    return null;
  }
}
function findContext() {
  const windows = candidateWindows();
  for (const [source, reader] of [
    ["host", hostContext],
    ["global", globalContext],
    ["modern-legacy", modernLegacyContext]
  ]) {
    for (const candidate of windows) {
      const found = reader(candidate);
      if (found) return { ...found, source, ownerWindow: candidate };
    }
  }
  return null;
}
function copyPageContext(found) {
  let pageContext;
  try {
    pageContext = JSON.parse(JSON.stringify(found.pageContext));
  } catch {
    pageContext = {};
    for (const key2 of SERIALIZABLE_FIELDS) {
      if (found.pageContext[key2] !== void 0) {
        pageContext[key2] = found.pageContext[key2];
      }
    }
  }
  for (const key2 of SERIALIZABLE_FIELDS) {
    if (pageContext[key2] === void 0 && found.raw[key2] !== void 0) {
      pageContext[key2] = found.raw[key2];
    }
  }
  pageContext.webAbsoluteUrl = found.webAbsoluteUrl;
  pageContext.webServerRelativeUrl ||= serverRelativeUrl(found.webAbsoluteUrl);
  pageContext.siteAbsoluteUrl ||= found.webAbsoluteUrl;
  pageContext.siteServerRelativeUrl ||= serverRelativeUrl(pageContext.siteAbsoluteUrl);
  try {
    const digest = found.ownerWindow.document.getElementById("__REQUESTDIGEST")?.value;
    if (digest) pageContext.formDigestValue = digest;
  } catch {
  }
  return pageContext;
}
function getSpContext({ refresh = false } = {}) {
  if (cached && !refresh) return cached;
  const found = findContext();
  if (found) {
    const pageContext = copyPageContext(found);
    cached = {
      live: true,
      source: found.source,
      capturedAt: Date.now(),
      pageContext,
      baseHref: `${pageContext.webAbsoluteUrl.replace(/\/$/, "")}/`,
      label: pageContext.webAbsoluteUrl,
      user: pageContext.userDisplayName || pageContext.userLoginName || ""
    };
    return cached;
  }
  cached = {
    live: false,
    source: "mock",
    capturedAt: Date.now(),
    pageContext: {
      isDcsPadMock: true,
      webAbsoluteUrl: location.origin,
      webServerRelativeUrl: "/",
      siteAbsoluteUrl: location.origin,
      siteServerRelativeUrl: "/",
      webTitle: "DCSPad Mock Web",
      userId: 1,
      userLoginName: "i:0#.f|membership|dev@mock.local",
      userDisplayName: "Mock Developer",
      currentLanguage: 1033,
      currentCultureName: "en-US",
      layoutsUrl: "_layouts/15",
      formDigestValue: "MOCK-DIGEST-0x0000",
      formDigestTimeoutSeconds: 1800
    },
    baseHref: null,
    label: "mock (not in SharePoint)",
    user: "Mock Developer"
  };
  return cached;
}

// ../src/sp-odata.js
var ACCEPT_JSON = "application/json;odata=nometadata";
var SpFileError = class extends Error {
  constructor(message, { code = "sharepoint", status = 0, cause } = {}) {
    super(message, { cause });
    this.name = "SpFileError";
    this.code = code;
    this.status = status;
  }
};
async function responseMessage(response) {
  try {
    const body = await response.clone().json();
    return body?.error?.message?.value || body?.error?.message || body?.["odata.error"]?.message?.value || "";
  } catch {
    try {
      return (await response.text()).trim();
    } catch {
      return "";
    }
  }
}
async function requireOk(response, fallback, code) {
  if (response.ok) return response;
  const detail = await responseMessage(response);
  let message = detail || `${fallback} (HTTP ${response.status})`;
  let normalizedCode = code;
  if (response.status === 401 || response.status === 403) {
    message = detail || "SharePoint denied this request. Check library permissions and try again.";
    normalizedCode = "permission";
  } else if (response.status === 404) {
    message = detail || "The SharePoint file or folder was not found.";
    normalizedCode = "not-found";
  } else if (response.status === 409) {
    message = detail || "A SharePoint file with that name already exists.";
    normalizedCode = "conflict";
  }
  throw new SpFileError(message, {
    code: normalizedCode,
    status: response.status
  });
}

// ../src/workbench/sp-rest.js
var PAGE_CAP = 5e3;
var MAX_CONCURRENT = 3;
var RETRY_STATUSES = /* @__PURE__ */ new Set([429, 503]);
function buildQuery({ select, expand, filter, orderby, top } = {}) {
  const parts = [];
  const join = (v) => Array.isArray(v) ? v.join(",") : String(v);
  if (select) parts.push(`$select=${join(select)}`);
  if (expand) parts.push(`$expand=${join(expand)}`);
  if (filter) parts.push(`$filter=${encodeURIComponent(String(filter))}`);
  if (orderby) parts.push(`$orderby=${join(orderby)}`);
  if (top) parts.push(`$top=${top}`);
  return parts.length ? `?${parts.join("&")}` : "";
}
function collectionOf(data) {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.value)) return data.value;
  if (Array.isArray(data?.d?.results)) return data.d.results;
  if (Array.isArray(data?.results)) return data.results;
  return null;
}
function nextLinkOf(data) {
  return data?.["odata.nextLink"] || data?.["@odata.nextLink"] || data?.d?.__next || "";
}
function entityOf(data) {
  return data?.d ?? data;
}
function createSpRestClient({
  getContext = getSpContext,
  fetchImpl = (...args) => fetch(...args),
  mockResolver: mockResolver2 = null
} = {}) {
  let inFlight = 0;
  const waiters = [];
  async function withSlot(work) {
    if (inFlight >= MAX_CONCURRENT) {
      await new Promise((resolve) => waiters.push(resolve));
    }
    inFlight++;
    try {
      return await work();
    } finally {
      inFlight--;
      waiters.shift()?.();
    }
  }
  function context() {
    const ctx2 = getContext();
    if (!ctx2?.live && !mockResolver2) {
      throw new SpFileError(
        "The SP Workbench needs an SP: Live context (or an injected mock).",
        { code: "not-live" }
      );
    }
    return ctx2;
  }
  function webUrl() {
    return context().pageContext.webAbsoluteUrl.replace(/\/+$/, "");
  }
  function apiUrl(path, opts) {
    const clean = String(path).replace(/^\/+/, "");
    return `${webUrl()}/_api/${clean}${buildQuery(opts)}`;
  }
  async function rawGet(url) {
    if (mockResolver2 && !getContext().live) {
      const data = mockResolver2(url);
      if (data == null) {
        throw new SpFileError(`No mock data for ${url}`, { code: "not-found", status: 404 });
      }
      return structuredClone(data);
    }
    return withSlot(async () => {
      const attempt = async () => {
        let response2;
        try {
          response2 = await fetchImpl(url, {
            credentials: "same-origin",
            headers: { Accept: ACCEPT_JSON }
          });
        } catch (cause) {
          throw new SpFileError(
            `Could not reach SharePoint (${cause.message || cause}).`,
            { code: "network", cause }
          );
        }
        return response2;
      };
      let response = await attempt();
      if (RETRY_STATUSES.has(response.status)) {
        const after = Number(response.headers.get("Retry-After")) || 2;
        await new Promise((r) => setTimeout(r, Math.min(after, 30) * 1e3));
        response = await attempt();
      }
      await requireOk(response, "SharePoint request failed", "get");
      return response.json();
    });
  }
  async function get(path, opts) {
    return entityOf(await rawGet(apiUrl(path, opts)));
  }
  async function getAll(path, opts) {
    let url = apiUrl(path, opts);
    const items = [];
    let partial = false;
    while (url) {
      const data = await rawGet(url);
      const page = collectionOf(data);
      if (!page) {
        items.push(entityOf(data));
        break;
      }
      items.push(...page);
      const next = nextLinkOf(data);
      if (!next) break;
      if (items.length >= PAGE_CAP) {
        partial = true;
        break;
      }
      url = next;
    }
    return { items, partial };
  }
  return { context, webUrl, apiUrl, get, getAll };
}

// ../src/workbench/mock-data.js
var WEB_URL = location.origin;
var LISTS = [
  list("Documents", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01", 101, 1, 42, false, "/Shared Documents"),
  list("Site Pages", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02", 119, 1, 17, false, "/SitePages"),
  list("Projects", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03", 100, 0, 128, false, "/Lists/Projects"),
  list("Tasks", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e04", 171, 0, 260, false, "/Lists/Tasks"),
  list("Site Assets", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e05", 101, 1, 96, false, "/SiteAssets"),
  list("User Information List", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e06", 112, 0, 57, true, "/_catalogs/users"),
  list("Master Page Gallery", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e07", 116, 1, 12, true, "/_catalogs/masterpage"),
  list("DevPadData", "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e08", 100, 0, 3, true, "/Lists/DevPadData")
];
function list(title, id, template, baseType, itemCount, hidden, url) {
  return {
    Id: id,
    Title: title,
    BaseTemplate: template,
    BaseType: baseType,
    ItemCount: itemCount,
    Hidden: hidden,
    Created: "2025-11-02T15:04:00Z",
    LastItemModifiedDate: "2026-07-20T09:30:00Z",
    EntityTypeName: title.replaceAll(" ", "_x0020_"),
    Description: hidden ? "" : `${title} for the mock web.`,
    DefaultViewUrl: `${url}/Forms/AllItems.aspx`,
    RootFolder: { ServerRelativeUrl: url }
  };
}
var fieldSeq = 0;
var FIELDS = {
  "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03": [
    field("Title", "Title", "Text", 2, { Required: true }),
    field("Project Status", "ProjectStatus", "Choice", 6, {
      Choices: ["Planned", "Active", "Blocked", "Done"],
      DefaultValue: "Planned"
    }),
    field("Due Date", "DueDate", "DateTime", 4),
    field("Owner", "ProjectOwner", "User", 20),
    field("Budget", "Budget", "Currency", 10),
    field("ID", "ID", "Counter", 5, { ReadOnlyField: true, Hidden: false }),
    field("Content Type", "ContentType", "Computed", 12, { Hidden: true, ReadOnlyField: true })
  ]
};
var DEFAULT_FIELDS = [
  field("Title", "Title", "Text", 2, { Required: true }),
  field("ID", "ID", "Counter", 5, { ReadOnlyField: true }),
  field("Modified", "Modified", "DateTime", 4, { ReadOnlyField: true }),
  field("Created", "Created", "DateTime", 4, { ReadOnlyField: true })
];
function field(title, internal, type, kind, extra = {}) {
  fieldSeq++;
  return {
    Id: `af0e2c1d-2222-4444-8888-${String(fieldSeq).padStart(12, "0")}`,
    Title: title,
    InternalName: internal,
    TypeAsString: type,
    FieldTypeKind: kind,
    Required: false,
    Hidden: false,
    ReadOnlyField: false,
    Group: "Custom Columns",
    DefaultValue: null,
    Description: "",
    EnforceUniqueValues: false,
    Indexed: false,
    SchemaXml: `<Field Name="${internal}" Type="${type}" DisplayName="${title}"/>`,
    ...extra
  };
}
var VIEWS = [
  { Id: "bb0e2c1d-3333-4444-8888-000000000001", Title: "All Items", DefaultView: true, PersonalView: false, Hidden: false, ServerRelativeUrl: "/Lists/Projects/AllItems.aspx", RowLimit: 30, Paged: true, ViewQuery: '<OrderBy><FieldRef Name="ID"/></OrderBy>' },
  { Id: "bb0e2c1d-3333-4444-8888-000000000002", Title: "Active only", DefaultView: false, PersonalView: false, Hidden: false, ServerRelativeUrl: "/Lists/Projects/Active.aspx", RowLimit: 100, Paged: true, ViewQuery: '<Where><Eq><FieldRef Name="ProjectStatus"/><Value Type="Choice">Active</Value></Eq></Where>' }
];
var CONTENT_TYPES = [
  { Id: { StringValue: "0x0100A1B2C3D4E5F601" }, Name: "Item", Group: "List Content Types", Hidden: false, ReadOnly: false, Sealed: false, Description: "Create a new list item." },
  { Id: { StringValue: "0x0120001122334455" }, Name: "Folder", Group: "_Hidden", Hidden: true, ReadOnly: false, Sealed: true, Description: "" }
];
var GROUPS = [
  { Id: 3, Title: "Mock Site Owners", Description: "Full control of the mock site.", OwnerTitle: "System Account", PrincipalType: 8, LoginName: "Mock Site Owners", OnlyAllowMembersViewMembership: false },
  { Id: 5, Title: "Mock Site Members", Description: "Contribute to the mock site.", OwnerTitle: "Mock Site Owners", PrincipalType: 8, LoginName: "Mock Site Members", OnlyAllowMembersViewMembership: false },
  { Id: 7, Title: "Mock Site Visitors", Description: "Read-only visitors.", OwnerTitle: "Mock Site Owners", PrincipalType: 8, LoginName: "Mock Site Visitors", OnlyAllowMembersViewMembership: true }
];
var GROUP_USERS = {
  3: [user(11, "Mock Developer", "dev@mock.local", true)],
  5: [user(11, "Mock Developer", "dev@mock.local", true), user(14, "Pat Example", "pat@mock.local", false)],
  7: [user(19, "Ronnie Reader", "ronnie@mock.local", false)]
};
function user(id, title, email, admin) {
  return {
    Id: id,
    Title: title,
    LoginName: `i:0#.f|membership|${email}`,
    Email: email,
    IsSiteAdmin: admin,
    PrincipalType: 1
  };
}
var ROLE_DEFINITIONS = [
  roleDef(1073741829, "Full Control", "Has full control.", 5, "2147483647", "4294967295"),
  roleDef(1073741827, "Contribute", "Can view, add, update, and delete list items and documents.", 3, "432", "1011028719"),
  roleDef(1073741826, "Read", "Can view pages and list items and download documents.", 2, "176", "138612833"),
  roleDef(1073741825, "Limited Access", "Can view specific lists when given access.", 1, "176", "138612801")
];
function roleDef(id, name, description, kind, high, low) {
  return {
    Id: id,
    Name: name,
    Description: description,
    RoleTypeKind: kind,
    Hidden: kind === 1,
    BasePermissions: { High: high, Low: low }
  };
}
var ROLE_ASSIGNMENTS = [
  assignment(3, "Mock Site Owners", 8, ["Full Control"]),
  assignment(5, "Mock Site Members", 8, ["Contribute"]),
  assignment(7, "Mock Site Visitors", 8, ["Read"])
];
function assignment(principalId, title, principalType, roleNames2) {
  return {
    PrincipalId: principalId,
    Member: { Id: principalId, Title: title, LoginName: title, PrincipalType: principalType },
    RoleDefinitionBindings: roleNames2.map((name) => ({
      Id: ROLE_DEFINITIONS.find((r) => r.Name === name)?.Id || 0,
      Name: name
    }))
  };
}
var WEB = {
  Id: "c0ffee00-1111-2222-3333-444455556666",
  Title: "Mock Web",
  Description: "Local workbench mock web.",
  Url: WEB_URL,
  ServerRelativeUrl: "/",
  WebTemplate: "SITEPAGEPUBLISHING",
  Configuration: 0,
  Created: "2025-10-01T12:00:00Z",
  LastItemModifiedDate: "2026-07-25T08:00:00Z",
  Language: 1033,
  UIVersion: 15,
  QuickLaunchEnabled: true,
  MembersCanShare: true
};
var SITE = {
  Id: "deadbeef-7777-8888-9999-aaaabbbbcccc",
  Url: WEB_URL,
  ServerRelativeUrl: "/",
  ReadOnly: false,
  ShareByEmailEnabled: false
};
var FEATURES = {
  site: [
    { DefinitionId: "b50e3104-6812-424f-a011-cc90e6327318", DisplayName: "BasicWebParts" },
    { DefinitionId: "8c6a6980-c3d9-440e-944c-77f93bc65a7e", DisplayName: "" }
  ],
  web: [
    { DefinitionId: "00bfea71-4ea5-48d4-a4ad-7ea5c011abe5", DisplayName: "TeamCollab" },
    { DefinitionId: "f151bb39-7c3b-414f-bb36-6bf18872052f", DisplayName: "" }
  ]
};
var SUBWEBS = [
  { Id: "aaaa1111-0000-0000-0000-000000000001", Title: "Archive", ServerRelativeUrl: "/archive", WebTemplate: "STS", Created: "2025-12-01T00:00:00Z", Language: 1033 }
];
var ALL_PROPERTIES = {
  vti_x005f_defaultlanguage: "en-us",
  vti_x005f_extenderversion: "16.0.0.26000",
  taxonomyhiddenlist: "{5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e99}",
  dcspad_x005f_deployfolder: "/SiteAssets/Code/dcspad-live"
};
var REGIONAL_SETTINGS = {
  LocaleId: 1033,
  Time24: false,
  FirstDayOfWeek: 0,
  WorkDays: 62,
  AdjustHijriDays: 0,
  TimeZone: { Id: 10, Description: "(UTC-05:00) Eastern Time (US and Canada)" }
};
var CURRENT_USER = user(11, "Mock Developer", "dev@mock.local", true);
var listIdOf = (url) => /lists\(guid'([0-9a-f-]+)'\)/i.exec(url)?.[1]?.toLowerCase();
var groupIdOf = (url) => /sitegroups\((\d+)\)/i.exec(url)?.[1];
function mockResolver(rawUrl) {
  const url = String(rawUrl);
  const path = url.slice(url.indexOf("/_api/") + 6).toLowerCase();
  if (/^web\/lists\(guid'/.test(path)) {
    const id = listIdOf(path);
    const found = LISTS.find((l) => l.Id.toLowerCase() === id);
    if (!found) return null;
    if (path.includes("/fields")) return { value: FIELDS[found.Id] || DEFAULT_FIELDS };
    if (/\/views\(guid'/.test(path) && path.includes("/viewfields")) {
      return { Items: ["LinkTitle", "ProjectStatus", "DueDate"] };
    }
    if (path.includes("/views")) return { value: VIEWS };
    if (path.includes("/contenttypes")) return { value: CONTENT_TYPES };
    if (path.includes("/roleassignments")) return { value: ROLE_ASSIGNMENTS };
    return found;
  }
  if (path.startsWith("web/lists")) {
    if (path.includes("hasuniqueroleassignments")) {
      return {
        value: LISTS.map((l, i) => ({
          Id: l.Id,
          Title: l.Title,
          Hidden: l.Hidden,
          BaseTemplate: l.BaseTemplate,
          HasUniqueRoleAssignments: i === 2
        }))
      };
    }
    return { value: LISTS };
  }
  if (/^web\/sitegroups\(\d+\)\/users/.test(path)) {
    const users = GROUP_USERS[groupIdOf(path)];
    return users ? { value: users } : { value: [] };
  }
  if (path.startsWith("web/sitegroups")) return { value: GROUPS };
  if (path.startsWith("web/roledefinitions")) return { value: ROLE_DEFINITIONS };
  if (path.startsWith("web/roleassignments")) return { value: ROLE_ASSIGNMENTS };
  if (path.startsWith("web/allproperties")) return ALL_PROPERTIES;
  if (path.startsWith("web/regionalsettings")) return REGIONAL_SETTINGS;
  if (path.startsWith("web/currentuser")) return CURRENT_USER;
  if (path.startsWith("web/webs")) return { value: SUBWEBS };
  if (path.startsWith("web/features")) return { value: FEATURES.web };
  if (path.startsWith("site/features")) return { value: FEATURES.site };
  if (path.startsWith("site")) return SITE;
  if (path.startsWith("web")) return WEB;
  return null;
}

// ../src/workbench/shell.js
var ROUTE_KEY = "dcspad.workbench.route";
var el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function createShell({ mount, deps, views }) {
  const instances = /* @__PURE__ */ new Map();
  let currentRoute = null;
  const rail = el("nav", "wb-rail");
  rail.setAttribute("aria-label", "Workbench sections");
  const host = el("main", "wb-host");
  const buttons = /* @__PURE__ */ new Map();
  for (const view of views) {
    const btn = el("button", "wb-rail-btn");
    btn.type = "button";
    btn.dataset.view = view.id;
    if (view.glyph) {
      const glyph = el("span", "wb-rail-glyph");
      glyph.innerHTML = view.glyph;
      btn.append(glyph);
    }
    btn.append(el("span", "wb-rail-label", view.label));
    btn.addEventListener("click", () => navigate({ view: view.id }));
    buttons.set(view.id, btn);
    rail.append(btn);
  }
  mount.append(rail, host);
  function instance(id) {
    if (instances.has(id)) return instances.get(id);
    const def = views.find((v) => v.id === id);
    if (!def) return null;
    const inst = def.create({ ...deps, navigate });
    instances.set(id, inst);
    return inst;
  }
  function navigate(route) {
    const def = views.find((v) => v.id === route?.view) ? route : { view: views[0].id };
    currentRoute = def;
    try {
      sessionStorage.setItem(ROUTE_KEY, JSON.stringify(def));
    } catch {
    }
    for (const [id, btn] of buttons) {
      btn.classList.toggle("active", id === def.view);
      btn.setAttribute("aria-current", id === def.view ? "page" : "false");
    }
    const inst = instance(def.view);
    host.textContent = "";
    host.append(inst.el);
    inst.load?.(def);
  }
  function restore() {
    let saved = null;
    try {
      saved = JSON.parse(sessionStorage.getItem(ROUTE_KEY) || "null");
    } catch {
    }
    navigate(saved || { view: views[0].id });
  }
  return { navigate, restore, getRoute: () => currentRoute };
}

// ../src/io.js?v=2
var MAX_IMPORT_BYTES = 5 * 1024 * 1024;
function downloadText(filename, text, type = "application/json") {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1e3);
}

// ../src/workbench/export.js
function copyText(text, flashEl) {
  const done = () => {
    if (!flashEl) return;
    flashEl.classList.add("copied");
    setTimeout(() => flashEl.classList.remove("copied"), 900);
  };
  if (navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(text).then(done, done);
  } else {
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
    } catch {
    }
    ta.remove();
    done();
  }
}
var cellValue = (row, col) => typeof col.value === "function" ? col.value(row) : row[col.key];
function cellText(row, col) {
  const v = cellValue(row, col);
  if (typeof col.format === "function") return String(col.format(v, row) ?? "");
  if (v === null || v === void 0) return "";
  if (typeof v === "boolean") return v ? "Yes" : "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}
function toCsv(rows, columns) {
  const quote = (s) => /[",\r\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
  const lines = [columns.map((c) => quote(String(c.label ?? c.key))).join(",")];
  for (const row of rows) {
    lines.push(columns.map((c) => quote(cellText(row, c))).join(","));
  }
  return `\uFEFF${lines.join("\r\n")}`;
}
function toJson(rows, columns) {
  const out = rows.map((row) => {
    const record = {};
    for (const c of columns) {
      const v = cellValue(row, c);
      record[c.key] = v === void 0 ? null : v;
    }
    return record;
  });
  return JSON.stringify(out, null, 2);
}
function toMarkdown(rows, columns) {
  const esc = (s) => s.replaceAll("|", "\\|").replaceAll("\r", "").replaceAll("\n", " ");
  const lines = [
    `| ${columns.map((c) => esc(String(c.label ?? c.key))).join(" | ")} |`,
    `| ${columns.map(() => "---").join(" | ")} |`
  ];
  for (const row of rows) {
    lines.push(`| ${columns.map((c) => esc(cellText(row, c))).join(" | ")} |`);
  }
  return lines.join("\n");
}
function downloadCsv(name, rows, columns) {
  downloadText(`${name}.csv`, toCsv(rows, columns), "text/csv;charset=utf-8");
}
function downloadJson(name, rows, columns) {
  downloadText(`${name}.json`, toJson(rows, columns), "application/json");
}

// ../src/workbench/grid.js
var el2 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var cellValue2 = (row, col) => typeof col.value === "function" ? col.value(row) : row[col.key];
function displayValue(row, col) {
  const v = cellValue2(row, col);
  if (typeof col.format === "function") return col.format(v, row);
  if (v === null || v === void 0) return "";
  if (typeof v === "boolean") return v ? "Yes" : "";
  return String(v);
}
function createGrid({
  columns,
  rowKey = "Id",
  onOpen = null,
  emptyText = "No rows.",
  filterPlaceholder = "Filter\u2026",
  exportName = ""
} = {}) {
  let rows = [];
  let visible = [];
  let sortKey = null;
  let sortDir = 1;
  let filterText = "";
  const root = el2("div", "wb-grid");
  const toolbar = el2("div", "wb-grid-toolbar");
  const count = el2("span", "wb-grid-count", "\u2014");
  const filter = el2("input", "wb-grid-filter");
  filter.type = "search";
  filter.placeholder = filterPlaceholder;
  filter.setAttribute("aria-label", "Filter rows");
  const actions = el2("span", "wb-grid-actions");
  toolbar.append(count, filter, actions);
  if (exportName) {
    const wrap = el2("span", "wb-menu-wrap");
    const btn = el2("button", "btn btn-xs", "Export \u25BE");
    btn.type = "button";
    btn.title = "Export the visible rows";
    const menu = el2("div", "wb-menu");
    menu.hidden = true;
    const items = [
      ["Download CSV", () => downloadCsv(exportName, visible, columns)],
      ["Download JSON", () => downloadJson(exportName, visible, columns)],
      ["Copy CSV", () => copyText(toCsv(visible, columns), btn)],
      ["Copy JSON", () => copyText(toJson(visible, columns), btn)],
      ["Copy Markdown", () => copyText(toMarkdown(visible, columns), btn)]
    ];
    for (const [label, run] of items) {
      const item = el2("button", "wb-menu-item", label);
      item.type = "button";
      item.addEventListener("click", () => {
        menu.hidden = true;
        run();
      });
      menu.append(item);
    }
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      menu.hidden = !menu.hidden;
    });
    document.addEventListener("click", () => {
      menu.hidden = true;
    });
    wrap.append(btn, menu);
    actions.append(wrap);
  }
  const scroller = el2("div", "wb-grid-scroll");
  const table = el2("table", "wb-table");
  const thead = el2("thead");
  const headRow = el2("tr");
  for (const col of columns) {
    const th = el2("th", "", col.label ?? col.key);
    if (col.width) th.style.width = col.width;
    th.tabIndex = 0;
    th.title = `Sort by ${col.label ?? col.key}`;
    const arrow = el2("span", "wb-sort-arrow", "");
    th.append(arrow);
    const sortBy = () => {
      if (sortKey === col.key) sortDir = -sortDir;
      else {
        sortKey = col.key;
        sortDir = 1;
      }
      render();
    };
    th.addEventListener("click", sortBy);
    th.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        sortBy();
      }
    });
    headRow.append(th);
  }
  thead.append(headRow);
  const tbody = el2("tbody");
  table.append(thead, tbody);
  scroller.append(table);
  const notice = el2("div", "wb-grid-notice");
  notice.hidden = true;
  const status = el2("div", "wb-grid-status");
  status.hidden = true;
  root.append(toolbar, notice, scroller, status);
  filter.addEventListener("input", () => {
    filterText = filter.value.trim().toLowerCase();
    render();
  });
  function matches(row) {
    if (!filterText) return true;
    return columns.some((col) => displayValue(row, col).toLowerCase().includes(filterText));
  }
  function compare(a, b) {
    const col = columns.find((c) => c.key === sortKey);
    if (!col) return 0;
    const va = cellValue2(a, col);
    const vb = cellValue2(b, col);
    if (va === vb) return 0;
    if (va === null || va === void 0) return 1;
    if (vb === null || vb === void 0) return -1;
    if (typeof va === "number" && typeof vb === "number") return (va - vb) * sortDir;
    return String(va).localeCompare(String(vb), void 0, { sensitivity: "base" }) * sortDir;
  }
  function render() {
    visible = rows.filter(matches);
    if (sortKey) visible = [...visible].sort(compare);
    count.textContent = filterText || visible.length !== rows.length ? `${visible.length} / ${rows.length}` : String(rows.length);
    for (const th of headRow.children) {
      const col = columns[[...headRow.children].indexOf(th)];
      th.querySelector(".wb-sort-arrow").textContent = col && col.key === sortKey ? sortDir === 1 ? " \u25B2" : " \u25BC" : "";
    }
    tbody.textContent = "";
    if (!visible.length) {
      const tr = el2("tr");
      const td = el2("td", "wb-empty", rows.length ? "No rows match the filter." : emptyText);
      td.colSpan = columns.length;
      tr.append(td);
      tbody.append(tr);
      return;
    }
    for (const row of visible) {
      const tr = el2("tr");
      if (onOpen) {
        tr.className = "wb-row-openable";
        tr.tabIndex = 0;
        tr.addEventListener("click", () => onOpen(row));
        tr.addEventListener("keydown", (e) => {
          if (e.key === "Enter") onOpen(row);
        });
      }
      tr.dataset.key = String(row[rowKey] ?? "");
      for (const col of columns) {
        const td = el2("td", col.mono ? "wb-mono" : "");
        const text = displayValue(row, col);
        if (col.copyable && text) {
          const span = el2("span", "sp-copy", text);
          span.title = "Click to copy";
          span.addEventListener("click", (e) => {
            e.stopPropagation();
            copyText(text, span);
          });
          td.append(span);
        } else {
          td.textContent = text;
        }
        tr.append(td);
      }
      tbody.append(tr);
    }
  }
  return {
    el: root,
    actionsEl: actions,
    setRows(next, { partial = false } = {}) {
      rows = Array.isArray(next) ? next : [];
      status.hidden = true;
      notice.hidden = !partial;
      if (partial) {
        notice.textContent = "\u26A0 Partial result set \u2014 the server returned more pages than the workbench cap.";
      }
      render();
    },
    setLoading(message = "Loading\u2026") {
      status.textContent = message;
      status.className = "wb-grid-status";
      status.hidden = false;
    },
    setError(err) {
      status.textContent = err?.message || String(err);
      status.className = "wb-grid-status wb-error";
      status.hidden = false;
    },
    getVisibleRows: () => [...visible],
    getColumns: () => columns
  };
}

// ../src/workbench/perm-kinds.js
var FLAGS = [
  // [name, bit] — bit as BigInt exponent in the combined 64-bit mask.
  ["ViewListItems", 0n],
  ["AddListItems", 1n],
  ["EditListItems", 2n],
  ["DeleteListItems", 3n],
  ["ApproveItems", 4n],
  ["OpenItems", 5n],
  ["ViewVersions", 6n],
  ["DeleteVersions", 7n],
  ["CancelCheckout", 8n],
  ["ManagePersonalViews", 9n],
  ["ManageLists", 11n],
  ["ViewFormPages", 12n],
  ["AnonymousSearchAccessList", 13n],
  ["Open", 16n],
  ["ViewPages", 17n],
  ["AddAndCustomizePages", 18n],
  ["ApplyThemeAndBorder", 19n],
  ["ApplyStyleSheets", 20n],
  ["ViewUsageData", 21n],
  ["CreateSSCSite", 22n],
  ["ManageSubwebs", 23n],
  ["CreateGroups", 24n],
  ["ManagePermissions", 25n],
  ["BrowseDirectories", 26n],
  ["BrowseUserInfo", 27n],
  ["AddDelPrivateWebParts", 28n],
  ["UpdatePersonalWebParts", 29n],
  ["ManageWeb", 30n],
  ["AnonymousSearchAccessWebLists", 32n],
  ["UseClientIntegration", 36n],
  ["UseRemoteAPIs", 37n],
  ["ManageAlerts", 38n],
  ["CreateAlerts", 39n],
  ["EditMyUserInfo", 40n],
  ["EnumeratePermissions", 62n]
];
var FULL_MASK = 0x7FFFFFFFFFFFFFFFn;
function combineBasePermissions(basePermissions) {
  const high = BigInt(String(basePermissions?.High ?? "0"));
  const low = BigInt(String(basePermissions?.Low ?? "0"));
  return high << 32n | low;
}
function decodeBasePermissions(basePermissions) {
  const mask = combineBasePermissions(basePermissions);
  if ((mask & FULL_MASK) === FULL_MASK) {
    return { flags: ["FullMask (all permissions)"], isFullControl: true, isEmpty: false };
  }
  const flags = FLAGS.filter(([, bit]) => (mask & 1n << bit) !== 0n).map(([name]) => name);
  return { flags, isFullControl: false, isEmpty: flags.length === 0 };
}
var PRINCIPAL_TYPE_NAMES = {
  1: "User",
  2: "Distribution list",
  4: "Security group",
  8: "SharePoint group"
};
var principalTypeName = (v) => PRINCIPAL_TYPE_NAMES[v] || String(v ?? "");

// ../src/inspect/tree-view.js
var el3 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function renderValue(node, opts = {}) {
  if (!node) return el3("span", "t-undef", "undefined");
  switch (node.t) {
    case "str": {
      const s = el3("span", "t-str", opts.bare ? node.v : JSON.stringify(node.v));
      if (node.trunc) s.append(el3("span", "t-truncated", ` \u2026(${node.trunc} chars)`));
      return s;
    }
    case "num":
      return el3("span", "t-num", String(node.v));
    case "bool":
      return el3("span", "t-bool", String(node.v));
    case "null":
      return el3("span", "t-null", "null");
    case "undef":
      return el3("span", "t-undef", "undefined");
    case "sym":
      return el3("span", "t-str", node.v);
    case "fn":
      return el3("span", "t-fn", `\u0192 ${node.v}()`);
    case "date":
      return el3("span", "t-node", node.v);
    case "regex":
      return el3("span", "t-str", node.v);
    case "node":
      return el3("span", "t-node", node.v);
    case "circ":
      return el3("span", "t-circular", "[circular]");
    case "maxdepth":
      return el3("span", "t-preview", node.v);
    case "err":
      return renderError(node);
    case "arr":
      return renderExpandable(node, `Array(${node.n})`, node.items.map((item, i) => [String(i), item]), opts);
    case "obj":
      return renderExpandable(node, node.cls === "Object" ? "" : node.cls, node.keys, opts);
    default:
      return el3("span", "t-preview", JSON.stringify(node));
  }
}
function renderError(node) {
  const wrap = el3("span");
  let head = `${node.name}: ${node.msg}`;
  if (node.status !== void 0) head += ` (HTTP ${node.status}${node.statusText ? " " + node.statusText : ""})`;
  wrap.append(el3("span", "t-err", head));
  if (node.stack) {
    const stack = el3("div", "stack-frame");
    stack.textContent = node.stack.split("\n").slice(1, 6).join("\n");
    wrap.append(stack);
  }
  return wrap;
}
function previewOf(node) {
  switch (node.t) {
    case "str": {
      const v = node.v.length > 24 ? node.v.slice(0, 24) + "\u2026" : node.v;
      return JSON.stringify(v);
    }
    case "num":
    case "bool":
      return String(node.v);
    case "null":
      return "null";
    case "undef":
      return "undefined";
    case "fn":
      return "\u0192";
    case "arr":
      return `Array(${node.n})`;
    case "obj":
      return node.cls === "Object" ? "{\u2026}" : `${node.cls}`;
    case "err":
      return node.name;
    case "node":
      return node.v;
    case "date":
      return node.v;
    case "maxdepth":
      return node.v;
    case "circ":
      return "[circular]";
    default:
      return "\u2026";
  }
}
function renderExpandable(node, label, entries, opts = {}) {
  const wrap = el3("span", "tree-node");
  const row = el3("span", "tree-row expandable");
  row.append(el3("span", "twist", "\u25B6"));
  if (label) row.append(el3("span", "", label + " "));
  const parts = entries.slice(0, 5).map(([k, v]) => (node.t === "arr" ? "" : `${k}: `) + previewOf(v));
  const openBrace = node.t === "arr" ? "[" : "{";
  const closeBrace = node.t === "arr" ? "]" : "}";
  const more = entries.length > 5 || node.trunc ? ", \u2026" : "";
  row.append(el3("span", "t-preview", `${openBrace}${parts.join(", ")}${more}${closeBrace}`));
  wrap.append(row);
  const children = el3("div", "tree-children");
  wrap.append(children);
  let built = false;
  row.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = row.classList.toggle("open");
    if (open && !built) {
      built = true;
      for (const [key2, val] of entries) {
        const line = el3("div");
        const keySpan = el3("span", "tree-key" + (opts.dimKeys?.has?.(key2) ? " dim-key" : ""), key2);
        line.append(keySpan, el3("span", "", ": "), renderValue(val, opts));
        children.append(line);
      }
      if (node.trunc) children.append(el3("div", "t-truncated", "\u2026 truncated"));
    }
  });
  return wrap;
}
function renderTable(dataNode, columns) {
  if (!dataNode || dataNode.t !== "arr" && dataNode.t !== "obj") {
    return renderValue(dataNode);
  }
  const rows = dataNode.t === "arr" ? dataNode.items.map((item, i) => [String(i), item]) : dataNode.keys;
  let cols = columns ? [...columns] : [];
  if (!cols.length) {
    const seen = /* @__PURE__ */ new Set();
    for (const [, v] of rows) {
      if (v.t === "obj") for (const [k] of v.keys) seen.add(k);
      else if (v.t === "arr") v.items.forEach((_, i) => seen.add(String(i)));
      else seen.add("Value");
    }
    cols = [...seen].slice(0, 20);
  }
  const wrap = el3("div", "console-table-wrap");
  const table = el3("table", "console-table");
  const thead = el3("thead");
  const hr = el3("tr");
  hr.append(el3("th", "", "(index)"));
  cols.forEach((c) => hr.append(el3("th", "", c)));
  thead.append(hr);
  table.append(thead);
  const tbody = el3("tbody");
  for (const [key2, v] of rows) {
    const tr = el3("tr");
    tr.append(el3("td", "", key2));
    for (const c of cols) {
      const td = el3("td");
      let cell;
      if (v.t === "obj") cell = v.keys.find(([k]) => k === c)?.[1];
      else if (v.t === "arr") cell = v.items[Number(c)];
      else if (c === "Value") cell = v;
      td.textContent = cell ? previewOf(cell) : "";
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  wrap.append(table);
  return wrap;
}

// ../src/inspect/sp-shapes.js
var NOISE_KEYS = /* @__PURE__ */ new Set(["__metadata", "__deferred", "odata.metadata", "odata.type", "odata.id", "odata.etag", "odata.editLink", "@odata.context", "@odata.type", "@odata.id", "@odata.etag", "@odata.editLink", "FirstUniqueAncestorSecurableObject", "RoleAssignments"]);
var key = (node, k) => node?.t === "obj" ? node.keys.find(([n]) => n === k)?.[1] : void 0;
var keyNames = (node) => node?.t === "obj" ? node.keys.map(([n]) => n) : [];
var str = (node) => node && (node.t === "str" || node.t === "num") ? String(node.v) : void 0;
function enhance(node) {
  if (!node || node.t !== "obj" && node.t !== "arr") return null;
  const d = key(node, "d");
  if (d && node.keys.length === 1) {
    return envelope("OData verbose", d, node);
  }
  const value = key(node, "value");
  if (value?.t === "arr" && keyNames(node).every((k) => k === "value" || k.startsWith("odata.") || k.startsWith("@odata."))) {
    return envelope("OData", value, node);
  }
  const results = key(node, "results");
  if (results?.t === "arr") {
    return collection(results, node);
  }
  if (node.t === "arr" && node.items.length && node.items.every(looksLikeSpObject)) {
    return collection(node, null);
  }
  if (looksLikeSpObject(node)) {
    return entity(node);
  }
  return null;
}
function looksLikeSpObject(node) {
  if (node?.t !== "obj") return false;
  if (key(node, "__metadata")) return true;
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  return has("InternalName", "TypeAsString") || has("BaseTemplate", "EntityTypeName") || has("ServerRelativeUrl", "WebTemplate") || has("LoginName", "PrincipalType") || names.includes("odata.type") || names.includes("@odata.type");
}
function spType(node) {
  const meta = key(node, "__metadata");
  return str(key(meta, "type")) || str(key(node, "odata.type")) || str(key(node, "@odata.type")) || detectShape(node);
}
function detectShape(node) {
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  if (has("InternalName", "TypeAsString")) return "SP.Field";
  if (has("BaseTemplate", "EntityTypeName")) return "SP.List";
  if (has("ServerRelativeUrl", "WebTemplate")) return "SP.Web";
  if (has("LoginName", "PrincipalType")) {
    return str(key(node, "OwnerTitle")) !== void 0 ? "SP.Group" : "SP.User";
  }
  return null;
}
function envelope(label, inner, outer) {
  const wrap = el3("div", "tree-node");
  const head = el3("div");
  head.append(badge(label));
  wrap.append(head);
  const enhanced = enhance(inner);
  wrap.append(enhanced ?? renderValue(inner, { dimKeys: NOISE_KEYS }));
  const metaKeys = outer.keys.filter(([k]) => k !== "d" && k !== "value");
  if (metaKeys.length) {
    const fold = el3("div", "sp-meta-fold");
    fold.append(renderValue({ t: "obj", cls: "envelope metadata", keys: metaKeys }, { dimKeys: NOISE_KEYS }));
    wrap.append(fold);
  }
  return wrap;
}
function collection(arrNode, parentNode) {
  const wrap = el3("div", "tree-node");
  const head = el3("div");
  const type = arrNode.items.length ? spType(arrNode.items[0]) : null;
  head.append(badge(`${arrNode.n} item${arrNode.n === 1 ? "" : "s"}`));
  if (type) head.append(el3("span", "sp-entity-head", shortType(type)));
  const toggle = el3("span", "table-toggle", "\u229E table view");
  head.append(toggle);
  wrap.append(head);
  const treeEl = el3("div");
  if (arrNode.items.length && arrNode.items.every((i) => i.t === "obj")) {
    const list2 = el3("div");
    arrNode.items.forEach((item, i) => {
      const row = el3("div");
      row.append(el3("span", "tree-key dim-key", `${i}: `));
      row.append(enhance(item) ?? renderValue(item, { dimKeys: NOISE_KEYS }));
      list2.append(row);
    });
    if (arrNode.trunc) list2.append(el3("div", "t-truncated", `\u2026 showing first ${arrNode.items.length} of ${arrNode.n}`));
    treeEl.append(list2);
  } else {
    treeEl.append(renderValue(arrNode, { dimKeys: NOISE_KEYS }));
  }
  const tableEl = el3("div");
  tableEl.hidden = true;
  let tableBuilt = false;
  toggle.addEventListener("click", () => {
    const showTable = tableEl.hidden;
    if (showTable && !tableBuilt) {
      tableBuilt = true;
      tableEl.append(renderTable(filterNoise(arrNode)));
    }
    tableEl.hidden = !showTable;
    treeEl.hidden = showTable;
    toggle.textContent = showTable ? "\u2261 tree view" : "\u229E table view";
  });
  wrap.append(treeEl, tableEl);
  const next = str(key(parentNode, "__next")) || str(key(parentNode, "odata.nextLink")) || str(key(parentNode, "@odata.nextLink"));
  if (next) {
    const warn = el3("div", "sp-next-link");
    warn.append(el3("span", "", "\u26A0 partial result set \u2014 next page: "));
    warn.append(copySpan(next, next.length > 80 ? next.slice(0, 80) + "\u2026" : next));
    wrap.append(warn);
  }
  return wrap;
}
function filterNoise(arrNode) {
  return {
    ...arrNode,
    items: arrNode.items.map((item) => item.t === "obj" ? { ...item, keys: item.keys.filter(([k]) => !NOISE_KEYS.has(k)) } : item)
  };
}
var ENTITY_FIELDS = {
  "SP.List": [
    ["Title", false],
    ["Id", true],
    ["EntityTypeName", true],
    ["BaseTemplate", false],
    ["ItemCount", false]
  ],
  "SP.Field": [
    ["Title", false],
    ["InternalName", true],
    ["TypeAsString", false],
    ["Required", false]
  ],
  "SP.Web": [
    ["Title", false],
    ["ServerRelativeUrl", true],
    ["WebTemplate", false]
  ],
  "SP.User": [
    ["Title", false],
    ["LoginName", true],
    ["Email", true]
  ],
  "SP.Group": [
    ["Title", false],
    ["Id", true],
    ["OwnerTitle", false]
  ],
  "SP.ListItem": [
    ["Title", false],
    ["Id", false]
  ]
};
function shortType(type) {
  if (!type) return "";
  if (type.startsWith("SP.Data.") && type.endsWith("Item")) return "SP.ListItem \xB7 " + type.slice(8);
  return type;
}
function entityKind(type) {
  if (!type) return null;
  if (ENTITY_FIELDS[type]) return type;
  for (const known of Object.keys(ENTITY_FIELDS)) {
    if (type.startsWith(known)) return known;
  }
  if (type.startsWith("SP.Data.")) return "SP.ListItem";
  return null;
}
function entity(node) {
  const type = spType(node);
  const kind = entityKind(type);
  const wrap = el3("div", "tree-node");
  const head = el3("div", "sp-entity-head");
  head.append(badge(shortType(type) || "SP"));
  if (kind) {
    for (const [field2, copyable] of ENTITY_FIELDS[kind]) {
      const v = key(node, field2);
      if (v === void 0) continue;
      const fieldEl = el3("span", "sp-field");
      fieldEl.append(el3("span", "dim-key tree-key", `${field2}: `));
      const text = v.t === "str" || v.t === "num" || v.t === "bool" ? String(v.v) : previewOf(v);
      fieldEl.append(copyable ? copySpan(text, text) : el3("span", "", text));
      head.append(fieldEl);
    }
  }
  wrap.append(head);
  wrap.append(renderValue(node, { dimKeys: NOISE_KEYS }));
  return wrap;
}
function badge(text) {
  return el3("span", "sp-badge", text);
}
function copySpan(copyText2, displayText) {
  const s = el3("span", "sp-copy", displayText);
  s.title = "Click to copy";
  s.addEventListener("click", async (e) => {
    e.stopPropagation();
    try {
      await navigator.clipboard.writeText(copyText2);
      s.classList.add("copied");
      setTimeout(() => s.classList.remove("copied"), 800);
    } catch {
    }
  });
  return s;
}

// ../src/inspect/to-node.js
function toNode(v, depth = 0, { maxDepth = 6, maxItems = 100 } = {}) {
  if (v === null) return { t: "null" };
  switch (typeof v) {
    case "string":
      return { t: "str", v };
    case "number":
      return { t: "num", v };
    case "boolean":
      return { t: "bool", v };
    case "undefined":
      return { t: "undef" };
  }
  if (depth >= maxDepth) return { t: "maxdepth", v: Array.isArray(v) ? `Array(${v.length})` : "{\u2026}" };
  const opts = { maxDepth, maxItems };
  if (Array.isArray(v)) {
    return { t: "arr", n: v.length, items: v.slice(0, maxItems).map((x) => toNode(x, depth + 1, opts)), trunc: v.length > maxItems };
  }
  const keys = Object.keys(v);
  return {
    t: "obj",
    cls: "Object",
    keys: keys.slice(0, maxItems).map((k) => [k, toNode(v[k], depth + 1, opts)]),
    trunc: keys.length > maxItems
  };
}

// ../src/workbench/views/lists.js
var BASE_TEMPLATE_NAMES = {
  100: "Generic list",
  101: "Document library",
  102: "Survey",
  103: "Links",
  104: "Announcements",
  105: "Contacts",
  106: "Events",
  107: "Tasks (classic)",
  108: "Discussion board",
  109: "Picture library",
  110: "Data sources",
  112: "User information",
  116: "Master page gallery",
  119: "Site pages",
  120: "Custom grid",
  140: "Workflow history",
  160: "Access requests",
  171: "Tasks",
  850: "Publishing pages"
};
var LIST_SELECT = [
  "Id",
  "Title",
  "BaseTemplate",
  "BaseType",
  "ItemCount",
  "Hidden",
  "Created",
  "LastItemModifiedDate",
  "EntityTypeName",
  "Description",
  "DefaultViewUrl",
  "RootFolder/ServerRelativeUrl"
];
var FIELD_SELECT = [
  "Id",
  "Title",
  "InternalName",
  "TypeAsString",
  "FieldTypeKind",
  "Required",
  "Hidden",
  "ReadOnlyField",
  "Group",
  "DefaultValue",
  "Choices",
  "Description",
  "EnforceUniqueValues",
  "Indexed"
];
var VIEW_SELECT = [
  "Id",
  "Title",
  "DefaultView",
  "PersonalView",
  "Hidden",
  "ServerRelativeUrl",
  "RowLimit",
  "Paged",
  "ViewQuery"
];
var CT_SELECT = ["Id", "Name", "Group", "Hidden", "ReadOnly", "Sealed", "Description"];
var fmtDate = (v) => v ? String(v).slice(0, 10) : "";
var choicesText = (v) => {
  const arr = Array.isArray(v) ? v : v?.results;
  return Array.isArray(arr) ? arr.join(" | ") : "";
};
var fileStem = (s) => String(s || "list").toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "list";
var el4 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var guidPath = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
function createListsView({ client: client2, navigate }) {
  const root = el4("section", "wb-view wb-view-lists");
  const gridPane = el4("div", "wb-pane");
  const head = el4("div", "wb-view-head");
  head.innerHTML = '<h2>Lists &amp; libraries</h2><p class="wb-view-hint">Every list in this web, hidden ones included. Click a row for fields, views, and content types.</p>';
  const grid = createGrid({
    columns: [
      { key: "Title", label: "Title" },
      { key: "BaseTemplate", label: "Template", format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? "") },
      { key: "ItemCount", label: "Items" },
      { key: "Hidden", label: "Hidden" },
      { key: "Url", label: "Url", value: (row) => row.RootFolder?.ServerRelativeUrl || "", mono: true, copyable: true },
      { key: "Id", label: "Id", mono: true, copyable: true },
      { key: "LastItemModifiedDate", label: "Modified", format: fmtDate }
    ],
    onOpen: (row) => navigate({ view: "lists", listId: row.Id, listTitle: row.Title }),
    emptyText: "No lists in this web.",
    filterPlaceholder: "Filter lists\u2026",
    exportName: "sp-lists"
  });
  gridPane.append(head, grid.el);
  const detailPane = el4("div", "wb-pane");
  detailPane.hidden = true;
  root.append(gridPane, detailPane);
  let listsLoaded = false;
  const tabCache = /* @__PURE__ */ new Map();
  async function loadLists() {
    if (listsLoaded) return;
    grid.setLoading("Loading lists\u2026");
    try {
      const { items, partial } = await client2.getAll("web/lists", {
        select: LIST_SELECT,
        expand: "RootFolder",
        orderby: "Title",
        top: 5e3
      });
      grid.setRows(items, { partial });
      listsLoaded = true;
    } catch (err) {
      grid.setError(err);
    }
  }
  function cached2(listId, tab, fetcher) {
    const key2 = `${listId}::${tab}`;
    if (!tabCache.has(key2)) {
      tabCache.set(key2, fetcher().catch((err) => {
        tabCache.delete(key2);
        throw err;
      }));
    }
    return tabCache.get(key2);
  }
  const TABS = [
    {
      id: "fields",
      label: "Fields",
      grid: (listId, title) => ({
        columns: [
          { key: "Title", label: "Title" },
          { key: "InternalName", label: "Internal name", mono: true, copyable: true },
          { key: "TypeAsString", label: "Type" },
          { key: "Required", label: "Required" },
          { key: "Hidden", label: "Hidden" },
          { key: "ReadOnlyField", label: "Read-only" },
          { key: "Choices", label: "Choices", format: choicesText },
          { key: "DefaultValue", label: "Default" },
          { key: "Group", label: "Group" }
        ],
        exportName: `fields-${fileStem(title)}`,
        fetch: () => client2.getAll(guidPath(listId, "/fields"), { select: FIELD_SELECT })
      })
    },
    {
      id: "views",
      label: "Views",
      grid: (listId, title) => ({
        columns: [
          { key: "Title", label: "Title" },
          { key: "DefaultView", label: "Default" },
          { key: "Hidden", label: "Hidden" },
          { key: "PersonalView", label: "Personal" },
          { key: "RowLimit", label: "Row limit" },
          { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true },
          { key: "ViewQuery", label: "CAML query", mono: true, copyable: true }
        ],
        exportName: `views-${fileStem(title)}`,
        fetch: () => client2.getAll(guidPath(listId, "/views"), { select: VIEW_SELECT })
      })
    },
    {
      id: "contenttypes",
      label: "Content types",
      grid: (listId, title) => ({
        columns: [
          { key: "Name", label: "Name" },
          { key: "Id", label: "Id", value: (row) => row.Id?.StringValue || String(row.Id ?? ""), mono: true, copyable: true },
          { key: "Group", label: "Group" },
          { key: "Hidden", label: "Hidden" },
          { key: "ReadOnly", label: "Read-only" },
          { key: "Sealed", label: "Sealed" },
          { key: "Description", label: "Description" }
        ],
        exportName: `contenttypes-${fileStem(title)}`,
        fetch: () => client2.getAll(guidPath(listId, "/contenttypes"), { select: CT_SELECT })
      })
    },
    {
      id: "permissions",
      label: "Permissions",
      grid: (listId, title) => ({
        columns: [
          { key: "Member", label: "Principal", value: (row) => row.Member?.Title || "" },
          { key: "LoginName", label: "Login", value: (row) => row.Member?.LoginName || "", mono: true, copyable: true },
          { key: "PrincipalType", label: "Type", value: (row) => row.Member?.PrincipalType, format: principalTypeName },
          {
            key: "Roles",
            label: "Roles",
            value: (row) => (row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || []).map((r) => r.Name).filter(Boolean).join(", ")
          }
        ],
        exportName: `permissions-${fileStem(title)}`,
        fetch: () => client2.getAll(guidPath(listId, "/roleassignments"), {
          expand: ["Member", "RoleDefinitionBindings"],
          select: [
            "PrincipalId",
            "Member/Id",
            "Member/Title",
            "Member/LoginName",
            "Member/PrincipalType",
            "RoleDefinitionBindings/Id",
            "RoleDefinitionBindings/Name"
          ]
        })
      })
    },
    { id: "raw", label: "Raw" }
  ];
  function showDetail(route) {
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = "";
    const listId = route.listId;
    const back = el4("button", "btn btn-xs wb-back", "\u2190 All lists");
    back.type = "button";
    back.addEventListener("click", () => navigate({ view: "lists" }));
    const title = el4("h2", "", route.listTitle || "List");
    const sub = el4("span", "wb-detail-id sp-copy", listId);
    sub.title = "Click to copy the list id";
    sub.addEventListener("click", () => copyText(listId, sub));
    const headRow = el4("div", "wb-detail-head");
    headRow.append(back, title, sub);
    const tabsBar = el4("div", "wb-tabs");
    tabsBar.setAttribute("role", "tablist");
    const body = el4("div", "wb-tab-body");
    const panes = /* @__PURE__ */ new Map();
    let activeTab = null;
    function activate(tab) {
      activeTab = tab.id;
      for (const btn of tabsBar.children) {
        btn.classList.toggle("active", btn.dataset.tab === tab.id);
        btn.setAttribute("aria-selected", btn.dataset.tab === tab.id ? "true" : "false");
      }
      body.textContent = "";
      body.append(pane(tab));
    }
    function pane(tab) {
      if (panes.has(tab.id)) return panes.get(tab.id);
      const wrap = el4("div", "wb-tab-pane");
      panes.set(tab.id, wrap);
      if (tab.id === "raw") {
        const status = el4("div", "wb-grid-status", "Loading raw list entity\u2026");
        wrap.append(status);
        cached2(listId, "raw", () => client2.get(guidPath(listId))).then((json) => {
          status.remove();
          const node = toNode(json, 0, { maxDepth: 8, maxItems: 250 });
          const inspector = el4("div", "wb-raw");
          inspector.append(enhance(node) ?? renderValue(node));
          wrap.append(inspector);
        }).catch((err) => {
          status.textContent = err?.message || String(err);
          status.classList.add("wb-error");
        });
        return wrap;
      }
      const spec = tab.grid(listId, route.listTitle);
      const tabGrid = createGrid({
        columns: spec.columns,
        emptyText: "Nothing here.",
        filterPlaceholder: `Filter ${tab.label.toLowerCase()}\u2026`,
        exportName: spec.exportName
      });
      wrap.append(tabGrid.el);
      tabGrid.setLoading(`Loading ${tab.label.toLowerCase()}\u2026`);
      cached2(listId, tab.id, spec.fetch).then(({ items, partial }) => tabGrid.setRows(items, { partial })).catch((err) => tabGrid.setError(err));
      return wrap;
    }
    for (const tab of TABS) {
      const btn = el4("button", "wb-tab", tab.label);
      btn.type = "button";
      btn.dataset.tab = tab.id;
      btn.setAttribute("role", "tab");
      btn.addEventListener("click", () => activate(tab));
      tabsBar.append(btn);
    }
    detailPane.append(headRow, tabsBar, body);
    activate(TABS.find((t) => t.id === route.tab) || TABS[0]);
    void activeTab;
  }
  function load(route) {
    if (route?.listId) {
      showDetail(route);
    } else {
      detailPane.hidden = true;
      gridPane.hidden = false;
      loadLists();
    }
  }
  return { el: root, load, grid };
}

// ../src/workbench/views/security.js
var el5 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var roleNames = (row) => (row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || []).map((r) => r.Name).filter(Boolean).join(", ");
function createSecurityView({ client: client2 }) {
  const root = el5("section", "wb-view wb-view-security");
  const head = el5("div", "wb-view-head");
  head.innerHTML = '<h2>Users, groups &amp; permissions</h2><p class="wb-view-hint">Site groups, role definitions, and who holds what on this web. The inheritance scan is on-demand \u2014 it makes SharePoint evaluate security per list.</p>';
  const tabsBar = el5("div", "wb-tabs");
  const body = el5("div", "wb-tab-body");
  root.append(head, tabsBar, body);
  const panes = /* @__PURE__ */ new Map();
  function groupsPane() {
    const wrap = el5("div", "wb-tab-pane");
    const grid = createGrid({
      columns: [
        { key: "Title", label: "Group" },
        { key: "Id", label: "Id" },
        { key: "OwnerTitle", label: "Owner" },
        { key: "OnlyAllowMembersViewMembership", label: "Members-only view" },
        { key: "Description", label: "Description" }
      ],
      onOpen: openMembers,
      emptyText: "No site groups.",
      filterPlaceholder: "Filter groups\u2026",
      exportName: "sp-groups"
    });
    const membersBox = el5("div", "wb-subpanel");
    membersBox.hidden = true;
    const membersTitle = el5("h3", "wb-subpanel-title", "");
    const membersHost = el5("div", "wb-subpanel-body");
    membersBox.append(membersTitle, membersHost);
    wrap.append(grid.el, membersBox);
    grid.setLoading("Loading site groups\u2026");
    client2.getAll("web/sitegroups", {
      select: ["Id", "Title", "Description", "OwnerTitle", "PrincipalType", "OnlyAllowMembersViewMembership"]
    }).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
    function openMembers(group) {
      membersBox.hidden = false;
      membersTitle.textContent = `Members of ${group.Title}`;
      membersHost.textContent = "";
      const membersGrid = createGrid({
        columns: [
          { key: "Title", label: "Name" },
          { key: "LoginName", label: "Login", mono: true, copyable: true },
          { key: "Email", label: "Email", copyable: true },
          { key: "IsSiteAdmin", label: "Site admin" },
          { key: "PrincipalType", label: "Type", format: principalTypeName }
        ],
        emptyText: "No members.",
        filterPlaceholder: "Filter members\u2026",
        exportName: `members-${group.Id}`
      });
      membersHost.append(membersGrid.el);
      membersGrid.setLoading("Loading members\u2026");
      client2.getAll(`web/sitegroups(${group.Id})/users`, {
        select: ["Id", "Title", "LoginName", "Email", "IsSiteAdmin", "PrincipalType"]
      }).then(({ items }) => membersGrid.setRows(items)).catch((err) => membersGrid.setError(err));
    }
    return wrap;
  }
  function roleDefsPane() {
    const wrap = el5("div", "wb-tab-pane");
    const grid = createGrid({
      columns: [
        { key: "Name", label: "Role" },
        { key: "RoleTypeKind", label: "Kind" },
        { key: "Hidden", label: "Hidden" },
        {
          key: "BasePermissions",
          label: "Permissions",
          value: (row) => decodeBasePermissions(row.BasePermissions).flags.length,
          format: (v, row) => {
            const d = decodeBasePermissions(row.BasePermissions);
            if (d.isFullControl) return "Full control";
            if (d.isEmpty) return "None";
            return `${d.flags.length} flags`;
          }
        },
        { key: "Description", label: "Description" }
      ],
      onOpen: openDecode,
      emptyText: "No role definitions.",
      filterPlaceholder: "Filter roles\u2026",
      exportName: "sp-roledefinitions"
    });
    const decodeBox = el5("div", "wb-subpanel");
    decodeBox.hidden = true;
    const decodeTitle = el5("h3", "wb-subpanel-title", "");
    const decodeBody = el5("div", "wb-subpanel-body wb-flags");
    decodeBox.append(decodeTitle, decodeBody);
    wrap.append(grid.el, decodeBox);
    grid.setLoading("Loading role definitions\u2026");
    client2.getAll("web/roledefinitions", {
      select: ["Id", "Name", "Description", "RoleTypeKind", "Hidden", "BasePermissions"]
    }).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
    function openDecode(role) {
      decodeBox.hidden = false;
      const d = decodeBasePermissions(role.BasePermissions);
      decodeTitle.textContent = `${role.Name} \u2014 ${d.isFullControl ? "full control" : `${d.flags.length} permission flags`}`;
      decodeBody.textContent = "";
      for (const flag of d.flags) decodeBody.append(el5("span", "wb-flag", flag));
      if (d.isEmpty) decodeBody.append(el5("span", "wb-view-hint", "No permission bits set."));
    }
    return wrap;
  }
  function assignmentsPane() {
    const wrap = el5("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "PrincipalId",
      columns: [
        { key: "Member", label: "Principal", value: (row) => row.Member?.Title || "" },
        { key: "LoginName", label: "Login", value: (row) => row.Member?.LoginName || "", mono: true, copyable: true },
        { key: "PrincipalType", label: "Type", value: (row) => row.Member?.PrincipalType, format: principalTypeName },
        { key: "Roles", label: "Roles", value: roleNames }
      ],
      emptyText: "No role assignments.",
      filterPlaceholder: "Filter assignments\u2026",
      exportName: "sp-roleassignments"
    });
    wrap.append(grid.el);
    grid.setLoading("Loading role assignments\u2026");
    client2.getAll("web/roleassignments", {
      expand: ["Member", "RoleDefinitionBindings"],
      select: [
        "PrincipalId",
        "Member/Id",
        "Member/Title",
        "Member/LoginName",
        "Member/PrincipalType",
        "RoleDefinitionBindings/Id",
        "RoleDefinitionBindings/Name"
      ]
    }).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
    return wrap;
  }
  function inheritancePane() {
    const wrap = el5("div", "wb-tab-pane");
    const bar = el5("div", "wb-scan-bar");
    const btn = el5("button", "btn", "Scan lists for unique permissions");
    btn.type = "button";
    const hint = el5(
      "span",
      "wb-view-hint",
      "Asks SharePoint for HasUniqueRoleAssignments on every list \u2014 slow on large sites, so it only runs on demand."
    );
    bar.append(btn, hint);
    const grid = createGrid({
      columns: [
        { key: "Title", label: "List" },
        { key: "HasUniqueRoleAssignments", label: "Unique permissions" },
        { key: "BaseTemplate", label: "Template", format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? "") },
        { key: "Hidden", label: "Hidden" },
        { key: "Id", label: "Id", mono: true, copyable: true }
      ],
      emptyText: "Run the scan to see results.",
      filterPlaceholder: "Filter results\u2026",
      exportName: "sp-unique-permissions"
    });
    wrap.append(bar, grid.el);
    grid.setRows([]);
    btn.addEventListener("click", async () => {
      btn.disabled = true;
      grid.setLoading("Scanning\u2026 (per-list security evaluation)");
      try {
        const { items, partial } = await client2.getAll("web/lists", {
          select: ["Id", "Title", "Hidden", "BaseTemplate", "HasUniqueRoleAssignments"],
          top: 5e3
        });
        const broken = items.filter((l) => l.HasUniqueRoleAssignments);
        grid.setRows(broken.length ? broken : items, { partial });
        hint.textContent = broken.length ? `${broken.length} of ${items.length} lists break inheritance (showing them).` : `No list breaks inheritance (showing all ${items.length} scanned).`;
      } catch (err) {
        grid.setError(err);
      } finally {
        btn.disabled = false;
      }
    });
    return wrap;
  }
  const TABS = [
    { id: "groups", label: "Groups", build: groupsPane },
    { id: "roledefs", label: "Role definitions", build: roleDefsPane },
    { id: "assignments", label: "Role assignments", build: assignmentsPane },
    { id: "inheritance", label: "Inheritance scan", build: inheritancePane }
  ];
  function activate(tab) {
    for (const btn of tabsBar.children) {
      btn.classList.toggle("active", btn.dataset.tab === tab.id);
    }
    if (!panes.has(tab.id)) panes.set(tab.id, tab.build());
    body.textContent = "";
    body.append(panes.get(tab.id));
  }
  for (const tab of TABS) {
    const btn = el5("button", "wb-tab", tab.label);
    btn.type = "button";
    btn.dataset.tab = tab.id;
    btn.addEventListener("click", () => activate(tab));
    tabsBar.append(btn);
  }
  function load() {
    if (!tabsBar.querySelector(".wb-tab.active")) activate(TABS[0]);
  }
  return { el: root, load };
}

// ../src/workbench/main.js
var GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>',
  security: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 1.8 13 3.6v3.6c0 3.2-2.1 5.6-5 6.9-2.9-1.3-5-3.7-5-6.9V3.6z"/><path d="m5.8 7.8 1.6 1.6 2.9-3"/></svg>'
};
function applyWorkbenchContext(ctx2) {
  const chip = document.getElementById("wb-chip");
  const chipText = document.getElementById("wb-chip-text");
  const statusCtx = document.getElementById("wb-status-context");
  chip.classList.toggle("sp-chip-live", ctx2.live);
  chip.classList.toggle("sp-chip-mock", !ctx2.live);
  chipText.textContent = ctx2.live ? "SP: Live" : "SP: Mock";
  chip.title = ctx2.live ? `Connected to ${ctx2.label}${ctx2.user ? ` as ${ctx2.user}` : ""} \xB7 context: ${ctx2.source}` : "Not connected to a SharePoint web \u2014 showing built-in mock data";
  statusCtx.textContent = ctx2.live ? `SP: ${ctx2.label}${ctx2.user ? ` \xB7 ${ctx2.user}` : ""}` : "SP: mock data (deploy to SharePoint for live inspection)";
}
var ctx = getSpContext();
applyWorkbenchContext(ctx);
var client = createSpRestClient({
  mockResolver: ctx.live ? null : mockResolver
});
var shell = createShell({
  mount: document.getElementById("wb-main"),
  deps: { client },
  views: [
    { id: "lists", label: "Lists", glyph: GLYPHS.lists, create: createListsView },
    { id: "security", label: "Security", glyph: GLYPHS.security, create: createSecurityView }
    // M4: { id: 'site', label: 'Site', ... }
  ]
});
shell.restore();
