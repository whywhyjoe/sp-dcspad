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
  for (const key of ["parent", "top"]) {
    try {
      const candidate = window[key];
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
    for (const key of SERIALIZABLE_FIELDS) {
      if (found.pageContext[key] !== void 0) {
        pageContext[key] = found.pageContext[key];
      }
    }
  }
  for (const key of SERIALIZABLE_FIELDS) {
    if (pageContext[key] === void 0 && found.raw[key] !== void 0) {
      pageContext[key] = found.raw[key];
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
function assignment(principalId, title, principalType, roleNames) {
  return {
    PrincipalId: principalId,
    Member: { Id: principalId, Title: title, LoginName: title, PrincipalType: principalType },
    RoleDefinitionBindings: roleNames.map((name) => ({
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

// ../src/workbench/grid.js
var el2 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
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
function displayValue(row, col) {
  const v = cellValue(row, col);
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
  filterPlaceholder = "Filter\u2026"
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
    const va = cellValue(a, col);
    const vb = cellValue(b, col);
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
var fmtDate = (v) => v ? String(v).slice(0, 10) : "";
function createListsView({ client: client2, navigate }) {
  const root = document.createElement("section");
  root.className = "wb-view wb-view-lists";
  const head = document.createElement("div");
  head.className = "wb-view-head";
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
    filterPlaceholder: "Filter lists\u2026"
  });
  root.append(head, grid.el);
  let loaded = false;
  async function load(route) {
    void route;
    if (loaded) return;
    grid.setLoading("Loading lists\u2026");
    try {
      const { items, partial } = await client2.getAll("web/lists", {
        select: LIST_SELECT,
        expand: "RootFolder",
        orderby: "Title",
        top: 5e3
      });
      grid.setRows(items, { partial });
      loaded = true;
    } catch (err) {
      grid.setError(err);
    }
  }
  return { el: root, load, grid };
}

// ../src/workbench/main.js
var GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>'
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
    { id: "lists", label: "Lists", glyph: GLYPHS.lists, create: createListsView }
    // M3: { id: 'security', label: 'Security', ... }
    // M4: { id: 'site', label: 'Site', ... }
  ]
});
shell.restore();
