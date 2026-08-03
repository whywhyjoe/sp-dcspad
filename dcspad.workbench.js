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
function odataPathLiteral(value) {
  return encodeURIComponent(String(value)).replaceAll("'", "''");
}
function resultArray(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.results)) return value.results;
  return [];
}
function unwrapJson(data) {
  return data?.d?.GetContextWebInformation || data?.GetContextWebInformation || data?.d || data;
}
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

// ../src/workbench/sp-rest.js?v=2
var PAGE_CAP = 5e3;
var MAX_CONCURRENT = 3;
var RETRY_STATUSES = /* @__PURE__ */ new Set([429, 503]);
function buildQuery({ select, expand, filter, orderby, top } = {}) {
  const parts = [];
  const join2 = (v) => Array.isArray(v) ? v.join(",") : String(v);
  if (select) parts.push(`$select=${join2(select)}`);
  if (expand) parts.push(`$expand=${join2(expand)}`);
  if (filter) parts.push(`$filter=${encodeURIComponent(String(filter))}`);
  if (orderby) parts.push(`$orderby=${join2(orderby)}`);
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
  let targetWebUrl = "";
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
  function hostWebUrl() {
    return context().pageContext.webAbsoluteUrl.replace(/\/+$/, "");
  }
  function webUrl() {
    return targetWebUrl || hostWebUrl();
  }
  function normalizeTarget(input) {
    const raw = String(input || "").trim();
    if (!raw) return "";
    const host = hostWebUrl();
    let candidate;
    try {
      candidate = new URL(raw, host);
    } catch {
      throw new SpFileError(
        "Enter a site URL on this tenant, such as /sites/ProjectName.",
        { code: "invalid-web-url" }
      );
    }
    if (!/^https?:$/.test(candidate.protocol) || candidate.origin !== new URL(host).origin) {
      throw new SpFileError(
        "That URL is on a different tenant \u2014 the workbench can only inspect sites on its own origin.",
        { code: "invalid-web-url" }
      );
    }
    candidate.hash = "";
    candidate.search = "";
    return candidate.href.replace(/\/+$/, "");
  }
  async function connectWeb(input) {
    const candidate = normalizeTarget(input);
    if (!candidate) {
      targetWebUrl = "";
      return entityOf(await rawGet(`${hostWebUrl()}/_api/web?$select=Id,Title,Url,ServerRelativeUrl`));
    }
    const web = entityOf(await rawGet(`${candidate}/_api/web?$select=Id,Title,Url,ServerRelativeUrl`));
    targetWebUrl = normalizeTarget(web?.Url) || candidate;
    return web;
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
      const remaining = PAGE_CAP - items.length;
      if (page.length > remaining) {
        items.push(...page.slice(0, remaining));
        partial = true;
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
  return { context, webUrl, hostWebUrl, connectWeb, apiUrl, get, getAll };
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
  ],
  // Documents library: the full editor-type spread for the Files browser,
  // including a read-only User field displayed via FieldValuesAsText.
  "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01": [
    field("Title", "Title", "Text", 2),
    field("Document category", "DocCategory", "Choice", 6, {
      Choices: ["Contract", "Report", "Misc"],
      DefaultValue: "Misc"
    }),
    field("Confidential", "Confidential", "Boolean", 8),
    field("Published date", "PublishedDate", "DateTime", 4),
    field("Source link", "SourceLink", "URL", 11),
    field("DocVersion", "DocVersion", "Text", 2),
    field("Author", "Author", "User", 20, { ReadOnlyField: true }),
    field("ID", "ID", "Counter", 5, { ReadOnlyField: true })
  ],
  // Site Pages: one field per editor type the metadata form supports, plus
  // the content fields the editor must refuse to touch.
  "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02": [
    field("Title", "Title", "Text", 2, { Required: true }),
    field("Description", "Description", "Note", 3),
    field("Page category", "PageCategory", "Choice", 6, {
      Choices: ["Announcement", "How-to", "Reference"],
      DefaultValue: "Reference"
    }),
    field("Review date", "ReviewDate", "DateTime", 4),
    field("Show in navigation", "ShowInNav", "Boolean", 8),
    field("Related link", "RelatedLink", "URL", 11),
    field("Promoted state", "PromotedState", "Number", 9, { ReadOnlyField: true }),
    field("Editor", "Editor", "User", 20, { ReadOnlyField: true }),
    field("Canvas content", "CanvasContent1", "Note", 3),
    field("ID", "ID", "Counter", 5, { ReadOnlyField: true })
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
var PROJECT_ITEMS = [
  item(1, "Intranet refresh", { ProjectStatus: "Active", DueDate: "2026-09-15T00:00:00Z", Budget: 12e3 }),
  item(2, "Records migration", { ProjectStatus: "Planned", DueDate: "2026-11-01T00:00:00Z", Budget: 4e4 }),
  item(3, "Team site cleanup", { ProjectStatus: "Done", DueDate: "2026-03-30T00:00:00Z", Budget: 1500 }),
  item(4, "Permission audit", { ProjectStatus: "Blocked", DueDate: "2026-08-05T00:00:00Z", Budget: 0 }),
  item(5, "Search tuning", { ProjectStatus: "Active", DueDate: "2026-10-20T00:00:00Z", Budget: 8e3 }),
  item(6, "Archive rollout", { ProjectStatus: "Planned", DueDate: "2027-01-10T00:00:00Z", Budget: 22e3 })
];
function item(id, title, extra = {}) {
  return {
    Id: id,
    ID: id,
    Title: title,
    Modified: "2026-07-18T10:00:00Z",
    Created: "2026-05-02T09:00:00Z",
    ...extra
  };
}
var HOME_CANVAS = JSON.stringify([
  {
    controlType: 4,
    id: "a1000000-0000-4000-8000-000000000001",
    position: { zoneIndex: 1, sectionIndex: 1, controlIndex: 1, sectionFactor: 12, layoutIndex: 1 },
    emphasis: {},
    innerHTML: "<h2>Welcome</h2><p>Welcome to the mock intranet home page.</p>"
  },
  {
    controlType: 3,
    id: "a1000000-0000-4000-8000-000000000002",
    position: { zoneIndex: 2, sectionIndex: 1, controlIndex: 1, sectionFactor: 6, layoutIndex: 1 },
    emphasis: { zoneEmphasis: 1 },
    webPartId: "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
    webPartData: {
      id: "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
      title: "Quick links",
      description: "Mock quick links",
      properties: { items: [{ title: "Docs" }, { title: "Pad" }] },
      serverProcessedContent: {
        htmlStrings: {},
        searchablePlainTexts: { "items[0].title": "Docs", "items[1].title": "Pad" },
        imageSources: {},
        links: { baseUrl: "/SitePages" }
      }
    }
  },
  {
    controlType: 3,
    id: "a1000000-0000-4000-8000-000000000003",
    position: { zoneIndex: 2, sectionIndex: 2, controlIndex: 1, sectionFactor: 6, layoutIndex: 1 },
    emphasis: { zoneEmphasis: 1 },
    webPartId: "ffff0000-1111-2222-3333-444455556666",
    webPartData: {
      id: "ffff0000-1111-2222-3333-444455556666",
      title: "Mystery part",
      properties: {},
      serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} }
    }
  },
  { horrible: "shape", with: ["no", "controlType"] },
  { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } }
]);
var LEGACY_CANVAS = '<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="{&quot;controlType&quot;:4,&quot;id&quot;:&quot;b2000000-0000-4000-8000-000000000001&quot;,&quot;position&quot;:{&quot;zoneIndex&quot;:1,&quot;sectionIndex&quot;:1,&quot;controlIndex&quot;:1,&quot;sectionFactor&quot;:12}}"><div data-sp-rte=""><p>Legacy formatted news body.</p></div></div></div>';
var SITEPAGES_ITEMS = [
  {
    ...item(1, "Home", {
      FileLeafRef: "Home.aspx",
      FileRef: "/SitePages/Home.aspx",
      FileDirRef: "/SitePages",
      PromotedState: 0,
      UniqueId: "ee000000-0000-4000-8000-000000000001",
      Author: { Title: "Mock Developer" },
      Editor: { Title: "Mock Developer" },
      CanvasContent1: HOME_CANVAS,
      LayoutWebpartsContent: null,
      Description: "Mock landing page.",
      BannerImageUrl: null,
      PageCategory: "Announcement",
      ReviewDate: "2026-08-01T00:00:00Z",
      ShowInNav: true,
      RelatedLink: { Url: "https://example.com", Description: "Example" },
      FieldValuesAsText: { Editor: "Mock Developer", CanvasContent1: "(canvas markup)" }
    })
  },
  {
    ...item(2, "Release notes", {
      FileLeafRef: "News-Update.aspx",
      FileRef: "/SitePages/News-Update.aspx",
      FileDirRef: "/SitePages",
      PromotedState: 2,
      UniqueId: "ee000000-0000-4000-8000-000000000002",
      Author: { Title: "Pat Example" },
      Editor: { Title: "Pat Example" },
      CanvasContent1: LEGACY_CANVAS,
      FieldValuesAsText: { Editor: "Pat Example" }
    })
  },
  {
    ...item(3, "Blank page", {
      FileLeafRef: "Blank.aspx",
      FileRef: "/SitePages/Blank.aspx",
      FileDirRef: "/SitePages",
      PromotedState: 0,
      UniqueId: "ee000000-0000-4000-8000-000000000003",
      Author: { Title: "Mock Developer" },
      Editor: { Title: "Mock Developer" },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: "Mock Developer" }
    })
  },
  // Pages in subfolders — the Pages view surfaces and sorts by folder.
  {
    ...item(4, "Weekly roundup", {
      FileLeafRef: "Weekly.aspx",
      FileRef: "/SitePages/news/Weekly.aspx",
      FileDirRef: "/SitePages/news",
      PromotedState: 2,
      UniqueId: "ee000000-0000-4000-8000-000000000004",
      Author: { Title: "Pat Example" },
      Editor: { Title: "Pat Example" },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: "Pat Example" }
    })
  },
  {
    ...item(5, "R\xE9sum\xE9 hebdo", {
      FileLeafRef: "Hebdo.aspx",
      FileRef: "/SitePages/news/fr/Hebdo.aspx",
      FileDirRef: "/SitePages/news/fr",
      PromotedState: 2,
      UniqueId: "ee000000-0000-4000-8000-000000000005",
      Author: { Title: "Mock Developer" },
      Editor: { Title: "Mock Developer" },
      CanvasContent1: null,
      FieldValuesAsText: { Editor: "Mock Developer" }
    })
  }
];
var ITEMS = {
  "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e03": PROJECT_ITEMS,
  "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e02": SITEPAGES_ITEMS
};
function mockFile(name, length, modified = "2026-07-10T09:00:00Z") {
  return {
    Name: name,
    ServerRelativeUrl: `__FOLDER__/${name}`,
    Length: length,
    TimeLastModified: modified,
    UIVersionLabel: "1.0",
    CheckOutType: 2
  };
}
var MOCK_TREE = {
  "/shared documents": {
    folders: [
      { Name: "Reports", ServerRelativeUrl: "/Shared Documents/Reports", ItemCount: 2, TimeLastModified: "2026-07-01T12:00:00Z" }
    ],
    files: [
      mockFile("proposal.docx", 48230),
      mockFile("logo.png", 15872),
      mockFile("archive.zip", 1048576),
      mockFile("notes.txt", 812),
      mockFile("widget.js", 2048),
      mockFile("data.csv", 5300)
    ]
  },
  "/shared documents/reports": {
    folders: [],
    files: [mockFile("q1-report.docx", 91e3), mockFile("q2-report.docx", 87e3)]
  }
};
for (const [folderPath, listing] of Object.entries(MOCK_TREE)) {
  for (const f of listing.files) {
    f.ServerRelativeUrl = f.ServerRelativeUrl.replace(
      "__FOLDER__",
      folderPath === "/shared documents" ? "/Shared Documents" : "/Shared Documents/Reports"
    );
  }
}
var FILE_ITEMS = {
  "/shared documents/proposal.docx": {
    Id: 201,
    Title: "Project proposal",
    DocCategory: "Contract",
    Confidential: true,
    PublishedDate: "2026-06-01T00:00:00Z",
    SourceLink: { Url: "https://example.com/spec", Description: "Spec" },
    DocVersion: "1.4",
    FieldValuesAsText: { Author: "Mock Developer" }
  }
};
var DOC_LIB_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5e01";
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
    const itemId = /\/items\((\d+)\)/.exec(path)?.[1];
    if (itemId) {
      const single = (ITEMS[found.Id] || []).find((i) => i.Id === Number(itemId));
      return single ?? null;
    }
    if (path.includes("/items")) return { value: ITEMS[found.Id] || [] };
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
  const folderPathOf = /getfolderbyserverrelativepath\(decodedurl='([^']*)'\)/.exec(path)?.[1];
  if (folderPathOf !== void 0) {
    let decoded = folderPathOf;
    try {
      decoded = decodeURIComponent(folderPathOf);
    } catch {
    }
    const listing = MOCK_TREE[decoded];
    if (path.includes("/folders")) return { value: listing?.folders || [] };
    if (path.includes("/files")) return { value: listing?.files || [] };
    if (path.includes("parentlist")) {
      return { ListItemAllFields: { ParentList: { Id: DOC_LIB_ID } } };
    }
    return { Name: decoded.split("/").pop() || "", ServerRelativeUrl: decoded };
  }
  const filePathOf = /getfilebyserverrelativepath\(decodedurl='([^']*)'\)/.exec(path)?.[1];
  if (filePathOf !== void 0) {
    let decoded = filePathOf;
    try {
      decoded = decodeURIComponent(filePathOf);
    } catch {
    }
    if (path.includes("/listitemallfields")) {
      return FILE_ITEMS[decoded] || { Id: 0, Title: "" };
    }
    return null;
  }
  if (path.startsWith("web/allproperties")) return ALL_PROPERTIES;
  if (path.startsWith("web/regionalsettings")) return REGIONAL_SETTINGS;
  if (path.startsWith("web/currentuser")) return CURRENT_USER;
  if (path.startsWith("web/webs")) return { value: SUBWEBS };
  if (path.startsWith("web/features")) return { value: FEATURES.web };
  if (path.startsWith("site/features")) return { value: FEATURES.site };
  if (path.startsWith("site")) return SITE;
  if (path.startsWith("web")) {
    const base = url.slice(0, url.indexOf("/_api/")).replace(/\/+$/, "");
    let rel = "/";
    try {
      rel = decodeURIComponent(new URL(base).pathname) || "/";
    } catch {
    }
    return {
      ...WEB,
      Url: base || WEB.Url,
      ServerRelativeUrl: rel,
      Title: rel === "/" ? WEB.Title : `Mock Web (${rel})`
    };
  }
  return null;
}

// ../src/workbench/shell.js?v=2
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
  let lastGroup = null;
  for (const view of views) {
    if (view.group !== void 0 && view.group !== lastGroup) {
      if (lastGroup !== null) rail.append(el("div", "wb-rail-sep"));
      rail.append(el("div", "wb-rail-group", view.group));
    }
    lastGroup = view.group ?? lastGroup;
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
  function reset() {
    for (const inst of instances.values()) inst.destroy?.();
    instances.clear();
    navigate({ view: currentRoute?.view || views[0].id });
  }
  return { navigate, restore, reset, getRoute: () => currentRoute };
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
function paneForFileName(fileName) {
  const match = /\.([^.]+)$/i.exec(String(fileName || "").trim());
  const extension = match?.[1]?.toLowerCase();
  if (extension === "html" || extension === "htm") return "html";
  if (extension === "css") return "css";
  if (extension === "js") return "js";
  return "";
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

// ../src/workbench/scriptgen.js
var join = (v) => Array.isArray(v) ? v.join(",") : String(v);
function queryString({ select, expand, filter, orderby, top } = {}) {
  const parts = [];
  if (select) parts.push(`$select=${join(select)}`);
  if (expand) parts.push(`$expand=${join(expand)}`);
  if (filter) parts.push(`$filter=${encodeURIComponent(String(filter))}`);
  if (orderby) parts.push(`$orderby=${join(orderby)}`);
  if (top) parts.push(`$top=${top}`);
  return parts.length ? `?${parts.join("&")}` : "";
}
var PNPJS_ROUTES = [
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/items\((\d+)\)$/i, (id, m) => `sp.web.lists.getById("${id}").items.getById(${m[2]})`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/items$/i, (id) => `sp.web.lists.getById("${id}").items`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/fields$/i, (id) => `sp.web.lists.getById("${id}").fields`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/views$/i, (id) => `sp.web.lists.getById("${id}").views`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/contenttypes$/i, (id) => `sp.web.lists.getById("${id}").contentTypes`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/roleassignments$/i, (id) => `sp.web.lists.getById("${id}").roleAssignments`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)$/i, (id) => `sp.web.lists.getById("${id}")`],
  [/^web\/lists$/i, () => "sp.web.lists"],
  [/^web\/sitegroups\((\d+)\)\/users$/i, (id) => `sp.web.siteGroups.getById(${id}).users`],
  [/^web\/sitegroups$/i, () => "sp.web.siteGroups"],
  [/^web\/roledefinitions$/i, () => "sp.web.roleDefinitions"],
  [/^web\/roleassignments$/i, () => "sp.web.roleAssignments"],
  [/^web\/webs$/i, () => "sp.web.webs"],
  [/^web\/features$/i, () => "sp.web.features"],
  [/^site\/features$/i, () => "sp.site.features"],
  [/^web\/allproperties$/i, () => "sp.web.allProperties"],
  [/^web\/regionalsettings$/i, () => "sp.web.regionalSettings"],
  [/^web\/currentuser$/i, () => "sp.web.currentUser"],
  [/^web$/i, () => "sp.web"],
  [/^site$/i, () => "sp.site"]
];
function toPnpjs2({ path, options = {} }) {
  const clean = String(path).replace(/^\/+/, "");
  const route = PNPJS_ROUTES.find(([re2]) => re2.test(clean));
  if (!route) {
    return [
      "// No direct PnPjs 2 fluent route for this endpoint; raw call:",
      `const data = await sp.web.getParentWeb(); // placeholder \u2014 see REST tab`,
      `// REST: /_api/${clean}${queryString(options)}`
    ].join("\n");
  }
  const [re, root] = route;
  const match = clean.match(re);
  let chain = root(match?.[1], match);
  if (options.select) chain += `
  .select(${join(options.select).split(",").map((s) => `"${s}"`).join(", ")})`;
  if (options.expand) chain += `
  .expand(${join(options.expand).split(",").map((s) => `"${s}"`).join(", ")})`;
  if (options.filter) chain += `
  .filter("${String(options.filter).replaceAll('"', '\\"')}")`;
  if (options.orderby) chain += `
  .orderBy("${join(options.orderby)}")`;
  if (options.top) chain += `
  .top(${options.top})`;
  return [
    "// PnPjs 2.x \u2014 paste into the DCSPad JS pane (pnpjs2 framework enabled)",
    `const data = await ${chain}
  .get();`,
    "console.table(data);"
  ].join("\n");
}
function toRestFetch({ path, options = {} }, webUrl = "") {
  const clean = String(path).replace(/^\/+/, "");
  const base = webUrl ? `"${webUrl}/_api/${clean}${queryString(options)}"` : `\`\${_spPageContextInfo.webAbsoluteUrl}/_api/${clean}${queryString(options)}\``;
  return [
    "// Raw SharePoint REST (GET) \u2014 same-origin cookies authenticate",
    `const response = await fetch(${base}, {`,
    "  credentials: 'same-origin',",
    "  headers: { Accept: 'application/json;odata=nometadata' },",
    "});",
    "const data = await response.json();",
    "console.table(data.value ?? data);"
  ].join("\n");
}
var POWERSHELL_ROUTES = [
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/fields$/i, (id) => `Get-PnPField -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/views$/i, (id) => `Get-PnPView -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)\/contenttypes$/i, (id) => `Get-PnPContentType -List (Get-PnPList -Identity "${id}")`],
  [/^web\/lists\(guid'([0-9a-f-]+)'\)$/i, (id) => `Get-PnPList -Identity "${id}" -Includes HasUniqueRoleAssignments`],
  [/^web\/lists$/i, () => "Get-PnPList -Includes Hidden, ItemCount"],
  [/^web\/sitegroups\((\d+)\)\/users$/i, (id) => `Get-PnPGroupMember -Group (Get-PnPGroup -Identity ${id})`],
  [/^web\/sitegroups$/i, () => "Get-PnPGroup"],
  [/^web\/roledefinitions$/i, () => "Get-PnPRoleDefinition"],
  [/^web\/webs$/i, () => "Get-PnPSubWeb"],
  [/^web\/features$/i, () => "Get-PnPFeature -Scope Web"],
  [/^site\/features$/i, () => "Get-PnPFeature -Scope Site"],
  [/^web\/allproperties$/i, () => "Get-PnPPropertyBag"],
  [/^web$/i, () => "Get-PnPWeb"],
  [/^site$/i, () => "Get-PnPSite"]
];
function toPnpPowerShell({ path, options = {} }, webUrl = "") {
  const clean = String(path).replace(/^\/+/, "");
  const connect = `Connect-PnPOnline -Url "${webUrl || "https://tenant.sharepoint.com/sites/yoursite"}" -Interactive`;
  const route = POWERSHELL_ROUTES.find(([re]) => re.test(clean));
  if (route) {
    const [re, cmd] = route;
    const id = clean.match(re)?.[1];
    return [`# PnP.PowerShell`, connect, cmd(id)].join("\n");
  }
  return [
    "# PnP.PowerShell \u2014 no direct cmdlet; raw REST via Invoke-PnPSPRestMethod",
    connect,
    `Invoke-PnPSPRestMethod -Url "/_api/${clean}${queryString(options).replaceAll('"', '`"')}"`
  ].join("\n");
}

// ../src/workbench/grid.js?v=2
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
  exportName = "",
  descriptor = null
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
  function menuButton(label, title, items) {
    const wrap = el2("span", "wb-menu-wrap");
    const btn = el2("button", "btn btn-xs", label);
    btn.type = "button";
    btn.title = title;
    const menu = el2("div", "wb-menu");
    menu.hidden = true;
    for (const [itemLabel, run] of items) {
      const item2 = el2("button", "wb-menu-item", itemLabel);
      item2.type = "button";
      item2.addEventListener("click", () => {
        menu.hidden = true;
        run(btn);
      });
      menu.append(item2);
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
  if (descriptor) {
    menuButton("Copy as \u25BE", "Copy this query as a runnable script", [
      ["PnPjs 2 (DCSPad pane)", (btn) => copyText(toPnpjs2(descriptor), btn)],
      ["REST fetch", (btn) => copyText(toRestFetch(descriptor, descriptor.webUrl), btn)],
      ["PnP.PowerShell", (btn) => copyText(toPnpPowerShell(descriptor, descriptor.webUrl), btn)]
    ]);
  }
  if (exportName) {
    menuButton("Export \u25BE", "Export the visible rows", [
      ["Download CSV", () => downloadCsv(exportName, visible, columns)],
      ["Download JSON", () => downloadJson(exportName, visible, columns)],
      ["Copy CSV", (btn) => copyText(toCsv(visible, columns), btn)],
      ["Copy JSON", (btn) => copyText(toJson(visible, columns), btn)],
      ["Copy Markdown", (btn) => copyText(toMarkdown(visible, columns), btn)]
    ]);
  }
  const scroller = el2("div", "wb-grid-scroll");
  const table = el2("table", "wb-table");
  const thead = el2("thead");
  const headRow = el2("tr");
  for (const col of columns) {
    const th = el2("th", "", col.label ?? col.key);
    if (col.num) th.classList.add("wb-num");
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
      const selected = Boolean(col && col.key === sortKey);
      th.querySelector(".wb-sort-arrow").textContent = selected ? sortDir === 1 ? " \u25B2" : " \u25BC" : "";
      th.classList.toggle("is-sorted", selected);
      th.setAttribute("aria-sort", selected ? sortDir === 1 ? "ascending" : "descending" : "none");
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
        const td = el2("td", [col.mono ? "wb-mono" : "", col.num ? "wb-num" : ""].filter(Boolean).join(" "));
        const text = displayValue(row, col);
        if (typeof col.render === "function") {
          const node = col.render(cellValue2(row, col), row);
          if (node) td.append(node);
          tr.append(td);
          continue;
        }
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

// ../src/workbench/config-links.js
var LIST_SETTINGS = {
  label: "List settings",
  path: "/_layouts/15/listedit.aspx?List={guid}"
};
var LINK_GROUPS = [
  {
    title: "General",
    links: [
      { label: "Site settings", path: "/_layouts/15/settings.aspx" },
      { label: "Site contents", path: "/_layouts/15/viewlsts.aspx" },
      { label: "Site information", path: "/_layouts/15/prjsetng.aspx", hint: "Classic title/description/logo page" },
      { label: "Site usage", path: "/_layouts/15/siteanalytics.aspx" }
    ]
  },
  {
    title: "Permissions & people",
    links: [
      { label: "Site permissions", path: "/_layouts/15/user.aspx" },
      { label: "People and groups", path: "/_layouts/15/people.aspx?MembershipGroupId=0" },
      { label: "Groups", path: "/_layouts/15/groups.aspx" },
      { label: "Access requests", path: "/Access%20Requests/pendingreq.aspx", hint: "Only exists once an access request has been made" }
    ]
  },
  {
    title: "Recycle bins",
    links: [
      { label: "Recycle bin", path: "/_layouts/15/RecycleBin.aspx" },
      { label: "Site collection recycle bin", path: "/_layouts/15/AdminRecycleBin.aspx", hint: "Site-collection scope \u2014 needs admin rights" },
      { label: "Second-stage recycle bin", path: "/_layouts/15/AdminRecycleBin.aspx?View=5", hint: "Deleted-from-end-user-bin view; needs admin rights" }
    ]
  },
  {
    title: "Galleries",
    links: [
      { label: "Site columns", path: "/_layouts/15/mngfield.aspx" },
      { label: "Site content types", path: "/_layouts/15/mngctype.aspx" },
      { label: "Themes gallery", path: "/_catalogs/theme/Forms/AllItems.aspx" }
    ]
  },
  {
    title: "Search",
    links: [
      { label: "Search settings", path: "/_layouts/15/enhancedSearch.aspx?level=site", hint: "Verify level param on your tenant" },
      { label: "Search schema (managed properties)", path: "/_layouts/15/listmanagedproperties.aspx?level=site", hint: "Verify level param on your tenant" },
      { label: "Result sources", path: "/_layouts/15/manageresultsources.aspx?level=site", hint: "Verify level param on your tenant" },
      { label: "Query rules", path: "/_layouts/15/listqueryrules.aspx?level=site", hint: "Verify level param on your tenant" },
      { label: "Searchable columns", path: "/_layouts/15/NoCrawlSettings.aspx" }
    ]
  }
];
function linkUrl(webUrl, link, params = {}) {
  let path = String(link?.path || "");
  for (const [key2, value] of Object.entries(params)) {
    path = path.replaceAll(`{${key2}}`, encodeURIComponent(String(value)));
  }
  return `${String(webUrl || "").replace(/\/+$/, "")}${path}`;
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
      return renderExpandable(node, `Array(${node.n})`, node.items.map((item2, i) => [String(i), item2]), opts);
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
  const rows = dataNode.t === "arr" ? dataNode.items.map((item2, i) => [String(i), item2]) : dataNode.keys;
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
    arrNode.items.forEach((item2, i) => {
      const row = el3("div");
      row.append(el3("span", "tree-key dim-key", `${i}: `));
      row.append(enhance(item2) ?? renderValue(item2, { dimKeys: NOISE_KEYS }));
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
    items: arrNode.items.map((item2) => item2.t === "obj" ? { ...item2, keys: item2.keys.filter(([k]) => !NOISE_KEYS.has(k)) } : item2)
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
      { key: "LastItemModifiedDate", label: "Modified", format: fmtDate },
      // Appended last: tests address earlier columns positionally.
      {
        key: "Settings",
        label: "",
        value: (row) => row.Id,
        format: () => "",
        // keep filter/sort/export free of the glyph
        render: (id) => {
          const a = document.createElement("a");
          a.className = "wb-cell-link";
          a.href = linkUrl(client2.webUrl(), LIST_SETTINGS, { guid: id });
          a.target = "_blank";
          a.rel = "noopener";
          a.title = "Open list settings in a new tab";
          a.textContent = "\u2699";
          a.addEventListener("click", (e) => e.stopPropagation());
          return a;
        }
      }
    ],
    onOpen: (row) => navigate({ view: "lists", listId: row.Id, listTitle: row.Title }),
    emptyText: "No lists in this web.",
    filterPlaceholder: "Filter lists\u2026",
    exportName: "sp-lists",
    descriptor: {
      path: "web/lists",
      options: { select: LIST_SELECT, expand: "RootFolder", orderby: "Title", top: 5e3 },
      webUrl: client2.webUrl()
    }
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
        query: { path: guidPath(listId, "/fields"), options: { select: FIELD_SELECT } }
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
        query: { path: guidPath(listId, "/views"), options: { select: VIEW_SELECT } }
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
        query: { path: guidPath(listId, "/contenttypes"), options: { select: CT_SELECT } }
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
        query: {
          path: guidPath(listId, "/roleassignments"),
          options: {
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
          }
        }
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
    const settingsLink = el4("a", "btn btn-xs wb-detail-settings", "List settings \u2197");
    settingsLink.href = linkUrl(client2.webUrl(), LIST_SETTINGS, { guid: listId });
    settingsLink.target = "_blank";
    settingsLink.rel = "noopener";
    settingsLink.title = "Open this list\u2019s settings page in a new tab";
    const headRow = el4("div", "wb-detail-head");
    headRow.append(back, title, sub, settingsLink);
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
        exportName: spec.exportName,
        descriptor: { ...spec.query, webUrl: client2.webUrl() }
      });
      wrap.append(tabGrid.el);
      tabGrid.setLoading(`Loading ${tab.label.toLowerCase()}\u2026`);
      cached2(listId, tab.id, () => client2.getAll(spec.query.path, spec.query.options)).then(({ items, partial }) => tabGrid.setRows(items, { partial })).catch((err) => tabGrid.setError(err));
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
  function load2(route) {
    if (route?.listId) {
      showDetail(route);
    } else {
      detailPane.hidden = true;
      gridPane.hidden = false;
      loadLists();
    }
  }
  return { el: root, load: load2, grid };
}

// ../src/sp-files.js
var DIGEST_SAFETY_MS = 6e4;
var FILE_METADATA_SPECS = Object.freeze([
  { key: "title", label: "Title", internalName: "Title", types: ["Text"] },
  { key: "description", label: "Description", internalName: "_ExtendedDescription", types: ["Note", "Text"] },
  { key: "docVersion", label: "DocVersion", internalName: "DocVersion", types: ["Text"] }
]);
function normalizedPath(value) {
  let path = String(value || "").trim().replaceAll("\\", "/");
  if (!path.startsWith("/")) path = `/${path}`;
  path = path.replace(/\/{2,}/g, "/");
  if (path.length > 1) path = path.replace(/\/+$/, "");
  return path;
}
function pathFromWebUrl(webUrl) {
  try {
    return normalizedPath(decodeURIComponent(new URL(webUrl).pathname));
  } catch {
    return "/";
  }
}
function browserTypeForFileName(fileName) {
  const name = String(fileName || "");
  if (/\.html?$/i.test(name)) return "html";
  if (/\.(?:md|markdown)$/i.test(name)) return "markdown";
  if (/\.css$/i.test(name)) return "css";
  if (/\.js$/i.test(name)) return "javascript";
  if (/\.json$/i.test(name)) return "json";
  if (/\.csv$/i.test(name)) return "csv";
  if (/\.txt$/i.test(name)) return "text";
  return "";
}
function createSpFilesClient({
  fetchImpl = (...args) => fetch(...args),
  getContext = getSpContext
} = {}) {
  const digestCache = /* @__PURE__ */ new Map();
  function context({ refresh = false } = {}) {
    const ctx2 = getContext({ refresh });
    if (!ctx2?.live || !ctx2.pageContext?.webAbsoluteUrl) {
      throw new SpFileError(
        "SharePoint file transfer requires an SP: Live context.",
        { code: "not-live" }
      );
    }
    return ctx2;
  }
  function webInfo(targetWebUrl = "") {
    const ctx2 = context({ refresh: true });
    const hostWebUrl = ctx2.pageContext.webAbsoluteUrl.replace(/\/+$/, "");
    let webUrl = hostWebUrl;
    if (targetWebUrl) {
      try {
        const candidate = new URL(String(targetWebUrl).trim(), hostWebUrl);
        if (!/^https?:$/.test(candidate.protocol) || candidate.origin !== new URL(hostWebUrl).origin) {
          throw new Error("origin");
        }
        candidate.hash = "";
        candidate.search = "";
        webUrl = candidate.href.replace(/\/+$/, "");
      } catch {
        throw new SpFileError(
          "Enter a SharePoint site URL on this tenant, such as /sites/ProjectName.",
          { code: "invalid-web-url" }
        );
      }
    }
    const rootPath = normalizedPath(
      webUrl === hostWebUrl && ctx2.pageContext.webServerRelativeUrl ? ctx2.pageContext.webServerRelativeUrl : pathFromWebUrl(webUrl)
    );
    return { ctx: ctx2, webUrl, rootPath, hostWebUrl };
  }
  function checkedPath(path, rootPath) {
    const normalized = normalizedPath(path || rootPath);
    if (rootPath !== "/" && normalized !== rootPath && !normalized.startsWith(`${rootPath}/`)) {
      throw new SpFileError(
        "That path is outside the current SharePoint web.",
        { code: "outside-web" }
      );
    }
    return normalized;
  }
  async function request(url, options = {}) {
    try {
      return await fetchImpl(url, {
        credentials: "same-origin",
        ...options
      });
    } catch (cause) {
      throw new SpFileError(
        `Could not reach SharePoint (${cause.message || cause}).`,
        { code: "network", cause }
      );
    }
  }
  async function fetchContextInfo(targetWebUrl = "") {
    const requested = webInfo(targetWebUrl);
    const { webUrl } = requested;
    const response = await request(`${webUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: ACCEPT_JSON }
    });
    await requireOk(response, "Could not obtain SharePoint request context", "context");
    const info = unwrapJson(await response.json()) || {};
    const value = info.FormDigestValue || info.formDigestValue;
    if (!value) {
      throw new SpFileError(
        "SharePoint contextinfo did not return a request digest.",
        { code: "context" }
      );
    }
    const timeoutSeconds = Number(info.FormDigestTimeoutSeconds || info.formDigestTimeoutSeconds) || 1800;
    const canonicalWebUrl = webInfo(
      info.WebFullUrl || info.webFullUrl || webUrl
    ).webUrl;
    const cached2 = {
      value,
      expiresAt: Date.now() + timeoutSeconds * 1e3,
      webFullUrl: canonicalWebUrl,
      siteFullUrl: info.SiteFullUrl || info.siteFullUrl || ""
    };
    digestCache.set(webUrl.toLowerCase(), cached2);
    digestCache.set(canonicalWebUrl.toLowerCase(), cached2);
    return {
      ...cached2,
      webUrl: canonicalWebUrl,
      rootPath: pathFromWebUrl(canonicalWebUrl)
    };
  }
  async function connectWeb(targetWebUrl = "") {
    const info = await fetchContextInfo(targetWebUrl);
    return {
      webUrl: info.webUrl,
      rootPath: info.rootPath,
      siteFullUrl: info.siteFullUrl
    };
  }
  async function getDigest2({ force = false, webUrl: targetWebUrl = "" } = {}) {
    const target = webInfo(targetWebUrl);
    const cacheKey = target.webUrl.toLowerCase();
    const cached2 = digestCache.get(cacheKey);
    if (!force && cached2?.expiresAt - DIGEST_SAFETY_MS > Date.now()) {
      return cached2.value;
    }
    if (!force && !cached2 && target.webUrl === target.hostWebUrl) {
      const ctx2 = context({ refresh: true });
      const value = ctx2.pageContext.formDigestValue;
      const timeoutSeconds = Number(ctx2.pageContext.formDigestTimeoutSeconds) || 0;
      if (value && !ctx2.pageContext.isDcsPadMock && timeoutSeconds > 0) {
        const pageDigest = {
          value,
          expiresAt: (ctx2.capturedAt || Date.now()) + timeoutSeconds * 1e3,
          webFullUrl: ctx2.pageContext.webAbsoluteUrl,
          siteFullUrl: ctx2.pageContext.siteAbsoluteUrl || ""
        };
        digestCache.set(cacheKey, pageDigest);
        if (pageDigest.expiresAt - DIGEST_SAFETY_MS > Date.now()) return value;
      }
    }
    return (await fetchContextInfo(target.webUrl)).value;
  }
  async function listFolder(serverRelativePath, { webUrl: targetWebUrl = "", purpose = "code" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')?$select=Name,ServerRelativeUrl,Folders/Name,Folders/ServerRelativeUrl,Files/Name,Files/ServerRelativeUrl,Files/Length,Files/TimeLastModified&$expand=Folders,Files`;
    const response = await request(endpoint, {
      headers: { Accept: ACCEPT_JSON }
    });
    await requireOk(response, "Could not list the SharePoint folder", "list");
    const data = unwrapJson(await response.json()) || {};
    const folders = resultArray(data.Folders).map((item2) => ({
      kind: "folder",
      name: String(item2.Name || ""),
      serverRelativeUrl: checkedPath(item2.ServerRelativeUrl, rootPath)
    })).filter((item2) => item2.name).sort((a, b) => a.name.localeCompare(b.name, void 0, { sensitivity: "base" }));
    const files = resultArray(data.Files).map((item2) => ({
      kind: "file",
      name: String(item2.Name || ""),
      pane: paneForFileName(item2.Name),
      browserType: browserTypeForFileName(item2.Name),
      serverRelativeUrl: checkedPath(item2.ServerRelativeUrl, rootPath),
      length: Number(item2.Length) || 0,
      modified: item2.TimeLastModified || ""
    })).filter((item2) => item2.name && (purpose === "browser" ? item2.browserType : item2.pane)).sort((a, b) => a.name.localeCompare(b.name, void 0, { sensitivity: "base" }));
    return {
      path: checkedPath(data.ServerRelativeUrl || path, rootPath),
      rootPath,
      folders,
      files
    };
  }
  async function readTextFile(serverRelativePath, { webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const pane = paneForFileName(path);
    if (!pane) {
      throw new SpFileError(
        "Only HTML, CSS, and JavaScript files can be imported.",
        { code: "unsupported-file" }
      );
    }
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')/$value`;
    const response = await request(endpoint);
    await requireOk(response, "Could not download the SharePoint file", "read");
    const length = Number(response.headers.get("content-length")) || 0;
    if (length > MAX_IMPORT_BYTES) {
      throw new SpFileError(
        "The selected SharePoint file is larger than the 5 MB import limit.",
        { code: "too-large" }
      );
    }
    const text = await response.text();
    if (new Blob([text]).size > MAX_IMPORT_BYTES) {
      throw new SpFileError(
        "The selected SharePoint file is larger than the 5 MB import limit.",
        { code: "too-large" }
      );
    }
    return {
      fileName: path.slice(path.lastIndexOf("/") + 1),
      pane,
      text,
      serverRelativeUrl: path
    };
  }
  async function inspectFileMetadata(folderPath, { filePath = "", webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const folder = checkedPath(folderPath, rootPath);
    const libraryEndpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(folder)}')?$select=ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList`;
    const libraryResponse = await request(libraryEndpoint, {
      headers: { Accept: ACCEPT_JSON }
    });
    const libraryData = libraryResponse.ok ? unwrapJson(await libraryResponse.json()) || {} : {};
    let libraryId = String(
      libraryData.ListItemAllFields?.ParentList?.Id || libraryData.ListItemAllFields?.ParentList?.ID || ""
    ).replace(/[{}]/g, "").trim();
    if (!/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(libraryId)) {
      const rootLibraryEndpoint = `${webUrl}/_api/web/GetList(@listUrl)?@listUrl='${odataPathLiteral(folder)}'&$select=Id`;
      const rootLibraryResponse = await request(rootLibraryEndpoint, {
        headers: { Accept: ACCEPT_JSON }
      });
      await requireOk(
        rootLibraryResponse,
        "Could not resolve the destination SharePoint library",
        "metadata-library"
      );
      const rootLibraryData = unwrapJson(await rootLibraryResponse.json()) || {};
      libraryId = String(rootLibraryData.Id || rootLibraryData.ID || "").replace(/[{}]/g, "").trim();
    }
    if (!/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(libraryId)) {
      throw new SpFileError(
        "SharePoint did not identify the destination document library.",
        { code: "metadata-library" }
      );
    }
    const fieldsEndpoint = `${webUrl}/_api/web/lists(guid'${libraryId}')/Fields?$select=InternalName,EntityPropertyName,Title,TypeAsString,ReadOnlyField,Hidden`;
    const fieldsResponse = await request(fieldsEndpoint, {
      headers: { Accept: ACCEPT_JSON }
    });
    await requireOk(
      fieldsResponse,
      "Could not inspect the destination library metadata fields",
      "metadata-fields"
    );
    const fieldsData = unwrapJson(await fieldsResponse.json()) || {};
    const libraryFields = resultArray(fieldsData.value || fieldsData);
    const fields = {};
    for (const spec of FILE_METADATA_SPECS) {
      const match = libraryFields.find((field2) => String(field2.InternalName || "").toLowerCase() === spec.internalName.toLowerCase());
      let reason = "";
      if (!match) reason = `${spec.internalName} is not available in this library.`;
      else if (match.ReadOnlyField) reason = `${spec.internalName} is read-only.`;
      else if (match.Hidden) reason = `${spec.internalName} is hidden in this library.`;
      else if (!spec.types.includes(String(match.TypeAsString || ""))) {
        reason = `${spec.internalName} is not a supported text field.`;
      }
      fields[spec.key] = {
        label: spec.label,
        internalName: match?.InternalName || spec.internalName,
        entityPropertyName: match?.EntityPropertyName || match?.InternalName || spec.internalName,
        available: !reason,
        reason,
        value: ""
      };
    }
    if (filePath) {
      const path = checkedPath(filePath, rootPath);
      const selected = Object.values(fields).filter((field2) => field2.available).map((field2) => field2.entityPropertyName);
      if (selected.length) {
        const valuesEndpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')/ListItemAllFields?$select=${selected.map(encodeURIComponent).join(",")}`;
        const valuesResponse = await request(valuesEndpoint, {
          headers: { Accept: ACCEPT_JSON }
        });
        await requireOk(
          valuesResponse,
          "Could not read the destination file metadata",
          "metadata-read"
        );
        const values = unwrapJson(await valuesResponse.json()) || {};
        for (const field2 of Object.values(fields)) {
          if (field2.available) {
            field2.value = String(
              values[field2.entityPropertyName] ?? values[field2.internalName] ?? ""
            );
          }
        }
      }
    }
    return { fields };
  }
  async function writeFileMetadata(serverRelativePath, fields, values, { webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const formValues = Object.entries(fields || {}).filter(([key2, field2]) => field2?.available && Object.hasOwn(values || {}, key2)).map(([key2, field2]) => ({
      FieldName: field2.internalName,
      FieldValue: String(values[key2] ?? "")
    }));
    if (!formValues.length) return { updated: [] };
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')/ListItemAllFields/ValidateUpdateListItem`;
    const update = async (forceDigest) => {
      const digest = await getDigest2({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: "POST",
        headers: {
          Accept: ACCEPT_JSON,
          "Content-Type": "application/json;odata=nometadata",
          "X-RequestDigest": digest
        },
        body: JSON.stringify({
          formValues,
          bNewDocumentUpdate: true
        })
      });
    };
    let response = await update(false);
    if (response.status === 403) response = await update(true);
    await requireOk(response, "Could not save the SharePoint file metadata", "metadata-write");
    const data = unwrapJson(await response.json()) || {};
    const results = resultArray(data.value || data.ValidateUpdateListItem || data);
    const failures = results.filter((result) => result.HasException || String(result.ErrorMessage || "").trim());
    if (failures.length) {
      const detail = failures.map((result) => `${result.FieldName || "Field"}: ${result.ErrorMessage || "SharePoint rejected the value."}`).join(" ");
      throw new SpFileError(
        `SharePoint rejected the file metadata. ${detail}`,
        { code: "metadata-write" }
      );
    }
    return { updated: formValues.map((value) => value.FieldName) };
  }
  async function writeTextFile(folderPath, fileName, text, { overwrite = false, webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const folder = checkedPath(folderPath, rootPath);
    const safeName = String(fileName || "").trim();
    if (!safeName || safeName === "." || safeName === ".." || /[\\/]/.test(safeName)) {
      throw new SpFileError(
        "Enter a file name without folder separators.",
        { code: "invalid-name" }
      );
    }
    const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(folder)}')/Files/AddUsingPath(decodedUrl='${odataPathLiteral(safeName)}',overwrite=${overwrite ? "true" : "false"})`;
    const upload = async (forceDigest) => {
      const digest = await getDigest2({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: "POST",
        headers: {
          Accept: ACCEPT_JSON,
          "Content-Type": "text/plain; charset=utf-8",
          "X-RequestDigest": digest
        },
        body: text
      });
    };
    let response = await upload(false);
    if (response.status === 403) response = await upload(true);
    await requireOk(response, "Could not upload the SharePoint file", "write");
    let result = {};
    try {
      result = unwrapJson(await response.json()) || {};
    } catch {
    }
    return {
      fileName: safeName,
      serverRelativeUrl: result.ServerRelativeUrl || `${folder.replace(/\/$/, "")}/${safeName}`
    };
  }
  return {
    webInfo,
    connectWeb,
    getDigest: getDigest2,
    listFolder,
    readTextFile,
    inspectFileMetadata,
    writeFileMetadata,
    writeTextFile
  };
}
var defaultClient = createSpFilesClient();
var getDigest = (options) => defaultClient.getDigest(options);

// ../src/workbench/sp-write.js
var MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
function defaultMockWriter(url, body, contentType) {
  const writes = globalThis.__DCSPAD_WB_WRITES__ ||= [];
  writes.push({ url, body, contentType });
  const lower = String(url).toLowerCase();
  if (lower.includes("validateupdatelistitem")) {
    let formValues = [];
    try {
      formValues = JSON.parse(body)?.formValues || [];
    } catch {
    }
    return {
      value: formValues.map((fv) => ({
        FieldName: fv.FieldName,
        HasException: false,
        ErrorMessage: null
      }))
    };
  }
  if (lower.includes("addusingpath")) {
    const name = /addusingpath\(decodedurl='([^']*)'/.exec(lower)?.[1] || "file";
    const folder = /getfolderbyserverrelativepath\(decodedurl='([^']*)'/.exec(lower)?.[1] || "";
    return { ServerRelativeUrl: `${decodeURIComponent(folder)}/${decodeURIComponent(name)}` };
  }
  return { ok: true };
}
function createSpWriteClient({
  client: client2,
  // the workbench sp-rest client
  fetchImpl = (...args) => fetch(...args),
  mockWriter = null
} = {}) {
  const isMock = () => !client2.context().live;
  async function post(url, { body, contentType = "application/json;odata=nometadata" } = {}, {
    fallback = "SharePoint write failed",
    code = "write"
  } = {}) {
    if (isMock()) {
      return structuredClone((mockWriter || defaultMockWriter)(url, body, contentType));
    }
    const attempt = async (forceDigest) => {
      const digest = await getDigest({ force: forceDigest, webUrl: client2.webUrl() });
      try {
        return await fetchImpl(url, {
          method: "POST",
          credentials: "same-origin",
          headers: {
            Accept: ACCEPT_JSON,
            "Content-Type": contentType,
            "X-RequestDigest": digest
          },
          body
        });
      } catch (cause) {
        throw new SpFileError(
          `Could not reach SharePoint (${cause.message || cause}).`,
          { code: "network", cause }
        );
      }
    };
    let response = await attempt(false);
    if (response.status === 403) response = await attempt(true);
    await requireOk(response, fallback, code);
    try {
      return unwrapJson(await response.json()) || {};
    } catch {
      return {};
    }
  }
  async function validateUpdateListItem(pathKind, formValues, { newDocumentUpdate = false } = {}) {
    if (!Array.isArray(formValues) || !formValues.length) return { updated: [] };
    const base = `${client2.webUrl()}/_api/web`;
    const endpoint = pathKind.fileServerRelativeUrl ? `${base}/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(pathKind.fileServerRelativeUrl)}')/ListItemAllFields/ValidateUpdateListItem` : `${base}/lists(guid'${pathKind.listId}')/items(${Number(pathKind.itemId)})/ValidateUpdateListItem`;
    const data = await post(endpoint, {
      body: JSON.stringify({ formValues, bNewDocumentUpdate: Boolean(newDocumentUpdate) })
    }, { fallback: "Could not save the item metadata", code: "metadata-write" });
    const results = resultArray(data.value || data.ValidateUpdateListItem || data);
    const failures = results.filter((result) => result.HasException || String(result.ErrorMessage || "").trim());
    if (failures.length) {
      const fieldErrors = {};
      for (const failure of failures) {
        fieldErrors[failure.FieldName || ""] = failure.ErrorMessage || "SharePoint rejected the value.";
      }
      const detail = failures.map((f) => `${f.FieldName || "Field"}: ${f.ErrorMessage || "SharePoint rejected the value."}`).join(" ");
      const err = new SpFileError(
        `SharePoint rejected the metadata. ${detail}`,
        { code: "metadata-write" }
      );
      err.fieldErrors = fieldErrors;
      throw err;
    }
    return { updated: formValues.map((fv) => fv.FieldName) };
  }
  async function uploadFile(folderServerRelativeUrl, fileName, data, { overwrite = false } = {}) {
    const safeName = String(fileName || "").trim();
    if (!safeName || safeName === "." || safeName === ".." || /[\\/]/.test(safeName)) {
      throw new SpFileError(
        "Enter a file name without folder separators.",
        { code: "invalid-name" }
      );
    }
    const size = data?.byteLength ?? data?.size ?? (typeof data === "string" ? data.length : 0);
    if (size > MAX_UPLOAD_BYTES) {
      throw new SpFileError(
        `The file is larger than the ${Math.round(MAX_UPLOAD_BYTES / (1024 * 1024))} MB upload limit.`,
        { code: "too-large" }
      );
    }
    const folder = String(folderServerRelativeUrl || "/").replace(/\/+$/, "") || "/";
    const endpoint = `${client2.webUrl()}/_api/web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(folder)}')/Files/AddUsingPath(decodedUrl='${odataPathLiteral(safeName)}',overwrite=${overwrite ? "true" : "false"})`;
    const result = await post(endpoint, {
      body: data,
      contentType: "application/octet-stream"
    }, { fallback: "Could not upload the file", code: "write" });
    return {
      fileName: safeName,
      serverRelativeUrl: result.ServerRelativeUrl || `${folder === "/" ? "" : folder}/${safeName}`
    };
  }
  async function postJson(path, body = {}, { fallback = "SharePoint write failed", code = "write" } = {}) {
    const url = `${client2.webUrl()}/_api/${String(path).replace(/^\/+/, "")}`;
    return post(url, { body: JSON.stringify(body) }, { fallback, code });
  }
  return { validateUpdateListItem, uploadFile, postJson, isMock };
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
  const spWrite = createSpWriteClient({ client: client2 });
  const head = el5("div", "wb-view-head");
  head.innerHTML = '<h2>Permissions</h2><p class="wb-view-hint">Site groups, membership, role definitions, and who holds what on this web. The inheritance scan is on-demand \u2014 it makes SharePoint evaluate security per list.</p>';
  const headLinks = el5("div", "wb-head-links");
  const permGroup = LINK_GROUPS.find((g) => g.title === "Permissions & people");
  for (const link of (permGroup?.links || []).filter((l) => l.label !== "Access requests")) {
    const a = el5("a", "btn btn-xs wb-head-link", `${link.label} \u2197`);
    a.target = "_blank";
    a.rel = "noopener";
    a.dataset.path = link.path;
    headLinks.append(a);
  }
  head.append(headLinks);
  const tabsBar = el5("div", "wb-tabs");
  const body = el5("div", "wb-tab-body");
  root.append(head, tabsBar, body);
  const panes = /* @__PURE__ */ new Map();
  function groupsPane() {
    const wrap = el5("div", "wb-tab-pane");
    const groupsQuery = {
      path: "web/sitegroups",
      options: { select: ["Id", "Title", "Description", "OwnerTitle", "PrincipalType", "OnlyAllowMembersViewMembership"] }
    };
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
      exportName: "sp-groups",
      descriptor: { ...groupsQuery, webUrl: client2.webUrl() }
    });
    const membersBox = el5("div", "wb-subpanel");
    membersBox.hidden = true;
    const membersTitle = el5("h3", "wb-subpanel-title", "");
    const membersHost = el5("div", "wb-subpanel-body");
    membersBox.append(membersTitle, membersHost);
    wrap.append(grid.el, membersBox);
    grid.setLoading("Loading site groups\u2026");
    client2.getAll(groupsQuery.path, groupsQuery.options).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
    function openMembers(group) {
      membersBox.hidden = false;
      membersTitle.textContent = `Members of ${group.Title}`;
      membersHost.textContent = "";
      const membersQuery = {
        path: `web/sitegroups(${group.Id})/users`,
        options: { select: ["Id", "Title", "LoginName", "Email", "IsSiteAdmin", "PrincipalType"] }
      };
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
        exportName: `members-${group.Id}`,
        descriptor: { ...membersQuery, webUrl: client2.webUrl() }
      });
      membersHost.append(membersGrid.el);
      membersGrid.setLoading("Loading members\u2026");
      client2.getAll(membersQuery.path, membersQuery.options).then(({ items }) => membersGrid.setRows(items)).catch((err) => membersGrid.setError(err));
    }
    return wrap;
  }
  function membersPane() {
    const wrap = el5("div", "wb-tab-pane");
    const notice = el5("div", "wb-consent");
    notice.hidden = true;
    function showNotice(message, { isError = false, confirm = null } = {}) {
      notice.textContent = "";
      notice.hidden = false;
      notice.classList.toggle("wb-consent-error", isError);
      notice.append(el5("span", "wb-consent-text", message));
      if (confirm) {
        const yes = el5("button", "btn btn-xs", confirm.label);
        yes.type = "button";
        yes.addEventListener("click", () => {
          notice.hidden = true;
          confirm.run();
        });
        notice.append(yes);
      }
      const dismiss = el5("button", "btn btn-xs", confirm ? "Cancel" : "Dismiss");
      dismiss.type = "button";
      dismiss.addEventListener("click", () => {
        notice.hidden = true;
      });
      notice.append(dismiss);
    }
    const addBar = el5("div", "wb-members-add");
    const groupSelect = el5("select", "wb-members-group");
    groupSelect.setAttribute("aria-label", "Group to add the user to");
    const loginInput = el5("input", "wb-members-login");
    loginInput.type = "text";
    loginInput.placeholder = "user@tenant.com or i:0#.f|membership|\u2026";
    const addBtn = el5("button", "btn btn-xs", "Add to group");
    addBtn.type = "button";
    addBar.append(el5("span", "wb-qb-label", "Add user"), groupSelect, loginInput, addBtn);
    const grid = createGrid({
      rowKey: "Key",
      columns: [
        { key: "GroupTitle", label: "Group" },
        { key: "GroupId", label: "Group id" },
        { key: "Title", label: "User" },
        { key: "Email", label: "Email", copyable: true },
        { key: "LoginName", label: "Login", mono: true, copyable: true },
        { key: "IsSiteAdmin", label: "Site admin" },
        {
          key: "Remove",
          label: "",
          value: (row) => row.Key,
          format: () => "",
          render: (key2, row) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "wb-cell-link wb-cell-copylink";
            btn.title = `Remove ${row.Title} from ${row.GroupTitle}`;
            btn.textContent = "\u2715";
            btn.addEventListener("click", (e) => {
              e.stopPropagation();
              showNotice(`Remove ${row.Title} from ${row.GroupTitle}?`, {
                isError: false,
                confirm: { label: "Remove", run: () => removeMember(row) }
              });
            });
            return btn;
          }
        }
      ],
      emptyText: "No group members.",
      filterPlaceholder: "Filter members\u2026",
      exportName: "sp-group-members"
    });
    wrap.append(addBar, notice, grid.el);
    let groups = [];
    async function loadMembers() {
      grid.setLoading("Loading group membership\u2026");
      try {
        const { items } = await client2.getAll("web/sitegroups", {
          select: ["Id", "Title"]
        });
        groups = items;
        groupSelect.textContent = "";
        for (const group of groups) {
          const opt = el5("option", "", group.Title);
          opt.value = String(group.Id);
          groupSelect.append(opt);
        }
        const memberLists = await Promise.all(groups.map((group) => client2.getAll(`web/sitegroups(${group.Id})/users`, {
          select: ["Id", "Title", "LoginName", "Email", "IsSiteAdmin"]
        }).then(({ items: users }) => users.map((user2) => ({
          Key: `${group.Id}:${user2.Id}`,
          GroupTitle: group.Title,
          GroupId: group.Id,
          UserId: user2.Id,
          Title: user2.Title,
          Email: user2.Email || "",
          LoginName: user2.LoginName || "",
          IsSiteAdmin: Boolean(user2.IsSiteAdmin)
        })), () => [])));
        const rows = memberLists.flat().sort((a, b) => a.GroupTitle.localeCompare(b.GroupTitle) || a.Title.localeCompare(b.Title));
        grid.setRows(rows);
      } catch (err) {
        grid.setError(err);
      }
    }
    const toLoginName = (input) => {
      const raw = String(input || "").trim();
      if (!raw) return "";
      return raw.includes("|") ? raw : `i:0#.f|membership|${raw}`;
    };
    addBtn.addEventListener("click", async () => {
      const loginName = toLoginName(loginInput.value);
      const groupId = Number(groupSelect.value);
      if (!loginName || !groupId) return;
      addBtn.disabled = true;
      try {
        await spWrite.postJson(`web/sitegroups(${groupId})/users`, { LoginName: loginName }, {
          fallback: "Could not add the user to the group",
          code: "group-add"
        });
        loginInput.value = "";
        showNotice(spWrite.isMock() ? "Added (mock mode \u2014 the fixture roster does not change)." : "User added.");
        await loadMembers();
      } catch (err) {
        showNotice(err?.message || String(err), { isError: true });
      } finally {
        addBtn.disabled = false;
      }
    });
    async function removeMember(row) {
      try {
        await spWrite.postJson(
          `web/sitegroups(${row.GroupId})/users/removebyid(${row.UserId})`,
          {},
          { fallback: "Could not remove the user from the group", code: "group-remove" }
        );
        showNotice(spWrite.isMock() ? "Removed (mock mode \u2014 the fixture roster does not change)." : `Removed ${row.Title} from ${row.GroupTitle}.`);
        await loadMembers();
      } catch (err) {
        showNotice(err?.message || String(err), { isError: true });
      }
    }
    loadMembers();
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
      exportName: "sp-roledefinitions",
      descriptor: {
        path: "web/roledefinitions",
        options: { select: ["Id", "Name", "Description", "RoleTypeKind", "Hidden", "BasePermissions"] },
        webUrl: client2.webUrl()
      }
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
      exportName: "sp-roleassignments",
      descriptor: {
        path: "web/roleassignments",
        options: { expand: ["Member", "RoleDefinitionBindings"] },
        webUrl: client2.webUrl()
      }
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
      exportName: "sp-unique-permissions",
      descriptor: {
        path: "web/lists",
        options: { select: ["Id", "Title", "Hidden", "BaseTemplate", "HasUniqueRoleAssignments"], top: 5e3 },
        webUrl: client2.webUrl()
      }
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
    { id: "members", label: "Members", build: membersPane },
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
  function load2() {
    for (const a of headLinks.querySelectorAll("a")) {
      a.href = linkUrl(client2.webUrl(), { path: a.dataset.path });
    }
    if (!tabsBar.querySelector(".wb-tab.active")) activate(TABS[0]);
  }
  return { el: root, load: load2 };
}

// ../src/workbench/views/site.js
var el6 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var decodeODataKey = (key2) => String(key2).replace(/_x([0-9a-f]{4})_/gi, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
var flatten = (v) => {
  if (v === null || v === void 0) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
};
var entityRows = (entity2) => Object.entries(entity2 || {}).filter(([k]) => !k.startsWith("odata.") && !k.startsWith("@odata") && k !== "__metadata").map(([k, v]) => ({ Property: k, Value: flatten(v) }));
var WEB_SELECT = [
  "Id",
  "Title",
  "Description",
  "Url",
  "ServerRelativeUrl",
  "WebTemplate",
  "Configuration",
  "Created",
  "LastItemModifiedDate",
  "Language",
  "UIVersion",
  "QuickLaunchEnabled",
  "MembersCanShare"
];
var SITE_SELECT = ["Id", "Url", "ServerRelativeUrl", "ReadOnly", "ShareByEmailEnabled"];
function createSiteView({ client: client2 }) {
  const root = el6("section", "wb-view wb-view-site");
  const head = el6("div", "wb-view-head");
  head.innerHTML = '<h2>Site overview</h2><p class="wb-view-hint">Web and site collection properties, features, subwebs, and the property bag.</p>';
  const tabsBar = el6("div", "wb-tabs");
  const body = el6("div", "wb-tab-body");
  root.append(head, tabsBar, body);
  const panes = /* @__PURE__ */ new Map();
  function sheetPane(query, extraSections = []) {
    const wrap = el6("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "Property",
      columns: [
        { key: "Property", label: "Property", mono: true, copyable: true },
        { key: "Value", label: "Value", copyable: true }
      ],
      emptyText: "Nothing returned.",
      filterPlaceholder: "Filter properties\u2026",
      exportName: query.exportName,
      descriptor: { path: query.path, options: query.options, webUrl: client2.webUrl() }
    });
    wrap.append(grid.el);
    grid.setLoading("Loading\u2026");
    client2.get(query.path, query.options).then((entity2) => grid.setRows(entityRows(entity2))).catch((err) => grid.setError(err));
    for (const extra of extraSections) {
      const box = el6("div", "wb-subpanel");
      box.hidden = false;
      box.append(el6("h3", "wb-subpanel-title", extra.title));
      const hostEl = el6("div", "wb-subpanel-body");
      box.append(hostEl);
      const extraGrid = createGrid({
        rowKey: "Property",
        columns: [
          { key: "Property", label: "Property", mono: true, copyable: true },
          { key: "Value", label: "Value", copyable: true }
        ],
        emptyText: "Nothing returned.",
        filterPlaceholder: "Filter\u2026",
        exportName: extra.exportName
      });
      hostEl.append(extraGrid.el);
      extraGrid.setLoading("Loading\u2026");
      client2.get(extra.path, extra.options).then((entity2) => extraGrid.setRows(entityRows(extra.map ? extra.map(entity2) : entity2))).catch((err) => extraGrid.setError(err));
      wrap.append(box);
    }
    return wrap;
  }
  function featuresPane() {
    const wrap = el6("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "DefinitionId",
      columns: [
        { key: "Scope", label: "Scope" },
        { key: "DisplayName", label: "Feature", format: (v) => v || "(no display name)" },
        { key: "DefinitionId", label: "Definition id", mono: true, copyable: true }
      ],
      emptyText: "No activated features.",
      filterPlaceholder: "Filter features\u2026",
      exportName: "sp-features",
      descriptor: {
        path: "web/features",
        options: { select: ["DefinitionId", "DisplayName"] },
        webUrl: client2.webUrl()
      }
    });
    wrap.append(grid.el);
    grid.setLoading("Loading features (site + web scope)\u2026");
    const options = { select: ["DefinitionId", "DisplayName"] };
    Promise.all([
      client2.getAll("site/features", options),
      client2.getAll("web/features", options)
    ]).then(([site, web]) => {
      const rows = [
        ...site.items.map((f) => ({ ...f, Scope: "Site" })),
        ...web.items.map((f) => ({ ...f, Scope: "Web" }))
      ];
      grid.setRows(rows, { partial: site.partial || web.partial });
    }).catch((err) => grid.setError(err));
    return wrap;
  }
  function subwebsPane() {
    const wrap = el6("div", "wb-tab-pane");
    const query = {
      path: "web/webs",
      options: { select: ["Id", "Title", "ServerRelativeUrl", "WebTemplate", "Created", "Language"] }
    };
    const grid = createGrid({
      columns: [
        { key: "Title", label: "Title" },
        { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true },
        { key: "WebTemplate", label: "Template" },
        { key: "Language", label: "Language" },
        { key: "Created", label: "Created", format: (v) => v ? String(v).slice(0, 10) : "" },
        { key: "Id", label: "Id", mono: true, copyable: true }
      ],
      emptyText: "No subwebs.",
      filterPlaceholder: "Filter subwebs\u2026",
      exportName: "sp-subwebs",
      descriptor: { ...query, webUrl: client2.webUrl() }
    });
    wrap.append(grid.el);
    grid.setLoading("Loading subwebs\u2026");
    client2.getAll(query.path, query.options).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
    return wrap;
  }
  function propertyBagPane() {
    const wrap = el6("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "RawKey",
      columns: [
        { key: "Key", label: "Key (decoded)", mono: true },
        { key: "RawKey", label: "Raw key", mono: true, copyable: true },
        { key: "Value", label: "Value", copyable: true }
      ],
      emptyText: "Empty property bag.",
      filterPlaceholder: "Filter keys\u2026",
      exportName: "sp-propertybag",
      descriptor: { path: "web/allproperties", options: {}, webUrl: client2.webUrl() }
    });
    wrap.append(grid.el);
    grid.setLoading("Loading property bag\u2026");
    client2.get("web/allproperties").then((bag) => {
      const rows = Object.entries(bag || {}).filter(([k]) => !k.startsWith("odata.") && !k.startsWith("@odata") && k !== "__metadata").map(([k, v]) => ({ Key: decodeODataKey(k), RawKey: k, Value: flatten(v) }));
      grid.setRows(rows);
    }).catch((err) => grid.setError(err));
    return wrap;
  }
  const TABS = [
    {
      id: "web",
      label: "Web",
      build: () => sheetPane(
        { path: "web", options: { select: WEB_SELECT }, exportName: "sp-web" },
        [
          {
            title: "Regional settings",
            path: "web/regionalsettings",
            options: { expand: "TimeZone" },
            exportName: "sp-regionalsettings"
          },
          {
            title: "Current user",
            path: "web/currentuser",
            options: {},
            exportName: "sp-currentuser"
          }
        ]
      )
    },
    {
      id: "site",
      label: "Site collection",
      build: () => sheetPane({ path: "site", options: { select: SITE_SELECT }, exportName: "sp-site" })
    },
    { id: "features", label: "Features", build: featuresPane },
    { id: "subwebs", label: "Subwebs", build: subwebsPane },
    { id: "propertybag", label: "Property bag", build: propertyBagPane }
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
    const btn = el6("button", "wb-tab", tab.label);
    btn.type = "button";
    btn.dataset.tab = tab.id;
    btn.addEventListener("click", () => activate(tab));
    tabsBar.append(btn);
  }
  function load2() {
    if (!tabsBar.querySelector(".wb-tab.active")) activate(TABS[0]);
  }
  return { el: root, load: load2 };
}

// ../src/workbench/views/site-home.js
var el7 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var fmtDate2 = (v) => v ? String(v).slice(0, 10) : "";
function createSiteHomeView({ client: client2, navigate, inspectSite: inspectSite2 }) {
  const root = el7("section", "wb-view wb-view-sitehome");
  const head = el7("div", "wb-view-head");
  head.innerHTML = '<h2>Site</h2><p class="wb-view-hint">The inspected web at a glance. Full property sheets are under Advanced.</p>';
  const cards = el7("div", "wb-home-cards");
  const webCard = el7("div", "wb-home-card");
  const userCard = el7("div", "wb-home-card");
  cards.append(webCard, userCard);
  const subwebsBox = el7("div", "wb-home-subwebs");
  root.append(head, cards, subwebsBox);
  let loadedForWeb = "";
  function factRow(label, value, { copyFull = "" } = {}) {
    const row = el7("div", "wb-home-fact");
    row.append(el7("span", "wb-home-fact-label", label));
    const v = el7("span", copyFull ? "wb-home-fact-value sp-copy" : "wb-home-fact-value", value || "\u2014");
    if (copyFull) {
      v.title = "Click to copy the full URL";
      v.addEventListener("click", () => copyText(copyFull, v));
    }
    row.append(v);
    return row;
  }
  async function load2() {
    const webUrl = client2.webUrl();
    if (loadedForWeb === webUrl) return;
    loadedForWeb = webUrl;
    webCard.textContent = "";
    webCard.append(el7("h3", "wb-home-card-title", "This web"));
    userCard.textContent = "";
    userCard.append(el7("h3", "wb-home-card-title", "You"));
    subwebsBox.textContent = "";
    try {
      const web = await client2.get("web", {
        select: [
          "Title",
          "Description",
          "Url",
          "ServerRelativeUrl",
          "WebTemplate",
          "Created",
          "LastItemModifiedDate",
          "Language"
        ]
      });
      webCard.append(
        factRow("Title", web.Title),
        factRow("Description", web.Description),
        factRow("URL", web.ServerRelativeUrl || "/", { copyFull: web.Url || webUrl }),
        factRow("Template", web.WebTemplate),
        factRow("Created", fmtDate2(web.Created)),
        factRow("Last modified", fmtDate2(web.LastItemModifiedDate))
      );
    } catch (err) {
      webCard.append(el7("div", "wb-grid-status wb-error", err?.message || String(err)));
    }
    try {
      const user2 = await client2.get("web/currentuser", {
        select: ["Title", "Email", "LoginName", "IsSiteAdmin"]
      });
      userCard.append(
        factRow("Name", user2.Title),
        factRow("Email", user2.Email),
        factRow("Login", user2.LoginName)
      );
      const roleRow = el7("div", "wb-home-fact");
      roleRow.append(el7("span", "wb-home-fact-label", "Role"));
      roleRow.append(el7(
        "span",
        user2.IsSiteAdmin ? "wb-role-chip wb-role-admin" : "wb-role-chip wb-role-user",
        user2.IsSiteAdmin ? "Site admin" : "Site user"
      ));
      userCard.append(roleRow);
    } catch (err) {
      userCard.append(el7("div", "wb-grid-status wb-error", err?.message || String(err)));
    }
    const grid = createGrid({
      columns: [
        { key: "Title", label: "Subweb" },
        { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true },
        { key: "Created", label: "Created", format: fmtDate2 },
        {
          key: "Inspect",
          label: "",
          value: (row) => row.ServerRelativeUrl,
          format: () => "",
          render: (url) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "btn btn-xs";
            btn.textContent = "Inspect";
            btn.addEventListener("click", (e) => {
              e.stopPropagation();
              inspectSite2?.(url);
            });
            return btn;
          }
        }
      ],
      emptyText: "No subwebs under this web.",
      filterPlaceholder: "Filter subwebs\u2026",
      exportName: "sp-subwebs",
      descriptor: {
        path: "web/webs",
        options: { select: ["Id", "Title", "ServerRelativeUrl", "Created", "WebTemplate"] },
        webUrl: client2.webUrl()
      }
    });
    subwebsBox.append(el7("h3", "wb-home-card-title", "Subwebs"), grid.el);
    grid.setLoading("Loading subwebs\u2026");
    client2.getAll("web/webs", {
      select: ["Id", "Title", "ServerRelativeUrl", "Created", "WebTemplate"]
    }).then(({ items, partial }) => grid.setRows(items, { partial })).catch((err) => grid.setError(err));
  }
  function loadRoute() {
    if (loadedForWeb && loadedForWeb !== client2.webUrl()) loadedForWeb = "";
    load2();
  }
  return { el: root, load: loadRoute };
}

// ../src/workbench/views/links.js
var el8 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function createLinksView({ client: client2 }) {
  const root = el8("section", "wb-view wb-view-links");
  const head = el8("div", "wb-view-head");
  head.innerHTML = '<h2>Panels</h2><p class="wb-view-hint">Quick jumps to the SharePoint configuration panels you actually reach for. Links open in a new tab; hover for the underlying page.</p>';
  const body = el8("div", "wb-links");
  root.append(head, body);
  function load2() {
    const webUrl = client2.webUrl();
    body.textContent = "";
    for (const group of LINK_GROUPS) {
      const card = el8("div", "wb-linkgroup");
      card.append(el8("h3", "", group.title));
      for (const link of group.links) {
        const row = el8("a", "wb-link");
        row.href = linkUrl(webUrl, link);
        row.target = "_blank";
        row.rel = "noopener";
        row.append(el8("span", "wb-link-label", link.label));
        row.append(el8("span", "wb-link-go", "\u2197"));
        row.title = link.hint ? `${link.path}
${link.hint}` : link.path;
        card.append(row);
      }
      body.append(card);
    }
  }
  return { el: root, load: load2 };
}

// ../src/workbench/views/query.js?v=2
var QUERY_KEY = "dcspad.workbench.query";
var DEFAULT_TOP = 100;
var MAX_TOP = 5e3;
var FIELD_SELECT2 = [
  "Id",
  "Title",
  "InternalName",
  "TypeAsString",
  "FieldTypeKind",
  "Hidden",
  "ReadOnlyField",
  "Choices"
];
var NUMERIC_TYPES = /* @__PURE__ */ new Set(["Number", "Currency", "Counter", "Integer"]);
var EXPANDABLE_TYPES = /* @__PURE__ */ new Set(["User", "UserMulti", "Lookup", "LookupMulti"]);
var OPERATORS = [
  ["eq", "="],
  ["ne", "\u2260"],
  ["gt", ">"],
  ["ge", "\u2265"],
  ["lt", "<"],
  ["le", "\u2264"],
  ["startswith", "starts with"],
  ["substringof", "contains"]
];
var el9 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function filterClause({ field: field2, type, op, value }) {
  const name = String(field2 || "").trim();
  if (!name) return "";
  const raw = String(value ?? "").trim();
  let literal;
  if (NUMERIC_TYPES.has(type)) literal = raw === "" ? "0" : String(Number(raw.replace(",", ".")));
  else if (type === "Boolean") literal = /^(1|true|yes)$/i.test(raw) ? "1" : "0";
  else if (type === "DateTime") {
    const d = new Date(raw);
    literal = `datetime'${Number.isNaN(d.getTime()) ? raw : d.toISOString()}'`;
  } else literal = `'${raw.replaceAll("'", "''")}'`;
  if (op === "startswith") return `startswith(${name},${literal})`;
  if (op === "substringof") return `substringof(${literal},${name})`;
  return `${name} ${op} ${literal}`;
}
function composeFilter(rows) {
  let out = "";
  for (const row of rows) {
    const clause = filterClause(row);
    if (!clause) continue;
    out = out ? `${out} ${row.join === "or" ? "or" : "and"} ${clause}` : clause;
  }
  return out;
}
function descriptorToRaw(descriptor) {
  const parts = [];
  const join2 = (v) => Array.isArray(v) ? v.join(",") : String(v);
  const o = descriptor.options || {};
  if (o.select) parts.push(`$select=${join2(o.select)}`);
  if (o.expand) parts.push(`$expand=${join2(o.expand)}`);
  if (o.filter) parts.push(`$filter=${o.filter}`);
  if (o.orderby) parts.push(`$orderby=${join2(o.orderby)}`);
  if (o.top) parts.push(`$top=${o.top}`);
  return `${descriptor.path}${parts.length ? `?${parts.join("&")}` : ""}`;
}
function rawToDescriptor(raw) {
  const s = String(raw || "").trim().replace(/^\/+/, "").replace(/^_api\//, "");
  if (!s) return null;
  const q = s.indexOf("?");
  if (q === -1) return { path: s, options: {} };
  const path = s.slice(0, q);
  const options = {};
  for (const pair of s.slice(q + 1).split("&")) {
    const eq = pair.indexOf("=");
    if (eq === -1) return null;
    const key2 = pair.slice(0, eq);
    const value = pair.slice(eq + 1);
    if (key2 === "$select") options.select = value.split(",");
    else if (key2 === "$expand") options.expand = value.split(",");
    else if (key2 === "$filter") {
      try {
        options.filter = decodeURIComponent(value);
      } catch {
        return null;
      }
    } else if (key2 === "$orderby") options.orderby = value;
    else if (key2 === "$top") options.top = Number(value) || void 0;
    else return null;
  }
  return { path, options };
}
function columnsForSelect(select) {
  return (select || []).map((entry) => {
    const path = String(entry).split("/");
    return {
      key: entry,
      label: entry,
      value: (row) => {
        let v = row;
        for (const seg of path) v = v?.[seg];
        if (v === null || v === void 0) return "";
        if (typeof v === "object") {
          const arr = Array.isArray(v) ? v : v.results;
          if (Array.isArray(arr)) return arr.map((x) => typeof x === "object" ? JSON.stringify(x) : x).join(", ");
          return JSON.stringify(v);
        }
        return v;
      }
    };
  });
}
function readSaved(webUrl) {
  try {
    const all = JSON.parse(sessionStorage.getItem(QUERY_KEY) || "{}");
    return all[webUrl.toLowerCase()] || null;
  } catch {
    return null;
  }
}
function writeSaved(webUrl, state2) {
  try {
    const all = JSON.parse(sessionStorage.getItem(QUERY_KEY) || "{}");
    all[webUrl.toLowerCase()] = state2;
    sessionStorage.setItem(QUERY_KEY, JSON.stringify(all));
  } catch {
  }
}
function createQueryView({ client: client2 }) {
  const root = el9("section", "wb-view wb-view-query");
  const head = el9("div", "wb-view-head");
  head.innerHTML = '<h2>Query builder</h2><p class="wb-view-hint">Compose an OData query against any list \u2014 or any /_api endpoint \u2014 and run it. \u201CCopy as\u201D turns the query into a script.</p>';
  const composer = el9("div", "wb-qb");
  const targetRow = el9("div", "wb-qb-row");
  const listSelect = el9("select", "wb-qb-list");
  listSelect.setAttribute("aria-label", "Query target list");
  const endpointInput = el9("input", "wb-qb-endpoint");
  endpointInput.type = "text";
  endpointInput.placeholder = "web/currentuser \u2014 path after /_api/";
  endpointInput.hidden = true;
  targetRow.append(el9("span", "wb-qb-label", "Target"), listSelect, endpointInput);
  const fieldsBox = el9("div", "wb-qb-fields");
  const fieldsList = el9("div", "wb-qb-fieldlist");
  fieldsBox.append(el9("span", "wb-qb-label", "$select"), fieldsList);
  const filtersBox = el9("div", "wb-qb-filters");
  const filterRows = el9("div", "wb-qb-filterrows");
  const addFilter = el9("button", "btn btn-xs", "+ Filter");
  addFilter.type = "button";
  filtersBox.append(el9("span", "wb-qb-label", "$filter"), filterRows, addFilter);
  const optionsRow = el9("div", "wb-qb-row");
  const orderSelect = el9("select", "wb-qb-order");
  const orderDir = el9("select", "wb-qb-orderdir");
  for (const [v, label] of [["asc", "ascending"], ["desc", "descending"]]) {
    const opt = el9("option", "", label);
    opt.value = v;
    orderDir.append(opt);
  }
  const topInput = el9("input", "wb-qb-top");
  topInput.type = "number";
  topInput.min = "1";
  topInput.max = String(MAX_TOP);
  topInput.value = String(DEFAULT_TOP);
  const expandInput = el9("input", "wb-qb-expand");
  expandInput.type = "text";
  expandInput.placeholder = "extra $expand (comma-separated)";
  optionsRow.append(
    el9("span", "wb-qb-label", "$orderby"),
    orderSelect,
    orderDir,
    el9("span", "wb-qb-label", "$top"),
    topInput,
    el9("span", "wb-qb-label", "$expand"),
    expandInput
  );
  const rawRow = el9("div", "wb-qb-rawrow");
  const rawArea = el9("textarea", "wb-qb-raw");
  rawArea.spellcheck = false;
  rawArea.setAttribute("aria-label", "Raw query");
  const rawNote = el9("span", "wb-qb-rawnote", "Editing the raw query overrides the builder.");
  rawNote.hidden = true;
  const runBtn = el9("button", "btn btn-xs wb-qb-run wb-primary", "Run \u25B6");
  runBtn.type = "button";
  const backToBuilder = el9("button", "btn btn-xs", "Back to builder");
  backToBuilder.type = "button";
  backToBuilder.hidden = true;
  rawRow.append(rawArea, rawNote, runBtn, backToBuilder);
  composer.append(targetRow, fieldsBox, filtersBox, optionsRow, rawRow);
  const results = el9("div", "wb-qb-results");
  root.append(head, composer, results);
  let lists = [];
  let fields = [];
  let rawMode = false;
  let loadedForWeb = "";
  const fieldByName = (name) => fields.find((f) => f.InternalName === name);
  const guidPath3 = (listId) => `web/lists(guid'${listId}')/items`;
  function pickedListId() {
    return listSelect.value === "::endpoint" ? "" : listSelect.value;
  }
  function addFilterRow(saved = {}) {
    const row = el9("div", "wb-qb-filterrow");
    const join2 = el9("select", "wb-qb-join");
    for (const [v, label] of [["and", "AND"], ["or", "OR"]]) {
      const opt = el9("option", "", label);
      opt.value = v;
      join2.append(opt);
    }
    join2.value = saved.join || "and";
    if (!filterRows.childElementCount) join2.classList.add("wb-qb-join-first");
    const fieldSel = el9("select", "wb-qb-field");
    for (const f of fields) {
      const opt = el9("option", "", f.InternalName);
      opt.value = f.InternalName;
      fieldSel.append(opt);
    }
    if (saved.field) fieldSel.value = saved.field;
    const opSel = el9("select", "wb-qb-op");
    for (const [v, label] of OPERATORS) {
      const opt = el9("option", "", label);
      opt.value = v;
      opSel.append(opt);
    }
    if (saved.op) opSel.value = saved.op;
    const valueInput = el9("input", "wb-qb-value");
    valueInput.type = "text";
    valueInput.placeholder = "value";
    valueInput.value = saved.value || "";
    const remove = el9("button", "btn btn-xs", "\xD7");
    remove.type = "button";
    remove.title = "Remove this filter";
    remove.addEventListener("click", () => {
      row.remove();
      onBuilderChange();
    });
    for (const control of [join2, fieldSel, opSel]) {
      control.addEventListener("change", onBuilderChange);
    }
    valueInput.addEventListener("input", onBuilderChange);
    row.append(join2, fieldSel, opSel, valueInput, remove);
    filterRows.append(row);
  }
  function readFilterRows() {
    return [...filterRows.children].map((row) => ({
      join: row.querySelector(".wb-qb-join").value,
      field: row.querySelector(".wb-qb-field").value,
      type: fieldByName(row.querySelector(".wb-qb-field").value)?.TypeAsString || "Text",
      op: row.querySelector(".wb-qb-op").value,
      value: row.querySelector(".wb-qb-value").value
    }));
  }
  function selectedFields() {
    return [...fieldsList.querySelectorAll("input:checked")].map((box) => box.value);
  }
  function composeDescriptor() {
    if (rawMode) {
      const parsed = rawToDescriptor(rawArea.value);
      return parsed ? { ...parsed, webUrl: client2.webUrl() } : null;
    }
    const listId = pickedListId();
    const select = selectedFields();
    const expand = new Set(
      expandInput.value.split(",").map((s) => s.trim()).filter(Boolean)
    );
    for (const entry of select) {
      if (entry.includes("/")) expand.add(entry.split("/")[0]);
    }
    const options = {};
    if (select.length) options.select = select;
    if (expand.size) options.expand = [...expand];
    const filter = composeFilter(readFilterRows());
    if (filter) options.filter = filter;
    if (orderSelect.value) {
      options.orderby = orderDir.value === "desc" ? `${orderSelect.value} desc` : orderSelect.value;
    }
    const top = Math.min(Math.max(Number(topInput.value) || DEFAULT_TOP, 1), MAX_TOP);
    options.top = top;
    const path = listId ? guidPath3(listId) : String(endpointInput.value || "").trim().replace(/^\/+/, "");
    if (!path) return null;
    return { path, options, webUrl: client2.webUrl() };
  }
  function onBuilderChange() {
    if (rawMode) return;
    const descriptor = composeDescriptor();
    rawArea.value = descriptor ? descriptorToRaw(descriptor) : "";
  }
  function enterRawMode() {
    if (rawMode) return;
    rawMode = true;
    composer.classList.add("wb-qb-rawmode");
    rawNote.hidden = false;
    backToBuilder.hidden = false;
  }
  function leaveRawMode() {
    rawMode = false;
    composer.classList.remove("wb-qb-rawmode");
    rawNote.hidden = true;
    backToBuilder.hidden = true;
    onBuilderChange();
  }
  rawArea.addEventListener("input", enterRawMode);
  backToBuilder.addEventListener("click", leaveRawMode);
  function renderFieldList(savedSelect = null) {
    fieldsList.textContent = "";
    orderSelect.textContent = "";
    const blank = el9("option", "", "(no ordering)");
    blank.value = "";
    orderSelect.append(blank);
    const wanted = new Set(savedSelect || ["Id", "Title"]);
    for (const f of fields) {
      const entry = EXPANDABLE_TYPES.has(f.TypeAsString) ? `${f.InternalName}/Title` : f.InternalName;
      const label = el9("label", "wb-qb-fieldopt");
      const box = el9("input");
      box.type = "checkbox";
      box.value = entry;
      box.checked = wanted.has(entry);
      box.addEventListener("change", onBuilderChange);
      label.append(box, document.createTextNode(entry));
      label.append(el9("span", "wb-qb-fieldtype", f.TypeAsString));
      fieldsList.append(label);
      const opt = el9("option", "", f.InternalName);
      opt.value = f.InternalName;
      orderSelect.append(opt);
    }
  }
  async function loadFieldsForList(listId, saved = null) {
    fields = [];
    fieldsList.textContent = "";
    filterRows.textContent = "";
    if (!listId) {
      renderFieldList();
      onBuilderChange();
      return;
    }
    fieldsList.append(el9("div", "wb-qb-loading", "Loading fields\u2026"));
    try {
      const { items } = await client2.getAll(`web/lists(guid'${listId}')/fields`, {
        select: FIELD_SELECT2
      });
      fields = items.filter((f) => !f.Hidden);
      renderFieldList(saved?.select);
      for (const savedRow of saved?.filters || []) addFilterRow(savedRow);
      if (saved?.orderby) orderSelect.value = saved.orderby;
      if (saved?.orderdir) orderDir.value = saved.orderdir;
      onBuilderChange();
    } catch (err) {
      fieldsList.textContent = "";
      fieldsList.append(el9("div", "wb-qb-loading wb-error", err?.message || String(err)));
    }
  }
  function renderListPicker(savedListId = "") {
    listSelect.textContent = "";
    for (const list2 of lists) {
      const opt = el9("option", "", list2.Hidden ? `${list2.Title} (hidden)` : list2.Title);
      opt.value = list2.Id;
      listSelect.append(opt);
    }
    const endpoint = el9("option", "", "\u2014 arbitrary endpoint \u2014");
    endpoint.value = "::endpoint";
    listSelect.append(endpoint);
    if (savedListId && lists.some((l) => l.Id === savedListId)) listSelect.value = savedListId;
    else if (savedListId === "::endpoint") listSelect.value = "::endpoint";
    endpointInput.hidden = listSelect.value !== "::endpoint";
  }
  listSelect.addEventListener("change", () => {
    endpointInput.hidden = listSelect.value !== "::endpoint";
    loadFieldsForList(pickedListId());
  });
  endpointInput.addEventListener("input", onBuilderChange);
  addFilter.addEventListener("click", () => {
    if (!fields.length) return;
    addFilterRow();
    onBuilderChange();
  });
  for (const control of [orderSelect, orderDir, topInput, expandInput]) {
    control.addEventListener("change", onBuilderChange);
    control.addEventListener("input", onBuilderChange);
  }
  let grid = null;
  async function run() {
    const descriptor = composeDescriptor();
    if (!descriptor || !descriptor.path) return;
    writeSaved(client2.webUrl(), {
      listId: listSelect.value,
      endpoint: endpointInput.value,
      select: selectedFields(),
      filters: readFilterRows().map(({ join: join2, field: field2, op, value }) => ({ join: join2, field: field2, op, value })),
      orderby: orderSelect.value,
      orderdir: orderDir.value,
      top: topInput.value,
      expand: expandInput.value,
      raw: rawMode ? rawArea.value : ""
    });
    const select = descriptor.options?.select;
    results.textContent = "";
    grid = createGrid({
      columns: Array.isArray(select) && select.length ? columnsForSelect(select) : [{ key: "__json", label: "Result", value: (row) => JSON.stringify(row), mono: true }],
      rowKey: "Id",
      emptyText: "The query returned no rows.",
      filterPlaceholder: "Filter results\u2026",
      exportName: "sp-query",
      // Raw-mode strings that don't round-trip get no Copy-as menu — a
      // wrong script is worse than none.
      descriptor: rawMode && !rawToDescriptor(rawArea.value) ? null : descriptor
    });
    results.append(grid.el);
    grid.setLoading("Running query\u2026");
    try {
      const { items, partial } = await client2.getAll(descriptor.path, descriptor.options);
      const rows = Array.isArray(select) && select.length ? items : items.map((item2, i) => ({ Id: item2?.Id ?? i, ...item2 }));
      grid.setRows(rows, { partial });
    } catch (err) {
      grid.setError(err);
    }
  }
  runBtn.addEventListener("click", run);
  async function load2() {
    const webUrl = client2.webUrl();
    if (loadedForWeb === webUrl) return;
    loadedForWeb = webUrl;
    try {
      const { items } = await client2.getAll("web/lists", {
        select: ["Id", "Title", "Hidden", "BaseTemplate"],
        orderby: "Title",
        top: 5e3
      });
      lists = items;
    } catch (err) {
      lists = [];
      results.textContent = "";
      results.append(el9("div", "wb-grid-status wb-error", err?.message || String(err)));
    }
    const saved = readSaved(webUrl);
    renderListPicker(saved?.listId || "");
    endpointInput.value = saved?.endpoint || "";
    if (saved?.top) topInput.value = saved.top;
    if (saved?.expand) expandInput.value = saved.expand;
    await loadFieldsForList(pickedListId(), saved);
    if (saved?.raw) {
      rawArea.value = saved.raw;
      enterRawMode();
    }
  }
  return { el: root, load: load2 };
}

// ../src/workbench/canvas.js
var WEBPART_NAMES = {
  "d1d91016-032f-456d-98a4-721247c305e8": "Image",
  "daf0b71c-6de8-4ef7-b511-faae7c388708": "Highlighted content",
  "490d7c76-1824-45b2-9de3-676421c997fa": "Embed",
  "b7dd04e1-19ce-4b24-9132-b60a1c2b910d": "File viewer",
  "af8be689-990e-492a-81f7-ba3e4cd3ed9c": "Image gallery",
  "6410b3b6-d440-4663-8744-378976dc041e": "Link",
  "0ef418ba-5d19-4ade-9db0-b339873291d0": "News feed",
  "a5df8fdf-b508-4b66-98a6-d83bc2597f63": "News",
  "8c88f208-6c77-4bdb-86a0-0c47b4316588": "News reel",
  "58fcd18b-e1af-4b0a-b23b-422c2c52d5a2": "Power BI",
  "91a50c94-865f-4f5c-8b4e-e49659e69772": "Quick chart",
  "eb95c819-ab8f-4689-bd03-0c2d65d47b1f": "Site activity",
  "275c0095-a77e-4f6d-a2a0-6a7626911518": "Stream",
  "31e9537e-f9dc-40a4-8834-0e3b7df418bc": "Yammer embed",
  "20745d7d-8581-4a6c-bf26-68279bc123fc": "Events",
  "6676088b-e28e-4a90-b9cb-d0d0303cd2eb": "Group calendar",
  "c4bd7b2f-7b6e-4599-8485-16504575f590": "Hero",
  "f92bf067-bc19-489e-a556-7fe95f508720": "List",
  "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788": "Page title",
  "7f718435-ee4d-431c-bdbf-9c4ff326f46e": "People",
  "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd": "Quick links",
  "e377ea37-9047-43b9-8cdb-a761be2f8e09": "Bing maps",
  "2161a1c6-db61-4731-b97c-3cdb303f7cbb": "Divider",
  "8654b779-4886-46d4-8ffb-b5ed960ee986": "Spacer",
  "b19b3b9e-8d13-4fec-a93c-401a091c0707": "Microsoft Forms",
  // ---- verify on live tenant (lower confidence) ----
  "f6fdf4f8-4a24-437b-a127-32e66a5dd9b4": "Twitter",
  "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823": "Weather",
  "cf91cf5d-ac23-4a7a-9dbc-cd9ea1a095eb": "Saved for later",
  "7cba020c-5ccb-42e8-b6fc-75b3149aba7b": "Document library",
  "0f087d7f-520e-42b7-89c0-496aaf979d58": "Button",
  "df8e44e7-edd5-46d5-90da-aca1539313b8": "Call to action",
  "62cac389-787f-495d-beca-e11786162ef4": "Countdown timer",
  "9d7e898c-f1bb-473a-9ace-8b415036578b": "Organization chart",
  "71c19a43-d08c-4178-8218-4df8554c0b0e": "Country/region web part",
  "e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b": "Sites",
  "544dd15b-cf3c-441b-96da-004d5a8cea1d": "YouTube",
  "a8cd4347-f996-48c1-bcfb-75373fed2a27": "World clock",
  "46698648-fcd5-41fc-9526-c7f7b2ace919": "Markdown",
  "1ef5ed11-ce7b-44be-bc5e-4abd55101d16": "Code snippet"
};
function webPartName(webPartId) {
  const key2 = String(webPartId || "").toLowerCase().replace(/[{}]/g, "");
  return WEBPART_NAMES[key2] || String(webPartId || "");
}
var asObject = (v) => v && typeof v === "object" && !Array.isArray(v) ? v : {};
function normalizeControl(entry) {
  if (entry === null || entry === void 0 || typeof entry !== "object" || Array.isArray(entry)) {
    return { kind: "unknown", raw: entry };
  }
  const position = asObject(entry.position);
  const base = {
    id: entry.id || entry.controlId || "",
    controlType: entry.controlType,
    position: {
      zoneIndex: position.zoneIndex,
      sectionIndex: position.sectionIndex,
      controlIndex: position.controlIndex,
      sectionFactor: position.sectionFactor,
      layoutIndex: position.layoutIndex,
      zoneId: position.zoneId
    },
    emphasis: asObject(entry.emphasis),
    zoneGroupMetadata: entry.zoneGroupMetadata || null,
    raw: entry
  };
  if (entry.pageSettingsSlice) {
    return { ...base, kind: "pageSettings", pageSettingsSlice: entry.pageSettingsSlice };
  }
  if (entry.controlType === 4) {
    return { ...base, kind: "text", innerHTML: String(entry.innerHTML ?? "") };
  }
  if (entry.controlType === 3) {
    const webPartData = asObject(entry.webPartData);
    const spc = asObject(webPartData.serverProcessedContent);
    return {
      ...base,
      kind: "webpart",
      webPartId: String(entry.webPartId || webPartData.id || ""),
      webPartData: {
        title: String(webPartData.title ?? ""),
        description: String(webPartData.description ?? ""),
        properties: asObject(webPartData.properties),
        serverProcessedContent: {
          htmlStrings: asObject(spc.htmlStrings),
          searchablePlainTexts: asObject(spc.searchablePlainTexts),
          imageSources: asObject(spc.imageSources),
          links: asObject(spc.links)
        }
      }
    };
  }
  if (entry.position && entry.controlType === void 0) {
    return { ...base, kind: "section" };
  }
  return { ...base, kind: "unknown" };
}
function parseHtmlFormat(raw, errors) {
  const doc = new DOMParser().parseFromString(raw, "text/html");
  const nodes = doc.querySelectorAll("[data-sp-canvascontrol], [data-sp-controldata]");
  const entries = [];
  for (const node of nodes) {
    const data = node.getAttribute("data-sp-controldata");
    if (!data) continue;
    let entry;
    try {
      entry = JSON.parse(data);
    } catch {
      errors.push("An HTML canvas control carried unparseable control data.");
      entries.push({ __unparseable: true, raw: data });
      continue;
    }
    if (!entry.webPartData) {
      const wpNode = node.querySelector("[data-sp-webpartdata]");
      if (wpNode) {
        try {
          entry.webPartData = JSON.parse(wpNode.getAttribute("data-sp-webpartdata"));
        } catch {
          errors.push("An HTML canvas web part carried unparseable web-part data.");
        }
      }
    }
    if (entry.controlType === 4 && entry.innerHTML === void 0) {
      entry.innerHTML = node.querySelector("[data-sp-rte]")?.innerHTML ?? "";
    }
    entries.push(entry);
  }
  if (!entries.length) errors.push("No canvas controls were found in the HTML markup.");
  return entries;
}
function parseCanvasContent(raw) {
  const errors = [];
  const text = String(raw ?? "").trim();
  if (!text) return { ok: true, controls: [], pageSettings: null, errors };
  let entries = null;
  if (text.startsWith("<")) {
    try {
      entries = parseHtmlFormat(text, errors);
    } catch {
      errors.push("The HTML canvas markup could not be parsed.");
      entries = [];
    }
  } else {
    try {
      const parsed = JSON.parse(text);
      entries = Array.isArray(parsed) ? parsed : [parsed];
    } catch {
      errors.push("CanvasContent1 is neither valid JSON nor recognizable HTML markup.");
      return { ok: false, controls: [], pageSettings: null, errors };
    }
  }
  const controls = [];
  let pageSettings = null;
  for (const entry of entries) {
    let control;
    try {
      control = entry?.__unparseable ? { kind: "unknown", raw: entry.raw } : normalizeControl(entry);
    } catch {
      control = { kind: "unknown", raw: entry };
      errors.push("A canvas entry could not be normalized.");
    }
    if (control.kind === "unknown" && !entry?.__unparseable) {
      errors.push("A canvas entry had an unrecognized shape \u2014 shown raw.");
    }
    if (control.kind === "pageSettings" && !pageSettings) {
      pageSettings = control.pageSettingsSlice;
    }
    controls.push(control);
  }
  return { ok: true, controls, pageSettings, errors };
}
function buildSectionTree(controls) {
  const sections = /* @__PURE__ */ new Map();
  const unplaced = [];
  for (const control of controls || []) {
    if (control.kind === "pageSettings") continue;
    const { zoneIndex, sectionIndex } = control.position || {};
    if (typeof zoneIndex !== "number" || typeof sectionIndex !== "number") {
      unplaced.push(control);
      continue;
    }
    if (!sections.has(zoneIndex)) {
      sections.set(zoneIndex, {
        zoneIndex,
        emphasis: 0,
        vertical: false,
        collapsible: null,
        columns: /* @__PURE__ */ new Map()
      });
    }
    const section = sections.get(zoneIndex);
    if (typeof control.emphasis?.zoneEmphasis === "number") {
      section.emphasis = control.emphasis.zoneEmphasis;
    }
    if (control.position.layoutIndex === 2) section.vertical = true;
    if (control.zoneGroupMetadata) section.collapsible = control.zoneGroupMetadata;
    if (!section.columns.has(sectionIndex)) {
      section.columns.set(sectionIndex, {
        sectionIndex,
        sectionFactor: control.position.sectionFactor,
        controls: []
      });
    }
    const column = section.columns.get(sectionIndex);
    if (typeof control.position.sectionFactor === "number") {
      column.sectionFactor = control.position.sectionFactor;
    }
    if (control.kind !== "section") column.controls.push(control);
  }
  const ordered = [...sections.values()].sort((a, b) => a.zoneIndex - b.zoneIndex).map((section) => ({
    ...section,
    columns: [...section.columns.values()].sort((a, b) => a.sectionIndex - b.sectionIndex).map((column) => ({
      ...column,
      controls: [...column.controls].sort(
        (a, b) => (a.position.controlIndex ?? 0) - (b.position.controlIndex ?? 0)
      )
    }))
  }));
  return { sections: ordered, unplaced };
}
function textOfControl(control) {
  if (!control) return "";
  if (control.kind === "text") {
    const doc = new DOMParser().parseFromString(control.innerHTML || "", "text/html");
    return (doc.body.textContent || "").replace(/\s+/g, " ").trim();
  }
  if (control.kind === "webpart") {
    const texts = control.webPartData?.serverProcessedContent?.searchablePlainTexts || {};
    return Object.values(texts).filter((v) => typeof v === "string").join(" \xB7 ");
  }
  return "";
}
function sanitizeHtml(html) {
  const doc = new DOMParser().parseFromString(String(html || ""), "text/html");
  for (const node of doc.querySelectorAll("script, iframe, object, embed, form")) {
    node.remove();
  }
  for (const node of doc.body.querySelectorAll("*")) {
    for (const attr of [...node.attributes]) {
      const name = attr.name.toLowerCase();
      if (name.startsWith("on")) node.removeAttribute(attr.name);
      else if ((name === "href" || name === "src" || name === "xlink:href") && /^\s*javascript:/i.test(attr.value)) {
        node.removeAttribute(attr.name);
      }
    }
  }
  return doc.body.innerHTML;
}

// ../src/workbench/field-editor.js
var el10 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var EDITABLE_TYPES = /* @__PURE__ */ new Set([
  "Text",
  "Note",
  "Choice",
  "MultiChoice",
  "Boolean",
  "Number",
  "Currency",
  "DateTime",
  "URL"
]);
var NO_EDIT_INTERNAL = /* @__PURE__ */ new Set([
  "CanvasContent1",
  "LayoutWebpartsContent",
  "ContentType",
  "Attachments"
]);
function isEditable(field2) {
  return !field2.ReadOnlyField && !field2.Hidden && EDITABLE_TYPES.has(String(field2.TypeAsString || "")) && !NO_EDIT_INTERNAL.has(String(field2.InternalName || ""));
}
var choicesOf = (field2) => {
  const v = field2?.Choices;
  const arr = Array.isArray(v) ? v : v?.results;
  return Array.isArray(arr) ? arr : [];
};
function toFormValue(field2, uiValue) {
  switch (String(field2?.TypeAsString || "")) {
    case "MultiChoice": {
      const arr = Array.isArray(uiValue) ? uiValue.filter(Boolean) : [];
      return arr.length ? `;#${arr.join(";#")};#` : "";
    }
    case "Boolean":
      return uiValue ? "1" : "0";
    case "Number":
    case "Currency": {
      const s = String(uiValue ?? "").trim();
      return s === "" ? "" : String(Number(s.replace(",", ".")));
    }
    case "DateTime": {
      const s = String(uiValue ?? "").trim();
      if (!s) return "";
      const d = new Date(s);
      return Number.isNaN(d.getTime()) ? s : d.toISOString();
    }
    case "URL": {
      const url = String(uiValue?.url ?? "").trim();
      const description = String(uiValue?.description ?? "").trim();
      if (!url) return "";
      return description ? `${url}, ${description}` : url;
    }
    default:
      return String(uiValue ?? "");
  }
}
function fromItemValue(field2, itemValue) {
  switch (String(field2?.TypeAsString || "")) {
    case "MultiChoice": {
      if (Array.isArray(itemValue)) return itemValue;
      if (Array.isArray(itemValue?.results)) return itemValue.results;
      return String(itemValue ?? "").split(";#").filter(Boolean);
    }
    case "Boolean":
      return itemValue === true || itemValue === 1 || /^(1|true|yes)$/i.test(String(itemValue ?? ""));
    case "Number":
    case "Currency":
      return itemValue === null || itemValue === void 0 ? "" : String(itemValue);
    case "DateTime": {
      const s = String(itemValue ?? "").trim();
      if (!s) return "";
      const d = new Date(s);
      if (Number.isNaN(d.getTime())) return s;
      const pad = (n) => String(n).padStart(2, "0");
      return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
    case "URL":
      return {
        url: String(itemValue?.Url ?? itemValue?.url ?? "").trim(),
        description: String(itemValue?.Description ?? itemValue?.description ?? "").trim()
      };
    default:
      return itemValue === null || itemValue === void 0 ? "" : String(itemValue);
  }
}
function createFieldEditor(field2, initialValue) {
  const type = String(field2.TypeAsString || "");
  const initial = fromItemValue(field2, initialValue);
  const row = el10("div", "wb-editor-row");
  row.dataset.internal = field2.InternalName || "";
  const label = el10("label", "wb-editor-label", field2.Title || field2.InternalName);
  const typeBadge = el10("span", "wb-editor-type", type);
  label.append(typeBadge);
  const control = el10("div", "wb-editor-control");
  const error = el10("div", "wb-editor-error");
  error.hidden = true;
  row.append(label, control, error);
  let getValue = () => "";
  const textInput = (tag, value) => {
    const input = el10(tag === "textarea" ? "textarea" : "input");
    if (tag !== "textarea") input.type = tag;
    input.value = value ?? "";
    control.append(input);
    return input;
  };
  switch (type) {
    case "Note": {
      const input = textInput("textarea", initial);
      getValue = () => input.value;
      break;
    }
    case "Choice": {
      const select = el10("select");
      const options = choicesOf(field2);
      const blank = el10("option", "", "\u2014");
      blank.value = "";
      select.append(blank);
      for (const choice of options) {
        const opt = el10("option", "", choice);
        opt.value = choice;
        select.append(opt);
      }
      if (initial && !options.includes(initial)) {
        const opt = el10("option", "", `${initial} (current)`);
        opt.value = initial;
        select.append(opt);
      }
      select.value = initial ?? "";
      control.append(select);
      if (field2.FillInChoice) {
        const fillIn = textInput("text", "");
        fillIn.placeholder = "Fill-in value\u2026";
        getValue = () => fillIn.value.trim() || select.value;
      } else {
        getValue = () => select.value;
      }
      break;
    }
    case "MultiChoice": {
      const listBox = el10("div", "wb-editor-choices");
      const initialSet = new Set(Array.isArray(initial) ? initial : []);
      const boxes = [];
      for (const choice of choicesOf(field2)) {
        const lab = el10("label", "wb-editor-choice");
        const box = el10("input");
        box.type = "checkbox";
        box.value = choice;
        box.checked = initialSet.has(choice);
        lab.append(box, document.createTextNode(choice));
        listBox.append(lab);
        boxes.push(box);
      }
      control.append(listBox);
      getValue = () => boxes.filter((b) => b.checked).map((b) => b.value);
      break;
    }
    case "Boolean": {
      const box = el10("input");
      box.type = "checkbox";
      box.checked = Boolean(initial);
      control.append(box);
      getValue = () => box.checked;
      break;
    }
    case "Number":
    case "Currency": {
      const input = textInput("number", initial);
      input.step = "any";
      getValue = () => input.value;
      break;
    }
    case "DateTime": {
      const input = textInput("datetime-local", initial);
      getValue = () => input.value;
      break;
    }
    case "URL": {
      const url = textInput("text", initial?.url);
      url.placeholder = "https://\u2026";
      const description = textInput("text", initial?.description);
      description.placeholder = "Description";
      getValue = () => ({ url: url.value, description: description.value });
      break;
    }
    default: {
      const input = textInput("text", initial);
      getValue = () => input.value;
    }
  }
  let baselineForm = toFormValue(field2, getValue());
  return {
    el: row,
    field: field2,
    getValue,
    isDirty: () => toFormValue(field2, getValue()) !== baselineForm,
    markClean() {
      baselineForm = toFormValue(field2, getValue());
    },
    setError(message) {
      error.textContent = message || "";
      error.hidden = !message;
      row.classList.toggle("wb-editor-invalid", Boolean(message));
    }
  };
}
function readOnlyRow(field2, displayText) {
  const row = el10("div", "wb-editor-row wb-editor-readonly");
  row.dataset.internal = field2.InternalName || "";
  const label = el10("label", "wb-editor-label", field2.Title || field2.InternalName);
  label.append(el10("span", "wb-editor-type", String(field2.TypeAsString || "")));
  const value = el10("div", "wb-editor-static", displayText || "");
  value.title = field2.ReadOnlyField ? "Read-only field" : "Not editable in the workbench";
  row.append(label, value);
  return row;
}
function createFieldEditorForm({ fields, item: item2 = {}, itemAsText = {}, onSave }) {
  const root = el10("div", "wb-editor-form");
  const rows = el10("div", "wb-editor-rows");
  const editors = [];
  const shown = (fields || []).filter((f) => !f.Hidden);
  for (const field2 of shown) {
    const internal = field2.InternalName;
    if (isEditable(field2)) {
      const editor = createFieldEditor(field2, item2[internal]);
      editors.push(editor);
      rows.append(editor.el);
    } else {
      const display = itemAsText?.[internal] ?? (item2[internal] === null || item2[internal] === void 0 || typeof item2[internal] === "object" ? "" : String(item2[internal]));
      rows.append(readOnlyRow(field2, String(display ?? "")));
    }
  }
  const bar = el10("div", "wb-editor-bar");
  const save = el10("button", "btn btn-xs", "Save metadata");
  save.type = "button";
  const status = el10("span", "wb-editor-status");
  bar.append(save, status);
  root.append(rows, bar);
  function dirtyFormValues() {
    return editors.filter((e) => e.isDirty()).map((e) => ({
      FieldName: e.field.InternalName,
      FieldValue: toFormValue(e.field, e.getValue())
    }));
  }
  save.addEventListener("click", async () => {
    for (const editor of editors) editor.setError("");
    const formValues = dirtyFormValues();
    if (!formValues.length) {
      status.textContent = "No changes to save.";
      status.className = "wb-editor-status";
      return;
    }
    save.disabled = true;
    status.textContent = "Saving\u2026";
    status.className = "wb-editor-status";
    try {
      await onSave(formValues);
      for (const editor of editors) editor.markClean();
      status.textContent = `Saved ${formValues.length} field${formValues.length === 1 ? "" : "s"}.`;
      status.className = "wb-editor-status wb-editor-saved";
    } catch (err) {
      const fieldErrors = err?.fieldErrors || {};
      let mapped = false;
      for (const editor of editors) {
        const message = fieldErrors[editor.field.InternalName];
        if (message) {
          editor.setError(message);
          mapped = true;
        }
      }
      status.textContent = mapped ? "Some fields were rejected \u2014 see the messages above." : err?.message || String(err);
      status.className = "wb-editor-status wb-editor-failed";
    } finally {
      save.disabled = false;
    }
  });
  return { el: root, getDirtyFormValues: dirtyFormValues, editors };
}

// ../src/workbench/page-export.js
var fmtDate3 = (v) => v ? String(v).slice(0, 10) : "";
function pageLocation({ siteTitle, libraryTitle, fileDirRef, libraryRootPath }) {
  const parts = [siteTitle, libraryTitle].filter(Boolean);
  const dir = String(fileDirRef || "");
  const root = String(libraryRootPath || "").replace(/\/+$/, "");
  let folder = "";
  if (root && dir.toLowerCase().startsWith(root.toLowerCase())) {
    folder = dir.slice(root.length).replace(/^\/+/, "");
  }
  if (folder) parts.push(folder);
  return parts.join(" | ");
}
function contentBlocks(controls) {
  const blocks = [];
  for (const control of controls || []) {
    if (control.kind === "text") {
      const html = String(control.innerHTML || "").trim();
      if (html) blocks.push(html);
    } else if (control.kind === "webpart") {
      const name = webPartName(control.webPartId);
      const title = control.webPartData?.title;
      const label = title && title !== name ? `**[Web part: ${name} \u2014 \u201C${title}\u201D]**` : `**[Web part: ${name}]**`;
      const texts = Object.values(
        control.webPartData?.serverProcessedContent?.searchablePlainTexts || {}
      ).filter((v) => typeof v === "string" && v.trim());
      blocks.push(texts.length ? `${label}

${texts.map((t) => `- ${t}`).join("\n")}` : label);
    } else if (control.kind === "unknown") {
      blocks.push("*[Unparsed canvas entry \u2014 see the raw export]*");
    }
  }
  return blocks;
}
var METADATA_SKIP = /* @__PURE__ */ new Set([
  "CanvasContent1",
  "LayoutWebpartsContent",
  "FieldValuesAsText",
  "Author",
  "Editor"
  // flattened into Created/Modified lines
]);
function buildContentExport({
  item: item2 = {},
  controls = [],
  siteTitle = "",
  webUrl = "",
  libraryTitle = "",
  libraryRootPath = ""
}) {
  const title = item2.Title || item2.FileLeafRef || "Untitled page";
  const author = item2.Author?.Title || "";
  const editor = item2.Editor?.Title || "";
  const location2 = pageLocation({
    siteTitle,
    libraryTitle,
    fileDirRef: item2.FileDirRef,
    libraryRootPath
  });
  let fullUrl = "";
  if (item2.FileRef) {
    try {
      fullUrl = `${new URL(webUrl).origin}${encodeURI(item2.FileRef)}`;
    } catch {
      fullUrl = item2.FileRef;
    }
  }
  const top = [`# ${title}`, ""];
  if (item2.Description) top.push(`> ${String(item2.Description).replace(/\r?\n/g, " ")}`, "");
  top.push(`Created ${fmtDate3(item2.Created)}${author ? ` by ${author}` : ""}  `);
  if (location2) top.push(`Location: ${location2}`, "");
  const meta = ["## Metadata", ""];
  const metaLine = (label, value) => {
    if (value !== "" && value !== null && value !== void 0) {
      meta.push(`- ${label}: ${value}`);
    }
  };
  metaLine("Title", item2.Title);
  metaLine("Description", item2.Description);
  metaLine("Created", item2.Created ? `${item2.Created}${author ? ` by ${author}` : ""}` : "");
  metaLine("Modified", item2.Modified ? `${item2.Modified}${editor ? ` by ${editor}` : ""}` : "");
  metaLine("Site", siteTitle);
  metaLine("Library", libraryTitle);
  metaLine("URL", fullUrl);
  for (const [key2, value] of Object.entries(item2)) {
    if (METADATA_SKIP.has(key2) || key2.startsWith("odata") || key2.startsWith("__")) continue;
    if (["Title", "Description", "Created", "Modified", "FileRef"].includes(key2)) continue;
    if (value === null || value === void 0 || typeof value === "object") continue;
    meta.push(`- ${key2}: ${value}`);
  }
  return [
    ...top,
    "---",
    "",
    contentBlocks(controls).join("\n\n"),
    "",
    "---",
    "",
    ...meta,
    ""
  ].join("\n");
}
function buildRawExport({ item: item2 = {}, controls = [] }) {
  return JSON.stringify({ item: item2, controls }, null, 2);
}
function exportFileStem(item2) {
  const name = String(item2.FileLeafRef || item2.Title || "page").replace(/\.aspx$/i, "");
  return name.toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "page";
}

// ../src/workbench/views/pages.js?v=2
var PAGE_SELECT = [
  "Id",
  "Title",
  "FileLeafRef",
  "FileRef",
  "FileDirRef",
  "PromotedState",
  "Modified",
  "UniqueId",
  "Editor/Title"
];
var DETAIL_SELECT = [
  "Id",
  "Title",
  "FileLeafRef",
  "FileRef",
  "FileDirRef",
  "Description",
  "BannerImageUrl",
  "PromotedState",
  "Created",
  "Modified",
  "Author/Title",
  "Editor/Title",
  "CanvasContent1",
  "LayoutWebpartsContent"
];
var FIELD_SELECT3 = [
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
  "FillInChoice"
];
var SITE_PAGES_BASE_TEMPLATE = 119;
var promotedLabel = (v) => ({ 0: "", 1: "News (pending)", 2: "News" })[v] ?? String(v ?? "");
var fmtDate4 = (v) => v ? String(v).slice(0, 10) : "";
var el11 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var encodedServerPath = (path) => String(path || "").split("/").map((segment) => {
  try {
    return encodeURIComponent(decodeURIComponent(segment));
  } catch {
    return encodeURIComponent(segment);
  }
}).join("/");
var guidPath2 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
function createPagesView({ client: client2, navigate }) {
  const root = el11("section", "wb-view wb-view-pages");
  const spWrite = createSpWriteClient({ client: client2 });
  const gridPane = el11("div", "wb-pane");
  const head = el11("div", "wb-view-head");
  head.innerHTML = '<h2>Pages</h2><p class="wb-view-hint">Modern pages in this web\u2019s Site Pages library, subfolders included. Click a row to inspect content, metadata, and structure.</p>';
  const libraryLink = el11("a", "btn btn-xs wb-head-link", "Open Site Pages library \u2197");
  libraryLink.target = "_blank";
  libraryLink.rel = "noopener";
  libraryLink.hidden = true;
  head.append(libraryLink);
  const masterStatus = el11("div", "wb-grid-status");
  masterStatus.hidden = true;
  gridPane.append(head, masterStatus);
  const detailPane = el11("div", "wb-pane");
  detailPane.hidden = true;
  root.append(gridPane, detailPane);
  let sitePagesPromise = null;
  let grid = null;
  let pagesLoaded = false;
  const detailCache = /* @__PURE__ */ new Map();
  let fieldsPromise = null;
  let detailRun = 0;
  function sitePagesList() {
    if (!sitePagesPromise) {
      sitePagesPromise = client2.getAll("web/lists", {
        select: ["Id", "Title", "BaseTemplate", "Hidden", "DefaultViewUrl", "RootFolder/ServerRelativeUrl"],
        expand: "RootFolder",
        top: 5e3
      }).then(({ items }) => {
        const found = items.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE && !l.Hidden) || items.find((l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE);
        return found ? {
          listId: found.Id,
          title: found.Title,
          rootPath: found.RootFolder?.ServerRelativeUrl || "",
          viewUrl: found.DefaultViewUrl || found.RootFolder?.ServerRelativeUrl || ""
        } : null;
      }).catch((err) => {
        sitePagesPromise = null;
        throw err;
      });
    }
    return sitePagesPromise;
  }
  let webInfoPromise = null;
  function webIdentity() {
    if (!webInfoPromise) {
      webInfoPromise = client2.get("web", { select: ["Title", "Url"] }).catch(() => ({ Title: "", Url: client2.webUrl() }));
    }
    return webInfoPromise;
  }
  function folderOf(fileDirRef, rootPath) {
    const dir = String(fileDirRef || "");
    const root2 = String(rootPath || "").replace(/\/+$/, "");
    if (!root2 || !dir.toLowerCase().startsWith(root2.toLowerCase())) return "";
    return dir.slice(root2.length).replace(/^\/+/, "");
  }
  async function loadPages() {
    if (pagesLoaded) return;
    masterStatus.hidden = true;
    try {
      const sitePages = await sitePagesList();
      if (!sitePages) {
        masterStatus.textContent = "This web has no Site Pages library (BaseTemplate 119).";
        masterStatus.hidden = false;
        return;
      }
      if (sitePages.viewUrl) {
        libraryLink.href = sitePages.viewUrl;
        libraryLink.hidden = false;
      }
      if (!grid) {
        const query = {
          path: guidPath2(sitePages.listId, "/items"),
          options: {
            select: PAGE_SELECT,
            expand: "Editor",
            orderby: "FileLeafRef",
            top: 5e3
          }
        };
        grid = createGrid({
          columns: [
            { key: "FileLeafRef", label: "Name", mono: true },
            { key: "Title", label: "Title" },
            {
              key: "Folder",
              label: "Folder",
              value: (row) => folderOf(row.FileDirRef, sitePages.rootPath),
              format: (v) => v ? `/${v}` : ""
            },
            { key: "PromotedState", label: "Promoted", format: promotedLabel },
            { key: "Modified", label: "Modified", format: fmtDate4 },
            { key: "Editor", label: "Editor", value: (row) => row.Editor?.Title || "" },
            {
              key: "FileRef",
              label: "",
              format: () => "",
              render: (fileRef) => {
                if (!fileRef) return null;
                const a = document.createElement("a");
                a.className = "wb-cell-link";
                a.href = fileRef;
                a.target = "_blank";
                a.rel = "noopener";
                a.title = "Open the page in a new tab";
                a.textContent = "\u2197";
                a.addEventListener("click", (e) => e.stopPropagation());
                return a;
              }
            }
          ],
          onOpen: (row) => navigate({
            view: "pages",
            pageId: row.Id,
            pageName: row.FileLeafRef || row.Title
          }),
          emptyText: "No pages in this library.",
          filterPlaceholder: "Filter pages\u2026",
          exportName: "sp-pages",
          descriptor: { ...query, webUrl: client2.webUrl() }
        });
        gridPane.append(grid.el);
        grid.setLoading("Loading pages\u2026");
        const { items, partial } = await client2.getAll(query.path, query.options);
        grid.setRows(items, { partial });
        pagesLoaded = true;
      }
    } catch (err) {
      if (grid) grid.setError(err);
      else {
        masterStatus.textContent = err?.message || String(err);
        masterStatus.classList.add("wb-error");
        masterStatus.hidden = false;
      }
    }
  }
  function pageItem(listId, pageId) {
    if (!detailCache.has(pageId)) {
      detailCache.set(pageId, client2.get(guidPath2(listId, `/items(${pageId})`), {
        select: DETAIL_SELECT,
        expand: ["Author", "Editor"]
      }).catch((err) => {
        detailCache.delete(pageId);
        throw err;
      }));
    }
    return detailCache.get(pageId);
  }
  function listFields(listId) {
    if (!fieldsPromise) {
      fieldsPromise = client2.getAll(guidPath2(listId, "/fields"), { select: FIELD_SELECT3 }).then(({ items }) => items).catch((err) => {
        fieldsPromise = null;
        throw err;
      });
    }
    return fieldsPromise;
  }
  function structurePane(parsed) {
    const wrap = el11("div", "wb-tab-pane");
    const tree = el11("div", "wb-canvas-tree");
    const { sections, unplaced } = buildSectionTree(parsed.controls);
    if (!sections.length && !unplaced.length) {
      tree.append(el11("div", "wb-grid-status", "No canvas sections on this page."));
    }
    sections.forEach((section, i) => {
      const bits = [`${section.columns.length} column${section.columns.length === 1 ? "" : "s"}`];
      if (section.emphasis) bits.push(`emphasis ${section.emphasis}`);
      if (section.vertical) bits.push("vertical");
      if (section.collapsible) bits.push("collapsible");
      tree.append(el11("div", "wb-canvas-section", `Section ${i + 1} \u2014 ${bits.join(", ")}`));
      for (const column of section.columns) {
        const row = el11("div", "wb-canvas-column");
        const width = typeof column.sectionFactor === "number" ? `${column.sectionFactor}/12` : "auto";
        row.append(el11("span", "wb-canvas-width", width));
        if (!column.controls.length) row.append(el11("span", "wb-canvas-chip wb-canvas-empty", "empty"));
        for (const control of column.controls) {
          const chipLabel = control.kind === "text" ? "Text" : control.kind === "webpart" ? control.webPartData.title || webPartName(control.webPartId) : control.kind;
          const chip = el11("span", "wb-canvas-chip", chipLabel);
          chip.title = control.kind === "webpart" ? `${webPartName(control.webPartId)} \xB7 ${control.webPartId}` : textOfControl(control).slice(0, 200);
          row.append(chip);
        }
        tree.append(row);
      }
    });
    if (unplaced.length) {
      tree.append(el11("div", "wb-canvas-section", `Unplaced entries (${unplaced.length})`));
      for (const control of unplaced) {
        const row = el11("div", "wb-canvas-column");
        row.append(el11("span", "wb-canvas-chip", control.kind));
        tree.append(row);
      }
    }
    wrap.append(tree);
    return wrap;
  }
  function webPartsPane(parsed) {
    const wrap = el11("div", "wb-tab-pane");
    const rows = parsed.controls.filter((c) => c.kind === "webpart").map((c, i) => ({
      Id: c.id || String(i),
      Title: c.webPartData.title,
      Type: webPartName(c.webPartId),
      WebPartId: c.webPartId,
      ControlId: c.id,
      Text: textOfControl(c).slice(0, 160)
    }));
    const partsGrid = createGrid({
      columns: [
        { key: "Title", label: "Title" },
        { key: "Type", label: "Type" },
        { key: "WebPartId", label: "Web part id", mono: true, copyable: true },
        { key: "ControlId", label: "Control id", mono: true, copyable: true },
        { key: "Text", label: "Text" }
      ],
      emptyText: "No client-side web parts on this page.",
      filterPlaceholder: "Filter web parts\u2026",
      exportName: "sp-page-webparts"
    });
    wrap.append(partsGrid.el);
    partsGrid.setRows(rows);
    return wrap;
  }
  function textPane(parsed) {
    const wrap = el11("div", "wb-tab-pane wb-text-pane");
    const texts = parsed.controls.filter((c) => c.kind === "text");
    if (!texts.length) {
      wrap.append(el11("div", "wb-grid-status", "No text web parts on this page."));
      return wrap;
    }
    texts.forEach((control, i) => {
      const block = el11("div", "wb-text-block");
      block.append(el11("div", "wb-subpanel-title", `Text web part ${i + 1}`));
      const rendered = el11("div", "wb-text-rendered");
      rendered.innerHTML = sanitizeHtml(control.innerHTML);
      block.append(rendered);
      const details = document.createElement("details");
      details.append(el11("summary", "", "Raw HTML"));
      const pre = el11("pre", "wb-text-raw", control.innerHTML);
      details.append(pre);
      block.append(details);
      wrap.append(block);
    });
    return wrap;
  }
  function metadataPane(listId, pageId) {
    const wrap = el11("div", "wb-tab-pane");
    const status = el11("div", "wb-grid-status", "Loading metadata\u2026");
    wrap.append(status);
    (async () => {
      const fields = await listFields(listId);
      let item2;
      let itemAsText = {};
      try {
        item2 = await client2.get(guidPath2(listId, `/items(${pageId})`), {
          expand: "FieldValuesAsText"
        });
        itemAsText = item2.FieldValuesAsText || {};
      } catch {
        item2 = await client2.get(guidPath2(listId, `/items(${pageId})`));
        try {
          itemAsText = await client2.get(guidPath2(listId, `/items(${pageId})/FieldValuesAsText`));
        } catch {
          itemAsText = {};
        }
      }
      status.remove();
      const form = createFieldEditorForm({
        fields,
        item: item2,
        itemAsText,
        onSave: (formValues) => spWrite.validateUpdateListItem({ listId, itemId: pageId }, formValues)
      });
      wrap.append(form.el);
    })().catch((err) => {
      status.textContent = err?.message || String(err);
      status.classList.add("wb-error");
    });
    return wrap;
  }
  function rawPane(item2, parsed) {
    const wrap = el11("div", "wb-tab-pane");
    const node = toNode({ item: item2, parsedCanvas: parsed.controls }, 0, { maxDepth: 10, maxItems: 400 });
    const inspector = el11("div", "wb-raw");
    inspector.append(enhance(node) ?? renderValue(node));
    wrap.append(inspector);
    return wrap;
  }
  async function showDetail(route) {
    const run = ++detailRun;
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = "";
    const back = el11("button", "btn btn-xs wb-back", "\u2190 All pages");
    back.type = "button";
    back.addEventListener("click", () => navigate({ view: "pages" }));
    const title = el11("h2", "", route.pageName || `Page ${route.pageId}`);
    const headRow = el11("div", "wb-detail-head");
    headRow.append(back, title);
    detailPane.append(headRow);
    const status = el11("div", "wb-grid-status", "Loading page\u2026");
    detailPane.append(status);
    let sitePages;
    let item2;
    try {
      sitePages = await sitePagesList();
      if (!sitePages) throw new Error("This web has no Site Pages library.");
      item2 = await pageItem(sitePages.listId, route.pageId);
    } catch (err) {
      if (run !== detailRun) return;
      status.textContent = err?.message || String(err);
      status.classList.add("wb-error");
      return;
    }
    if (run !== detailRun) return;
    status.remove();
    if (item2.FileRef) {
      const origin = (() => {
        try {
          return new URL(client2.webUrl()).origin;
        } catch {
          return "";
        }
      })();
      const fullUrl = `${origin}${encodedServerPath(item2.FileRef)}`;
      const frag = el11("span", "wb-detail-id sp-copy", item2.FileRef);
      frag.title = `Click to copy the full URL
${fullUrl}`;
      frag.addEventListener("click", () => copyText(fullUrl, frag));
      headRow.append(frag);
    }
    const actions = el11("span", "wb-detail-actions");
    const exportContent = el11("button", "btn btn-xs", "Export content");
    exportContent.type = "button";
    exportContent.title = "One human-readable file: metadata, merged web-part content, full metadata";
    const exportRaw = el11("button", "btn btn-xs", "Export raw");
    exportRaw.type = "button";
    exportRaw.title = "Item + parsed canvas controls as JSON, for scripts";
    actions.append(exportContent, exportRaw);
    if (item2.FileRef) {
      const open = el11("a", "btn btn-xs", "Open page \u2197");
      open.href = item2.FileRef;
      open.target = "_blank";
      open.rel = "noopener";
      actions.append(open);
    }
    headRow.append(actions);
    const parsed = parseCanvasContent(item2.CanvasContent1);
    exportContent.addEventListener("click", async () => {
      const web = await webIdentity();
      downloadText(`${exportFileStem(item2)}-content.md`, buildContentExport({
        item: item2,
        controls: parsed.controls,
        siteTitle: web.Title || "",
        webUrl: web.Url || client2.webUrl(),
        libraryTitle: sitePages.title,
        libraryRootPath: sitePages.rootPath
      }), "text/markdown;charset=utf-8");
    });
    exportRaw.addEventListener("click", () => {
      downloadText(
        `${exportFileStem(item2)}-raw.json`,
        buildRawExport({ item: item2, controls: parsed.controls }),
        "application/json"
      );
    });
    if (parsed.errors.length) {
      const notice = el11(
        "div",
        "wb-grid-notice",
        `\u26A0 ${parsed.errors.length} canvas entr${parsed.errors.length === 1 ? "y" : "ies"} could not be fully parsed \u2014 shown raw where possible.`
      );
      notice.title = parsed.errors.join("\n");
      detailPane.append(notice);
    }
    const tabsBar = el11("div", "wb-tabs");
    tabsBar.setAttribute("role", "tablist");
    const body = el11("div", "wb-tab-body");
    const panes = /* @__PURE__ */ new Map();
    const TABS = [
      { id: "text", label: "Extract", build: () => textPane(parsed) },
      { id: "metadata", label: "Metadata", build: () => metadataPane(sitePages.listId, route.pageId) },
      { id: "structure", label: "Structure", build: () => structurePane(parsed) },
      { id: "webparts", label: "Web parts", build: () => webPartsPane(parsed) },
      { id: "raw", label: "Raw", build: () => rawPane(item2, parsed) }
    ];
    function activate(tab) {
      for (const btn of tabsBar.children) {
        btn.classList.toggle("active", btn.dataset.tab === tab.id);
        btn.setAttribute("aria-selected", btn.dataset.tab === tab.id ? "true" : "false");
      }
      if (!panes.has(tab.id)) panes.set(tab.id, tab.build());
      body.textContent = "";
      body.append(panes.get(tab.id));
    }
    for (const tab of TABS) {
      const btn = el11("button", "wb-tab", tab.label);
      btn.type = "button";
      btn.dataset.tab = tab.id;
      btn.setAttribute("role", "tab");
      btn.addEventListener("click", () => activate(tab));
      tabsBar.append(btn);
    }
    detailPane.append(tabsBar, body);
    activate(TABS.find((t) => t.id === route.tab) || TABS[0]);
  }
  function load2(route) {
    if (route?.pageId) {
      showDetail(route);
    } else {
      detailRun += 1;
      detailPane.hidden = true;
      gridPane.hidden = false;
      loadPages();
    }
  }
  return { el: root, load: load2 };
}

// ../src/workbench/views/browser.js?v=2
var FIELD_SELECT4 = [
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
  "FillInChoice"
];
var DOCUMENT_LIBRARY_BASE_TYPE = 1;
var GUID = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
var FOLDER_SELECT = ["Name", "ServerRelativeUrl", "ItemCount", "TimeLastModified"];
var FILE_SELECT = [
  "Name",
  "ServerRelativeUrl",
  "Length",
  "TimeLastModified",
  "UIVersionLabel",
  "CheckOutType"
];
var el12 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var icon = (name, size = 15) => {
  const paths = {
    folder: [
      '<path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/>'
    ],
    file: [
      '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>',
      '<path d="M14 2v4a2 2 0 0 0 2 2h4"/>'
    ],
    download: [
      '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>',
      '<path d="m7 10 5 5 5-5"/>',
      '<path d="M12 15V3"/>'
    ],
    link: [
      '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>',
      '<path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
    ]
  };
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", String(size));
  svg.setAttribute("height", String(size));
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "2");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = (paths[name] || []).join("");
  return svg;
};
var fileRole = (row) => {
  if (row.kind === "folder") return /^forms$|^_/i.test(row.Name || "") ? "sys" : "user";
  const ext = extOf(row.Name);
  if (["js", "mjs", "cjs", "ts", "tsx"].includes(ext)) return "js";
  if (["html", "htm", "svg"].includes(ext)) return "html";
  if (["css", "scss", "less"].includes(ext)) return "css";
  if (["json", "csv", "tsv", "xml", "xlsx", "xls"].includes(ext)) return "json";
  if (["doc", "docx", "pdf", "ppt", "pptx", "rtf"].includes(ext)) return "doc";
  return "file";
};
var encodedServerPath2 = (path) => String(path || "").split("/").map((segment) => {
  try {
    return encodeURIComponent(decodeURIComponent(segment));
  } catch {
    return encodeURIComponent(segment);
  }
}).join("/");
var fmtDate5 = (v) => v ? String(v).slice(0, 10) : "";
function formatBytes(n) {
  const bytes = Number(n);
  if (!Number.isFinite(bytes) || bytes < 0) return "";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}
var extOf = (name) => {
  const dot = String(name || "").lastIndexOf(".");
  return dot > 0 ? String(name).slice(dot + 1).toLowerCase() : "";
};
function normalizedPath2(value) {
  let path = String(value || "").trim().replaceAll("\\", "/");
  if (!path.startsWith("/")) path = `/${path}`;
  path = path.replace(/\/{2,}/g, "/");
  if (path.length > 1) path = path.replace(/\/+$/, "");
  return path;
}
function createBrowserView({ client: client2, navigate }) {
  const root = el12("section", "wb-view wb-view-files");
  const spWrite = createSpWriteClient({ client: client2 });
  const head = el12("div", "wb-view-head");
  head.innerHTML = '<h2>Files</h2><p class="wb-view-hint">Browse any library or folder of this web \u2014 every file type, with download, binary upload, and full metadata editing.</p>';
  const bar = el12("div", "wb-crumbs-bar");
  const librarySelect = el12("select", "wb-lib-select");
  librarySelect.setAttribute("aria-label", "Jump to a document library");
  const crumbs = el12("div", "wb-crumbs");
  bar.append(librarySelect, crumbs);
  const consent = el12("div", "wb-consent");
  consent.hidden = true;
  const gridWrap = el12("div", "wb-files-grid");
  const metaPanel = el12("div", "wb-subpanel wb-file-meta");
  metaPanel.hidden = true;
  root.append(head, bar, consent, gridWrap, metaPanel);
  let libraries = [];
  let currentPath = "";
  let currentListing = { folders: [], files: [] };
  let grid = null;
  let librariesLoaded = false;
  let listingRun = 0;
  const parentListCache = /* @__PURE__ */ new Map();
  const fieldsCache = /* @__PURE__ */ new Map();
  function webRootPath() {
    try {
      return normalizedPath2(decodeURIComponent(new URL(client2.webUrl()).pathname)) || "/";
    } catch {
      return "/";
    }
  }
  function checkedPath(path) {
    const rootPath = webRootPath();
    const normalized = normalizedPath2(path || rootPath);
    if (rootPath !== "/" && normalized !== rootPath && !normalized.startsWith(`${rootPath}/`)) {
      throw new Error("That path is outside the inspected web.");
    }
    return normalized;
  }
  const folderApi = (path, sub) => `web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')${sub}`;
  const fileApi = (path, sub) => `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')${sub}`;
  function downloadHref(serverRelativeUrl2) {
    if (spWrite.isMock()) return serverRelativeUrl2;
    return `${client2.webUrl()}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl2)}`;
  }
  function renderCrumbs() {
    crumbs.textContent = "";
    const rootPath = webRootPath();
    const segments = currentPath === "/" ? [] : currentPath.slice(1).split("/");
    let acc = "";
    const rootBtn = el12("button", "wb-crumb", rootPath === "/" ? "/" : rootPath);
    rootBtn.type = "button";
    rootBtn.addEventListener("click", () => navigate({ view: "files", path: rootPath }));
    let started = rootPath === "/";
    if (started) crumbs.append(rootBtn);
    for (const segment of segments) {
      acc += `/${segment}`;
      if (!started) {
        if (normalizedPath2(acc) === rootPath) {
          started = true;
          const btn2 = el12("button", "wb-crumb", rootPath);
          btn2.type = "button";
          btn2.addEventListener("click", () => navigate({ view: "files", path: rootPath }));
          crumbs.append(btn2);
        }
        continue;
      }
      crumbs.append(el12("span", "wb-crumb-sep", "/"));
      const target = acc;
      const btn = el12("button", "wb-crumb", segment);
      btn.type = "button";
      btn.addEventListener("click", () => navigate({ view: "files", path: target }));
      crumbs.append(btn);
    }
  }
  async function loadLibraries() {
    if (librariesLoaded) return;
    try {
      const { items } = await client2.getAll("web/lists", {
        select: ["Id", "Title", "BaseType", "Hidden", "RootFolder/ServerRelativeUrl"],
        expand: "RootFolder",
        orderby: "Title",
        top: 5e3
      });
      libraries = items.filter((l) => l.BaseType === DOCUMENT_LIBRARY_BASE_TYPE && !l.Hidden);
      librariesLoaded = true;
    } catch {
      libraries = [];
    }
    librarySelect.textContent = "";
    const blank = el12("option", "", "Libraries\u2026");
    blank.value = "";
    librarySelect.append(blank);
    for (const lib of libraries) {
      const url = lib.RootFolder?.ServerRelativeUrl;
      if (!url) continue;
      const opt = el12("option", "", lib.Title);
      opt.value = url;
      librarySelect.append(opt);
    }
  }
  librarySelect.addEventListener("change", () => {
    if (librarySelect.value) navigate({ view: "files", path: librarySelect.value });
  });
  function makeGrid() {
    grid = createGrid({
      columns: [
        {
          key: "Name",
          label: "Name",
          value: (row) => row.Name,
          render: (name, row) => {
            const wrap = el12("span", `wb-file-name wb-node-${fileRole(row)}`);
            const glyph = icon(row.kind === "folder" ? "folder" : "file");
            glyph.classList.add("wb-node");
            wrap.append(glyph, el12("span", "wb-file-name-text", name));
            return wrap;
          }
        },
        { key: "Type", label: "Type", value: (row) => row.kind === "folder" ? "Folder" : extOf(row.Name) },
        { key: "Length", label: "Size", num: true, value: (row) => row.kind === "folder" ? null : Number(row.Length) || 0, format: (v, row) => row.kind === "folder" ? "" : formatBytes(v) },
        { key: "TimeLastModified", label: "Modified", format: fmtDate5 },
        { key: "UIVersionLabel", label: "Version", value: (row) => row.kind === "folder" ? "" : row.UIVersionLabel || "" },
        {
          key: "Actions",
          label: "",
          value: (row) => row.ServerRelativeUrl,
          format: () => "",
          render: (serverRelativeUrl2, row) => {
            if (row.kind === "folder") return null;
            const span = document.createElement("span");
            span.className = "wb-file-actions";
            const dl = document.createElement("a");
            dl.className = "wb-cell-link";
            dl.href = downloadHref(serverRelativeUrl2);
            dl.title = "Download";
            dl.setAttribute("aria-label", `Download ${row.Name}`);
            dl.append(icon("download", 13));
            if (!spWrite.isMock()) dl.setAttribute("download", row.Name);
            dl.addEventListener("click", (e) => e.stopPropagation());
            span.append(dl);
            const link = document.createElement("button");
            link.type = "button";
            link.className = "wb-cell-link wb-cell-copylink";
            link.title = "Copy the direct URL";
            link.setAttribute("aria-label", `Copy the direct URL for ${row.Name}`);
            link.append(icon("link", 13));
            link.addEventListener("click", (e) => {
              e.stopPropagation();
              const origin = new URL(client2.webUrl()).origin;
              copyText(`${origin}${encodedServerPath2(serverRelativeUrl2)}`, link);
            });
            span.append(link);
            return span;
          }
        }
      ],
      rowKey: "ServerRelativeUrl",
      onOpen: (row) => {
        if (row.kind === "folder") navigate({ view: "files", path: row.ServerRelativeUrl });
        else openMetadata(row);
      },
      emptyText: "This folder is empty.",
      filterPlaceholder: "Filter files\u2026",
      exportName: "sp-files"
    });
    const uploadBtn = el12("button", "btn btn-xs wb-primary", "Upload\u2026");
    uploadBtn.type = "button";
    const fileInput = el12("input");
    fileInput.type = "file";
    fileInput.hidden = true;
    fileInput.setAttribute("aria-label", "Choose a file to upload");
    uploadBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", () => {
      if (fileInput.files?.length) startUpload(fileInput.files[0]);
      fileInput.value = "";
    });
    const refreshBtn = el12("button", "btn btn-xs", "Refresh");
    refreshBtn.type = "button";
    refreshBtn.addEventListener("click", () => listFolder(currentPath, { force: true }));
    grid.actionsEl.prepend(uploadBtn, fileInput, refreshBtn);
    gridWrap.append(grid.el);
  }
  async function listFolder(path, { force = false } = {}) {
    void force;
    const run = ++listingRun;
    currentPath = checkedPath(path);
    renderCrumbs();
    metaPanel.hidden = true;
    consent.hidden = true;
    if (!grid) makeGrid();
    grid.setLoading("Loading folder\u2026");
    try {
      const [folders, files] = await Promise.all([
        client2.getAll(folderApi(currentPath, "/Folders"), { select: FOLDER_SELECT, top: 5e3 }),
        client2.getAll(folderApi(currentPath, "/Files"), { select: FILE_SELECT, top: 5e3 })
      ]);
      const sortByName = (a, b) => String(a.Name).localeCompare(String(b.Name), void 0, { sensitivity: "base" });
      if (run !== listingRun) return;
      currentListing = {
        folders: folders.items.map((f) => ({ ...f, kind: "folder" })).sort(sortByName),
        files: files.items.map((f) => ({ ...f, kind: "file" })).sort(sortByName)
      };
      grid.setRows([...currentListing.folders, ...currentListing.files], {
        partial: folders.partial || files.partial
      });
      const matching = [...librarySelect.options].find((o) => o.value && (currentPath === o.value || currentPath.startsWith(`${o.value}/`)));
      librarySelect.value = matching ? matching.value : "";
    } catch (err) {
      if (run !== listingRun) return;
      grid.setError(err);
    }
  }
  function showConsent(message, onConfirm) {
    consent.textContent = "";
    consent.hidden = false;
    consent.append(el12("span", "wb-consent-text", message));
    const replace = el12("button", "btn btn-xs", "Replace");
    replace.type = "button";
    replace.addEventListener("click", () => {
      consent.hidden = true;
      onConfirm();
    });
    const cancel = el12("button", "btn btn-xs", "Cancel");
    cancel.type = "button";
    cancel.addEventListener("click", () => {
      consent.hidden = true;
    });
    consent.append(replace, cancel);
  }
  function uploadNotice(message, isError = false) {
    consent.textContent = "";
    consent.hidden = false;
    consent.classList.toggle("wb-consent-error", isError);
    consent.append(el12("span", "wb-consent-text", message));
    const dismiss = el12("button", "btn btn-xs", "Dismiss");
    dismiss.type = "button";
    dismiss.addEventListener("click", () => {
      consent.hidden = true;
    });
    consent.append(dismiss);
  }
  async function startUpload(file) {
    const folderPath = currentPath;
    consent.classList.remove("wb-consent-error");
    if (file.size > MAX_UPLOAD_BYTES) {
      uploadNotice(
        `\u201C${file.name}\u201D is ${formatBytes(file.size)} \u2014 above the ${formatBytes(MAX_UPLOAD_BYTES)} upload limit.`,
        true
      );
      return;
    }
    const existing = currentListing.files.find(
      (f) => String(f.Name).toLowerCase() === file.name.toLowerCase()
    );
    if (existing) {
      showConsent(
        `\u201C${file.name}\u201D already exists in this folder. Replace it?`,
        () => doUpload(file, { overwrite: true, folderPath })
      );
      return;
    }
    await doUpload(file, { overwrite: false, folderPath });
  }
  async function doUpload(file, { overwrite, folderPath }) {
    uploadNotice(`Uploading \u201C${file.name}\u201D\u2026`);
    let data;
    try {
      data = await file.arrayBuffer();
    } catch (err) {
      uploadNotice(`Could not read the file: ${err?.message || err}`, true);
      return;
    }
    try {
      const result = await spWrite.uploadFile(folderPath, file.name, data, { overwrite });
      consent.hidden = true;
      if (currentPath !== folderPath) {
        uploadNotice(`Uploaded \u201C${file.name}\u201D to ${folderPath}.`);
        return;
      }
      await listFolder(folderPath, { force: true });
      const uploaded = currentListing.files.find(
        (f) => String(f.Name).toLowerCase() === file.name.toLowerCase()
      ) || { kind: "file", Name: result.fileName, ServerRelativeUrl: result.serverRelativeUrl };
      openMetadata(uploaded, { justUploaded: true });
    } catch (err) {
      if (err?.code === "conflict" && !overwrite) {
        showConsent(
          `\u201C${file.name}\u201D already exists in this folder. Replace it?`,
          () => doUpload(file, { overwrite: true, folderPath })
        );
        return;
      }
      uploadNotice(`Upload failed: ${err?.message || err}`, true);
    }
  }
  function parentListId(folderPath) {
    const key2 = folderPath.toLowerCase();
    if (!parentListCache.has(key2)) {
      parentListCache.set(key2, client2.get(folderApi(folderPath, ""), {
        select: "ListItemAllFields/ParentList/Id",
        expand: "ListItemAllFields,ListItemAllFields/ParentList"
      }).catch(() => ({})).then(async (data) => {
        let id = String(
          data?.ListItemAllFields?.ParentList?.Id || data?.ListItemAllFields?.ParentList?.ID || ""
        ).replace(/[{}]/g, "").trim();
        if (!GUID.test(id)) {
          const aliasPath = `web/GetList(@listUrl)?@listUrl='${odataPathLiteral(folderPath)}'&$select=Id`;
          const viaUrl = await client2.get(aliasPath);
          id = String(viaUrl?.Id || viaUrl?.ID || "").replace(/[{}]/g, "").trim();
        }
        if (!GUID.test(id)) {
          throw new Error("SharePoint did not identify this folder\u2019s document library.");
        }
        return id;
      }).catch((err) => {
        parentListCache.delete(key2);
        throw err;
      }));
    }
    return parentListCache.get(key2);
  }
  function listFields(listId) {
    if (!fieldsCache.has(listId)) {
      fieldsCache.set(listId, client2.getAll(`web/lists(guid'${listId}')/fields`, {
        select: FIELD_SELECT4
      }).then(({ items }) => items).catch((err) => {
        fieldsCache.delete(listId);
        throw err;
      }));
    }
    return fieldsCache.get(listId);
  }
  async function openMetadata(row, { justUploaded = false } = {}) {
    metaPanel.hidden = false;
    metaPanel.textContent = "";
    const titleRow = el12("div", "wb-file-meta-head");
    titleRow.append(el12(
      "h3",
      "wb-subpanel-title",
      `${justUploaded ? "Uploaded \u2713 \u2014 metadata for" : "Metadata for"} ${row.Name}`
    ));
    const close = el12("button", "btn btn-xs", justUploaded ? "Keep without metadata" : "Close");
    close.type = "button";
    close.addEventListener("click", () => {
      metaPanel.hidden = true;
    });
    titleRow.append(close);
    metaPanel.append(titleRow);
    const body = el12("div", "wb-subpanel-body");
    metaPanel.append(body);
    const status = el12("div", "wb-grid-status", "Loading metadata\u2026");
    body.append(status);
    try {
      const listId = await parentListId(currentPath);
      const fields = await listFields(listId);
      let item2 = {};
      let itemAsText = {};
      try {
        item2 = await client2.get(fileApi(row.ServerRelativeUrl, "/ListItemAllFields"), {
          expand: "FieldValuesAsText"
        });
        itemAsText = item2.FieldValuesAsText || {};
      } catch {
        try {
          item2 = await client2.get(fileApi(row.ServerRelativeUrl, "/ListItemAllFields"));
        } catch {
          item2 = {};
        }
      }
      status.remove();
      const form = createFieldEditorForm({
        fields,
        item: item2,
        itemAsText,
        onSave: (formValues) => spWrite.validateUpdateListItem(
          { fileServerRelativeUrl: row.ServerRelativeUrl },
          formValues,
          { newDocumentUpdate: true }
        )
      });
      body.append(form.el);
    } catch (err) {
      status.textContent = justUploaded ? `The file was uploaded, but its metadata could not be loaded: ${err?.message || err}` : err?.message || String(err);
      status.classList.add("wb-error");
    }
  }
  async function load2(route) {
    await loadLibraries();
    let path = route?.path;
    if (!path) {
      path = libraries[0]?.RootFolder?.ServerRelativeUrl || webRootPath();
    }
    try {
      await listFolder(path);
    } catch (err) {
      if (!grid) makeGrid();
      grid.setError(err);
    }
  }
  return { el: root, load: load2 };
}

// ../src/state.js
var STORAGE_KEY = "dcspad.v2.workspace";
var DEFAULTS = {
  projectName: "",
  html: '<div id="app">\n  <h2>Hello from DCSPad</h2>\n  <p>Edit HTML, CSS and JS, then press Run.</p>\n</div>\n',
  css: 'body {\n  font-family: "Segoe UI", sans-serif;\n  padding: 1rem;\n}\n',
  js: 'console.log("DCSPad ready", { when: new Date().toISOString() });\n',
  libraries: { enabled: [], pinned: ["pnpjs2"], custom: [] },
  settings: {
    autorun: false,
    jsAsModule: false,
    autoClearConsole: true,
    seenSplash: false,
    previewDark: true,
    diagFontSize: 12,
    editorFontSize: 13,
    wordWrap: false,
    spFilesWebUrl: "",
    spFilesFolder: "",
    browserHistory: [],
    browserFavorites: [],
    projectFileFingerprint: ""
  },
  layout: {
    sidebarW: 230,
    sidebarCollapsed: false,
    editorsFr: 1,
    runtimeFr: 1,
    previewFr: 1,
    diagH: 260,
    diagCollapsed: false,
    editorTab: "js",
    diagTab: "console",
    // Pane visibility (the topbar segmented toggles). sidebarCollapsed /
    // diagCollapsed above are legacy flags kept for shape stability: layout.js
    // reads them once to seed `panes` for pre-existing workspaces, then only
    // writes `panes`.
    panes: { resources: true, preview: true, console: true },
    snippetsPanelH: 210
  }
};
var state = load();
var saveTimer = null;
var listeners = /* @__PURE__ */ new Set();
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return structuredClone(DEFAULTS);
    const parsed = JSON.parse(raw);
    return {
      ...structuredClone(DEFAULTS),
      ...parsed,
      libraries: { ...structuredClone(DEFAULTS.libraries), ...parsed.libraries || {} },
      settings: { ...structuredClone(DEFAULTS.settings), ...parsed.settings || {} },
      layout: { ...structuredClone(DEFAULTS.layout), ...parsed.layout || {} }
    };
  } catch {
    return structuredClone(DEFAULTS);
  }
}
function persist() {
  saveTimer = null;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    for (const fn of listeners) fn("saved");
  } catch (e) {
    console.warn("DCSPad: autosave failed", e);
    for (const fn of listeners) fn("error");
  }
}
function saveNow() {
  clearTimeout(saveTimer);
  persist();
}
function loadDoc(key2) {
  try {
    const raw = localStorage.getItem(key2);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && Array.isArray(parsed.items) ? parsed : null;
  } catch {
    return null;
  }
}
function saveDoc(key2, doc) {
  try {
    localStorage.setItem(key2, JSON.stringify(doc));
    return true;
  } catch (e) {
    console.warn(`DCSPad: saving ${key2} failed`, e);
    return false;
  }
}
var idSeed = Math.random().toString(36).slice(2, 6);
var idCounter = 0;
function newId(prefix) {
  return `${prefix}_${Date.now().toString(36)}${idSeed}${(idCounter++).toString(36)}`;
}
window.addEventListener("pagehide", () => {
  if (saveTimer) saveNow();
});

// ../src/workbench/favorites.js
var FAVORITES_KEY = "dcspad.v2.wbsites";
var RECENTS_CAP = 8;
function emptyDoc() {
  return { kind: "dcspad-workbench-sites", version: 1, items: [], recents: [] };
}
function readDoc() {
  const doc = loadDoc(FAVORITES_KEY);
  if (!doc || doc.kind !== "dcspad-workbench-sites") return emptyDoc();
  return {
    ...emptyDoc(),
    ...doc,
    items: Array.isArray(doc.items) ? doc.items : [],
    recents: Array.isArray(doc.recents) ? doc.recents : []
  };
}
var canonical = (url) => String(url || "").replace(/\/+$/, "").toLowerCase();
var quotaListener = null;
function onQuotaError(fn) {
  quotaListener = fn;
}
function writeDoc(doc) {
  if (!saveDoc(FAVORITES_KEY, doc)) quotaListener?.();
  return doc;
}
function getFavorites() {
  return readDoc().items;
}
function getRecents() {
  return readDoc().recents;
}
function isFavorite(url) {
  const key2 = canonical(url);
  return readDoc().items.some((item2) => canonical(item2.url) === key2);
}
function addFavorite({ url = "", title = "" } = {}) {
  const doc = readDoc();
  const key2 = canonical(url);
  if (doc.items.some((item2) => canonical(item2.url) === key2)) return doc.items;
  doc.items.push({ id: newId("fav"), url: String(url || ""), title: String(title || ""), addedAt: (/* @__PURE__ */ new Date()).toISOString() });
  return writeDoc(doc).items;
}
function removeFavorite(url) {
  const doc = readDoc();
  const key2 = canonical(url);
  doc.items = doc.items.filter((item2) => canonical(item2.url) !== key2);
  return writeDoc(doc).items;
}
function pushRecent({ url = "", title = "" } = {}) {
  const doc = readDoc();
  const key2 = canonical(url);
  doc.recents = [
    { url: String(url || ""), title: String(title || ""), lastAt: (/* @__PURE__ */ new Date()).toISOString() },
    ...doc.recents.filter((item2) => canonical(item2.url) !== key2)
  ].slice(0, RECENTS_CAP);
  return writeDoc(doc).recents;
}

// ../src/workbench/main.js
var GLYPHS = {
  lists: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" aria-hidden="true"><path d="M5.5 4h8M5.5 8h8M5.5 12h8"/><circle cx="2.7" cy="4" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="8" r=".9" fill="currentColor" stroke="none"/><circle cx="2.7" cy="12" r=".9" fill="currentColor" stroke="none"/></svg>',
  security: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 1.8 13 3.6v3.6c0 3.2-2.1 5.6-5 6.9-2.9-1.3-5-3.7-5-6.9V3.6z"/><path d="m5.8 7.8 1.6 1.6 2.9-3"/></svg>',
  site: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" aria-hidden="true"><path d="M2.2 13.3V6.5L8 2.3l5.8 4.2v6.8z"/><path d="M6.2 13.3V9.4h3.6v3.9"/></svg>',
  links: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6.5 9.5 9.5 6.5"/><path d="M7.5 4.6 9 3.1a2.6 2.6 0 0 1 3.7 0l.2.2a2.6 2.6 0 0 1 0 3.7L11.4 8.5"/><path d="M8.5 11.4 7 12.9a2.6 2.6 0 0 1-3.7 0l-.2-.2a2.6 2.6 0 0 1 0-3.7L4.6 7.5"/></svg>',
  query: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="7" cy="7" r="4.4"/><path d="m13.5 13.5-3.2-3.2"/><path d="M5.2 7h3.6M7 5.2v3.6"/></svg>',
  pages: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.4 1.8h6.4l2.8 2.8v9.6H3.4z"/><path d="M9.6 1.8v3h3"/><path d="M5.4 8h5.2M5.4 10.4h5.2"/></svg>',
  files: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1.8 4.2v8.2h12.4V5.8H7.6L6.2 4.2H1.8z"/><path d="M1.8 4.2V2.9h4.4l1.4 1.6"/></svg>',
  advanced: '<svg width="15" height="15" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="8" cy="8" r="2.2"/><path d="M8 1.9v2M8 12.1v2M1.9 8h2M12.1 8h2M3.7 3.7l1.4 1.4M10.9 10.9l1.4 1.4M12.3 3.7l-1.4 1.4M5.1 10.9l-1.4 1.4"/></svg>'
};
var SITE_KEY = "dcspad.workbench.site";
function applyWorkbenchContext(ctx2, inspecting = "") {
  const chip = document.getElementById("wb-chip");
  const chipText = document.getElementById("wb-chip-text");
  const statusCtx2 = document.getElementById("wb-status-context");
  chip.classList.toggle("sp-chip-live", ctx2.live);
  chip.classList.toggle("sp-chip-mock", !ctx2.live);
  chipText.textContent = ctx2.live ? "SP" : "SP: Mock";
  chip.title = ctx2.live ? `Connected to ${ctx2.label}${ctx2.user ? ` as ${ctx2.user}` : ""} \xB7 context: ${ctx2.source}` : "Not connected to a SharePoint web \u2014 showing built-in mock data";
  const inspectingNote = inspecting ? ` \xB7 inspecting ${inspecting}` : "";
  statusCtx2.textContent = ctx2.live ? `SP: ${ctx2.label}${ctx2.user ? ` \xB7 ${ctx2.user}` : ""}${inspectingNote}` : `SP: mock data (deploy to SharePoint for live inspection)${inspectingNote}`;
}
var ctx = getSpContext();
applyWorkbenchContext(ctx);
var client = createSpRestClient({
  mockResolver: ctx.live ? null : mockResolver
});
var shell = createShell({
  mount: document.getElementById("wb-main"),
  // inspectSite is a hoisted declaration below; views get a late-bound ref.
  deps: { client, inspectSite: (url) => inspectSite(url) },
  views: [
    // Nav order and grouping are Joe's spec (2026-07-31): identity first,
    // then content, then query, then jump-off/diagnostic sections.
    { id: "site", label: "Site", glyph: GLYPHS.site, group: "Site", create: createSiteHomeView },
    { id: "security", label: "Permissions", glyph: GLYPHS.security, group: "Site", create: createSecurityView },
    { id: "lists", label: "Lists", glyph: GLYPHS.lists, group: "Content", create: createListsView },
    { id: "pages", label: "Pages", glyph: GLYPHS.pages, group: "Content", create: createPagesView },
    { id: "files", label: "Files", glyph: GLYPHS.files, group: "Content", create: createBrowserView },
    { id: "query", label: "Query", glyph: GLYPHS.query, group: "Tools", create: createQueryView },
    { id: "links", label: "Panels", glyph: GLYPHS.links, group: "Tools", create: createLinksView },
    { id: "advanced", label: "Advanced", glyph: GLYPHS.advanced, group: "Tools", create: createSiteView }
  ]
});
var siteForm = document.getElementById("wb-site-form");
var siteInput = document.getElementById("wb-site-input");
var siteOpen = document.getElementById("wb-site-open");
var siteError = document.getElementById("wb-site-error");
function rememberSite(value) {
  try {
    if (value) sessionStorage.setItem(SITE_KEY, value);
    else sessionStorage.removeItem(SITE_KEY);
  } catch {
  }
}
async function inspectSite(input, { reset = true } = {}) {
  siteError.hidden = true;
  siteOpen.disabled = true;
  siteOpen.textContent = "Opening\u2026";
  try {
    const web = await client.connectWeb(input);
    const inspectingHost = client.webUrl() === client.hostWebUrl();
    siteInput.value = inspectingHost ? "" : client.webUrl();
    rememberSite(inspectingHost ? "" : client.webUrl());
    applyWorkbenchContext(ctx, inspectingHost ? "" : `${web?.Title || "web"} (${client.webUrl()})`);
    currentSite = {
      url: inspectingHost ? "" : client.webUrl(),
      title: web?.Title || (inspectingHost ? "This site" : client.webUrl())
    };
    pushRecent(currentSite);
    refreshFavStar();
    refreshCurrentUser();
    if (reset) shell.reset();
    return true;
  } catch (err) {
    siteError.textContent = err?.message || String(err);
    siteError.hidden = false;
    return false;
  } finally {
    siteOpen.disabled = false;
    siteOpen.textContent = "Inspect";
  }
}
siteForm.addEventListener("submit", (e) => {
  e.preventDefault();
  inspectSite(siteInput.value);
});
var favBtn = document.getElementById("wb-site-fav");
var favListBtn = document.getElementById("wb-site-favlist");
var favMenu = document.getElementById("wb-site-menu");
var statusCtx = document.getElementById("wb-status-context");
var currentSite = { url: "", title: "This site" };
onQuotaError(() => {
  const previous = statusCtx.textContent;
  statusCtx.textContent = "Could not save workbench favorites (storage quota).";
  setTimeout(() => {
    statusCtx.textContent = previous;
  }, 4e3);
});
function refreshFavStar() {
  const fav = isFavorite(currentSite.url);
  favBtn.textContent = fav ? "\u2605" : "\u2606";
  favBtn.classList.toggle("active", fav);
  favBtn.setAttribute("aria-pressed", fav ? "true" : "false");
  favBtn.title = fav ? "Remove the inspected site from favorites" : "Favorite the inspected site";
}
favBtn.addEventListener("click", () => {
  if (isFavorite(currentSite.url)) removeFavorite(currentSite.url);
  else addFavorite(currentSite);
  refreshFavStar();
});
function siteMenuItem(entry, hint) {
  const item2 = document.createElement("button");
  item2.type = "button";
  item2.className = "wb-menu-item";
  const label = entry.url ? entry.title || entry.url : entry.title || "This site";
  item2.textContent = entry.url ? `${label} \u2014 ${entry.url}` : `${label} (host web)`;
  if (hint) item2.title = hint;
  item2.addEventListener("click", () => {
    favMenu.hidden = true;
    siteInput.value = entry.url;
    inspectSite(entry.url);
  });
  return item2;
}
function menuHeading(text) {
  const h = document.createElement("div");
  h.className = "wb-menu-heading";
  h.textContent = text;
  return h;
}
function rebuildSiteMenu() {
  favMenu.textContent = "";
  const favorites = getFavorites();
  const recents = getRecents();
  if (favorites.length) {
    favMenu.append(menuHeading("Favorites"));
    for (const entry of favorites) favMenu.append(siteMenuItem(entry));
  }
  const favUrls = new Set(favorites.map((f) => (f.url || "").toLowerCase()));
  const rest = recents.filter((r) => !favUrls.has((r.url || "").toLowerCase()));
  if (rest.length) {
    favMenu.append(menuHeading("Recent"));
    for (const entry of rest) favMenu.append(siteMenuItem(entry));
  }
  if (!favMenu.childElementCount) {
    const empty = document.createElement("div");
    empty.className = "wb-menu-empty";
    empty.textContent = "No favorite or recent sites yet.";
    favMenu.append(empty);
  }
}
favListBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  if (favMenu.hidden) rebuildSiteMenu();
  favMenu.hidden = !favMenu.hidden;
});
document.addEventListener("click", () => {
  favMenu.hidden = true;
});
refreshFavStar();
var statusUser = document.getElementById("wb-status-user");
var statusRole = document.getElementById("wb-status-role");
async function refreshCurrentUser() {
  try {
    const user2 = await client.get("web/currentuser", {
      select: ["Title", "Email", "IsSiteAdmin"]
    });
    statusUser.textContent = user2.Email ? `${user2.Title} \xB7 ${user2.Email}` : user2.Title || "";
    statusRole.textContent = user2.IsSiteAdmin ? "Site admin" : "Site user";
    statusRole.className = user2.IsSiteAdmin ? "wb-role-chip wb-role-admin" : "wb-role-chip wb-role-user";
    statusRole.hidden = false;
  } catch {
    statusUser.textContent = "";
    statusRole.hidden = true;
  }
}
refreshCurrentUser();
(async () => {
  let saved = "";
  try {
    saved = sessionStorage.getItem(SITE_KEY) || "";
  } catch {
  }
  if (saved) {
    siteInput.value = saved;
    const ok = await inspectSite(saved, { reset: false });
    if (!ok) {
      rememberSite("");
      siteInput.value = "";
      try {
        sessionStorage.removeItem("dcspad.workbench.route");
      } catch {
      }
    }
  }
  shell.restore();
})();
