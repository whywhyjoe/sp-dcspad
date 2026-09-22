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

// ../src/build-info.js
var APP_VERSION = "1.0.0";
var injectedBuild = true ? "185" : "dev";
var injectedRevision = true ? "8c6f9088" : "";
var APP_BUILD_INFO = Object.freeze({
  version: APP_VERSION,
  build: injectedBuild,
  revision: injectedRevision
});
function buildTooltipFor(appName, info = APP_BUILD_INFO) {
  const revision = info.revision ? ` (${info.revision})` : "";
  return `${appName} \u2014 version ${info.version} \u2014 Build #${info.build}${revision}`;
}
function logBuildInfo(appName, info = APP_BUILD_INFO) {
  console.info(`[${appName}] version ${info.version} \u2014 Build #${info.build}${info.revision ? ` (${info.revision})` : ""}`);
}
function applyWorkbenchBuildMarker(root = document) {
  logBuildInfo("SP Workbench");
  const tooltip = buildTooltipFor("SP Workbench");
  const logo = root.querySelector(".wb-logo");
  if (logo) {
    logo.title = tooltip;
    logo.setAttribute("aria-label", tooltip);
  }
  document.documentElement.dataset.dcspadWbBuild = APP_BUILD_INFO.build;
  window.__DCSPAD_WB_BUILD_INFO__ = APP_BUILD_INFO;
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
async function responseErrorCode(response) {
  try {
    const body = await response.clone().json();
    return String(body?.error?.code || body?.["odata.error"]?.code || "");
  } catch {
    return "";
  }
}
function isCheckoutRefusal(detail, errorCode) {
  return /SPFileCheckOutException|-2147018029/i.test(String(errorCode || "")) || /is not checked out|must first check out|must be checked out|checked out for editing|currently checked out|is checked out (?:or locked )?(?:for editing )?by|locked for (?:shared|exclusive) use/i.test(String(detail || ""));
}
async function requireOk(response, fallback, code) {
  if (response.ok) return response;
  const detail = await responseMessage(response);
  let message = detail || `${fallback} (HTTP ${response.status})`;
  let normalizedCode = code;
  if (response.status === 401) {
    message = detail || "SharePoint could not authenticate this request. Reload the page to sign in again.";
    normalizedCode = "auth";
  } else if ([403, 409, 423].includes(response.status) && isCheckoutRefusal(detail, await responseErrorCode(response))) {
    message = detail || "This file must be checked out before it can be changed.";
    normalizedCode = "checkout-required";
  } else if (response.status === 403) {
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
var LARGE_PAGE_CAP = 1e5;
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
  async function getAll(path, opts, { cap = PAGE_CAP, allowLargeCap = false } = {}) {
    const ceiling = allowLargeCap ? LARGE_PAGE_CAP : PAGE_CAP;
    const limit = Math.min(Math.max(1, Number(cap) || ceiling), ceiling);
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
      const remaining = limit - items.length;
      if (page.length > remaining) {
        items.push(...page.slice(0, remaining));
        partial = true;
        break;
      }
      items.push(...page);
      const next = nextLinkOf(data);
      if (!next) break;
      if (items.length >= limit) {
        partial = true;
        break;
      }
      url = next;
    }
    return { items, partial };
  }
  return { context, webUrl, hostWebUrl, connectWeb, apiUrl, get, getAll };
}

// ../src/io.js?v=2
var MAX_IMPORT_BYTES = 5 * 1024 * 1024;
function saveBlob(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1e3);
}
function downloadText(filename, text, type = "application/json") {
  saveBlob(filename, new Blob([text], { type }));
}
function downloadBytes(filename, bytes, type = "application/octet-stream") {
  saveBlob(filename, new Blob([bytes], { type }));
}
var BUILT_IN_SHAREPOINT_FILE_TYPES = Object.freeze([
  Object.freeze({
    id: "html",
    label: "HTML",
    extensions: Object.freeze(["html", "htm"]),
    pane: "html",
    defaultExtension: "html"
  }),
  Object.freeze({
    id: "css",
    label: "CSS",
    extensions: Object.freeze(["css"]),
    pane: "css",
    defaultExtension: "css"
  }),
  Object.freeze({
    id: "javascript",
    label: "JavaScript",
    extensions: Object.freeze(["js"]),
    pane: "js",
    defaultExtension: "js"
  })
]);
function sharePointFileTypes(additionalTypes = []) {
  const usedExtensions = new Set(
    BUILT_IN_SHAREPOINT_FILE_TYPES.flatMap((type) => type.extensions)
  );
  const types = [...BUILT_IN_SHAREPOINT_FILE_TYPES];
  for (const [index, raw] of additionalTypes.entries()) {
    if (!raw || typeof raw !== "object") continue;
    const extensions = (Array.isArray(raw.extensions) ? raw.extensions : []).map((extension) => String(extension || "").trim().replace(/^\./, "").toLowerCase()).filter((extension) => extension && !usedExtensions.has(extension));
    if (!extensions.length || !["html", "css", "js"].includes(raw.pane)) continue;
    for (const extension of extensions) usedExtensions.add(extension);
    types.push(Object.freeze({
      id: `additional-${index}-${extensions[0]}`,
      label: String(raw.label || "").trim() || extensions[0].toUpperCase(),
      extensions: Object.freeze(extensions),
      pane: raw.pane,
      defaultExtension: extensions[0]
    }));
  }
  return types;
}
function fileTypeForFileName(fileName, additionalTypes = []) {
  const match = /\.([^.]+)$/i.exec(String(fileName || "").trim());
  const extension = match?.[1]?.toLowerCase();
  if (!extension) return null;
  return sharePointFileTypes(additionalTypes).find((type) => type.extensions.includes(extension)) || null;
}

// ../src/sp-files.js
var DIGEST_SAFETY_MS = 6e4;
var LIBRARY_GUID = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
var CHECK_IN_COMMENT = "Saved from DCSPad";
var CHECK_OUT_TYPE_NONE = 2;
function isCheckedOut(checkOutType) {
  const type = Number(checkOutType ?? CHECK_OUT_TYPE_NONE);
  return Number.isFinite(type) && type !== CHECK_OUT_TYPE_NONE;
}
function isCheckedOutByCurrentUser(user2, pageContext = {}, { sameWeb = true } = {}) {
  if (!user2) return false;
  const login = String(pageContext?.userLoginName || "").trim().toLowerCase();
  const claim = String(user2.LoginName || "").trim().toLowerCase();
  if (login && claim && (login === claim || claim.endsWith(`|${login}`) || login.endsWith(`|${claim}`))) {
    return true;
  }
  const email = String(pageContext?.userEmail || "").trim().toLowerCase();
  const userEmail = String(user2.Email || user2.UserPrincipalName || "").trim().toLowerCase();
  if (email && userEmail) return email === userEmail;
  const id = Number(pageContext?.userId);
  if (sameWeb && Number.isFinite(id) && id > 0) return id === Number(user2.Id);
  return false;
}
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
  async function listFolder(serverRelativePath, { webUrl: targetWebUrl = "", purpose = "code", additionalTypes = [] } = {}) {
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
    const files = resultArray(data.Files).map((item2) => {
      const fileType = fileTypeForFileName(item2.Name, additionalTypes);
      return {
        kind: "file",
        name: String(item2.Name || ""),
        pane: fileType?.pane || "",
        fileType,
        browserType: browserTypeForFileName(item2.Name),
        serverRelativeUrl: checkedPath(item2.ServerRelativeUrl, rootPath),
        length: Number(item2.Length) || 0,
        modified: item2.TimeLastModified || ""
      };
    }).filter((item2) => item2.name && (purpose === "browser" ? item2.browserType : item2.fileType)).sort((a, b) => a.name.localeCompare(b.name, void 0, { sensitivity: "base" }));
    return {
      path: checkedPath(data.ServerRelativeUrl || path, rootPath),
      rootPath,
      folders,
      files
    };
  }
  async function readTextFile(serverRelativePath, { webUrl: targetWebUrl = "", additionalTypes = [] } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const fileType = fileTypeForFileName(path, additionalTypes);
    if (!fileType) {
      throw new SpFileError(
        "That file type is not supported for SharePoint import.",
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
      pane: fileType.pane,
      fileType,
      text,
      serverRelativeUrl: path
    };
  }
  async function checkOutState({ webUrl, hostWebUrl, rootPath, libraryId, filePath, ctx: ctx2 }) {
    const state2 = {
      required: false,
      known: false,
      checkedOut: false,
      checkedOutByCurrentUser: false,
      checkedOutBy: "",
      reason: ""
    };
    try {
      const policyResponse = await request(
        `${webUrl}/_api/web/lists(guid'${libraryId}')?$select=ForceCheckout`,
        { headers: { Accept: ACCEPT_JSON } }
      );
      await requireOk(
        policyResponse,
        "Could not read the destination library check-out policy",
        "checkout-policy"
      );
      const list2 = unwrapJson(await policyResponse.json()) || {};
      state2.required = Boolean(list2.ForceCheckout ?? list2.forceCheckout);
      state2.known = true;
      if (!filePath) return state2;
      const path = checkedPath(filePath, rootPath);
      const fileResponse = await request(
        `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')?$select=CheckOutType,CheckedOutByUser/Id,CheckedOutByUser/Title,CheckedOutByUser/LoginName,CheckedOutByUser/Email&$expand=CheckedOutByUser`,
        { headers: { Accept: ACCEPT_JSON } }
      );
      await requireOk(
        fileResponse,
        "Could not read the destination file check-out state",
        "checkout-state"
      );
      const file = unwrapJson(await fileResponse.json()) || {};
      state2.checkedOut = isCheckedOut(file.CheckOutType ?? file.checkOutType);
      if (state2.checkedOut) {
        const user2 = file.CheckedOutByUser || file.checkedOutByUser || null;
        state2.checkedOutBy = String(user2?.Title || user2?.LoginName || "").trim();
        state2.checkedOutByCurrentUser = isCheckedOutByCurrentUser(
          user2,
          ctx2?.pageContext,
          { sameWeb: webUrl === hostWebUrl }
        );
      }
    } catch (error) {
      state2.known = false;
      state2.reason = String(error?.message || error);
    }
    return state2;
  }
  async function postFileMethod(webUrl, path, method, fallback, code) {
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')/${method}`;
    const attempt = async (forceDigest) => {
      const digest = await getDigest2({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: "POST",
        headers: { Accept: ACCEPT_JSON, "X-RequestDigest": digest }
      });
    };
    let response = await attempt(false);
    if (response.status === 403) response = await attempt(true);
    await requireOk(response, fallback, code);
  }
  async function checkOutFile(serverRelativePath, { webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    await postFileMethod(
      webUrl,
      path,
      "CheckOut()",
      "Could not check out the SharePoint file",
      "checkout"
    );
    return { serverRelativeUrl: path };
  }
  async function checkInFile(serverRelativePath, { comment = "", checkInType = 0, webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const stateResponse = await request(
      `${webUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(path)}')?$select=CheckOutType`,
      { headers: { Accept: ACCEPT_JSON } }
    );
    await requireOk(
      stateResponse,
      "Could not read the file check-out state",
      "checkout-state"
    );
    const file = unwrapJson(await stateResponse.json()) || {};
    if (!isCheckedOut(file.CheckOutType ?? file.checkOutType)) {
      return { serverRelativeUrl: path, checkedIn: false };
    }
    const type = [0, 1, 2].includes(Number(checkInType)) ? Number(checkInType) : 0;
    const safeComment = String(comment || "").slice(0, 1023);
    await postFileMethod(
      webUrl,
      path,
      `CheckIn(comment='${odataPathLiteral(safeComment)}',checkintype=${type})`,
      "Could not check in the SharePoint file",
      "checkin"
    );
    return { serverRelativeUrl: path, checkedIn: true };
  }
  async function undoCheckOutFile(serverRelativePath, { webUrl: targetWebUrl = "" } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    await postFileMethod(
      webUrl,
      path,
      "UndoCheckOut()",
      "Could not discard the check-out",
      "checkout-undo"
    );
    return { serverRelativeUrl: path };
  }
  async function inspectFileMetadata(folderPath, { filePath = "", webUrl: targetWebUrl = "" } = {}) {
    const { ctx: ctx2, webUrl, rootPath, hostWebUrl } = webInfo(targetWebUrl);
    const folder = checkedPath(folderPath, rootPath);
    const libraryEndpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${odataPathLiteral(folder)}')?$select=ListItemAllFields/ParentList/Id&$expand=ListItemAllFields,ListItemAllFields/ParentList`;
    const libraryResponse = await request(libraryEndpoint, {
      headers: { Accept: ACCEPT_JSON }
    });
    const libraryData = libraryResponse.ok ? unwrapJson(await libraryResponse.json()) || {} : {};
    let libraryId = String(
      libraryData.ListItemAllFields?.ParentList?.Id || libraryData.ListItemAllFields?.ParentList?.ID || ""
    ).replace(/[{}]/g, "").trim();
    if (!LIBRARY_GUID.test(libraryId)) {
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
    if (!LIBRARY_GUID.test(libraryId)) {
      throw new SpFileError(
        "SharePoint did not identify the destination document library.",
        { code: "metadata-library" }
      );
    }
    const checkout = await checkOutState({
      webUrl,
      hostWebUrl,
      rootPath,
      libraryId,
      filePath,
      ctx: ctx2
    });
    try {
      return {
        fields: await inspectFields(webUrl, rootPath, libraryId, filePath),
        checkout
      };
    } catch (error) {
      if (error && typeof error === "object") error.checkout = checkout;
      throw error;
    }
  }
  async function inspectFields(webUrl, rootPath, libraryId, filePath) {
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
    return fields;
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
        // bNewDocumentUpdate makes this the tail of the upload rather than a
        // new version — and on a checked-out file SharePoint checks it in as
        // part of the update, recording checkInComment.
        body: JSON.stringify({
          formValues,
          bNewDocumentUpdate: true,
          checkInComment: CHECK_IN_COMMENT
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
      serverRelativeUrl: result.ServerRelativeUrl || `${folder.replace(/\/$/, "")}/${safeName}`,
      // SP.File as returned by the upload: a new file in a ForceCheckout
      // library is born checked out. Undefined when the server doesn't say.
      checkOutType: result.CheckOutType
    };
  }
  return {
    webInfo,
    connectWeb,
    getDigest: getDigest2,
    listFolder,
    readTextFile,
    checkOutFile,
    checkInFile,
    undoCheckOutFile,
    inspectFileMetadata,
    writeFileMetadata,
    writeTextFile
  };
}
var defaultClient = createSpFilesClient();
var getDigest = (options) => defaultClient.getDigest(options);

// ../src/workbench/sp-write.js
var MAX_UPLOAD_BYTES = 50 * 1024 * 1024;
var mockNewItemId = 1e3;
var mockEnsuredUserId = 9e3;
function defaultMockWriter(url, body, contentType, headers) {
  const writes = globalThis.__DCSPAD_WB_WRITES__ ||= [];
  writes.push({ url, body, contentType, headers });
  const lower = String(url).toLowerCase();
  if (lower.includes("addvalidateupdateitemusingpath")) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const formValues = Array.isArray(data?.formValues) ? data.formValues : [];
    const id = String(mockNewItemId++);
    return {
      value: [
        ...formValues.map((fv) => ({
          FieldName: fv.FieldName,
          FieldValue: fv.FieldValue,
          HasException: false,
          ErrorMessage: null
        })),
        { FieldName: "Id", FieldValue: id, HasException: false, ErrorMessage: null }
      ]
    };
  }
  if (lower.includes("/ensureuser")) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const logon = String(data?.logonName || "");
    const email = logon.includes("@") ? logon : `${logon.replace(/[^a-z0-9.]+/gi, ".")}@mock.local`;
    return {
      Id: mockEnsuredUserId++,
      Title: logon,
      LoginName: `i:0#.f|membership|${email.toLowerCase()}`,
      Email: email
    };
  }
  if (lower.includes("attachmentfiles/add(")) {
    const name = /attachmentfiles\/add\(filename='([^']*)'\)/.exec(lower)?.[1] || "file";
    let decoded = name;
    try {
      decoded = decodeURIComponent(name);
    } catch {
    }
    return { FileName: decoded, ServerRelativeUrl: `/mock/attachments/${decoded}` };
  }
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
  mockWriter: mockWriter2 = null
} = {}) {
  const isMock = () => !client2.context().live;
  async function post(url, { body, contentType = "application/json;odata=nometadata", headers = {} } = {}, {
    fallback = "SharePoint write failed",
    code = "write"
  } = {}) {
    if (isMock()) {
      return structuredClone((mockWriter2 || defaultMockWriter)(url, body, contentType, headers));
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
            "X-RequestDigest": digest,
            ...headers
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
    if (response.status === 429 || response.status === 503) {
      const after = Number(response.headers.get("Retry-After")) || 2;
      await new Promise((r) => setTimeout(r, Math.min(after, 30) * 1e3));
      response = await attempt(false);
    }
    await requireOk(response, fallback, code);
    try {
      return unwrapJson(await response.json()) || {};
    } catch {
      return {};
    }
  }
  async function validateUpdateListItem(pathKind, formValues, { newDocumentUpdate = false, checkInComment = "" } = {}) {
    if (!Array.isArray(formValues) || !formValues.length) return { updated: [] };
    const base = `${client2.webUrl()}/_api/web`;
    const endpoint = pathKind.fileServerRelativeUrl ? `${base}/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(pathKind.fileServerRelativeUrl)}')/ListItemAllFields/ValidateUpdateListItem` : `${base}/lists(guid'${pathKind.listId}')/items(${Number(pathKind.itemId)})/ValidateUpdateListItem`;
    const data = await post(endpoint, {
      // With bNewDocumentUpdate SharePoint checks a checked-out file in as
      // part of the update; checkInComment is what it records when it does.
      body: JSON.stringify({
        formValues,
        bNewDocumentUpdate: Boolean(newDocumentUpdate),
        ...newDocumentUpdate && checkInComment ? { checkInComment } : {}
      })
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
  async function addValidateUpdateItem(listId, {
    folderPath,
    formValues,
    underlyingObjectType = 0,
    leafName
  } = {}) {
    const endpoint = `${client2.webUrl()}/_api/web/lists(guid'${listId}')/AddValidateUpdateItemUsingPath`;
    const data = await post(endpoint, {
      body: JSON.stringify({
        listItemCreateInfo: {
          FolderPath: { DecodedUrl: folderPath },
          UnderlyingObjectType: underlyingObjectType,
          ...leafName ? { LeafName: { DecodedUrl: leafName } } : {}
        },
        formValues,
        bNewDocumentUpdate: false
      })
    }, { fallback: "Could not create the item", code: "metadata-write" });
    const results = resultArray(data.value || data.AddValidateUpdateItemUsingPath || data);
    const failures = results.filter((result) => result.HasException || String(result.ErrorMessage || "").trim());
    if (failures.length) {
      const fieldErrors = {};
      for (const failure of failures) {
        fieldErrors[failure.FieldName || ""] = failure.ErrorMessage || "SharePoint rejected the value.";
      }
      const detail = failures.map((f) => `${f.FieldName || "Field"}: ${f.ErrorMessage || "SharePoint rejected the value."}`).join(" ");
      const err = new SpFileError(`SharePoint rejected the metadata. ${detail}`, { code: "metadata-write" });
      err.fieldErrors = fieldErrors;
      throw err;
    }
    const idRow = results.find((r) => /^id$/i.test(r.FieldName || ""));
    const id = idRow ? Number(idRow.FieldValue) : null;
    if (!id) {
      throw new SpFileError("SharePoint did not return the new item\u2019s id.", { code: "metadata-write" });
    }
    return { id };
  }
  const ensureUserCache = /* @__PURE__ */ new Map();
  async function ensureUser(logonName) {
    const key2 = String(logonName || "").toLowerCase();
    if (!key2) return null;
    if (ensureUserCache.has(key2)) return ensureUserCache.get(key2);
    const promise = (async () => {
      const data = await post(`${client2.webUrl()}/_api/web/ensureuser`, {
        body: JSON.stringify({ logonName })
      }, { fallback: "Could not resolve the user", code: "write" });
      return { id: data.Id, loginName: data.LoginName || "", email: data.Email || "", title: data.Title || "" };
    })();
    ensureUserCache.set(key2, promise);
    try {
      return await promise;
    } catch (err) {
      ensureUserCache.delete(key2);
      throw err;
    }
  }
  async function addAttachment(listId, itemId, fileName, bytes) {
    const endpoint = `${client2.webUrl()}/_api/web/lists(guid'${listId}')/items(${Number(itemId)})/AttachmentFiles/add(FileName='${odataPathLiteral(fileName)}')`;
    const data = await post(
      endpoint,
      { body: bytes, contentType: "application/octet-stream" },
      { fallback: "Could not add the attachment", code: "write" }
    );
    return { fileName, serverRelativeUrl: data.ServerRelativeUrl || "" };
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
      serverRelativeUrl: result.ServerRelativeUrl || `${folder === "/" ? "" : folder}/${safeName}`,
      // SP.File as returned by the upload: a new file in a ForceCheckout
      // library is born checked out, and an overwrite leaves a check-out
      // standing. Undefined when the server (or the mock) doesn't say.
      checkOutType: result.CheckOutType
    };
  }
  async function checkOutFile(fileServerRelativeUrl) {
    const endpoint = `${client2.webUrl()}/_api/web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(fileServerRelativeUrl)}')/CheckOut()`;
    await post(
      endpoint,
      { body: "" },
      { fallback: "Could not check out the file", code: "checkout" }
    );
    return { serverRelativeUrl: fileServerRelativeUrl };
  }
  async function checkInFile(fileServerRelativeUrl, { comment = "" } = {}) {
    const file = `web/GetFileByServerRelativePath(decodedUrl='${odataPathLiteral(fileServerRelativeUrl)}')`;
    if (!isMock()) {
      const state2 = await client2.get(file, { select: "CheckOutType" });
      if (!isCheckedOut(state2?.CheckOutType)) {
        return { serverRelativeUrl: fileServerRelativeUrl, checkedIn: false };
      }
    }
    const safeComment = String(comment || "").slice(0, 1023);
    await post(
      `${client2.webUrl()}/_api/${file}/CheckIn(comment='${odataPathLiteral(safeComment)}',checkintype=0)`,
      { body: "" },
      { fallback: "Could not check in the file", code: "checkin" }
    );
    return { serverRelativeUrl: fileServerRelativeUrl, checkedIn: true };
  }
  async function createFolder(parentServerRelativeUrl, name) {
    const clean = String(name || "").trim();
    if (!clean || /["*:<>?/\\|]/.test(clean) || clean.startsWith(".") || clean.endsWith(".")) {
      throw new SpFileError(
        'Folder names cannot contain " * : < > ? / \\ | or start or end with a dot.',
        { code: "invalid-name" }
      );
    }
    if (/^(CON|PRN|AUX|NUL|COM\d|LPT\d)(\..*)?$/i.test(clean) || /_vti_/i.test(clean)) {
      throw new SpFileError(
        "That folder name is reserved by SharePoint.",
        { code: "invalid-name" }
      );
    }
    const parent = String(parentServerRelativeUrl || "/").replace(/\/+$/, "") || "";
    const path = `${parent}/${clean}`;
    const endpoint = `${client2.webUrl()}/_api/web/Folders/AddUsingPath(decodedUrl='${odataPathLiteral(path)}')`;
    await post(
      endpoint,
      { body: "" },
      { fallback: "Could not create the folder", code: "write" }
    );
    return { name: clean, serverRelativeUrl: path };
  }
  async function postJson(path, body = {}, {
    fallback = "SharePoint write failed",
    code = "write",
    headers = {},
    contentType
  } = {}) {
    const url = `${client2.webUrl()}/_api/${String(path).replace(/^\/+/, "")}`;
    return post(url, {
      body: JSON.stringify(body),
      headers,
      ...contentType ? { contentType } : {}
    }, { fallback, code });
  }
  async function mergeJson(path, body = {}, {
    fallback = "SharePoint write failed",
    code = "write",
    headers = {},
    contentType
  } = {}) {
    const url = `${client2.webUrl()}/_api/${String(path).replace(/^\/+/, "")}`;
    return post(url, {
      body: JSON.stringify(body),
      headers: { "X-HTTP-Method": "MERGE", "IF-MATCH": "*", ...headers },
      ...contentType ? { contentType } : {}
    }, { fallback, code });
  }
  return {
    validateUpdateListItem,
    addValidateUpdateItem,
    ensureUser,
    addAttachment,
    uploadFile,
    checkOutFile,
    checkInFile,
    createFolder,
    postJson,
    mergeJson,
    isMock
  };
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
    field("Details", "ProjectDetails", "Note", 3, { RichText: true }),
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
function schemaField(title, internal, type, kind, { xmlAttrs = {}, xmlInner = "", ...extra } = {}) {
  fieldSeq++;
  const id = `bb1e2c1d-4444-5555-9999-${String(fieldSeq).padStart(12, "0")}`;
  const row = {
    Id: id,
    Title: title,
    InternalName: internal,
    StaticName: internal,
    TypeAsString: type,
    FieldTypeKind: kind,
    Required: false,
    Hidden: false,
    ReadOnlyField: false,
    FromBaseType: false,
    CanBeDeleted: true,
    Sealed: false,
    Group: "Custom Columns",
    DefaultValue: null,
    Description: "",
    EnforceUniqueValues: false,
    Indexed: false,
    CustomFormatter: "",
    LookupList: null,
    LookupField: null,
    IsDependentLookup: false,
    PrimaryFieldId: null,
    ...extra
  };
  const attrs = {
    ID: `{${id}}`,
    SourceID: "{deadbeef-0000-4000-8000-000000000000}",
    ColName: `tp_${internal}`,
    RowOrdinal: "0",
    Version: "1",
    Name: internal,
    StaticName: internal,
    DisplayName: title,
    Type: type,
    ...row.Required ? { Required: "TRUE" } : {},
    ...row.Indexed ? { Indexed: "TRUE" } : {},
    ...row.EnforceUniqueValues ? { EnforceUniqueValues: "TRUE" } : {},
    ...xmlAttrs
  };
  const attrText = Object.entries(attrs).map(([k, v]) => `${k}="${v}"`).join(" ");
  row.SchemaXml = xmlInner ? `<Field ${attrText}>${xmlInner}</Field>` : `<Field ${attrText} />`;
  return row;
}
var SCHEMA_REQUESTS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a01";
var SCHEMA_CLIENTS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a02";
var SCHEMA_REGIONS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a03";
var SCHEMA_DOCUMENTS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6a04";
var fRequestsTitle = schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false, Required: true });
var fRequestsId = schemaField("ID", "ID", "Counter", 5, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true });
var fStatus = schemaField("Status", "Status", "Choice", 6, { Choices: ["New", "Active", "Closed"], DefaultValue: "New" });
var fNotes = schemaField("Notes", "RequestNotes", "Note", 3, { xmlAttrs: { AppendOnly: "TRUE" } });
var fBudget = schemaField("Budget", "Budget", "Number", 9, {});
var fDue = schemaField("Due", "RequestDue", "DateTime", 4, {});
var fApproved = schemaField("Approved", "Approved", "Boolean", 8, {});
var fReference = schemaField("Reference", "Reference", "URL", 11, {});
var fOwner = schemaField("Owner", "RequestOwner", "User", 20, { xmlAttrs: { List: "UserInfo", ShowField: "Name" } });
var fClient = schemaField("Client", "Client", "Lookup", 7, {
  LookupList: SCHEMA_CLIENTS_ID,
  LookupField: "Title",
  xmlAttrs: { List: `{${SCHEMA_CLIENTS_ID}}`, ShowField: "Title" }
});
var fClientCode = schemaField("Client code", "ClientCode", "Lookup", 7, {
  LookupList: SCHEMA_CLIENTS_ID,
  LookupField: "ClientCode",
  IsDependentLookup: true,
  PrimaryFieldId: fClient.Id,
  xmlAttrs: { List: `{${SCHEMA_CLIENTS_ID}}`, ShowField: "ClientCode", FieldRef: `{${fClient.Id}}` }
});
var fParent = schemaField("Parent request", "ParentRequest", "Lookup", 7, {
  LookupList: SCHEMA_REQUESTS_ID,
  LookupField: "Title",
  xmlAttrs: { List: `{${SCHEMA_REQUESTS_ID}}`, ShowField: "Title" }
});
var fRegion = schemaField("Region", "Region", "Lookup", 7, {
  LookupList: SCHEMA_REGIONS_ID,
  LookupField: "Title",
  xmlAttrs: { List: `{${SCHEMA_REGIONS_ID}}`, ShowField: "Title" }
});
var fTotal = schemaField("Total", "Total", "Calculated", 12, {
  ReadOnlyField: true,
  xmlInner: `<Formula>=[Budget]*1</Formula><FieldRefs><FieldRef ID="{${fBudget.Id}}" Name="Budget"/></FieldRefs>`
});
var fTracking = schemaField("Tracking code", "TrackingCode", "Text", 2, { Indexed: true, EnforceUniqueValues: true });
var fFormatted = schemaField("Formatted note", "FormattedNote", "Text", 2, { CustomFormatter: '{"schema":"https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json"}' });
var fCategory = schemaField("Category", "RequestCategory", "TaxonomyFieldType", 26, {});
var fContentType = schemaField("Content Type", "ContentType", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, Hidden: true, ReadOnlyField: true });
var fAttachmentsF = schemaField("Attachments", "Attachments", "Attachments", 23, { FromBaseType: true, CanBeDeleted: false, Hidden: true });
var fAuthor = schemaField("Created By", "Author", "User", 20, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, xmlAttrs: { List: "UserInfo" } });
var fLinkTitle = schemaField("Title", "LinkTitle", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true });
var SCHEMA_REQUESTS_FIELDS = [
  fRequestsTitle,
  fRequestsId,
  fStatus,
  fNotes,
  fBudget,
  fDue,
  fApproved,
  fReference,
  fOwner,
  fClient,
  fClientCode,
  fParent,
  fRegion,
  fTotal,
  fTracking,
  fFormatted,
  fCategory,
  fContentType,
  fAttachmentsF,
  fAuthor,
  fLinkTitle
];
var SCHEMA_REQUESTS_LIST = {
  Id: SCHEMA_REQUESTS_ID,
  Title: "Requests",
  BaseTemplate: 100,
  BaseType: 0,
  ItemCount: 12,
  Hidden: false,
  Created: "2025-01-05T00:00:00Z",
  LastItemModifiedDate: "2026-08-01T00:00:00Z",
  EntityTypeName: "Requests",
  Description: "Schema source list.",
  DefaultViewUrl: "/sites/schema/Lists/Requests/AllItems.aspx",
  RootFolder: { ServerRelativeUrl: "/sites/schema/Lists/Requests", Name: "Requests" },
  ContentTypesEnabled: true,
  EnableVersioning: true,
  MajorVersionLimit: 50,
  EnableMinorVersions: false,
  ForceCheckout: false,
  EnableAttachments: true,
  EnableFolderCreation: false,
  EnableModeration: false,
  OnQuickLaunch: true,
  ValidationFormula: "=[Budget]>0",
  ValidationMessage: "Budget must be positive.",
  NoCrawl: false,
  DisableGridEditing: false,
  Ordered: false,
  ReadSecurity: 1,
  WriteSecurity: 1,
  ListExperienceOptions: 0,
  EnableRequestSignOff: false
};
var SCHEMA_CLIENTS_LIST = list("Clients", SCHEMA_CLIENTS_ID, 100, 0, 5, false, "/sites/schema/Lists/Clients");
var SCHEMA_REGIONS_LIST = list("Regions", SCHEMA_REGIONS_ID, 100, 0, 4, false, "/sites/schema/Lists/Regions");
var SCHEMA_DOCUMENTS_LIST = {
  Id: SCHEMA_DOCUMENTS_ID,
  Title: "Documents",
  BaseTemplate: 101,
  BaseType: 1,
  ItemCount: 6,
  Hidden: false,
  Created: "2025-02-01T00:00:00Z",
  LastItemModifiedDate: "2026-08-01T00:00:00Z",
  EntityTypeName: "Documents",
  Description: "Schema source library.",
  DefaultViewUrl: "/sites/schema/Documents/Forms/AllItems.aspx",
  RootFolder: { ServerRelativeUrl: "/sites/schema/Documents", Name: "Documents" },
  ContentTypesEnabled: true,
  EnableVersioning: true,
  MajorVersionLimit: 20,
  EnableMinorVersions: true,
  MajorWithMinorVersionsLimit: 10,
  DraftVersionVisibility: 0,
  ForceCheckout: true,
  EnableAttachments: true,
  EnableFolderCreation: true,
  EnableModeration: false,
  OnQuickLaunch: true,
  ValidationFormula: "",
  ValidationMessage: "",
  NoCrawl: false,
  DisableGridEditing: false,
  Ordered: false,
  ReadSecurity: 1,
  WriteSecurity: 1,
  ListExperienceOptions: 0,
  EnableRequestSignOff: false,
  DocumentTemplateUrl: "/sites/schema/Documents/Forms/custom-template.dotx",
  IrmEnabled: false
};
var SCHEMA_LISTS = [SCHEMA_REQUESTS_LIST, SCHEMA_CLIENTS_LIST, SCHEMA_REGIONS_LIST, SCHEMA_DOCUMENTS_LIST];
var DOC_PARENT_CT = "0x010100442912F2B6C7409A8FF25CE5504F1FE";
var DOC_LIST_CT = `${DOC_PARENT_CT}00${"B".repeat(32)}`;
var SCHEMA_DOCUMENTS_FIELDS = [
  schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false }),
  schemaField("ID", "ID", "Counter", 5, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true }),
  schemaField("Document category", "DocCategory", "Choice", 6, {
    Choices: ["Contract", "Report", "Misc"],
    DefaultValue: "Misc"
  }),
  schemaField("Name", "LinkFilename", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, Hidden: false }),
  schemaField("Type", "DocIcon", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, Hidden: true }),
  schemaField("File Size", "FileSizeDisplay", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, Hidden: true }),
  schemaField("Content Type", "ContentType", "Computed", 12, { FromBaseType: true, CanBeDeleted: false, Hidden: true, ReadOnlyField: true }),
  schemaField("Created By", "Author", "User", 20, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, xmlAttrs: { List: "UserInfo" } }),
  // Not Hidden on a real library (unlike DocIcon/FileSizeDisplay above), so
  // a real capture reaches these through doc.fields/baseFieldNames like any
  // other non-custom column — the buildApplyPlan view comment's own "base
  // Title/ID/Created/Modified/Author set already is assumed present" line.
  schemaField("Modified By", "Editor", "User", 20, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true, xmlAttrs: { List: "UserInfo" } }),
  schemaField("Created", "Created", "DateTime", 4, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true }),
  schemaField("Modified", "Modified", "DateTime", 4, { FromBaseType: true, CanBeDeleted: false, ReadOnlyField: true })
];
var SCHEMA_DOCUMENTS_VIEWS = [
  {
    Id: "cc1e2c1d-6666-7777-bbbb-000000000001",
    Title: "All Documents",
    DefaultView: true,
    PersonalView: false,
    Hidden: false,
    ServerRelativeUrl: "/sites/schema/Documents/Forms/AllItems.aspx",
    RowLimit: 30,
    Paged: true,
    ViewQuery: "",
    ViewFields: { Items: ["DocIcon", "LinkFilename", "DocCategory", "FileSizeDisplay", "Modified"] }
  },
  // A second, non-default view — grouped by kind of document, all columns
  // library base fields already carry. Exercises a plain (non-default)
  // view.upsert alongside the default-view match above so the library
  // Create check (tests/workbench-schema.mjs) can assert both recreated
  // views land with their exact field sequences and zero failed steps.
  {
    Id: "cc1e2c1d-6666-7777-bbbb-000000000002",
    Title: "By kind",
    DefaultView: false,
    PersonalView: false,
    Hidden: false,
    ServerRelativeUrl: "/sites/schema/Documents/Forms/ByKind.aspx",
    RowLimit: 30,
    Paged: true,
    ViewQuery: '<GroupBy Collapse="TRUE"><FieldRef Name="DocIcon"/></GroupBy>',
    ViewFields: { Items: ["DocIcon", "LinkFilename", "Modified", "Editor"] }
  }
];
var SCHEMA_DOCUMENTS_CTS = [
  { Id: { StringValue: DOC_LIST_CT }, Name: "Report", Group: "Custom Content Types", Hidden: false, ReadOnly: false, Sealed: false, Description: "A schema-source report document." },
  { Id: { StringValue: "0x0101" }, Name: "Document", Group: "Document Content Types", Hidden: false, ReadOnly: false, Sealed: false, Description: "Create a new document." }
];
var SCHEMA_FIELDS = {
  [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_FIELDS,
  [SCHEMA_CLIENTS_ID]: [
    schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false }),
    schemaField("Client code", "ClientCode", "Text", 2, {})
  ],
  [SCHEMA_REGIONS_ID]: [schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false })],
  [SCHEMA_DOCUMENTS_ID]: SCHEMA_DOCUMENTS_FIELDS
};
var SCHEMA_REQUESTS_VIEWS = [
  {
    Id: "cc1e2c1d-5555-6666-aaaa-000000000001",
    Title: "All Items",
    DefaultView: true,
    PersonalView: false,
    Hidden: false,
    ServerRelativeUrl: "/sites/schema/Lists/Requests/AllItems.aspx",
    RowLimit: 30,
    Paged: true,
    ViewQuery: '<OrderBy><FieldRef Name="ID"/></OrderBy>',
    ViewFields: { Items: ["LinkTitle", "Status", "Budget"] }
  },
  {
    Id: "cc1e2c1d-5555-6666-aaaa-000000000002",
    Title: "Active only",
    DefaultView: false,
    PersonalView: false,
    Hidden: false,
    ServerRelativeUrl: "/sites/schema/Lists/Requests/Active.aspx",
    RowLimit: 100,
    Paged: true,
    ViewQuery: '<Where><Eq><FieldRef Name="Status"/><Value Type="Choice">Active</Value></Eq></Where>',
    ViewFields: { Items: ["LinkTitle", "Status", "Client"] }
  }
];
var REQUEST_PARENT_CT = "0x0100442912F2B6C7409A8FF25CE5504F1FD";
var REQUEST_LIST_CT = `${REQUEST_PARENT_CT}00${"A".repeat(32)}`;
var SCHEMA_REQUESTS_CTS = [
  { Id: { StringValue: REQUEST_LIST_CT }, Name: "Request", Group: "Custom Content Types", Hidden: false, ReadOnly: false, Sealed: false, Description: "A schema-source request." },
  { Id: { StringValue: "0x01" }, Name: "Item", Group: "List Content Types", Hidden: false, ReadOnly: false, Sealed: false, Description: "Create a new list item." }
];
var VIEWS_BY_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_VIEWS, [SCHEMA_DOCUMENTS_ID]: SCHEMA_DOCUMENTS_VIEWS };
var CONTENT_TYPES_BY_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_CTS, [SCHEMA_DOCUMENTS_ID]: SCHEMA_DOCUMENTS_CTS };
var SCHEMA_REQUESTS_ITEMS = [
  item(1, "Server upgrade", {
    Status: "Active",
    Budget: 5e3,
    RequestDue: "2026-09-01T00:00:00Z",
    Approved: true,
    RequestOwnerId: 11,
    FSObjType: 0,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/1_.000`,
    FieldValuesAsText: { Status: "Active", RequestOwner: "Mock Developer", Budget: "5000", Approved: "Yes" }
  }),
  item(2, "Server upgrade \u2014 phase 2", {
    Status: "New",
    Budget: 2e3,
    ParentRequestId: 1,
    Attachments: true,
    FSObjType: 0,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/2_.000`,
    AttachmentFiles: [{
      FileName: "quote.pdf",
      ServerRelativeUrl: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/Attachments/2/quote.pdf`
    }],
    FieldValuesAsText: { Status: "New", ParentRequest: "Server upgrade" }
  }),
  item(3, "Archive", {
    FSObjType: 1,
    FileDirRef: SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl,
    FileRef: `${SCHEMA_REQUESTS_LIST.RootFolder.ServerRelativeUrl}/Archive`
  })
];
var ITEMS_BY_SCHEMA_LIST = { [SCHEMA_REQUESTS_ID]: SCHEMA_REQUESTS_ITEMS };
var SCHEMA_SITE_USERS = [
  user(11, "Mock Developer", "dev@mock.local", true),
  user(14, "Pat Example", "pat@mock.local", false)
];
var TARGET_CLIENTS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b01";
var TARGET_TASKS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b02";
var TARGET_DOCUMENTS_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b03";
var TARGET_ARCHIVE_ID = "5f8c6b7e-0d4a-4b6e-9f2e-1a2b3c4d6b04";
var TARGET_LISTS = [
  list("Clients", TARGET_CLIENTS_ID, 100, 0, 3, false, "/sites/target/Lists/Clients"),
  list("Tasks", TARGET_TASKS_ID, 171, 0, 10, false, "/sites/target/Lists/Tasks"),
  list("Documents", TARGET_DOCUMENTS_ID, 101, 1, 4, false, "/sites/target/Documents"),
  list("Archive Requests", TARGET_ARCHIVE_ID, 100, 0, 2, false, "/sites/target/Lists/ArchiveRequests")
];
var TARGET_FIELDS = {
  [TARGET_ARCHIVE_ID]: [
    schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false }),
    schemaField("Status", "Status", "Choice", 6, { Choices: ["New", "Active", "Closed"], DefaultValue: "New" })
  ],
  [TARGET_CLIENTS_ID]: [schemaField("Title", "Title", "Text", 2, { FromBaseType: true, CanBeDeleted: false })]
};
var TARGET_AVAILABLE_CTS = [
  { StringId: REQUEST_PARENT_CT, Name: "Request", Group: "Custom Content Types" },
  { StringId: "0x0101", Name: "Document", Group: "Document Content Types" },
  { StringId: DOC_PARENT_CT, Name: "Report", Group: "Custom Content Types" }
];
var writerState = {
  lists: /* @__PURE__ */ new Map(),
  // `${webBase}::${lower title}` -> list row
  fields: /* @__PURE__ */ new Map(),
  // listId -> Map(internalName -> field row)
  views: /* @__PURE__ */ new Map(),
  // listId -> Map(viewId -> { Id, Title, fields: [] })
  contentTypes: /* @__PURE__ */ new Map()
  // listId -> Set(parentId)
};
var webBaseOf = (url) => {
  const s = String(url);
  const i = s.indexOf("/_api/");
  return i === -1 ? "" : s.slice(0, i).replace(/\/+$/, "");
};
function xmlAttr(xml, name) {
  const m = new RegExp(`${name}="([^"]*)"`, "i").exec(String(xml || ""));
  return m ? m[1] : "";
}
var mockWriteSeq = 0;
var nextMockId = (prefix) => `${prefix}${String(++mockWriteSeq).padStart(8, "0")}`;
function baseFieldRow(internalName, title, type, extra = {}) {
  return {
    Id: nextMockId("bf"),
    InternalName: internalName,
    Title: title,
    TypeAsString: type,
    FromBaseType: true,
    CanBeDeleted: false,
    Hidden: false,
    ReadOnlyField: false,
    ...extra
  };
}
function baseFieldRows() {
  return [
    baseFieldRow("Title", "Title", "Text"),
    baseFieldRow("ID", "ID", "Counter", { ReadOnlyField: true }),
    baseFieldRow("LinkTitle", "Title", "Computed"),
    baseFieldRow("LinkTitleNoMenu", "Title", "Computed", { Hidden: true }),
    baseFieldRow("Attachments", "Attachments", "Attachments", { Hidden: true }),
    baseFieldRow("ContentType", "Content Type", "Computed", { Hidden: true }),
    baseFieldRow("Created", "Created", "DateTime", { ReadOnlyField: true }),
    baseFieldRow("Modified", "Modified", "DateTime", { ReadOnlyField: true }),
    baseFieldRow("Author", "Created By", "User", { ReadOnlyField: true }),
    baseFieldRow("Editor", "Modified By", "User", { ReadOnlyField: true })
  ];
}
function libraryBaseFieldRows() {
  return [
    baseFieldRow("FileLeafRef", "Name", "Computed", { ReadOnlyField: true }),
    baseFieldRow("LinkFilename", "Name", "Computed", { ReadOnlyField: true }),
    baseFieldRow("LinkFilenameNoMenu", "Name", "Computed", { Hidden: true, ReadOnlyField: true }),
    baseFieldRow("DocIcon", "Type", "Computed", { Hidden: true, ReadOnlyField: true }),
    baseFieldRow("FileSizeDisplay", "File Size", "Computed", { Hidden: true, ReadOnlyField: true }),
    baseFieldRow("Title", "Title", "Text", { Required: false }),
    baseFieldRow("_ExtendedDescription", "Description", "Note", { FromBaseType: false, CanBeDeleted: true }),
    baseFieldRow("Created", "Created", "DateTime", { ReadOnlyField: true }),
    baseFieldRow("Modified", "Modified", "DateTime", { ReadOnlyField: true }),
    baseFieldRow("Author", "Created By", "User", { ReadOnlyField: true }),
    baseFieldRow("Editor", "Modified By", "User", { ReadOnlyField: true }),
    baseFieldRow("ContentType", "Content Type", "Computed", { Hidden: true, ReadOnlyField: true }),
    baseFieldRow("ID", "ID", "Counter", { ReadOnlyField: true })
  ];
}
function mockWriter(url, body, contentType, headers = {}) {
  const webBase = webBaseOf(url);
  const path = String(url).slice(String(url).indexOf("/_api/") + 6).toLowerCase();
  const method = headers?.["X-HTTP-Method"] || headers?.["x-http-method"] || "";
  const record = () => {
    (globalThis.__DCSPAD_WB_WRITES__ ||= []).push({ url, body, contentType, headers });
  };
  if (/^web\/lists$/.test(path) && !method) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const id = nextMockId("cc00");
    let base = "";
    try {
      base = new URL(webBase).pathname.replace(/\/+$/, "");
    } catch {
    }
    const isLib = Number(data.BaseTemplate) === 101;
    const urlName = String(data.Title || "List").replace(/\s+/g, "");
    const rootUrl = isLib ? `${base}/${urlName}` : `${base}/Lists/${urlName}`;
    const entry = {
      Id: id,
      Title: data.Title,
      BaseTemplate: data.BaseTemplate ?? 100,
      // BaseType follows BaseTemplate: a document library (101) is BaseType
      // 1, so stage 1b-b's library gate (list-data-apply.js, still refusing
      // item import into a library) sees the created list correctly.
      BaseType: isLib ? 1 : 0,
      ContentTypesEnabled: !!data.ContentTypesEnabled,
      RootFolder: { ServerRelativeUrl: rootUrl }
    };
    writerState.lists.set(`${webBase}::${String(data.Title || "").toLowerCase()}`, entry);
    const fields = /* @__PURE__ */ new Map();
    for (const row of isLib ? libraryBaseFieldRows() : baseFieldRows()) fields.set(row.InternalName, row);
    writerState.fields.set(id, fields);
    const views = /* @__PURE__ */ new Map();
    const viewId = nextMockId("fa00");
    views.set(viewId, isLib ? { Id: viewId, Title: "All Documents", fields: ["DocIcon", "LinkFilename", "Modified", "Editor"], defaultView: true } : { Id: viewId, Title: "All Items", fields: ["LinkTitle"], defaultView: true });
    writerState.views.set(id, views);
    writerState.contentTypes.set(id, /* @__PURE__ */ new Set());
    record();
    return { Id: id, Title: entry.Title, RootFolder: entry.RootFolder };
  }
  const listIdMatch = /lists\(guid'([0-9a-f-]+)'\)/i.exec(path);
  const listId = listIdMatch?.[1];
  if (listId && method === "MERGE") {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    if (new RegExp(`^web/lists\\(guid'${listId}'\\)$`).test(path)) {
      const entry = [...writerState.lists.values()].find((l) => l.Id === listId);
      if (entry && data.Title !== void 0) entry.Title = data.Title;
      record();
      return {};
    }
    const fieldMerge = new RegExp(`^web/lists\\(guid'${listId}'\\)/fields\\(guid'([0-9a-f-]+)'\\)$`).exec(path);
    if (fieldMerge) {
      const fields = writerState.fields.get(listId);
      const field2 = fields && [...fields.values()].find((f) => f.Id === fieldMerge[1]);
      if (field2) {
        if (data.Indexed !== void 0) field2.Indexed = !!data.Indexed;
        if (data.EnforceUniqueValues !== void 0) field2.EnforceUniqueValues = !!data.EnforceUniqueValues;
        if (data.Title !== void 0) field2.Title = data.Title;
      }
      record();
      return {};
    }
    const viewMergeMatch = new RegExp(`^web/lists\\(guid'${listId}'\\)/views\\(guid'([0-9a-f-]+)'\\)$`).exec(path);
    if (viewMergeMatch) {
      const view = writerState.views.get(listId)?.get(viewMergeMatch[1]);
      if (view) {
        if (data.Title !== void 0) view.Title = data.Title;
        if (data.DefaultView !== void 0) view.defaultView = !!data.DefaultView;
      }
      record();
      return {};
    }
    record();
    return {};
  }
  if (listId && path.includes("createfieldasxml")) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const xml = data?.parameters?.SchemaXml || "";
    const optionsBits = Number(data?.parameters?.Options) || 0;
    let internalName = xmlAttr(xml, "Name") || xmlAttr(xml, "StaticName");
    const displayName = xmlAttr(xml, "DisplayName") || internalName;
    if (!(optionsBits & 8)) internalName = displayName.replace(/[^A-Za-z0-9]+/g, "_x0020_");
    const type = xmlAttr(xml, "Type");
    const fields = writerState.fields.get(listId) || /* @__PURE__ */ new Map();
    const id = nextMockId("ff00");
    fields.set(internalName, { Id: id, InternalName: internalName, Title: displayName, TypeAsString: type });
    writerState.fields.set(listId, fields);
    record();
    return { Id: id, InternalName: internalName };
  }
  if (listId && path.includes("addavailablecontenttype")) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const cts = writerState.contentTypes.get(listId) || /* @__PURE__ */ new Set();
    cts.add(data.contentTypeId);
    writerState.contentTypes.set(listId, cts);
    record();
    return {};
  }
  if (listId && /\/views$/.test(path) && !method) {
    let data = {};
    try {
      data = JSON.parse(body);
    } catch {
    }
    const id = nextMockId("fa00");
    const views = writerState.views.get(listId) || /* @__PURE__ */ new Map();
    views.set(id, { Id: id, Title: data.Title, fields: [], defaultView: !!data.DefaultView });
    writerState.views.set(listId, views);
    record();
    return { Id: id, Title: data.Title };
  }
  const viewMatch = /views\(guid'([0-9a-f-]+)'\)/i.exec(path);
  if (listId && viewMatch && path.includes("removeallviewfields")) {
    const view = writerState.views.get(listId)?.get(viewMatch[1]);
    if (view) view.fields = [];
    record();
    return {};
  }
  const addViewField = /addviewfield\('([^']*)'\)/i.exec(path);
  if (listId && viewMatch && addViewField) {
    const view = writerState.views.get(listId)?.get(viewMatch[1]);
    const castName = /addviewfield\('([^']*)'\)/i.exec(String(url))?.[1] ?? addViewField[1];
    if (view) view.fields.push(decodeURIComponent(castName));
    record();
    return {};
  }
  return defaultMockWriter(url, body, contentType, headers);
}
var PROJECT_ITEMS = [
  item(1, "Intranet refresh", {
    ProjectStatus: "Active",
    DueDate: "2026-09-15T00:00:00Z",
    Budget: 12e3,
    // Rich text + attachments: the item-export markdown path needs both.
    ProjectDetails: '<div><p>Kickoff <strong>done</strong>.</p><ul><li>Phase 1</li><li>Phase 2</li></ul><p>See the <a href="https://example.com/plan">plan</a>.</p></div>',
    AttachmentFiles: [
      { FileName: "kickoff.pptx", ServerRelativeUrl: "/Lists/Projects/Attachments/1/kickoff.pptx" }
    ],
    FieldValuesAsText: { Author: "Mock Developer", Editor: "Pat Example", ProjectOwner: "Mock Developer" }
  }),
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
  dcspad_x005f_deployfolder: "/Dev/tools/dcspad"
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
var CLASSIC_LISTS = [
  list("Documents", "7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f01", 101, 1, 8, false, "/sites/classic/Documents"),
  list("Pages", "7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f02", 850, 1, 4, false, "/sites/classic/Pages")
];
var CLASSIC_PAGE_ITEMS = [
  {
    Id: 1,
    Title: "Benefits overview",
    FileLeafRef: "Benefits.aspx",
    FileRef: "/sites/classic/Pages/Benefits.aspx",
    FileDirRef: "/sites/classic/Pages",
    UniqueId: "ef000000-0000-4000-8000-000000000001",
    Created: "2019-04-02T10:00:00Z",
    Modified: "2026-06-11T14:20:00Z",
    Author: { Title: "Pat Example" },
    Editor: { Title: "Pat Example" },
    PublishingPageContent: "<h2>Benefits</h2><p>Open enrollment runs through March.</p>",
    FieldValuesAsText: { Editor: "Pat Example" }
  },
  {
    // Body empty on purpose: everything readable is in the web parts.
    Id: 2,
    Title: "Rates",
    FileLeafRef: "Rates.aspx",
    FileRef: "/sites/classic/Pages/Rates.aspx",
    FileDirRef: "/sites/classic/Pages",
    UniqueId: "ef000000-0000-4000-8000-000000000002",
    Created: "2018-09-14T08:30:00Z",
    Modified: "2026-05-02T09:05:00Z",
    Author: { Title: "Mock Developer" },
    Editor: { Title: "Mock Developer" },
    PublishingPageContent: "",
    FieldValuesAsText: { Editor: "Mock Developer" }
  },
  {
    Id: 3,
    Title: "Empty",
    FileLeafRef: "Empty.aspx",
    FileRef: "/sites/classic/Pages/Empty.aspx",
    FileDirRef: "/sites/classic/Pages",
    UniqueId: "ef000000-0000-4000-8000-000000000003",
    Created: "2020-01-05T11:00:00Z",
    Modified: "2026-01-05T11:00:00Z",
    Author: { Title: "Mock Developer" },
    Editor: { Title: "Mock Developer" },
    PublishingPageContent: null,
    FieldValuesAsText: { Editor: "Mock Developer" }
  },
  {
    // Body with an EMBEDDED web part: rich bodies place web parts inside the
    // field as .ms-rte-wpbox markers, and reading order must interleave —
    // intro → web part → conclusion, never body-then-web-parts.
    Id: 4,
    Title: "Newsletter",
    FileLeafRef: "Newsletter.aspx",
    FileRef: "/sites/classic/Pages/Newsletter.aspx",
    FileDirRef: "/sites/classic/Pages",
    UniqueId: "ef000000-0000-4000-8000-000000000004",
    Created: "2021-03-09T09:00:00Z",
    Modified: "2026-07-30T16:45:00Z",
    Author: { Title: "Pat Example" },
    Editor: { Title: "Pat Example" },
    PublishingPageContent: '<p>Intro paragraph.</p><div class="ms-rtestate-read ms-rte-wpbox"><div id="div_c3000000-0000-4000-8000-000000000006"></div></div><p>Closing paragraph.</p>',
    FieldValuesAsText: { Editor: "Pat Example" }
  }
];
var CLASSIC_WEBPARTS = {
  "/sites/classic/pages/benefits.aspx": [
    {
      Id: "c3000000-0000-4000-8000-000000000001",
      WebPart: {
        Title: "Contact details",
        ZoneIndex: 2,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: "<p>Call the benefits desk on x4120.</p>", ContentLink: "" }
      }
    },
    {
      Id: "c3000000-0000-4000-8000-000000000002",
      WebPart: {
        Title: "Eligibility",
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: "<![CDATA[<p>All staff after 90 days.</p>]]>", ContentLink: "" }
      }
    }
  ],
  "/sites/classic/pages/rates.aspx": [
    {
      Id: "c3000000-0000-4000-8000-000000000003",
      WebPart: {
        Title: "Rate table",
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: "", ContentLink: "/sites/classic/Style Library/rates.html" }
      }
    },
    {
      Id: "c3000000-0000-4000-8000-000000000004",
      WebPart: {
        Title: "Rate calculator",
        ZoneIndex: 2,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: '<script>calcRates();<\/script><div id="calc">Rates</div>' }
      }
    },
    {
      // No Content and no ContentLink — an inventory row, not a reading part.
      Id: "c3000000-0000-4000-8000-000000000005",
      WebPart: { Title: "List view", ZoneIndex: 3, Hidden: true, IsClosed: false, Properties: { ListName: "Rates" } }
    }
  ],
  "/sites/classic/pages/newsletter.aspx": [
    {
      // Matched by the wpbox marker in the Newsletter body — must land
      // BETWEEN the intro and closing paragraphs, not after the body.
      Id: "c3000000-0000-4000-8000-000000000006",
      WebPart: {
        Title: "Signup form",
        ZoneIndex: 1,
        Hidden: false,
        IsClosed: false,
        Properties: { Content: "<p>Subscribe at the front desk.</p>", ContentLink: "" }
      }
    }
  ]
};
var CLASSIC_ITEMS = {
  "7a1c6b7e-0d4a-4b6e-9f2e-1a2b3c4d5f02": CLASSIC_PAGE_ITEMS
};
var BOTH_SITE_PAGES_ID = "9c2d4e6f-1111-4222-8333-44445555a001";
var BOTH_PAGES_ID = "9c2d4e6f-1111-4222-8333-44445555a002";
var BOTH_LISTS = [
  list("Site Pages", BOTH_SITE_PAGES_ID, 119, 1, 1, false, "/sites/both/SitePages"),
  list("Pages", BOTH_PAGES_ID, 850, 1, 2, false, "/sites/both/Pages")
];
var BOTH_FIELDS = {
  [BOTH_SITE_PAGES_ID]: [
    field("Title", "Title", "Text", 2, { Required: true }),
    field("Promoted state", "PromotedState", "Number", 9, { ReadOnlyField: true }),
    field("Canvas content", "CanvasContent1", "Note", 3),
    field("Editor", "Editor", "User", 20, { ReadOnlyField: true }),
    field("ID", "ID", "Counter", 5, { ReadOnlyField: true })
  ],
  [BOTH_PAGES_ID]: [
    field("Title", "Title", "Text", 2, { Required: true }),
    field("Page content", "PublishingPageContent", "Note", 3),
    field("Editor", "Editor", "User", 20, { ReadOnlyField: true }),
    field("ID", "ID", "Counter", 5, { ReadOnlyField: true })
  ]
};
var BOTH_ITEMS = {
  [BOTH_SITE_PAGES_ID]: [
    {
      Id: 1,
      Title: "Team news",
      FileLeafRef: "TeamNews.aspx",
      FileRef: "/sites/both/SitePages/TeamNews.aspx",
      FileDirRef: "/sites/both/SitePages",
      PromotedState: 0,
      Modified: "2026-07-18T10:00:00Z",
      Editor: { Title: "Mock Developer" },
      CanvasContent1: JSON.stringify([
        { controlType: 4, id: "c1", innerHTML: "<p>The one modern page.</p>" }
      ])
    }
  ],
  [BOTH_PAGES_ID]: [
    {
      Id: 1,
      Title: "Policies",
      FileLeafRef: "Policies.aspx",
      FileRef: "/sites/both/Pages/Policies.aspx",
      FileDirRef: "/sites/both/Pages",
      Modified: "2026-07-04T10:00:00Z",
      Editor: { Title: "Pat Example" },
      PublishingPageContent: "<h2>Policies</h2><p>Where the real content is.</p>"
    },
    {
      Id: 2,
      Title: "Handbook",
      FileLeafRef: "Handbook.aspx",
      FileRef: "/sites/both/Pages/Handbook.aspx",
      FileDirRef: "/sites/both/Pages",
      Modified: "2026-07-05T10:00:00Z",
      Editor: { Title: "Pat Example" },
      PublishingPageContent: "<p>Staff handbook.</p>"
    }
  ]
};
var listIdOf = (url) => /lists\(guid'([0-9a-f-]+)'\)/i.exec(url)?.[1]?.toLowerCase();
var groupIdOf = (url) => /sitegroups\((\d+)\)/i.exec(url)?.[1];
function mockResolver(rawUrl) {
  const url = String(rawUrl);
  const path = url.slice(url.indexOf("/_api/") + 6).toLowerCase();
  const webBase = url.slice(0, url.indexOf("/_api/")).replace(/[/]+$/, "");
  const classic = /[/]sites[/]classic$/i.test(webBase);
  const both = /[/]sites[/]both$/i.test(webBase);
  const schema = /[/]sites[/]schema$/i.test(webBase);
  const target = /[/]sites[/]target$/i.test(webBase);
  const staticLists = both ? BOTH_LISTS : classic ? CLASSIC_LISTS : schema ? SCHEMA_LISTS : target ? TARGET_LISTS : LISTS;
  const itemsByList = both ? BOTH_ITEMS : classic ? CLASSIC_ITEMS : schema ? ITEMS_BY_SCHEMA_LIST : ITEMS;
  const dynamicLists = [...writerState.lists.entries()].filter(([key2]) => key2.startsWith(`${webBase}::`)).map(([, entry]) => entry);
  const lists = [...staticLists, ...dynamicLists];
  const wpFile = /getfilebyserverrelativepath[(]decodedurl='([^']*)'[)][/]getlimitedwebpartmanager/.exec(path)?.[1];
  if (wpFile !== void 0 && path.includes("/webparts")) {
    let decoded = wpFile;
    try {
      decoded = decodeURIComponent(wpFile);
    } catch {
    }
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
    if (path.includes("/items")) {
      const rows = [...itemsByList[found.Id] || []];
      const order = /\$orderby=([a-z0-9_]+)(?:(?:%20| +)(asc|desc))?/.exec(path);
      if (order) {
        const key2 = Object.keys(rows[0] || {}).find((k) => k.toLowerCase() === order[1]) || order[1];
        const dir = order[2] === "desc" ? -1 : 1;
        rows.sort((a, b) => a[key2] > b[key2] ? dir : a[key2] < b[key2] ? -dir : 0);
      }
      return { value: rows };
    }
    if (path.includes("/fields")) {
      const dyn = writerState.fields.get(found.Id);
      if (dyn) return { value: [...dyn.values()] };
      const perWeb = both ? BOTH_FIELDS : schema ? SCHEMA_FIELDS : target ? TARGET_FIELDS : FIELDS;
      return { value: perWeb[found.Id] || DEFAULT_FIELDS };
    }
    if (/\/views\(guid'/.test(path) && path.includes("/viewfields")) {
      return { Items: ["LinkTitle", "ProjectStatus", "DueDate"] };
    }
    if (path.includes("/views")) {
      const dyn = writerState.views.get(found.Id);
      if (dyn) {
        return {
          value: [...dyn.values()].map((v) => ({
            Id: v.Id,
            Title: v.Title,
            DefaultView: !!v.defaultView,
            PersonalView: false,
            Hidden: false,
            RowLimit: 30,
            Paged: true,
            ViewQuery: "",
            ViewFields: { Items: v.fields }
          }))
        };
      }
      return { value: VIEWS_BY_LIST[found.Id] || VIEWS };
    }
    if (path.includes("/contenttypes")) {
      const dyn = writerState.contentTypes.get(found.Id);
      if (dyn) {
        return {
          value: [...dyn].map((id2) => ({
            Id: { StringValue: id2 },
            Name: id2,
            Group: "",
            Hidden: false,
            ReadOnly: false,
            Sealed: false,
            Description: ""
          }))
        };
      }
      return { value: CONTENT_TYPES_BY_LIST[found.Id] || CONTENT_TYPES };
    }
    if (path.includes("/roleassignments")) return { value: ROLE_ASSIGNMENTS };
    return found;
  }
  if (path.startsWith("web/lists")) {
    if (path.includes("hasuniqueroleassignments")) {
      return {
        value: lists.map((l, i) => ({
          Id: l.Id,
          Title: l.Title,
          Hidden: l.Hidden,
          BaseTemplate: l.BaseTemplate,
          HasUniqueRoleAssignments: i === 2
        }))
      };
    }
    return { value: lists };
  }
  if (/^web\/sitegroups\(\d+\)\/users/.test(path)) {
    const users = GROUP_USERS[groupIdOf(path)];
    return users ? { value: users } : { value: [] };
  }
  if (path.startsWith("web/sitegroups")) return { value: GROUPS };
  if (path.startsWith("web/roledefinitions")) return { value: ROLE_DEFINITIONS };
  if (path.startsWith("web/roleassignments")) return { value: ROLE_ASSIGNMENTS };
  const siteUserId = /^web\/siteusers\/getbyid\((\d+)\)/.exec(path)?.[1];
  if (siteUserId !== void 0) {
    const found = SCHEMA_SITE_USERS.find((u) => u.Id === Number(siteUserId));
    return found || null;
  }
  if (path.startsWith("web/siteusers")) return { value: SCHEMA_SITE_USERS };
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
  if (path.startsWith("web/availablecontenttypes")) {
    return { value: target ? TARGET_AVAILABLE_CTS : [] };
  }
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
    const inst = def.create({ ...deps, navigate, updateRoute });
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
  function updateRoute(patch) {
    if (!currentRoute || !patch) return;
    currentRoute = { ...currentRoute, ...patch };
    try {
      sessionStorage.setItem(ROUTE_KEY, JSON.stringify(currentRoute));
    } catch {
    }
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
  return { navigate, updateRoute, restore, reset, getRoute: () => currentRoute };
}

// ../src/workbench/denied.js
function isDeniedRead(err) {
  return err?.code === "permission" || err?.status === 403;
}
function isExpiredSession(err) {
  return err?.code === "auth" || err?.status === 401;
}
var EXPIRED_SESSION_NOTE = "Your SharePoint sign-in has expired \u2014 reload the page to sign in again.";
function deniedNote(subject = "") {
  return subject ? `Your account doesn\u2019t have permission to see ${subject} here.` : "Your account doesn\u2019t have permission to see this.";
}
function showFailure(node, err, subject = "") {
  const denied = isDeniedRead(err);
  const expired = !denied && isExpiredSession(err);
  node.textContent = denied ? deniedNote(subject) : expired ? EXPIRED_SESSION_NOTE : err?.message || String(err);
  node.classList.remove("wb-error", "wb-denied");
  node.classList.add(denied ? "wb-denied" : "wb-error");
  if ((denied || expired) && err?.message) node.title = err.message;
  else node.removeAttribute("title");
  node.hidden = false;
  return node;
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
function downloadCsv(name, rows, columns) {
  downloadText(`${name}.csv`, toCsv(rows, columns), "text/csv;charset=utf-8");
}
function downloadJson(name, rows, columns) {
  downloadText(`${name}.json`, toJson(rows, columns), "application/json");
}
function downloadMarkdown(name, text) {
  downloadText(`${name}.md`, text, "text/markdown;charset=utf-8");
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
  if (options.orderby) {
    for (const clause of join(options.orderby).split(",")) {
      const m = /^\s*(.+?)(?:\s+(asc|desc))?\s*$/i.exec(clause);
      if (!m || !m[1]) continue;
      chain += `
  .orderBy(${JSON.stringify(m[1])}, ${String(m[2]).toLowerCase() !== "desc"})`;
    }
  }
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
var encodeSpPath = (path) => String(path).split("/").map(encodeURIComponent).join("/");
function bindNewTab(a) {
  a.target = "_blank";
  a.rel = "noopener";
  a.addEventListener("click", (e) => {
    e.stopPropagation();
    if (e.ctrlKey || e.metaKey || e.shiftKey || e.altKey || e.button === 1) return;
    e.preventDefault();
    window.open(a.href, "_blank", "noopener");
  });
  return a;
}
function createMenuButton(label, title, items) {
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
  return wrap;
}
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
  descriptor = null,
  toolbarExtras = null,
  exportExtras = [],
  selectable = false,
  subject = ""
} = {}) {
  let rows = [];
  let visible = [];
  let sortKey = null;
  let sortDir = 1;
  let filterText = "";
  const selectedKeys = /* @__PURE__ */ new Set();
  const keyOf = (row) => String(row?.[rowKey] ?? "");
  const exportRows = () => {
    const chosen = visible.filter((row) => selectedKeys.has(keyOf(row)));
    return chosen.length ? chosen : visible;
  };
  const root = el2("div", "wb-grid");
  const toolbar = el2("div", "wb-grid-toolbar");
  const count = el2("span", "wb-grid-count", "\u2014");
  const filter = el2("input", "wb-grid-filter");
  filter.type = "search";
  filter.placeholder = filterPlaceholder;
  filter.setAttribute("aria-label", "Filter rows");
  const actions = el2("span", "wb-grid-actions");
  toolbar.append(count, filter, ...toolbarExtras ? [toolbarExtras] : [], actions);
  function menuButton(label, title, items) {
    actions.append(createMenuButton(label, title, items));
  }
  if (descriptor) {
    menuButton("Copy as \u25BE", "Copy this query as a runnable script", [
      ["PnPjs 2 (DCSPad pane)", (btn) => copyText(toPnpjs2(descriptor), btn)],
      ["REST fetch", (btn) => copyText(toRestFetch(descriptor, descriptor.webUrl), btn)],
      ["PnP.PowerShell", (btn) => copyText(toPnpPowerShell(descriptor, descriptor.webUrl), btn)]
    ]);
  }
  if (exportName) {
    menuButton("Export \u25BE", "Export the selected rows, or all visible rows", [
      ...exportExtras,
      ["Download CSV", () => downloadCsv(exportName, exportRows(), columns)],
      ["Download JSON", () => downloadJson(exportName, exportRows(), columns)],
      ["Copy CSV", (btn) => copyText(toCsv(exportRows(), columns), btn)],
      ["Copy JSON", (btn) => copyText(toJson(exportRows(), columns), btn)]
    ]);
  }
  const scroller = el2("div", "wb-grid-scroll");
  const table2 = el2("table", "wb-table");
  const thead = el2("thead");
  const headRow = el2("tr");
  let widthsFrozen = false;
  function freezeWidths() {
    if (widthsFrozen) return;
    widthsFrozen = true;
    for (const th of headRow.children) {
      th.style.width = `${Math.round(th.getBoundingClientRect().width)}px`;
    }
    table2.style.tableLayout = "fixed";
    syncTableWidth();
  }
  function syncTableWidth() {
    let total = 0;
    for (const th of headRow.children) total += parseFloat(th.style.width) || 0;
    if (total) table2.style.width = `${total}px`;
  }
  function attachResizer(th) {
    const handle = el2("span", "wb-col-resize");
    handle.title = "Drag to resize";
    handle.addEventListener("click", (e) => e.stopPropagation());
    handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      freezeWidths();
      const startX = e.clientX;
      const startW = parseFloat(th.style.width) || th.getBoundingClientRect().width;
      handle.classList.add("dragging");
      handle.setPointerCapture(e.pointerId);
      const move = (ev) => {
        th.style.width = `${Math.max(48, Math.round(startW + (ev.clientX - startX)))}px`;
        syncTableWidth();
      };
      const up = () => {
        handle.classList.remove("dragging");
        handle.removeEventListener("pointermove", move);
        handle.removeEventListener("pointerup", up);
      };
      handle.addEventListener("pointermove", move);
      handle.addEventListener("pointerup", up);
    });
    th.append(handle);
  }
  let selectAll = null;
  if (selectable) {
    const th = el2("th", "wb-select-col");
    selectAll = el2("input");
    selectAll.type = "checkbox";
    selectAll.className = "wb-select-all";
    selectAll.setAttribute("aria-label", "Select all visible rows");
    selectAll.title = "Select all visible rows";
    selectAll.addEventListener("change", () => {
      if (selectAll.checked) visible.forEach((row) => selectedKeys.add(keyOf(row)));
      else visible.forEach((row) => selectedKeys.delete(keyOf(row)));
      render();
    });
    th.append(selectAll);
    headRow.append(th);
  }
  columns.forEach((col, colIndex) => {
    const th = el2("th", "", col.label ?? col.key);
    th.dataset.colIndex = String(colIndex);
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
    attachResizer(th);
    headRow.append(th);
  });
  thead.append(headRow);
  const tbody = el2("tbody");
  table2.append(thead, tbody);
  scroller.append(table2);
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
    const selectedVisible = selectable ? visible.filter((row) => selectedKeys.has(keyOf(row))).length : 0;
    const base = filterText || visible.length !== rows.length ? `${visible.length} / ${rows.length}` : String(rows.length);
    count.textContent = selectedVisible ? `${base} \xB7 ${selectedVisible} selected` : base;
    if (selectAll) {
      selectAll.checked = visible.length > 0 && selectedVisible === visible.length;
      selectAll.indeterminate = selectedVisible > 0 && selectedVisible < visible.length;
    }
    for (const th of headRow.children) {
      if (th.dataset.colIndex === void 0) continue;
      const col = columns[Number(th.dataset.colIndex)];
      const selected = Boolean(col && col.key === sortKey);
      th.querySelector(".wb-sort-arrow").textContent = selected ? sortDir === 1 ? " \u25B2" : " \u25BC" : "";
      th.classList.toggle("is-sorted", selected);
      th.setAttribute("aria-sort", selected ? sortDir === 1 ? "ascending" : "descending" : "none");
    }
    tbody.textContent = "";
    if (!visible.length) {
      const tr = el2("tr");
      const td = el2("td", "wb-empty", rows.length ? "No rows match the filter." : emptyText);
      td.colSpan = columns.length + (selectable ? 1 : 0);
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
          if (e.key === "Enter" && e.target === tr) onOpen(row);
        });
      }
      tr.dataset.key = String(row[rowKey] ?? "");
      if (selectable) {
        const key2 = keyOf(row);
        tr.classList.toggle("wb-row-selected", selectedKeys.has(key2));
        const toggle = () => {
          if (selectedKeys.has(key2)) selectedKeys.delete(key2);
          else selectedKeys.add(key2);
          render();
        };
        const td = el2("td", "wb-select-cell");
        const box = el2("input");
        box.type = "checkbox";
        box.className = "wb-row-check";
        box.checked = selectedKeys.has(key2);
        box.setAttribute("aria-label", "Select row");
        box.addEventListener("click", (e) => e.stopPropagation());
        box.addEventListener("change", toggle);
        if (onOpen) {
          td.addEventListener("click", (e) => {
            e.stopPropagation();
            if (e.target !== box) toggle();
          });
        }
        td.append(box);
        tr.append(td);
        if (!onOpen) {
          tr.classList.add("wb-row-selectable");
          tr.addEventListener("click", (e) => {
            if (e.target.closest("a, button, input, .sp-copy")) return;
            toggle();
          });
        }
      }
      for (const col of columns) {
        const td = el2("td", [col.mono ? "wb-mono" : "", col.num ? "wb-num" : ""].filter(Boolean).join(" "));
        const text = displayValue(row, col);
        if (typeof col.render === "function") {
          const node = col.render(cellValue2(row, col), row);
          if (node) td.append(node);
          tr.append(td);
          continue;
        }
        const href = typeof col.link === "function" && text ? String(col.link(cellValue2(row, col), row) || "") : "";
        if (href) {
          const a = el2("a", "wb-cell-url", text);
          a.href = href;
          a.title = "Open in a new tab";
          bindNewTab(a);
          td.append(a);
          if (col.copyable) {
            const glyph = el2("span", "sp-copy wb-cell-copy", "\u29C9");
            glyph.title = "Click to copy";
            glyph.addEventListener("click", (e) => {
              e.stopPropagation();
              copyText(text, glyph);
            });
            td.append(glyph);
          }
        } else if (col.copyable && text) {
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
      selectedKeys.clear();
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
    // A denial is reported, not alarmed about: neutral register, plain
    // sentence, SharePoint's own words on the tooltip. Everything else stays
    // loud. See denied.js for why.
    setError(err) {
      status.className = "wb-grid-status";
      showFailure(status, err, subject);
    },
    getVisibleRows: () => [...visible],
    getExportRows: () => [...exportRows()],
    getColumns: () => columns
  };
}

// ../src/workbench/item-export.js
var EXCLUDED_TYPES = /* @__PURE__ */ new Set(["Computed", "Attachments"]);
var EXCLUDED_INTERNAL = /* @__PURE__ */ new Set([
  "ContentType",
  "Attachments",
  "ComplianceAssetId",
  "AppAuthor",
  "AppEditor",
  "Edit",
  "DocIcon",
  "ItemChildCount",
  "FolderChildCount",
  "_ColorTag",
  "_UIVersionString",
  "LinkTitle",
  "LinkTitleNoMenu",
  "LinkFilename",
  "LinkFilenameNoMenu"
]);
var SYSTEM_INTERNAL = /* @__PURE__ */ new Set(["ID", "Id", "Title", "Created", "Modified", "Author", "Editor"]);
var VIEW_FIELD_ALIAS = {
  LinkTitle: "Title",
  LinkTitleNoMenu: "Title",
  LinkFilename: "FileLeafRef",
  LinkFilenameNoMenu: "FileLeafRef"
};
var VIEW_ALLOWED_READONLY = /* @__PURE__ */ new Set(["FileLeafRef"]);
function isContentField(f) {
  return Boolean(f) && !f.Hidden && !f.ReadOnlyField && !EXCLUDED_TYPES.has(f.TypeAsString) && !EXCLUDED_INTERNAL.has(f.InternalName) && !SYSTEM_INTERNAL.has(f.InternalName);
}
function contentFields(fields) {
  return (fields || []).filter(isContentField);
}
function viewColumnFields(fields, viewFieldNames) {
  const byName = new Map((fields || []).filter((f) => isContentField(f) || f && !f.Hidden && VIEW_ALLOWED_READONLY.has(f.InternalName)).map((f) => [f.InternalName, f]));
  const out = [];
  for (const raw of viewFieldNames || []) {
    const name = VIEW_FIELD_ALIAS[raw] || String(raw);
    if (SYSTEM_INTERNAL.has(name) || EXCLUDED_INTERNAL.has(name)) continue;
    const field2 = byName.get(name);
    if (field2 && !out.includes(field2)) out.push(field2);
  }
  return out;
}
var asTextKey = (name) => String(name).replaceAll("_", "_x005f_");
function textOf(item2, internalName) {
  const fvt = item2?.FieldValuesAsText;
  if (!fvt || typeof fvt !== "object") return void 0;
  return fvt[internalName] ?? fvt[asTextKey(internalName)];
}
function scalarText(v) {
  if (v === null || v === void 0) return "";
  if (typeof v === "boolean") return v ? "Yes" : "No";
  if (Array.isArray(v)) return v.map(scalarText).filter(Boolean).join(", ");
  if (typeof v === "object") {
    if (Array.isArray(v.results)) return scalarText(v.results);
    if (v.Title) return String(v.Title);
    if (v.Url) {
      const desc = String(v.Description || "");
      return desc && desc !== v.Url ? `${desc} (${v.Url})` : String(v.Url);
    }
    return "";
  }
  return String(v);
}
function itemTitle(item2) {
  const title = textOf(item2, "Title") ?? scalarText(item2?.Title);
  return String(title || "").trim();
}
function personText(item2, internalName) {
  return String(textOf(item2, internalName) ?? scalarText(item2?.[internalName]) ?? "").trim();
}
function fieldText(item2, field2) {
  const name = field2.InternalName;
  const raw = item2?.[name];
  if (field2.TypeAsString === "Note") {
    const text2 = textOf(item2, name);
    if (text2 !== void 0) return String(text2).replace(/\s+/g, " ").trim();
    return htmlToMarkdown(typeof raw === "string" ? raw : "").replace(/\s+/g, " ").trim();
  }
  if (field2.TypeAsString === "URL") {
    return scalarText(raw) || String(textOf(item2, name) ?? "");
  }
  const text = textOf(item2, name);
  return String(text !== void 0 ? text : scalarText(raw)).trim();
}
var mdLink = (label, url) => `[${String(label).replace(/([[\]])/g, "\\$1")}](${String(url).replace(/\(/g, "%28").replace(/\)/g, "%29")})`;
var encodeSpPath2 = (path) => String(path).split("/").map(encodeURIComponent).join("/");
function fieldMarkdown(item2, field2) {
  const name = field2.InternalName;
  const raw = item2?.[name];
  if (field2.TypeAsString === "Note") {
    if (typeof raw === "string" && raw.includes("<")) return htmlToMarkdown(raw);
    return String(raw ?? textOf(item2, name) ?? "").trim();
  }
  if (field2.TypeAsString === "URL") {
    if (typeof raw === "string") return raw.trim();
    const url = String(raw?.Url ?? raw?.url ?? "").trim();
    if (url) {
      const desc = String(raw?.Description ?? raw?.description ?? "").trim();
      return desc && desc !== url ? mdLink(desc, url) : url;
    }
    return String(textOf(item2, name) ?? scalarText(raw)).trim();
  }
  return fieldText(item2, field2);
}
function attachmentLinks(item2, origin = "") {
  const value = item2?.AttachmentFiles;
  const files = Array.isArray(value) ? value : value?.results || [];
  return files.map((f) => {
    const rel = String(f?.ServerRelativeUrl || "");
    if (!rel) return "";
    const name = String(f?.FileName || rel.split("/").pop() || rel);
    return mdLink(name, `${origin}${encodeSpPath2(rel)}`);
  }).filter(Boolean);
}
var BLOCK_TAGS = /* @__PURE__ */ new Set([
  "p",
  "div",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "ul",
  "ol",
  "table",
  "blockquote",
  "pre",
  "section",
  "article",
  "header",
  "footer",
  "hr"
]);
var collapse = (s) => s.replace(/[ \t]*\n[ \t]*/g, "\n").replace(/[ \t]{2,}/g, " ").replace(/^[ \t]+|[ \t]+$/g, "").replace(/^\n+|\n+$/g, "");
function inlineChildren(node) {
  let out = "";
  for (const child of node.childNodes) out += inlineNode(child);
  return out;
}
function inlineNode(node) {
  if (node.nodeType === 3) return String(node.nodeValue).replace(/\s+/g, " ");
  if (node.nodeType !== 1) return "";
  const tag = node.tagName.toLowerCase();
  if (tag === "br") return "\n";
  if (tag === "script" || tag === "style") return "";
  if (BLOCK_TAGS.has(tag)) {
    const block = blockNode(node);
    return block ? `
${block}
` : "";
  }
  const body = inlineChildren(node);
  const core = body.trim();
  if (tag === "strong" || tag === "b") return core ? `**${core}**` : "";
  if (tag === "em" || tag === "i") return core ? `*${core}*` : "";
  if (tag === "a") {
    const href = String(node.getAttribute("href") || "");
    const label = core || href;
    return href && !/^javascript:/i.test(href) ? mdLink(label, href) : label;
  }
  if (tag === "img") {
    const src = String(node.getAttribute("src") || "");
    return src ? `!${mdLink(node.getAttribute("alt") || "", src)}` : "";
  }
  return body;
}
function blockNode(node) {
  const tag = node.tagName.toLowerCase();
  if (tag === "hr") return "---";
  if (tag === "ul" || tag === "ol") {
    const items = [...node.children].filter((c) => c.tagName?.toLowerCase() === "li");
    return items.map((li, i) => `${tag === "ol" ? `${i + 1}.` : "-"} ${collapse(inlineChildren(li)).replace(/\n+/g, " ")}`).join("\n");
  }
  if (/^h[1-6]$/.test(tag)) {
    const core = collapse(inlineChildren(node)).replace(/\n+/g, " ");
    return core ? `**${core}**` : "";
  }
  if (tag === "blockquote") {
    return blockChildren(node).split("\n").map((l) => `> ${l}`).join("\n");
  }
  if (tag === "pre") {
    return `\`\`\`
${String(node.textContent).replace(/\s+$/, "")}
\`\`\``;
  }
  if (tag === "table") {
    return [...node.querySelectorAll("tr")].map((tr) => [...tr.children].map((td) => collapse(inlineChildren(td)).replace(/\n+/g, " ")).join(" | ")).join("\n");
  }
  return blockChildren(node);
}
function blockChildren(container) {
  const parts = [];
  let run = "";
  const flush = () => {
    const text = collapse(run);
    if (text) parts.push(text);
    run = "";
  };
  for (const child of container.childNodes) {
    const tag = child.nodeType === 1 ? child.tagName.toLowerCase() : "";
    if (BLOCK_TAGS.has(tag)) {
      flush();
      const block = blockNode(child);
      if (block) parts.push(block);
    } else {
      run += inlineNode(child);
    }
  }
  flush();
  return parts.join("\n\n");
}
var sharedParser = null;
function htmlToMarkdown(html) {
  const source = String(html ?? "");
  if (!source.trim()) return "";
  if (typeof DOMParser === "undefined") {
    return source.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }
  sharedParser = sharedParser || new DOMParser();
  const doc = sharedParser.parseFromString(source, "text/html");
  return blockChildren(doc.body).replace(/\n{3,}/g, "\n\n").trim();
}
function buildItemsMarkdown({
  listTitle = "List",
  webUrl = "",
  viewTitle = "",
  items = [],
  fields = [],
  viewFieldNames = null,
  filter = "",
  orderby = ""
} = {}) {
  const columns = viewFieldNames ? viewColumnFields(fields, viewFieldNames) : contentFields(fields);
  let origin = "";
  try {
    origin = new URL(webUrl).origin;
  } catch {
  }
  const lines = [`# ${listTitle}`, ""];
  const source = [
    webUrl,
    viewTitle ? `view \u201C${viewTitle}\u201D` : "all columns",
    filter ? `filter: ${filter}` : "",
    orderby ? `order: ${orderby}` : "",
    `${items.length} item${items.length === 1 ? "" : "s"}`
  ].filter(Boolean).join(" \xB7 ");
  lines.push(source, "");
  const oneLine = (s) => String(s ?? "").replace(/\s+/g, " ").trim();
  for (const item2 of items) {
    const id = item2?.ID ?? item2?.Id;
    const title = oneLine(itemTitle(item2));
    lines.push(`## ${title || (id !== void 0 && id !== null ? `Item ${id}` : "Item")}`, "");
    const put = (label, value) => {
      const v = String(value ?? "").trim();
      if (!v) return;
      if (v.includes("\n")) {
        lines.push(
          `${oneLine(label)}:`,
          "",
          ...v.split("\n").map((l) => l.trim() ? `> ${l}` : ">"),
          ""
        );
      } else {
        lines.push(`${oneLine(label)}: ${v}  `);
      }
    };
    put("ID", id);
    for (const field2 of columns) put(field2.Title || field2.InternalName, fieldMarkdown(item2, field2));
    const attachments = attachmentLinks(item2, origin);
    if (attachments.length) put("Attachments", attachments.join(", "));
    put("Created", textOf(item2, "Created") ?? item2?.Created);
    put("Created By", personText(item2, "Author"));
    put("Modified", textOf(item2, "Modified") ?? item2?.Modified);
    put("Modified By", personText(item2, "Editor"));
    lines.push("");
  }
  return `${lines.join("\n").replace(/\n{3,}/g, "\n\n").trimEnd()}
`;
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
  const table2 = el3("table", "console-table");
  const thead = el3("thead");
  const hr = el3("tr");
  hr.append(el3("th", "", "(index)"));
  cols.forEach((c) => hr.append(el3("th", "", c)));
  thead.append(hr);
  table2.append(thead);
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
  table2.append(tbody);
  wrap.append(table2);
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

// ../src/workbench/list-schema.js
var SCHEMA_KIND = "dcspad-sputils-list-schema";
var DATA_KIND = "dcspad-sputils-list-data";
var SCHEMA_VERSION = 2;
var LOOKUP_TYPES = /* @__PURE__ */ new Set(["Lookup", "LookupMulti"]);
var USER_TYPES = /* @__PURE__ */ new Set(["User", "UserMulti"]);
var TAXONOMY_TYPES = /* @__PURE__ */ new Set(["TaxonomyFieldType", "TaxonomyFieldTypeMulti"]);
var NEVER_WRITE = /* @__PURE__ */ new Set([
  "ID",
  "Id",
  "Attachments",
  "ContentType",
  "ContentTypeId",
  "Author",
  "Editor",
  "Created",
  "Modified",
  "GUID",
  "FileRef",
  "FileDirRef",
  "FileLeafRef",
  "UniqueId",
  "_UIVersionString",
  "Order"
]);
var NEVER_WRITE_TYPES = /* @__PURE__ */ new Set([
  "Computed",
  "Counter",
  "Attachments",
  "File",
  "ContentTypeId",
  "Calculated",
  "Guid"
]);
var NOT_INDEXABLE_TYPES = /* @__PURE__ */ new Set(["Note", "Computed", "Attachments", "Calculated"]);
function cleanGuid(v) {
  const s = String(v ?? "").replace(/[{}]/g, "").trim();
  return s ? s.toLowerCase() : null;
}
var GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function isGuidLike(v) {
  const g = cleanGuid(v);
  return !!g && GUID_RE.test(g);
}
function lookupMapGet(lookupMap, title) {
  if (!title || !lookupMap) return void 0;
  if (Object.prototype.hasOwnProperty.call(lookupMap, title)) return lookupMap[title];
  const wanted = String(title).toLowerCase();
  const key2 = Object.keys(lookupMap).find((k) => k.toLowerCase() === wanted);
  return key2 !== void 0 ? lookupMap[key2] : void 0;
}
function isCustomField(f) {
  return !f?.FromBaseType && f?.CanBeDeleted === true && !f?.Hidden && f?.TypeAsString !== "Computed";
}
var LIBRARY_BASE_TEMPLATES = /* @__PURE__ */ new Set([101, 109, 115, 119, 850, 851]);
function schemaBaseType(doc) {
  const sourceBaseType = doc?.source?.baseType;
  if (sourceBaseType != null) return Number(sourceBaseType);
  return LIBRARY_BASE_TEMPLATES.has(Number(doc?.list?.baseTemplate)) ? 1 : 0;
}
function fieldTier(f) {
  if (f?.type === "Calculated") return 3;
  if (f?.isDependentLookup) return 2;
  if (LOOKUP_TYPES.has(f?.type)) return 1;
  return 0;
}
function orderFields(fields) {
  return (fields || []).map((f, i) => [f, i]).sort((a, b) => fieldTier(a[0]) - fieldTier(b[0]) || a[1] - b[1]).map(([f]) => f);
}
function indexableType(type) {
  return !NOT_INDEXABLE_TYPES.has(type);
}
function parentContentTypeId(listCtId) {
  const id = String(listCtId || "");
  const m = /^(.*?)00[0-9a-f]{32}$/i.exec(id);
  return m ? m[1] : id;
}
var LIST_DEFAULTS = {
  title: "",
  description: "",
  baseTemplate: 100,
  enableVersioning: false,
  majorVersionLimit: null,
  enableMinorVersions: false,
  majorWithMinorVersionsLimit: null,
  draftVersionVisibility: 0,
  forceCheckout: false,
  hidden: false,
  contentTypesEnabled: false,
  enableAttachments: true,
  enableFolderCreation: false,
  enableModeration: false,
  validationFormula: "",
  validationMessage: "",
  onQuickLaunch: false,
  noCrawl: false,
  disableGridEditing: false,
  ordered: false,
  readSecurity: null,
  writeSecurity: null,
  listExperienceOptions: null,
  enableRequestSignOff: false,
  library: null,
  contentTypeOrder: []
};
function normalizeList(raw = {}) {
  return {
    ...LIST_DEFAULTS,
    title: raw.Title,
    description: raw.Description || "",
    baseTemplate: raw.BaseTemplate,
    enableVersioning: !!raw.EnableVersioning,
    majorVersionLimit: raw.MajorVersionLimit ?? null,
    enableMinorVersions: !!raw.EnableMinorVersions,
    majorWithMinorVersionsLimit: raw.MajorWithMinorVersionsLimit ?? null,
    draftVersionVisibility: raw.DraftVersionVisibility ?? 0,
    forceCheckout: !!raw.ForceCheckout,
    hidden: !!raw.Hidden,
    contentTypesEnabled: !!raw.ContentTypesEnabled,
    enableAttachments: raw.EnableAttachments !== false,
    enableFolderCreation: !!raw.EnableFolderCreation,
    enableModeration: !!raw.EnableModeration,
    validationFormula: raw.ValidationFormula || "",
    validationMessage: raw.ValidationMessage || "",
    onQuickLaunch: !!raw.OnQuickLaunch,
    noCrawl: !!raw.NoCrawl,
    disableGridEditing: !!raw.DisableGridEditing,
    ordered: !!raw.Ordered,
    readSecurity: raw.ReadSecurity ?? null,
    writeSecurity: raw.WriteSecurity ?? null,
    listExperienceOptions: raw.ListExperienceOptions ?? null,
    enableRequestSignOff: !!raw.EnableRequestSignOff,
    library: raw.BaseType === 1 ? { documentTemplateUrl: raw.DocumentTemplateUrl || "", irmEnabled: !!raw.IrmEnabled } : null,
    contentTypeOrder: []
  };
}
var FIELD_DEFAULTS = {
  fieldTypeKind: null,
  group: "",
  hidden: false,
  sealed: false,
  canBeDeleted: true,
  formula: "",
  outputType: "",
  displayFormat: null,
  richText: false,
  validationFormula: "",
  validationMessage: "",
  jsLink: "",
  baseTweak: null
};
function baseTweakFor(raw, custom) {
  if (custom) return null;
  if (raw.InternalName !== "Title" && raw.InternalName !== "_ExtendedDescription") return null;
  return { title: raw.Title, required: !!raw.Required, description: raw.Description || "" };
}
function normalizeField(raw, { listId = "", lookupTitleById = () => null } = {}) {
  const type = raw.TypeAsString || "";
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
    description: raw.Description || "",
    defaultValue: raw.DefaultValue ?? null,
    choices: Array.isArray(raw.Choices) ? raw.Choices : raw.Choices?.results || [],
    maxLength: raw.MaxLength ?? null,
    indexed: !!raw.Indexed,
    enforceUniqueValues: !!raw.EnforceUniqueValues,
    allowMultipleValues: !!raw.AllowMultipleValues || type.endsWith("Multi"),
    customFormatter: raw.CustomFormatter || "",
    lookupListId,
    lookupList: lookupListId ? lookupTitleById(lookupListId) : null,
    lookupField: isLookup ? raw.LookupField || "Title" : null,
    isSelfLookup: !!lookupListId && lookupListId === cleanGuid(listId),
    isDependentLookup: !!raw.IsDependentLookup,
    primaryFieldId: raw.IsDependentLookup ? cleanGuid(raw.PrimaryFieldId) : null,
    schemaXml: raw.SchemaXml,
    // v2 additions.
    fieldTypeKind: raw.FieldTypeKind ?? null,
    group: raw.Group || "",
    hidden: !!raw.Hidden,
    sealed: !!raw.Sealed,
    canBeDeleted: raw.CanBeDeleted !== false,
    formula: raw.Formula || "",
    outputType: raw.OutputType || "",
    displayFormat: raw.DisplayFormat ?? null,
    richText: !!raw.RichText,
    validationFormula: raw.ValidationFormula || "",
    validationMessage: raw.ValidationMessage || "",
    jsLink: raw.JSLink || "",
    baseTweak: baseTweakFor(raw, custom)
  };
}
var VIEW_DEFAULTS = {
  viewTypeKind: null,
  scope: null,
  aggregations: "",
  aggregationsStatus: "",
  tabularView: true,
  mobileView: false,
  mobileDefaultView: false,
  viewData: "",
  viewJoins: "",
  readOnlyView: false,
  includeRootFolder: false,
  serverRelativeUrl: ""
};
function normalizeView(raw = {}, viewFieldNames = []) {
  return {
    ...VIEW_DEFAULTS,
    id: raw.Id,
    title: raw.Title,
    defaultView: !!raw.DefaultView,
    hidden: !!raw.Hidden,
    viewType: raw.ViewType || "HTML",
    viewQuery: raw.ViewQuery || "",
    rowLimit: raw.RowLimit ?? 30,
    paged: raw.Paged !== false,
    customFormatter: raw.CustomFormatter || "",
    jsLink: raw.JSLink || "",
    fields: viewFieldNames || [],
    viewTypeKind: raw.ViewTypeKind ?? null,
    scope: raw.Scope ?? null,
    aggregations: raw.Aggregations || "",
    aggregationsStatus: raw.AggregationsStatus || "",
    tabularView: raw.TabularView !== false,
    mobileView: !!raw.MobileView,
    mobileDefaultView: !!raw.MobileDefaultView,
    viewData: raw.ViewData || "",
    viewJoins: raw.ViewJoins || "",
    readOnlyView: !!raw.ReadOnlyView,
    includeRootFolder: !!raw.IncludeRootFolder,
    serverRelativeUrl: raw.ServerRelativeUrl || ""
  };
}
var CT_DEFAULTS = { parentId: "", sealed: false, documentTemplate: "", fieldLinks: [] };
function normalizeContentType(raw = {}, fieldLinks = []) {
  const id = raw.StringId || raw.Id?.StringValue || String(raw.Id || "");
  return {
    ...CT_DEFAULTS,
    name: raw.Name,
    id,
    description: raw.Description || "",
    group: raw.Group || "",
    hidden: !!raw.Hidden,
    readOnly: !!raw.ReadOnly,
    parentId: parentContentTypeId(id),
    sealed: !!raw.Sealed,
    documentTemplate: raw.DocumentTemplate || "",
    fieldLinks: fieldLinks || []
  };
}
function buildSchemaDoc({
  source = {},
  list: list2 = {},
  fields = [],
  views = [],
  contentTypes = [],
  warnings = [],
  generatorBuild = "dev"
} = {}) {
  return {
    kind: SCHEMA_KIND,
    version: SCHEMA_VERSION,
    exported: (/* @__PURE__ */ new Date()).toISOString(),
    generator: { tool: "dcspad-workbench", build: generatorBuild },
    source,
    list: list2,
    fields,
    views,
    contentTypes,
    warnings
  };
}
function normalizeSchemaDoc(doc) {
  if (!doc || typeof doc !== "object" || doc.kind !== SCHEMA_KIND) {
    throw new SpFileError(
      `Not a list schema document (expected kind ${SCHEMA_KIND}).`,
      { code: "bad-schema" }
    );
  }
  const sourceVersion = Number(doc.version) || 1;
  const list2 = { ...LIST_DEFAULTS, ...doc.list || {} };
  const fields = (doc.fields || []).map((f) => ({ ...FIELD_DEFAULTS, ...f }));
  const views = (doc.views || []).map((v) => ({ ...VIEW_DEFAULTS, ...v }));
  const contentTypes = (doc.contentTypes || []).map((c) => ({
    ...CT_DEFAULTS,
    ...c,
    parentId: c.parentId || parentContentTypeId(c.id)
  }));
  return {
    kind: doc.kind,
    version: SCHEMA_VERSION,
    exported: doc.exported || "",
    generator: doc.generator || { tool: "unknown", build: "" },
    source: doc.source || {},
    list: list2,
    fields,
    views,
    contentTypes,
    warnings: doc.warnings || [],
    _sourceVersion: sourceVersion
  };
}
var BASE_TEMPLATE_LABELS = { 100: "Generic list", 101: "Document library" };
function schemaSummary(doc) {
  const d = normalizeSchemaDoc(doc);
  const fields = d.fields || [];
  const custom = fields.filter((f) => f.custom).length;
  const views = d.views || [];
  return {
    kind: BASE_TEMPLATE_LABELS[d.list.baseTemplate] || `List (template ${d.list.baseTemplate})`,
    fieldsText: `${fields.length} field${fields.length === 1 ? "" : "s"} \xB7 ${custom} custom`,
    viewsText: `${views.length} view${views.length === 1 ? "" : "s"}`,
    contentTypesText: d.list.contentTypesEnabled ? "content types on" : "content types off",
    versioningText: d.list.enableVersioning ? "versioning on" : "versioning off",
    isLibrary: schemaBaseType(d) === 1
  };
}
var STRIP_ATTRS = [
  "ID",
  "SourceID",
  "ColName",
  "RowOrdinal",
  "Version",
  "WebId",
  "List",
  "Sealed",
  "Customization",
  "Indexed",
  "EnforceUniqueValues"
];
function scrubSchemaXml(xml, { lookupListId = null, primaryFieldId = null, fieldType = "" } = {}) {
  let doc;
  try {
    doc = new DOMParser().parseFromString(String(xml || ""), "text/xml");
  } catch (cause) {
    throw new SpFileError("SchemaXml could not be parsed.", { code: "bad-xml", cause });
  }
  const root = doc.documentElement;
  if (!root || root.nodeName === "parsererror" || doc.getElementsByTagName("parsererror").length) {
    throw new SpFileError("SchemaXml could not be parsed.", { code: "bad-xml" });
  }
  for (const attr of STRIP_ATTRS) root.removeAttribute(attr);
  if (LOOKUP_TYPES.has(fieldType) && lookupListId) {
    root.setAttribute("List", `{${cleanGuid(lookupListId)}}`);
  } else if (USER_TYPES.has(fieldType)) {
    root.setAttribute("List", "UserInfo");
  }
  if (LOOKUP_TYPES.has(fieldType) && primaryFieldId) {
    root.setAttribute("FieldRef", `{${cleanGuid(primaryFieldId)}}`);
  }
  const refs = root.getElementsByTagName("FieldRef");
  for (let i = 0; i < refs.length; i++) refs[i].removeAttribute("ID");
  return new XMLSerializer().serializeToString(root);
}
function textFallbackXml(field2) {
  const name = String(field2?.internalName || "").replaceAll('"', "&quot;");
  const display = String(field2?.displayName || field2?.internalName || "").replaceAll('"', "&quot;");
  const required = field2?.required ? ' Required="TRUE"' : "";
  return `<Field Type="Text" Name="${name}" DisplayName="${display}"${required} />`;
}
function xmlHasAttr(xml, name) {
  try {
    const doc = new DOMParser().parseFromString(String(xml || ""), "text/xml");
    return doc.documentElement?.hasAttribute(name) ?? false;
  } catch {
    return false;
  }
}
function step(id, kind, label, {
  dependsOn = [],
  payload = {},
  refs = {},
  optional = false,
  status = "planned",
  error = "",
  final = false
} = {}) {
  return { id, kind, label, dependsOn, payload, refs, optional, status, error, final, result: null };
}
function isBuiltinParent(parentId) {
  return parentId === "0x01" || parentId === "0x0120" || parentId === "0x0101";
}
function settingsPayload(list2, { reconcile = false, isLib = false } = {}) {
  const library = isLib;
  if (reconcile) {
    const groupA2 = {};
    if (list2.enableAttachments && !library) groupA2.EnableAttachments = true;
    if (list2.enableFolderCreation) groupA2.EnableFolderCreation = true;
    return { groupA: groupA2, groupB: {} };
  }
  const groupA = {
    EnableVersioning: list2.enableVersioning,
    EnableAttachments: list2.enableAttachments,
    EnableFolderCreation: list2.enableFolderCreation,
    EnableModeration: list2.enableModeration,
    Hidden: list2.hidden,
    OnQuickLaunch: list2.onQuickLaunch,
    NoCrawl: list2.noCrawl,
    DisableGridEditing: list2.disableGridEditing,
    // No Ordered: SP.List has no such REST property (a MERGE naming it is a
    // 400, verified live) — `ordered` stays in the doc as capture-only.
    ReadSecurity: list2.readSecurity,
    WriteSecurity: list2.writeSecurity,
    ListExperienceOptions: list2.listExperienceOptions,
    EnableRequestSignOff: list2.enableRequestSignOff,
    ForceCheckout: list2.forceCheckout
  };
  const groupB = {};
  if (list2.enableVersioning && list2.majorVersionLimit) groupB.MajorVersionLimit = list2.majorVersionLimit;
  if (library) {
    if (list2.enableVersioning && list2.enableMinorVersions) {
      groupA.EnableMinorVersions = true;
      if (list2.majorWithMinorVersionsLimit) groupB.MajorWithMinorVersionsLimit = list2.majorWithMinorVersionsLimit;
    }
  } else if (list2.enableVersioning && list2.enableMinorVersions) {
    groupB.EnableMinorVersions = true;
    if (list2.majorWithMinorVersionsLimit) groupB.MajorWithMinorVersionsLimit = list2.majorWithMinorVersionsLimit;
  }
  if (list2.enableModeration || list2.enableMinorVersions) groupB.DraftVersionVisibility = list2.draftVersionVisibility;
  if (library) delete groupA.EnableAttachments;
  const defined = (obj) => Object.fromEntries(Object.entries(obj).filter(([, v]) => v != null));
  return { groupA: defined(groupA), groupB: defined(groupB) };
}
function defaultTargetTitle(doc, targetLists = []) {
  const source = doc?.list?.title || doc?.source?.listTitle || "List";
  const taken = new Set((targetLists || []).map((l) => String(l.title || "").toLowerCase()));
  if (!taken.has(source.toLowerCase())) return source;
  for (let n = 1; ; n++) {
    const candidate = n === 1 ? `${source} Copy` : `${source} Copy ${n}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}
function fieldMergePayload(f, warnings) {
  const merges = {};
  if (f.indexed && indexableType(f.type)) merges.Indexed = true;
  if (f.enforceUniqueValues && indexableType(f.type)) merges.EnforceUniqueValues = true;
  if (f.customFormatter && !xmlHasAttr(f.schemaXml, "CustomFormatter")) {
    merges.CustomFormatter = f.customFormatter;
  }
  if (f.indexed && !indexableType(f.type)) {
    warnings.push(`Column \u2018${f.internalName}\u2019 (${f.type}) cannot be indexed \u2014 Indexed was not re-applied.`);
  }
  return merges;
}
function pushMergeStep(steps, mergeStepIds, f, dependsOnId, merges) {
  if (!Object.keys(merges).length) return;
  const mergeId = `merge:${f.internalName}`;
  mergeStepIds.push(mergeId);
  steps.push(step(mergeId, "field.merge", `Apply ${Object.keys(merges).join(", ")} on \u2018${f.internalName}\u2019`, {
    dependsOn: [dependsOnId],
    payload: { internalName: f.internalName, merges },
    optional: true
  }));
}
function buildApplyPlan(doc, options = {}, probe = {}) {
  const d = normalizeSchemaDoc(doc);
  const baseType = schemaBaseType(d);
  const isLib = baseType === 1;
  const isGenericList = baseType === 0 && Number(d.list.baseTemplate) === 100;
  if (!isGenericList && !isLib) {
    throw new SpFileError(
      "Only generic lists and document libraries can be created from a schema.",
      { code: "unsupported-template" }
    );
  }
  const warnings = [];
  const title = String(options.title || "").trim() || defaultTargetTitle(d, probe.targetLists || []);
  const existing = probe.existingList || null;
  if (existing && existing.baseType != null && Number(existing.baseType) !== baseType) {
    throw new SpFileError(
      `\u2018${title}\u2019 already exists on the target as a different type of list \u2014 its schema cannot be applied to it.`,
      { code: "base-type-mismatch" }
    );
  }
  const steps = [];
  if (existing) {
    if ((options.existing || "fail") === "fail") {
      throw new SpFileError(
        `A list named \u2018${title}\u2019 already exists on the target.`,
        { code: "exists" }
      );
    }
    steps.push(step("list", "list.adopt", `Add to existing list \u2018${title}\u2019`, {
      payload: { listId: existing.id, title },
      refs: { listId: { self: true } }
    }));
  } else {
    const createBaseTemplate = isLib ? 101 : d.list.baseTemplate || 100;
    if (isLib && Number(d.list.baseTemplate) !== 101) {
      warnings.push(`Recreated as a standard document library (source template ${d.list.baseTemplate}).`);
    }
    steps.push(step("list", "list.create", `Create list \u2018${title}\u2019 (${isLib ? "Document library" : BASE_TEMPLATE_LABELS[d.list.baseTemplate] || d.list.baseTemplate})`, {
      payload: {
        title,
        description: options.description ?? d.list.description ?? "",
        baseTemplate: createBaseTemplate,
        contentTypesEnabled: !!d.list.contentTypesEnabled,
        urlName: options.urlName || ""
      }
    }));
  }
  steps.push(step("settings", "list.settings", existing ? "Reconcile list settings" : "Apply list settings", {
    dependsOn: ["list"],
    payload: settingsPayload(d.list, { reconcile: Boolean(existing), isLib }),
    refs: { listId: { self: true } }
  }));
  const ctStepIds = [];
  if (d.list.contentTypesEnabled) {
    const already = new Set(probe.existingContentTypeIds || []);
    const available = new Set((probe.availableContentTypes || []).map((c) => c.id));
    const seen = /* @__PURE__ */ new Set();
    for (const ct of d.contentTypes) {
      const parentId = ct.parentId || parentContentTypeId(ct.id);
      if (isBuiltinParent(parentId) || already.has(parentId) || seen.has(parentId)) continue;
      seen.add(parentId);
      const id = `ct:${parentId}`;
      ctStepIds.push(id);
      const isAvailable = available.has(parentId);
      steps.push(step(id, "ct.attach", `Attach content type \u2018${ct.name}\u2019`, {
        dependsOn: ["list"],
        payload: { contentTypeId: parentId, name: ct.name },
        refs: { listId: { self: true } },
        optional: true,
        status: isAvailable ? "planned" : "failed",
        final: !isAvailable,
        error: isAvailable ? "" : `\u2018${ct.name}\u2019 is not available on the target web \u2014 publish the content type there first.`
      }));
    }
  }
  const preProvisioned = (f) => isLib && f.internalName === "_ExtendedDescription";
  const custom = orderFields(d.fields.filter((f) => f.custom && !preProvisioned(f)));
  const existingFields = probe.existingFields || [];
  const createdInternalNames = /* @__PURE__ */ new Set();
  const fieldStepIds = [];
  const mergeStepIds = [];
  const missingLookupPolicy = options.missingLookup === "text" ? "text" : "skip";
  const lookupMap = options.lookupMap || {};
  const baseFieldNames = new Set(d.fields.filter((f) => !f.custom && f.fromBaseType).map((f) => f.internalName));
  const fieldOptions = d.list.contentTypesEnabled && ctStepIds.length > 0 ? 8 | 4 : 8;
  for (const f of custom) {
    const id = `field:${f.internalName}`;
    fieldStepIds.push(id);
    const present = existingFields.find((ef) => ef.internalName === f.internalName);
    if (present) {
      const same = present.typeAsString === f.type;
      steps.push(step(id, "field.create", `Column \u2018${f.displayName}\u2019 (${f.internalName})`, {
        dependsOn: ["list", ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: same ? "skipped" : "failed",
        final: !same,
        error: same ? "" : `exists on the target as ${present.typeAsString}, source is ${f.type} \u2014 values will not import.`
      }));
      if (same) {
        createdInternalNames.add(f.internalName);
        pushMergeStep(steps, mergeStepIds, f, id, fieldMergePayload(f, warnings));
      }
      continue;
    }
    const titleTaken = existingFields.some((ef) => ef.internalName !== f.internalName && String(ef.title || "").toLowerCase() === String(f.displayName || "").toLowerCase());
    if (titleTaken) {
      steps.push(step(id, "field.create", `Column \u2018${f.displayName}\u2019 (${f.internalName})`, {
        dependsOn: ["list", ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: "failed",
        final: true,
        error: `a different column already uses the display name \u2018${f.displayName}\u2019.`
      }));
      continue;
    }
    if (TAXONOMY_TYPES.has(f.type)) {
      steps.push(step(id, "field.create", `Skip column \u2018${f.displayName}\u2019 (${f.internalName}) \u2014 managed metadata is not recreated`, {
        dependsOn: ["list", ...ctStepIds],
        payload: { field: f, options: fieldOptions },
        status: "skipped",
        error: "",
        final: true
      }));
      steps[steps.length - 1].skipReason = "managed-metadata";
      continue;
    }
    const refs = {};
    let asText = false;
    let blocked = false;
    let blockReason = "";
    if (f.isDependentLookup) {
      const primaryField = d.fields.find((pf) => pf.id === f.primaryFieldId);
      if (primaryField && createdInternalNames.has(primaryField.internalName)) {
        refs.primaryFieldId = { field: primaryField.internalName };
      } else {
        blocked = true;
        blockReason = "its primary lookup column was not created on the target.";
      }
    }
    if (!blocked && LOOKUP_TYPES.has(f.type)) {
      if (f.isSelfLookup) {
        refs.lookupListId = { self: true };
      } else {
        const mappedValue = lookupMapGet(lookupMap, f.lookupList) ?? lookupMapGet(lookupMap, f.lookupListId) ?? f.lookupList;
        if (isGuidLike(mappedValue)) {
          refs.lookupListId = { id: cleanGuid(mappedValue) };
        } else {
          const wanted = String(mappedValue || "").toLowerCase();
          const target = (probe.targetLists || []).find((l) => String(l.title || "").toLowerCase() === wanted);
          if (!target) {
            if (missingLookupPolicy === "text") {
              asText = true;
              warnings.push(`Lookup \u2018${f.internalName}\u2019 target \u2018${mappedValue}\u2019 is missing on the target \u2014 created as a single line of text (policy: text; a re-run cannot upgrade it).`);
            } else {
              steps.push(step(id, "field.create", `Skip column \u2018${f.displayName}\u2019 (${f.internalName}) \u2014 lookup target \u2018${mappedValue}\u2019 is missing (policy: skip)`, {
                dependsOn: ["list", ...ctStepIds],
                payload: { field: f, options: fieldOptions },
                status: "skipped",
                error: ""
              }));
              steps[steps.length - 1].skipReason = "lookup-target-missing";
              continue;
            }
          } else {
            refs.lookupListId = { list: target.title };
          }
        }
      }
    }
    steps.push(step(id, "field.create", `${blocked ? "Blocked" : "Add"} ${f.type} column \u2018${f.displayName}\u2019 (${f.internalName})${asText ? " as text" : ""}`, {
      dependsOn: ["list", ...ctStepIds],
      payload: { field: f, asText, options: fieldOptions },
      refs,
      status: blocked ? "blocked" : "planned",
      error: blockReason
    }));
    if (!blocked) createdInternalNames.add(f.internalName);
    if (!blocked && !asText) {
      pushMergeStep(steps, mergeStepIds, f, id, fieldMergePayload(f, warnings));
    }
  }
  const BASE_TWEAKABLE_NAMES = /* @__PURE__ */ new Set(["Title", "_ExtendedDescription"]);
  for (const f of d.fields) {
    if (f.custom && !preProvisioned(f) || !BASE_TWEAKABLE_NAMES.has(f.internalName)) continue;
    const isTitleField = f.internalName === "Title";
    const required = isTitleField && isLib ? false : Boolean(f.required);
    const defaultRequired = isTitleField && !isLib;
    const defaultName = isTitleField ? "Title" : "Description";
    const changed = f.displayName !== defaultName || required !== defaultRequired || Boolean(f.description);
    if (!changed) continue;
    steps.push(step(
      isTitleField ? "title" : `base:${f.internalName}`,
      "field.base",
      isTitleField ? "Set the Title column\u2019s display name and required flag" : `Set the ${f.displayName || f.internalName} column\u2019s display name${f.description ? " and description" : ""}`,
      {
        dependsOn: ["list"],
        payload: {
          internalName: f.internalName,
          displayName: f.displayName,
          required,
          description: f.description || ""
        },
        optional: true
      }
    ));
  }
  const ALWAYS_PRESENT_LIBRARY_VIEW_FIELDS = /* @__PURE__ */ new Set(["LinkFilename", "DocIcon", "FileSizeDisplay"]);
  for (const v of d.views.filter((view) => !view.hidden)) {
    const id = `view:${v.title}`;
    const targetViews = probe.existingViews || [];
    const existingView = targetViews.find((ev) => String(ev.title).toLowerCase() === String(v.title).toLowerCase()) || (v.defaultView ? targetViews.find((ev) => ev.defaultView) : null) || null;
    const wanted = v.fields.filter((name) => createdInternalNames.has(name) || existingFields.some((ef) => ef.internalName === name) || baseFieldNames.has(name) || isLib && ALWAYS_PRESENT_LIBRARY_VIEW_FIELDS.has(name));
    const missing = v.fields.filter((name) => !wanted.includes(name));
    if (missing.length) {
      warnings.push(`View \u2018${v.title}\u2019: columns not on the target were left out: ${missing.join(", ")}.`);
    }
    steps.push(step(id, "view.upsert", `View \u2018${v.title}\u2019 (${wanted.length} column${wanted.length === 1 ? "" : "s"})`, {
      // Only the list: a view that names a column which failed to create
      // simply leaves it out (warned above). Depending on every field step
      // let one managed-metadata column block every view on the list.
      dependsOn: ["list"],
      payload: {
        title: v.title,
        viewId: existingView?.id || null,
        fields: wanted,
        viewQuery: v.viewQuery,
        rowLimit: v.rowLimit,
        paged: v.paged,
        defaultView: v.defaultView,
        customFormatter: v.customFormatter,
        jsLink: v.jsLink,
        viewTypeKind: v.viewTypeKind,
        scope: v.scope,
        aggregations: v.aggregations,
        aggregationsStatus: v.aggregationsStatus,
        tabularView: v.tabularView,
        mobileView: v.mobileView,
        mobileDefaultView: v.mobileDefaultView,
        includeRootFolder: v.includeRootFolder,
        viewData: v.viewData,
        viewJoins: v.viewJoins
      }
    }));
  }
  if (d.list.validationFormula) {
    steps.push(step("validation", "list.validation", "Apply the list validation formula", {
      // Ordered last so every column the formula names exists; a formula
      // over a column that failed is refused by SharePoint on its own merits.
      dependsOn: ["list"],
      payload: { validationFormula: d.list.validationFormula, validationMessage: d.list.validationMessage },
      optional: true
    }));
  }
  return { title, targetWebUrl: options.targetWebUrl || "", existingListId: existing?.id || null, steps, warnings };
}
function newReport(plan) {
  return {
    dryRun: false,
    title: plan.title,
    listId: plan.existingListId || null,
    created: false,
    targetWebUrl: plan.targetWebUrl || "",
    rootFolder: "",
    adopted: Boolean(plan.existingListId),
    aborted: "",
    settings: { applied: [], failed: [] },
    validation: { applied: false, error: "" },
    contentTypes: { attached: 0, skipped: 0, failed: 0 },
    fields: { added: 0, skipped: 0, failed: [] },
    fieldMerges: { applied: 0, failed: 0 },
    views: { added: 0, updated: 0, failed: [] },
    fieldIdMap: {},
    warnings: [...plan.warnings || []],
    steps: (plan.steps || []).map((s) => ({ ...s }))
  };
}
function buildApplyReport({ report, doc, targetWebUrl }) {
  const title = report?.title || doc?.list?.title || "List";
  const dest = targetWebUrl || report?.targetWebUrl || "the target site";
  const lines = ["# List schema report", ""];
  lines.push(report?.created ? `Created \u2018${title}\u2019 on ${dest}.` : `Updated \u2018${title}\u2019 on ${dest}.`, "");
  lines.push(
    "## Fields",
    "",
    `- Added: ${report?.fields?.added ?? 0}`,
    `- Skipped: ${report?.fields?.skipped ?? 0}`,
    `- Failed: ${(report?.fields?.failed || []).length}`,
    ""
  );
  lines.push(
    "## Views",
    "",
    `- Added: ${report?.views?.added ?? 0}`,
    `- Updated: ${report?.views?.updated ?? 0}`,
    `- Failed: ${(report?.views?.failed || []).length}`,
    ""
  );
  if (report?.contentTypes) {
    lines.push(
      "## Content types",
      "",
      `- Attached: ${report.contentTypes.attached ?? 0}`,
      `- Skipped: ${report.contentTypes.skipped ?? 0}`,
      `- Failed: ${report.contentTypes.failed ?? 0}`,
      ""
    );
  }
  const ir = report?.itemsReport;
  if (ir) {
    lines.push(
      "## Items",
      "",
      `- Added: ${ir.items?.added ?? 0}`,
      `- Failed: ${(ir.items?.failed || []).length + (ir.items?.failedTruncated || 0)}` + (ir.items?.failedTruncated ? ` (${ir.items.failedTruncated} not itemised below)` : ""),
      `- Folders created: ${ir.folders?.created ?? 0}`,
      ""
    );
    if (ir.fieldErrors?.length) {
      lines.push("### Item field errors", "");
      for (const fe of ir.fieldErrors) lines.push(`- Source id ${fe.sourceId} \u2014 ${fe.field}: ${fe.message || ""}`);
      lines.push("");
    }
  } else if (report?.itemsHeld) {
    lines.push("## Items", "", `- ${report.itemsHeld}`, "");
  } else if (report?.itemsError) {
    lines.push("## Items", "", `- Could not be imported \u2014 ${report.itemsError}`, "");
  }
  const warnings = [...report?.warnings || [], ...ir?.warnings || []];
  if (warnings.length) {
    lines.push("## Warnings", "");
    for (const w of warnings) lines.push(`- ${w}`);
    lines.push("");
  }
  const failedSteps = (report?.steps || []).filter((s) => s.status === "failed");
  if (failedSteps.length) {
    lines.push("## Failed steps", "");
    for (const s of failedSteps) lines.push(`- ${s.label}: ${s.error || "failed"}`);
    lines.push("");
  }
  return lines.join("\n");
}

// ../src/workbench/list-schema-capture.js
var guidPath = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
var cleanGuid2 = (v) => String(v || "").replace(/[{}]/g, "").toLowerCase();
async function resolveLookupTitles(client2, guids, selfListId) {
  const map = /* @__PURE__ */ new Map();
  const selfId = selfListId ? cleanGuid2(selfListId) : null;
  for (const raw of guids) {
    const id = cleanGuid2(raw);
    if (!id || map.has(id)) continue;
    if (id === selfId) {
      map.set(id, null);
      continue;
    }
    try {
      const target = await client2.get(guidPath(id), { select: "Title" });
      map.set(id, target?.Title ?? null);
    } catch {
      map.set(id, null);
    }
  }
  return map;
}
async function captureListSchema(client2, listId, { includeHidden = false } = {}) {
  const warnings = [];
  const web = await client2.get("web", { select: ["Id", "Title", "Url", "ServerRelativeUrl", "Language"] });
  const list2 = await client2.get(guidPath(listId), { expand: "RootFolder" });
  const NAMED_ONLY = ["ValidationFormula", "ValidationMessage", "OnQuickLaunch", "ReadSecurity", "WriteSecurity"];
  const missing = NAMED_ONLY.filter((k) => !(k in list2));
  if (missing.length) {
    try {
      Object.assign(list2, await client2.get(guidPath(listId), { select: missing }));
    } catch (err) {
      if (isDeniedRead(err) || isExpiredSession(err)) throw err;
      for (const key2 of missing) {
        try {
          Object.assign(list2, await client2.get(guidPath(listId), { select: [key2] }));
        } catch (err2) {
          if (err2?.status !== 400) throw err2;
          warnings.push(`List property ${key2} is not available on this tenant \u2014 it was not captured.`);
        }
      }
    }
  }
  const { items: rawFields } = await client2.getAll(
    guidPath(listId, "/fields"),
    includeHidden ? {} : { filter: "Hidden eq false" }
  );
  const lookupGuids = [...new Set(
    rawFields.filter((f) => LOOKUP_TYPES.has(f.TypeAsString)).map((f) => f.LookupList).filter(Boolean)
  )];
  const lookupTitles = await resolveLookupTitles(client2, lookupGuids, list2.Id);
  const listIdClean = cleanGuid2(list2.Id);
  const unreadable = /* @__PURE__ */ new Set();
  const lookupTitleById = (guid) => {
    const id = cleanGuid2(guid);
    if (id === listIdClean) return list2.Title;
    const title = lookupTitles.get(id);
    if (title == null && lookupTitles.has(id) && !unreadable.has(id)) {
      unreadable.add(id);
      warnings.push(`Lookup target list ${guid} could not be read; it will need a lookupMap entry.`);
    }
    return title ?? null;
  };
  const fields = rawFields.map((raw) => normalizeField(raw, { listId: list2.Id, lookupTitleById }));
  for (const f of fields) {
    if (f.custom && TAXONOMY_TYPES.has(f.type)) {
      warnings.push(`Managed metadata column \u201C${f.internalName}\u201D cannot be recreated automatically (needs a term-set binding).`);
    }
  }
  let rawViews;
  try {
    ({ items: rawViews } = await client2.getAll(
      guidPath(listId, "/views"),
      { filter: "PersonalView eq false", expand: "ViewFields" }
    ));
  } catch (err) {
    if (isDeniedRead(err) || isExpiredSession(err)) throw err;
    ({ items: rawViews } = await client2.getAll(
      guidPath(listId, "/views"),
      { filter: "PersonalView eq false" }
    ));
  }
  const views = [];
  for (const raw of rawViews) {
    let names = raw.ViewFields?.Items?.results ?? raw.ViewFields?.Items ?? null;
    if (!Array.isArray(names)) {
      try {
        const vf = await client2.get(guidPath(listId, `/views(guid'${raw.Id}')/viewfields`));
        names = vf?.Items?.results ?? vf?.Items ?? [];
      } catch (err) {
        warnings.push(`View \u201C${raw.Title}\u201D: fields could not be read (${err?.message || err}).`);
        names = [];
      }
    }
    views.push(normalizeView(raw, names));
  }
  let contentTypes = [];
  try {
    const { items: rawCts } = await client2.getAll(guidPath(listId, "/contenttypes"), { expand: "FieldLinks" });
    contentTypes = rawCts.map((raw) => {
      const links = raw.FieldLinks?.Items?.results ?? raw.FieldLinks?.Items ?? [];
      const fieldLinks = links.map((l) => ({ id: l.Id, name: l.Name, required: !!l.Required, hidden: !!l.Hidden }));
      return normalizeContentType(raw, fieldLinks);
    });
  } catch (err) {
    warnings.push(`Content types could not be read (${err?.message || err}).`);
  }
  const source = {
    siteUrl: client2.webUrl(),
    listTitle: list2.Title,
    listId: listIdClean,
    rootFolder: list2.RootFolder?.ServerRelativeUrl ?? null,
    itemCount: list2.ItemCount ?? null,
    webId: web.Id,
    webTitle: web.Title,
    language: web.Language,
    baseType: list2.BaseType,
    entityTypeName: list2.EntityTypeName,
    rootFolderName: list2.RootFolder?.Name ?? null
  };
  const normalizedList = normalizeList(list2);
  normalizedList.contentTypeOrder = contentTypes.filter((ct) => !ct.hidden).map((ct) => ct.id);
  const templateUrl = normalizedList.library?.documentTemplateUrl || "";
  if (templateUrl && !/\/forms\/template\.dotx$/i.test(templateUrl)) {
    warnings.push("Per-library Forms template \u2014 upload it after the copy.");
  }
  for (const ct of contentTypes) {
    if (ct.documentTemplate) {
      warnings.push(`Per-library Forms template on content type \u2018${ct.name}\u2019 \u2014 upload it after the copy.`);
    }
  }
  const doc = buildSchemaDoc({
    source,
    list: normalizedList,
    fields,
    views,
    contentTypes,
    warnings,
    generatorBuild: APP_BUILD_INFO.build
  });
  return { doc, raw: { web, list: list2, fields: rawFields, views: rawViews } };
}
async function probeTarget(client2, { title, doc } = {}) {
  const { items: allLists } = await client2.getAll("web/lists", {
    select: ["Id", "Title", "BaseTemplate", "BaseType", "ContentTypesEnabled", "RootFolder/ServerRelativeUrl"],
    expand: "RootFolder"
  });
  const wanted = String(title || "").trim().toLowerCase();
  const match = wanted ? allLists.find((l) => String(l.Title || "").toLowerCase() === wanted) : null;
  const existingList = match ? {
    id: match.Id,
    title: match.Title,
    baseTemplate: match.BaseTemplate,
    // The base-type-mismatch check (list-schema.js buildApplyPlan) compares
    // this, not baseTemplate — the target's actual BaseType is what
    // decides list-vs-library, a v1-compatible fact templates alone don't
    // carry (a picture/asset/form library is base type 1 too).
    baseType: match.BaseType,
    contentTypesEnabled: !!match.ContentTypesEnabled,
    rootFolderUrl: match.RootFolder?.ServerRelativeUrl || ""
  } : null;
  let existingFields = [];
  let existingViews = [];
  let existingContentTypeIds = [];
  if (existingList) {
    const { items: rawFields } = await client2.getAll(guidPath(existingList.id, "/fields"));
    existingFields = rawFields.map((f) => ({
      internalName: f.InternalName,
      title: f.Title,
      typeAsString: f.TypeAsString,
      id: f.Id
    }));
    const { items: rawViews } = await client2.getAll(guidPath(existingList.id, "/views"), { select: ["Id", "Title", "DefaultView"] });
    existingViews = rawViews.map((v) => ({ id: v.Id, title: v.Title, defaultView: !!v.DefaultView }));
    if (existingList.contentTypesEnabled) {
      const { items: rawCts } = await client2.getAll(guidPath(existingList.id, "/contenttypes"), { select: ["StringId"] });
      existingContentTypeIds = rawCts.map((c) => parentContentTypeId(c.StringId));
    }
  }
  let availableContentTypes = null;
  if (doc?.list?.contentTypesEnabled) {
    const { items } = await client2.getAll("web/availablecontenttypes", {
      select: ["StringId", "Name", "Group"],
      top: 5e3
    });
    availableContentTypes = items.map((c) => ({ id: c.StringId, name: c.Name, group: c.Group }));
  }
  return {
    existingList,
    existingFields,
    existingViews,
    existingContentTypeIds,
    availableContentTypes,
    targetLists: allLists.map((l) => ({ id: l.Id, title: l.Title }))
  };
}

// ../src/workbench/list-data.js
var DATA_VERSION = 2;
function cleanGuid3(v) {
  const s = String(v ?? "").replace(/[{}]/g, "").trim();
  return s ? s.toLowerCase() : null;
}
function toIdList(v) {
  return (v == null ? [] : Array.isArray(v) ? v : [v]).filter((x) => x != null);
}
function deriveUsersFromResolved(rows, fields) {
  const byId = /* @__PURE__ */ new Map();
  for (const row of rows || []) {
    const resolved = row?._resolved;
    if (!resolved) continue;
    for (const [key2, val] of Object.entries(resolved)) {
      const isUserField = key2 === "Author" || key2 === "Editor" || USER_TYPES.has(fields?.[key2]?.type);
      if (!isUserField) continue;
      const arr = Array.isArray(val) ? val : [val];
      for (const u of arr) {
        if (!u || u.Id == null || byId.has(u.Id)) continue;
        byId.set(u.Id, { Id: u.Id, Email: u.Email || "", LoginName: u.LoginName || "", Title: u.Title || "" });
      }
    }
  }
  return [...byId.values()];
}
function buildDataDoc({
  source = {},
  fields = {},
  items = [],
  folders = [],
  users = [],
  warnings = [],
  generatorBuild = "dev"
} = {}) {
  return {
    kind: DATA_KIND,
    version: DATA_VERSION,
    exported: (/* @__PURE__ */ new Date()).toISOString(),
    generator: { tool: "dcspad-workbench", build: generatorBuild },
    source,
    fields,
    items: [...folders, ...items],
    warnings,
    // v2 additions.
    folders,
    users
  };
}
function normalizeDataDoc(doc) {
  if (!doc || typeof doc !== "object" || doc.kind !== DATA_KIND) {
    throw new SpFileError(
      `Not a list data document (expected kind ${DATA_KIND}).`,
      { code: "bad-data" }
    );
  }
  const sourceVersion = Number(doc.version) || 1;
  const allItems = Array.isArray(doc.items) ? doc.items : [];
  const folders = Array.isArray(doc.folders) ? doc.folders : allItems.filter((i) => i?._folder);
  const items = allItems.filter((i) => !i?._folder);
  const fields = doc.fields || {};
  const users = Array.isArray(doc.users) && doc.users.length ? doc.users : deriveUsersFromResolved([...folders, ...items], fields);
  return {
    kind: doc.kind,
    version: DATA_VERSION,
    exported: doc.exported || "",
    generator: doc.generator || { tool: "unknown", build: "" },
    source: doc.source || {},
    fields,
    items,
    folders,
    users,
    warnings: doc.warnings || [],
    _sourceVersion: sourceVersion
  };
}
function validateDataDoc(doc) {
  if (!doc || typeof doc !== "object") return "not a JSON object";
  const version = doc.version == null ? 1 : Number(doc.version);
  if (!(version === 1 || version === 2)) return `unsupported version ${doc.version}`;
  if (!Array.isArray(doc.items)) return '"items" must be an array';
  if (doc.fields != null && (typeof doc.fields !== "object" || Array.isArray(doc.fields))) {
    return '"fields" must be an object keyed by internal name';
  }
  if (doc.folders != null && !Array.isArray(doc.folders)) return '"folders" must be an array';
  return "";
}
function folderOrder(folders) {
  const depthOf = (f) => {
    const path = f?._folderPath || f?.folderPath || "";
    return path ? path.split("/").length : 0;
  };
  return (folders || []).map((f, i) => [f, i]).sort((a, b) => depthOf(a[0]) - depthOf(b[0]) || a[1] - b[1]).map(([f]) => f);
}
function writableFields(schemaFields, targetFields) {
  const byName = /* @__PURE__ */ new Map();
  for (const f of targetFields || []) {
    const name = f.InternalName ?? f.internalName;
    if (name) byName.set(name, f);
  }
  const out = [];
  for (const [name, meta] of Object.entries(schemaFields || {})) {
    const tf = byName.get(name);
    if (!tf) continue;
    const readOnly = tf.ReadOnlyField ?? tf.readOnly ?? false;
    const typeAsString = tf.TypeAsString ?? tf.type ?? "";
    if (readOnly) continue;
    if (NEVER_WRITE.has(name) || NEVER_WRITE_TYPES.has(typeAsString)) continue;
    if (!(meta?.custom || name === "Title")) continue;
    const lookupListId = cleanGuid3(tf.LookupList ?? tf.lookupListId);
    const sameLookupList = LOOKUP_TYPES.has(typeAsString) && !!meta?.lookupListId && lookupListId === cleanGuid3(meta.lookupListId);
    out.push({ name, tf, meta, sameLookupList, typeAsString });
  }
  return out;
}
function pad2(n) {
  return String(n).padStart(2, "0");
}
function toWebDateString(value, dateFormat, { dateOnly = false } = {}) {
  const f = dateFormat || { order: "mdy", sep: "/", timeSep: ":", offsetMinutes: 0 };
  const utc = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(utc.getTime())) return String(value);
  const local = new Date(utc.getTime() + (Number(f.offsetMinutes) || 0) * 6e4);
  const parts = {
    y: String(local.getUTCFullYear()),
    m: pad2(local.getUTCMonth() + 1),
    d: pad2(local.getUTCDate())
  };
  const dateText = (f.order || "mdy").split("").map((k) => parts[k]).join(f.sep ?? "/");
  if (dateOnly) return dateText;
  return `${dateText} ${pad2(local.getUTCHours())}${f.timeSep ?? ":"}${pad2(local.getUTCMinutes())}`;
}
function fieldDateFormat(dateFormat, name) {
  if (dateFormat?.perField?.has(name)) {
    return { ...dateFormat, offsetMinutes: dateFormat.perField.get(name) };
  }
  return dateFormat;
}
function userFormValue(users, userIds) {
  const keys = [];
  for (const u of users || []) {
    const key2 = String(u?.Email || u?.LoginName || "").toLowerCase();
    const login = key2 && userIds ? userIds.get(key2) : null;
    if (login) keys.push({ Key: login });
  }
  return keys.length ? JSON.stringify(keys) : "";
}
function resolveLookupByShownValue(shown, candidates, { sourceId = null, sameLookupList = false } = {}) {
  if (shown == null || shown === "") {
    return { id: null, error: `source id ${sourceId ?? "?"} had no shown value on the source; reference left empty` };
  }
  const list2 = candidates || [];
  if (list2.length === 1) return { id: list2[0], error: "" };
  if (list2.length > 1 && sameLookupList && sourceId != null && list2.includes(sourceId)) {
    return { id: sourceId, error: "" };
  }
  return {
    id: null,
    error: list2.length ? `"${shown}" matches ${list2.length} items on the target lookup list; reference left empty` : `"${shown}" (source id ${sourceId ?? "?"}) was not found on the target lookup list; reference left empty`
  };
}
function lookupFormValue(resolved, multi, lookupIdsByValue, sameLookupList, errs, fieldName) {
  const ids = [];
  for (const r of resolved || []) {
    const { id, error } = resolveLookupByShownValue(
      r?.value,
      lookupIdsByValue?.get(String(r?.value ?? "")) || [],
      { sourceId: r?.Id, sameLookupList }
    );
    if (id != null) ids.push(id);
    else errs.push({ field: fieldName, message: error });
  }
  if (!ids.length) return "";
  return multi ? `${ids.join(";#")};#` : String(ids[0]);
}
function selfLookupFormValue(resolved, idMap, multi, errs, fieldName) {
  const ids = [];
  for (const r of resolved || []) {
    const mapped = idMap?.[r?.Id];
    if (mapped != null) ids.push(mapped);
    else errs.push({ field: fieldName, message: `references source item ${r?.Id}, which was not imported; reference left empty` });
  }
  if (!ids.length) return "";
  return multi ? `${ids.join(";#")};#` : String(ids[0]);
}
function taxonomyFormValue(raw) {
  const terms = (raw == null ? [] : Array.isArray(raw) ? raw : [raw]).filter((t) => t && t.Label);
  return terms.map((t) => `${t.Label}|${cleanGuid3(t.TermGuid)};`).join("");
}
function toImportFormValues(item2, fields, { dateFormat, userIds, lookupIds, idMap } = {}) {
  const values = [];
  const errors = [];
  for (const w of fields || []) {
    const raw = item2?.[w.name];
    const type = w.typeAsString ?? w.tf?.TypeAsString ?? w.tf?.type ?? "";
    let value = "";
    switch (type) {
      case "Boolean":
        value = raw == null ? "" : raw ? "1" : "0";
        break;
      case "MultiChoice": {
        const arr = raw == null ? [] : Array.isArray(raw) ? raw : [raw];
        value = arr.length ? `;#${arr.join(";#")};#` : "";
        break;
      }
      case "URL":
        value = raw && raw.Url ? `${raw.Url}, ${raw.Description || raw.Url}` : "";
        break;
      case "DateTime":
        value = raw ? toWebDateString(raw, fieldDateFormat(dateFormat, w.name), { dateOnly: (w.tf?.DisplayFormat ?? w.tf?.displayFormat) === 0 }) : "";
        break;
      case "User":
      case "UserMulti":
        value = userFormValue(item2?._resolved?.[w.name], userIds);
        break;
      case "Lookup":
      case "LookupMulti": {
        const multi = type === "LookupMulti";
        value = w.meta?.isSelfLookup && idMap ? selfLookupFormValue(item2?._resolved?.[w.name], idMap, multi, errors, w.name) : lookupFormValue(item2?._resolved?.[w.name], multi, lookupIds?.get(w.name), w.sameLookupList, errors, w.name);
        break;
      }
      case "TaxonomyFieldType":
      case "TaxonomyFieldTypeMulti":
        value = taxonomyFormValue(raw);
        break;
      default:
        value = raw == null ? "" : typeof raw === "object" ? JSON.stringify(raw) : String(raw);
    }
    values.push({ FieldName: w.name, FieldValue: value });
  }
  return { values, errors };
}
function authorshipFormValues(item2, { dateFormat, userIds } = {}) {
  const values = [];
  const author = userFormValue(item2?._resolved?.Author, userIds);
  const editor = userFormValue(item2?._resolved?.Editor, userIds);
  if (author) values.push({ FieldName: "Author", FieldValue: author });
  if (editor) values.push({ FieldName: "Editor", FieldValue: editor });
  if (item2?.Created) values.push({ FieldName: "Created", FieldValue: toWebDateString(item2.Created, fieldDateFormat(dateFormat, "Created")) });
  if (item2?.Modified) values.push({ FieldName: "Modified", FieldValue: toWebDateString(item2.Modified, fieldDateFormat(dateFormat, "Modified")) });
  return values;
}
function partitionPasses(items, fields) {
  const selfLookupFields = (fields || []).filter((w) => LOOKUP_TYPES.has(w.typeAsString) && w.meta?.isSelfLookup);
  const selfLookupNames = new Set(selfLookupFields.map((w) => w.name));
  const pass1Fields = (fields || []).filter((w) => !selfLookupNames.has(w.name));
  const hasAuthorship = (item2) => Boolean(
    toIdList(item2?._resolved?.Author).length || toIdList(item2?._resolved?.Editor).length || item2?.Created || item2?.Modified
  );
  const hasSelfLookupRef = (item2) => selfLookupFields.some((w) => toIdList(item2?._resolved?.[w.name]).length);
  return {
    pass1: { items: items || [], fields: pass1Fields },
    pass2: selfLookupFields.length ? { items: (items || []).filter(hasSelfLookupRef), fields: selfLookupFields } : { items: [], fields: [] },
    pass3: { items: (items || []).filter(hasAuthorship), fields: ["Author", "Editor", "Created", "Modified"] }
  };
}

// ../src/workbench/list-schema-apply.js
var listPath = (id, sub = "") => `web/lists(guid'${id}')${sub}`;
function needsVerboseRetry(err) {
  if (err?.status !== 400) return false;
  const msg = String(err?.message || "");
  return /__metadata/i.test(msg) || /\bSP\.[A-Za-z]+(\.[A-Za-z]+)*\b/.test(msg) || /entity ?type/i.test(msg);
}
async function writeJson(spWrite, method, path, body, entityType, opts = {}) {
  try {
    return await spWrite[method](path, body, opts);
  } catch (err) {
    if (!needsVerboseRetry(err)) throw err;
    const verboseBody = entityType === "SP.XmlSchemaFieldCreationInformation" ? { ...body, parameters: { __metadata: { type: entityType }, ...body.parameters || {} } } : { __metadata: { type: entityType }, ...body };
    return spWrite[method](path, verboseBody, { ...opts, contentType: "application/json;odata=verbose" });
  }
}
async function readTargetFields(ctx2) {
  const { items } = await ctx2.client.getAll(listPath(ctx2.listId, "/fields"), {
    select: ["Id", "InternalName", "Title", "TypeAsString", "Hidden", "ReadOnlyField", "Indexed", "EnforceUniqueValues"]
  });
  return items.map((f) => ({
    id: f.Id,
    internalName: f.InternalName,
    title: f.Title,
    typeAsString: f.TypeAsString,
    hidden: !!f.Hidden,
    readOnly: !!f.ReadOnlyField,
    indexed: !!f.Indexed,
    enforceUniqueValues: !!f.EnforceUniqueValues
  }));
}
async function readTargetListsByTitle(ctx2) {
  const { items } = await ctx2.client.getAll("web/lists", { select: ["Id", "Title"] });
  const map = /* @__PURE__ */ new Map();
  for (const l of items) map.set(String(l.Title || "").toLowerCase(), l.Id);
  return map;
}
async function readTargetViews(ctx2) {
  const { items } = await ctx2.client.getAll(listPath(ctx2.listId, "/views"), {
    select: ["Id", "Title", "DefaultView"]
  });
  return items.map((v) => ({ id: v.Id, title: v.Title, defaultView: !!v.DefaultView }));
}
function seedFieldIdMap(ctx2) {
  for (const f of ctx2.targetFields || []) {
    if (!(f.internalName in ctx2.fieldIdMap)) ctx2.fieldIdMap[f.internalName] = f.id;
  }
}
async function readTargetState(ctx2) {
  ctx2.targetFields = await readTargetFields(ctx2);
  seedFieldIdMap(ctx2);
  ctx2.targetListsByTitle = await readTargetListsByTitle(ctx2);
  if (ctx2.planHasViews) ctx2.targetViews = await readTargetViews(ctx2);
}
async function safeRead(ctx2, report, fn) {
  try {
    await fn();
    return true;
  } catch (err) {
    if (isExpiredSession(err)) {
      report.aborted = "auth";
      return false;
    }
    report.aborted = "probe";
    report.warnings.push(err?.message || String(err));
    return false;
  }
}
function resolveListIdByTitle(ctx2, title) {
  const id = ctx2.targetListsByTitle?.get(String(title || "").toLowerCase());
  if (!id) {
    throw new SpFileError(
      `Lookup target list \u2018${title}\u2019 was not found on the target.`,
      { code: "not-found" }
    );
  }
  return id;
}
function assignViews(steps, targetViews) {
  const byId = /* @__PURE__ */ new Map();
  const claimed = /* @__PURE__ */ new Set();
  const views = steps.filter((s) => s.kind === "view.upsert" && !s.payload.viewId);
  for (const s of steps) if (s.kind === "view.upsert" && s.payload.viewId) claimed.add(s.payload.viewId);
  for (const s of views) {
    const t = targetViews.find((tv) => !claimed.has(tv.id) && String(tv.title).toLowerCase() === String(s.payload.title).toLowerCase());
    if (t) {
      byId.set(s.id, t.id);
      claimed.add(t.id);
    }
  }
  for (const s of views) {
    if (byId.has(s.id) || !s.payload.defaultView) continue;
    const t = targetViews.find((tv) => tv.defaultView && !claimed.has(tv.id));
    if (t) {
      byId.set(s.id, t.id);
      claimed.add(t.id);
    }
  }
  return byId;
}
async function runListCreate(step2, ctx2, report) {
  const { title, description, baseTemplate, contentTypesEnabled, urlName } = step2.payload;
  const proposedUrlName = urlName || String(title || "").replace(/\s+/g, "");
  const isLib = Number(baseTemplate) === 101;
  if (!ctx2.spWrite.isMock()) {
    let webRel = "";
    try {
      webRel = new URL(ctx2.client.webUrl()).pathname.replace(/\/+$/, "");
    } catch {
    }
    const listUrl = isLib ? `${webRel}/${proposedUrlName}` : `${webRel}/Lists/${proposedUrlName}`;
    try {
      const existing = await ctx2.client.get(
        `web/GetList(@listUrl)?@listUrl='${odataPathLiteral(listUrl)}'&$select=Id,Title`
      );
      if (existing?.Title && existing.Title !== title) {
        throw new SpFileError(
          `The list URL \u2018${listUrl}\u2019 is already taken by \u2018${existing.Title}\u2019 \u2014 it may be sitting in the recycle bin. Choose a different title, or empty the recycle bin first.`,
          { code: "url-taken" }
        );
      }
    } catch (err) {
      if (err?.status === 404 || err?.code === "not-found") {
      } else throw err;
    }
  }
  const body = {
    Title: urlName || title,
    Description: description || "",
    BaseTemplate: baseTemplate,
    ContentTypesEnabled: !!contentTypesEnabled,
    AllowContentTypes: true
  };
  const data = await writeJson(
    ctx2.spWrite,
    "postJson",
    "web/lists",
    body,
    "SP.List",
    { fallback: "Could not create the list", code: "write" }
  );
  ctx2.listId = data.Id;
  report.listId = data.Id;
  report.created = true;
  report.rootFolder = data.RootFolder?.ServerRelativeUrl || "";
  if (urlName) {
    try {
      await writeJson(ctx2.spWrite, "mergeJson", listPath(data.Id), { Title: title }, "SP.List", {});
    } catch (err) {
      if (isExpiredSession(err)) ctx2.abortAfterStep = "auth";
      report.warnings.push(
        `The list was created under the URL name \u2018${urlName}\u2019 \u2014 renaming its title to \u2018${title}\u2019 failed (${err.message || err}).`
      );
    }
  }
  return data;
}
async function runListAdopt(step2) {
  const { listId, title } = step2.payload;
  return { id: listId, title };
}
async function mergeSettingsGroup(ctx2, report, group) {
  const keys = Object.keys(group || {});
  if (!keys.length) return;
  try {
    await writeJson(
      ctx2.spWrite,
      "mergeJson",
      listPath(ctx2.listId),
      group,
      "SP.List",
      { fallback: "Could not apply list settings", code: "write" }
    );
    report.settings.applied.push(...keys);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    for (const key2 of keys) {
      try {
        await writeJson(
          ctx2.spWrite,
          "mergeJson",
          listPath(ctx2.listId),
          { [key2]: group[key2] },
          "SP.List",
          { fallback: `Could not apply ${key2}`, code: "write" }
        );
        report.settings.applied.push(key2);
      } catch (err2) {
        if (isExpiredSession(err2)) throw err2;
        report.settings.failed.push(key2);
        report.warnings.push(`Setting \u2018${key2}\u2019 was rejected on the target \u2014 ${err2.message}`);
      }
    }
  }
}
async function runListSettings(step2, ctx2, report) {
  const { groupA, groupB } = step2.payload;
  await mergeSettingsGroup(ctx2, report, groupA);
  await mergeSettingsGroup(ctx2, report, groupB);
}
async function runCtAttach(step2, ctx2) {
  const { contentTypeId } = step2.payload;
  return writeJson(
    ctx2.spWrite,
    "postJson",
    `${listPath(ctx2.listId)}/contenttypes/addAvailableContentType`,
    { contentTypeId },
    "SP.List",
    { fallback: "Could not attach the content type", code: "write" }
  );
}
function resolveFieldRefs(step2, ctx2) {
  const refs = step2.refs || {};
  let lookupListId = null;
  if (refs.lookupListId) {
    if (refs.lookupListId.self) lookupListId = ctx2.listId;
    else if (refs.lookupListId.id) lookupListId = refs.lookupListId.id;
    else if (refs.lookupListId.list) lookupListId = resolveListIdByTitle(ctx2, refs.lookupListId.list);
  }
  let primaryFieldId = null;
  if (refs.primaryFieldId?.field) {
    primaryFieldId = ctx2.fieldIdMap[refs.primaryFieldId.field] || null;
    if (!primaryFieldId) {
      throw new SpFileError(
        `Its primary lookup column \u2018${refs.primaryFieldId.field}\u2019 was not created on the target.`,
        { code: "blocked" }
      );
    }
  }
  return { lookupListId, primaryFieldId };
}
async function runFieldCreate(step2, ctx2) {
  const { field: f, asText, options } = step2.payload;
  const present = (ctx2.targetFields || []).find((tf) => tf.internalName === f.internalName);
  if (present) {
    if (present.typeAsString !== f.type) {
      throw new SpFileError(
        `already on the target as ${present.typeAsString}, source is ${f.type} \u2014 values will not import.`,
        { code: "write" }
      );
    }
    ctx2.fieldIdMap[f.internalName] = present.id;
    step2.status = "skipped";
    step2.skipReason = "already on the target";
    return { id: present.id, alreadyPresent: true };
  }
  const { lookupListId, primaryFieldId } = resolveFieldRefs(step2, ctx2);
  const xml = asText ? textFallbackXml(f) : scrubSchemaXml(f.schemaXml, { lookupListId, primaryFieldId, fieldType: f.type });
  const body = { parameters: { SchemaXml: xml, Options: options ?? 8 } };
  const data = await writeJson(
    ctx2.spWrite,
    "postJson",
    `${listPath(ctx2.listId)}/fields/createfieldasxml`,
    body,
    "SP.XmlSchemaFieldCreationInformation",
    { fallback: `Could not create the column \u2018${f.displayName}\u2019`, code: "write" }
  );
  if (data.InternalName && data.InternalName !== f.internalName) {
    throw new SpFileError(
      `Created as \u2018${data.InternalName}\u2019, expected \u2018${f.internalName}\u2019.`,
      { code: "name-mismatch" }
    );
  }
  ctx2.fieldIdMap[f.internalName] = data.Id;
  return data;
}
var FIELD_MERGE_ORDER = ["Indexed", "EnforceUniqueValues", "CustomFormatter"];
async function runFieldMerge(step2, ctx2, report) {
  const { internalName, merges } = step2.payload;
  const fid = ctx2.fieldIdMap[internalName];
  if (!fid) {
    throw new SpFileError(`Column \u2018${internalName}\u2019 has no id on the target to merge onto.`, { code: "write" });
  }
  const keys = FIELD_MERGE_ORDER.filter((k) => k in merges);
  const failures = [];
  for (const key2 of keys) {
    try {
      await writeJson(
        ctx2.spWrite,
        "mergeJson",
        `${listPath(ctx2.listId)}/fields(guid'${fid}')`,
        { [key2]: merges[key2] },
        "SP.Field",
        { fallback: `Could not apply ${key2}`, code: "write" }
      );
      report.fieldMerges.applied++;
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.fieldMerges.failed++;
      failures.push({ key: key2, message: err.message });
    }
  }
  if (keys.length && failures.length === keys.length) {
    throw new SpFileError(failures.map((f) => f.message).join(" "), { code: "write" });
  }
  for (const f of failures) {
    report.warnings.push(`\u2018${internalName}\u2019: ${f.key} was not applied \u2014 ${f.message}`);
  }
  return { failed: failures.map((f) => f.key) };
}
async function runFieldBase(step2, ctx2) {
  const { internalName, displayName, required, description } = step2.payload;
  const body = { Title: displayName };
  if (internalName === "Title") body.Required = !!required;
  if (description) body.Description = description;
  return writeJson(
    ctx2.spWrite,
    "mergeJson",
    `${listPath(ctx2.listId)}/fields/getbyinternalnameortitle('${odataPathLiteral(internalName)}')`,
    body,
    "SP.Field",
    { fallback: `Could not update the ${internalName} column`, code: "write" }
  );
}
async function runViewUpsert(step2, ctx2, report) {
  const {
    title,
    viewId: matchedViewId,
    fields: wanted,
    viewQuery,
    rowLimit,
    paged,
    defaultView,
    customFormatter,
    jsLink,
    viewTypeKind,
    scope,
    aggregations,
    aggregationsStatus,
    tabularView,
    mobileView,
    mobileDefaultView,
    includeRootFolder,
    viewData,
    viewJoins
  } = step2.payload;
  let viewId = matchedViewId;
  if (!viewId) viewId = ctx2.viewAssignment?.get(step2.id) || null;
  let created = false;
  if (viewId) {
    await writeJson(
      ctx2.spWrite,
      "mergeJson",
      `${listPath(ctx2.listId)}/views(guid'${viewId}')`,
      // No Title: a view matched as the target's (localized) default view
      // keeps its own name — the consent only promises rebuilt columns.
      { ViewQuery: viewQuery || "", RowLimit: rowLimit ?? 30, Paged: paged !== false },
      "SP.View",
      { fallback: `Could not update the view \u2018${title}\u2019`, code: "write" }
    );
  } else {
    const body = {
      Title: title,
      PersonalView: false,
      ViewQuery: viewQuery || "",
      RowLimit: rowLimit ?? 30,
      Paged: paged !== false,
      DefaultView: false
    };
    if (viewTypeKind != null && viewTypeKind !== 1) body.ViewTypeKind = viewTypeKind;
    const data = await writeJson(
      ctx2.spWrite,
      "postJson",
      `${listPath(ctx2.listId)}/views`,
      body,
      "SP.View",
      { fallback: `Could not create the view \u2018${title}\u2019`, code: "write" }
    );
    viewId = data.Id;
    created = true;
  }
  const available = new Set((ctx2.targetFields || []).map((f) => f.internalName));
  const toAdd = [];
  for (const name of wanted || []) {
    if (available.has(name)) toAdd.push(name);
    else report.warnings.push(`View \u2018${title}\u2019: column \u2018${name}\u2019 is not on the target \u2014 left out.`);
  }
  let previous = [];
  try {
    const vf = await ctx2.client.get(`${listPath(ctx2.listId)}/views(guid'${viewId}')/viewfields`);
    previous = vf?.Items?.results ?? vf?.Items ?? [];
  } catch {
  }
  await writeJson(
    ctx2.spWrite,
    "postJson",
    `${listPath(ctx2.listId)}/views(guid'${viewId}')/viewfields/removeallviewfields`,
    {},
    "SP.View",
    { fallback: `Could not clear the columns on \u2018${title}\u2019`, code: "write" }
  );
  try {
    for (const name of toAdd) {
      await writeJson(
        ctx2.spWrite,
        "postJson",
        `${listPath(ctx2.listId)}/views(guid'${viewId}')/viewfields/addviewfield('${odataPathLiteral(name)}')`,
        {},
        "SP.View",
        { fallback: `Could not add column \u2018${name}\u2019 to \u2018${title}\u2019`, code: "write" }
      );
    }
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    try {
      await ctx2.spWrite.postJson(`${listPath(ctx2.listId)}/views(guid'${viewId}')/viewfields/removeallviewfields`, {});
      for (const name of previous) {
        await ctx2.spWrite.postJson(
          `${listPath(ctx2.listId)}/views(guid'${viewId}')/viewfields/addviewfield('${odataPathLiteral(name)}')`,
          {}
        );
      }
    } catch {
    }
    throw err;
  }
  const propsMerge = {};
  if (!created) {
    Object.assign(propsMerge, {
      Scope: scope ?? 0,
      TabularView: tabularView !== false,
      MobileView: !!mobileView,
      MobileDefaultView: !!mobileDefaultView,
      IncludeRootFolder: !!includeRootFolder
    });
  }
  if (scope != null && scope !== 0) propsMerge.Scope = scope;
  if (aggregations) propsMerge.Aggregations = aggregations;
  if (aggregationsStatus) propsMerge.AggregationsStatus = aggregationsStatus;
  if (viewData) propsMerge.ViewData = viewData;
  if (viewJoins) propsMerge.ViewJoins = viewJoins;
  if (tabularView === false) propsMerge.TabularView = false;
  if (mobileView === true) propsMerge.MobileView = true;
  if (mobileDefaultView === true) propsMerge.MobileDefaultView = true;
  if (includeRootFolder === true) propsMerge.IncludeRootFolder = true;
  if (Object.keys(propsMerge).length) {
    try {
      await writeJson(ctx2.spWrite, "mergeJson", `${listPath(ctx2.listId)}/views(guid'${viewId}')`, propsMerge, "SP.View", {});
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.warnings.push(`View \u2018${title}\u2019: some view properties were not applied \u2014 ${err.message}`);
    }
  }
  const optionalMerge = async (body, label) => {
    try {
      await writeJson(ctx2.spWrite, "mergeJson", `${listPath(ctx2.listId)}/views(guid'${viewId}')`, body, "SP.View", {});
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.warnings.push(`View \u2018${title}\u2019: ${label} \u2014 ${err.message}`);
    }
  };
  if (customFormatter) await optionalMerge({ CustomFormatter: customFormatter }, "the column-formatting JSON was not applied");
  if (jsLink) await optionalMerge({ JSLink: jsLink }, "JSLink was not applied");
  if (defaultView) await optionalMerge({ DefaultView: true }, "could not be set as the default view");
  return { id: viewId, created };
}
async function runListValidation(step2, ctx2) {
  const { validationFormula, validationMessage } = step2.payload;
  return writeJson(
    ctx2.spWrite,
    "mergeJson",
    listPath(ctx2.listId),
    { ValidationFormula: validationFormula, ValidationMessage: validationMessage || "" },
    "SP.List",
    { fallback: "Could not apply the validation formula", code: "write" }
  );
}
var STEP_RUNNERS = {
  "list.create": runListCreate,
  "list.adopt": runListAdopt,
  "list.settings": runListSettings,
  "ct.attach": runCtAttach,
  "field.create": runFieldCreate,
  "field.merge": runFieldMerge,
  "field.base": runFieldBase,
  "view.upsert": runViewUpsert,
  "list.validation": runListValidation
};
function recordSkippedFieldId(step2, ctx2) {
  if (step2.kind !== "field.create" || step2.status !== "skipped") return;
  const name = step2.payload?.field?.internalName;
  if (!name) return;
  const found = (ctx2.targetFields || []).find((f) => f.internalName === name);
  if (found) ctx2.fieldIdMap[name] = found.id;
}
function tallyStep(step2, report) {
  switch (step2.kind) {
    case "field.create": {
      const name = step2.payload?.field?.internalName;
      if (step2.status === "done") report.fields.added++;
      else if (step2.status === "skipped") report.fields.skipped++;
      else if (step2.status === "failed") report.fields.failed.push({ internalName: name, error: step2.error });
      break;
    }
    case "view.upsert": {
      if (step2.status === "done") {
        if (step2.result?.created) report.views.added++;
        else report.views.updated++;
      } else if (step2.status === "failed") {
        report.views.failed.push({ title: step2.payload?.title, error: step2.error });
      }
      break;
    }
    case "ct.attach": {
      if (step2.status === "done") report.contentTypes.attached++;
      else if (step2.status === "failed") report.contentTypes.failed++;
      else if (step2.status === "skipped") report.contentTypes.skipped++;
      break;
    }
    case "list.validation": {
      report.validation.applied = step2.status === "done";
      if (step2.status === "failed") report.validation.error = step2.error;
      break;
    }
    default:
      break;
  }
}
async function runPlan(plan, ctx2 = {}) {
  if (plan.existingListId && !ctx2.listId) ctx2.listId = plan.existingListId;
  ctx2.fieldIdMap = ctx2.fieldIdMap || {};
  ctx2.targetFields = ctx2.targetFields || [];
  ctx2.targetListsByTitle = ctx2.targetListsByTitle || null;
  ctx2.targetViews = ctx2.targetViews || [];
  ctx2.planHasViews = (plan.steps || []).some((st) => st.kind === "view.upsert");
  ctx2.abortAfterStep = "";
  const report = newReport(plan);
  const steps = report.steps;
  const total = steps.length;
  const byId = new Map(steps.map((s) => [s.id, s]));
  if (ctx2.listId && !ctx2.targetFields.length) {
    const ok = await safeRead(ctx2, report, () => readTargetState(ctx2));
    if (!ok) return report;
  }
  let fieldsRereadForViews = false;
  const isBlockedByDeps = (s) => (s.dependsOn || []).some((depId) => {
    const dep = byId.get(depId);
    return dep && !dep.optional && (dep.status === "failed" || dep.status === "blocked");
  });
  for (let i = 0; i < steps.length; i++) {
    const s = steps[i];
    if (s.status === "skipped" || s.status === "failed" || s.status === "blocked") {
      recordSkippedFieldId(s, ctx2);
      tallyStep(s, report);
      ctx2.onStep?.(s, i, total);
      continue;
    }
    if (ctx2.signal?.aborted) {
      report.aborted = "user";
      break;
    }
    if (isBlockedByDeps(s)) {
      s.status = "blocked";
      tallyStep(s, report);
      ctx2.onStep?.(s, i, total);
      continue;
    }
    if (s.kind === "view.upsert" && !fieldsRereadForViews) {
      fieldsRereadForViews = true;
      const ok = await safeRead(ctx2, report, async () => {
        ctx2.targetFields = await readTargetFields(ctx2);
        ctx2.targetViews = await readTargetViews(ctx2);
      });
      if (!ok) return report;
      ctx2.viewAssignment = assignViews(steps, ctx2.targetViews);
    }
    s.status = "running";
    ctx2.onStep?.(s, i, total);
    let aborted = false;
    try {
      const runner = STEP_RUNNERS[s.kind];
      if (!runner) throw new SpFileError(`No executor for step kind \u2018${s.kind}\u2019.`, { code: "internal" });
      const result = await runner(s, ctx2, report);
      if (s.status === "running") s.status = "done";
      s.result = result ?? null;
    } catch (err) {
      s.status = err?.code === "blocked" ? "blocked" : "failed";
      s.error = err?.message || String(err);
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        aborted = true;
      } else if (s.optional) {
        report.warnings.push(`${s.label}: ${s.error}`);
      }
    }
    recordSkippedFieldId(s, ctx2);
    tallyStep(s, report);
    ctx2.onStep?.(s, i, total);
    if (aborted) break;
    if (ctx2.abortAfterStep) {
      report.aborted = ctx2.abortAfterStep;
      break;
    }
    if (s.status === "done" && (s.kind === "list.create" || s.kind === "list.adopt")) {
      const ok = await safeRead(ctx2, report, () => readTargetState(ctx2));
      if (!ok) return report;
    }
  }
  return report;
}

// ../src/workbench/list-data-capture.js
var guidPath2 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
var DATA_CAP = 1e5;
var ATTACHMENT_FILE_CAP = 10 * 1024 * 1024;
var ATTACHMENT_TOTAL_CAP = 50 * 1024 * 1024;
var toIdList2 = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]).filter((x) => x != null);
async function fetchAttachmentBytes(client2, serverRelativeUrl2) {
  let origin = "";
  try {
    origin = new URL(client2.webUrl()).origin;
  } catch {
  }
  const abs = /^https?:/i.test(serverRelativeUrl2) ? serverRelativeUrl2 : `${origin}${serverRelativeUrl2}`;
  let res;
  try {
    res = await fetch(abs, { credentials: "same-origin" });
  } catch (cause) {
    throw new SpFileError(`Could not reach the source site (${cause.message || cause}).`, { code: "network", cause });
  }
  if (!res.ok) {
    throw new SpFileError(`The source attachment could not be read (HTTP ${res.status}).`, {
      code: res.status === 401 ? "auth" : "network",
      status: res.status
    });
  }
  return res.arrayBuffer();
}
function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 32768) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
  }
  return btoa(binary);
}
async function captureAttachment(client2, rawFile, warnings, totalState) {
  const name = rawFile?.FileName || "";
  const url = rawFile?.ServerRelativeUrl || "";
  if (!name) return null;
  if (totalState.bytes >= ATTACHMENT_TOTAL_CAP) {
    warnings.push(`Attachment "${name}" was not embedded \u2014 the ${ATTACHMENT_TOTAL_CAP / (1024 * 1024)} MB total attachment cap was reached; exported as a link only.`);
    return { name, url };
  }
  let bytes;
  try {
    bytes = await fetchAttachmentBytes(client2, url);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    warnings.push(`Attachment "${name}" could not be read (${err.message || err}); exported as a link only.`);
    return { name, url };
  }
  if (bytes.byteLength > ATTACHMENT_FILE_CAP) {
    warnings.push(`Attachment "${name}" was not embedded \u2014 it is larger than the ${ATTACHMENT_FILE_CAP / (1024 * 1024)} MB per-file cap; exported as a link only.`);
    return { name, url };
  }
  if (totalState.bytes + bytes.byteLength > ATTACHMENT_TOTAL_CAP) {
    warnings.push(`Attachment "${name}" was not embedded \u2014 the ${ATTACHMENT_TOTAL_CAP / (1024 * 1024)} MB total attachment cap was reached; exported as a link only.`);
    return { name, url };
  }
  totalState.bytes += bytes.byteLength;
  return { name, base64: arrayBufferToBase64(bytes) };
}
async function captureListData(client2, listId, { schemaDoc = null, maxItems = null } = {}) {
  const warnings = [];
  const schema = schemaDoc && schemaDoc.kind === SCHEMA_KIND ? schemaDoc : (await captureListSchema(client2, listId)).doc;
  const rootFolder = String(schema.source?.rootFolder || "").replace(/\/+$/, "");
  const relPath = (serverRelative) => {
    const p = String(serverRelative || "");
    if (!rootFolder || !p.toLowerCase().startsWith(rootFolder.toLowerCase())) return "";
    return p.slice(rootFolder.length).replace(/^\/+/, "");
  };
  const cap = maxItems ? Math.max(1, Math.min(Number(maxItems) || DATA_CAP, DATA_CAP)) : DATA_CAP;
  const hasAttachments = schema.fields.some((f) => f.type === "Attachments");
  const expand = ["FieldValuesAsText", ...hasAttachments ? ["AttachmentFiles"] : []];
  const { items: rawRows, partial } = await client2.getAll(
    guidPath2(listId, "/items"),
    { select: ["*", "FSObjType", "FileDirRef", "FileRef", ...expand], expand, orderby: "ID asc" },
    { cap, allowLargeCap: true }
  );
  if (partial) {
    warnings.push(`Only the first ${rawRows.length} item(s) were read \u2014 the list has more than the ${cap}-item cap.`);
  }
  const userFields = schema.fields.filter((f) => USER_TYPES.has(f.type) && (f.custom || f.internalName === "Author" || f.internalName === "Editor"));
  const lookupFields = schema.fields.filter((f) => LOOKUP_TYPES.has(f.type) && f.custom);
  const lookupValues = /* @__PURE__ */ new Map();
  for (const f of lookupFields) {
    if (!f.lookupListId) continue;
    try {
      const { items: targetRows, partial: partial2 } = await client2.getAll(
        guidPath2(f.lookupListId, "/items"),
        { select: ["Id", f.lookupField || "Title"] },
        { cap: DATA_CAP, allowLargeCap: true }
      );
      lookupValues.set(f.internalName, new Map(targetRows.map((r) => [r.Id, r[f.lookupField || "Title"]])));
      if (partial2) {
        warnings.push(`Lookup \u201C${f.internalName}\u201D: the target list has more items than could be indexed at once \u2014 values for ids beyond that are looked up individually.`);
      }
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      warnings.push(`Lookup \u201C${f.internalName}\u201D: target list could not be read (${err?.message || err}); ids exported without values.`);
    }
  }
  let siteUsersById = null;
  const ensureSiteUsers = async () => {
    if (siteUsersById) return siteUsersById;
    siteUsersById = /* @__PURE__ */ new Map();
    try {
      const { items: items2 } = await client2.getAll("web/siteusers", { select: ["Id", "Email", "LoginName", "Title"] });
      for (const u of items2) siteUsersById.set(u.Id, u);
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      warnings.push(`Site users could not be read (${err?.message || err}); person fields resolve one id at a time instead.`);
    }
    return siteUsersById;
  };
  const referencedUserIds = /* @__PURE__ */ new Set();
  const resolveUser = async (id) => {
    const users2 = await ensureSiteUsers();
    referencedUserIds.add(id);
    if (!users2.has(id)) {
      try {
        const u2 = await client2.get(`web/siteusers/getbyid(${Number(id)})`, { select: ["Id", "Email", "LoginName", "Title"] });
        users2.set(id, u2);
      } catch (err) {
        if (isExpiredSession(err)) throw err;
        users2.set(id, null);
      }
    }
    const u = users2.get(id);
    return u ? { Id: u.Id, Email: u.Email || "", LoginName: u.LoginName || "", Title: u.Title || "" } : { Id: id, Email: "", LoginName: "", Title: "" };
  };
  const resolveLookupValue = async (f, id) => {
    const map = lookupValues.get(f.internalName);
    if (!map) return null;
    if (map.has(id)) return map.get(id) ?? null;
    try {
      const row = await client2.get(guidPath2(f.lookupListId, `/items(${id})`), { select: ["Id", f.lookupField || "Title"] });
      const value = row ? row[f.lookupField || "Title"] ?? null : null;
      map.set(id, value);
      return value;
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      return null;
    }
  };
  const items = [];
  const folders = [];
  const attachmentTotal = { bytes: 0 };
  for (const raw of rawRows) {
    const resolved = {};
    for (const f of userFields) {
      const ids = toIdList2(raw[`${f.internalName}Id`]);
      if (ids.length) resolved[f.internalName] = await Promise.all(ids.map(resolveUser));
    }
    for (const f of lookupFields) {
      const ids = toIdList2(raw[`${f.internalName}Id`]);
      if (ids.length) {
        resolved[f.internalName] = await Promise.all(
          ids.map(async (id) => ({ Id: id, value: await resolveLookupValue(f, id) }))
        );
      }
    }
    const rawFiles = Array.isArray(raw.AttachmentFiles) ? raw.AttachmentFiles : raw.AttachmentFiles?.results || [];
    const attachments = [];
    for (const f of rawFiles) {
      const a = await captureAttachment(client2, f, warnings, attachmentTotal);
      if (a) attachments.push(a);
    }
    const { AttachmentFiles: _omit, ...rest } = raw;
    const isFolder = Number(raw.FSObjType) === 1;
    const row = { ...rest, _resolved: resolved, _dir: relPath(raw.FileDirRef) };
    if (attachments.length) row._attachments = attachments;
    if (isFolder) {
      row._folder = true;
      row._folderPath = relPath(raw.FileRef);
      folders.push(row);
    } else {
      items.push(row);
    }
  }
  const fieldMap = {};
  for (const f of schema.fields) {
    fieldMap[f.internalName] = {
      type: f.type,
      custom: f.custom,
      readOnly: f.readOnly,
      lookupList: f.lookupList,
      lookupListId: f.lookupListId,
      lookupField: f.lookupField,
      isSelfLookup: f.isSelfLookup,
      allowMultipleValues: f.allowMultipleValues
    };
  }
  const users = siteUsersById ? [...referencedUserIds].map((id) => siteUsersById.get(id)).filter(Boolean).map((u) => ({ Id: u.Id, Email: u.Email || "", LoginName: u.LoginName || "", Title: u.Title || "" })) : [];
  const source = {
    siteUrl: client2.webUrl(),
    listTitle: schema.source.listTitle,
    listId: schema.source.listId,
    rootFolder: schema.source.rootFolder ?? null,
    itemCount: schema.source.itemCount ?? null
  };
  const doc = buildDataDoc({
    source,
    fields: fieldMap,
    items,
    folders,
    users,
    warnings,
    generatorBuild: APP_BUILD_INFO.build
  });
  return { doc, raw: { items: rawRows } };
}

// ../src/workbench/list-data-apply.js
var guidPath3 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
var FAILED_CAP = 50;
var LOOKUP_INDEX_CAP = 1e5;
function cleanGuid4(v) {
  const s = String(v ?? "").replace(/[{}]/g, "").trim();
  return s ? s.toLowerCase() : null;
}
function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}
function newDataReport(data) {
  return {
    listTitle: data.source?.listTitle || "",
    attempted: data.items.length,
    items: { added: 0, failed: [] },
    folders: { created: 0, failed: 0 },
    attachments: { added: 0, skipped: 0, failed: 0 },
    authorship: { applied: 0, failed: 0 },
    fieldErrors: [],
    idMap: {},
    warnings: [],
    aborted: ""
  };
}
function capItemFailures(report) {
  if (report.items.failed.length > FAILED_CAP) {
    report.items.failedTruncated = report.items.failed.length - FAILED_CAP;
    report.items.failed = report.items.failed.slice(0, FAILED_CAP);
  }
}
async function readTargetListRow(client2, listId) {
  return client2.get(guidPath3(listId), {
    select: ["BaseType", "Title", "RootFolder/ServerRelativeUrl"],
    expand: "RootFolder"
  });
}
async function readTargetFields2(client2, listId) {
  const { items } = await client2.getAll(guidPath3(listId, "/fields"), {
    select: ["Id", "InternalName", "TypeAsString", "ReadOnlyField", "LookupList", "LookupField", "DisplayFormat"]
  });
  return items;
}
async function getRegionalSettings(client2) {
  return client2.get("web/RegionalSettings", {
    select: ["DateFormat", "DateSeparator", "TimeSeparator", "Time24", "AM", "PM", "LocaleId"]
  });
}
function formatFromSettings(s) {
  return {
    order: s?.DateFormat === 2 ? "ymd" : s?.DateFormat === 1 ? "dmy" : "mdy",
    sep: s?.DateSeparator || "/",
    timeSep: s?.TimeSeparator || ":"
  };
}
function formatFromSample(sample) {
  const m = /(\d{1,4})([^\d\s])(\d{1,2})\2(\d{1,4})/.exec(String(sample || ""));
  if (!m) return null;
  const parts = [m[1], m[3], m[4]];
  const order = parts.map((p) => p.length === 4 ? "y" : Number(p) === 23 ? "d" : "m").join("");
  if (!/^(mdy|dmy|ymd)$/.test(order)) return null;
  const rest = sample.slice(m.index + m[0].length);
  const tm = /\d{1,2}([^\d\s])\d{2}/.exec(rest);
  return { order, sep: m[2], timeSep: tm ? tm[1] : ":" };
}
async function calibrateDateFormat(client2, spWrite, listId, rootFolder, probeField) {
  const settings = await getRegionalSettings(client2);
  const fallback = formatFromSettings(settings);
  try {
    await spWrite.addValidateUpdateItem(listId, {
      folderPath: rootFolder,
      underlyingObjectType: 0,
      formValues: [{ FieldName: probeField, FieldValue: "not-a-date" }]
    });
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    const sample = err?.fieldErrors?.[probeField];
    if (sample) {
      const parsed = formatFromSample(sample.split(":").slice(1).join(":"));
      if (parsed) return parsed;
    }
  }
  return fallback;
}
async function utcOffsetAtMs(client2, isoInstant) {
  const data = await client2.get(`web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${isoInstant}'`);
  const local = data?.value;
  if (!local) {
    throw new SpFileError("Web time zone could not be read; date not written.", { code: "write" });
  }
  return (/* @__PURE__ */ new Date(`${local}Z`)).getTime() - new Date(isoInstant).getTime();
}
async function webLocalOffsetMinutes(cache, client2, utc) {
  const dayOf = (d) => d.toISOString().slice(0, 10);
  const offsetForDay = async (day2) => {
    if (!cache.has(day2)) cache.set(day2, await utcOffsetAtMs(client2, `${day2}T12:00:00Z`));
    return cache.get(day2);
  };
  const day = dayOf(utc);
  const here = await offsetForDay(day);
  const prev = await offsetForDay(dayOf(new Date(utc.getTime() - 864e5)));
  const next = await offsetForDay(dayOf(new Date(utc.getTime() + 864e5)));
  let offsetMs = here;
  if (here !== prev || here !== next) {
    offsetMs = await utcOffsetAtMs(client2, utc.toISOString());
  }
  return offsetMs / 6e4;
}
async function buildUserIds(spWrite, users, report) {
  const userIds = /* @__PURE__ */ new Map();
  for (const u of users || []) {
    const key2 = String(u?.Email || u?.LoginName || "").toLowerCase();
    if (!key2) continue;
    if (!u.Email && !/^i:0#\.f\|membership\|/i.test(u.LoginName || "")) continue;
    try {
      const ensured = await spWrite.ensureUser(u.Email || u.LoginName);
      if (ensured?.loginName) userIds.set(key2, ensured.loginName);
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return userIds;
      }
      if (u.Email) userIds.set(key2, `i:0#.f|membership|${u.Email.toLowerCase()}`);
    }
  }
  return userIds;
}
async function runAttachments(ctx2, report, item2, newItemId) {
  for (const a of item2._attachments) {
    if (!ctx2.includeAttachments) {
      report.attachments.skipped++;
      continue;
    }
    try {
      let bytes;
      if (typeof a.base64 === "string") {
        bytes = base64ToArrayBuffer(a.base64);
      } else if (ctx2.sourceClient && ctx2.sourceClient.context().live && a.url) {
        bytes = await fetchAttachmentBytes(ctx2.sourceClient, a.url);
      } else {
        report.attachments.skipped++;
        if (!ctx2.warnedNoSourceClient) {
          ctx2.warnedNoSourceClient = true;
          report.warnings.push(
            "Some attachments were not copied \u2014 the imported data document has no attachment bytes for them; use Copy to\u2026 from the source site to bring them across."
          );
        }
        continue;
      }
      await ctx2.spWrite.addAttachment(ctx2.listId, newItemId, a.name, bytes);
      report.attachments.added++;
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.attachments.failed++;
      report.warnings.push(`Item ${item2.Id} \u2192 ${newItemId}: attachment "${a.name}" failed (${err.message || err}).`);
    }
  }
}
async function createItemWithRetry(ctx2, report, sourceId, spec) {
  try {
    return await ctx2.spWrite.addValidateUpdateItem(ctx2.listId, spec);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    if (err?.code === "metadata-write" && err.fieldErrors && Object.keys(err.fieldErrors).length) {
      for (const [field2, message] of Object.entries(err.fieldErrors)) {
        report.fieldErrors.push({ sourceId, field: field2, message });
      }
      const drop = new Set(Object.keys(err.fieldErrors));
      const retryValues = spec.formValues.filter((v) => !drop.has(v.FieldName));
      const retried = await ctx2.spWrite.addValidateUpdateItem(ctx2.listId, { ...spec, formValues: retryValues });
      return { ...retried, droppedFields: [...drop] };
    }
    throw err;
  }
}
async function resolveItemFields(ctx2, report, item2, fields) {
  const dateNames = fields.filter((w) => w.typeAsString === "DateTime").map((w) => w.name);
  if (!dateNames.length) return { dateFormat: ctx2.baseFormat, fields };
  const dateFormat = await ctx2.dateFormatForFields(item2, dateNames);
  const failed = dateNames.filter((n) => dateFormat.perField.get(n) === null);
  for (const n of failed) {
    if (!ctx2.warnedDateFields.has(n)) {
      ctx2.warnedDateFields.add(n);
      report.warnings.push(`\u201C${n}\u201D not written \u2014 the target web's date format could not be learned.`);
    }
  }
  return { dateFormat, fields: failed.length ? fields.filter((w) => !failed.includes(w.name)) : fields };
}
async function runFolders(ctx2, data, report, pass1Fields, userIds, lookupIds, onStep) {
  const ordered = folderOrder(data.folders);
  let i = 0;
  for (const f of ordered) {
    if (ctx2.signal?.aborted) {
      report.aborted = "user";
      return false;
    }
    i++;
    onStep?.({ phase: "folders", index: i, total: ordered.length, label: `Creating folder ${i} of ${ordered.length}\u2026` });
    const path = String(f._folderPath || f.folderPath || "");
    const slash = path.lastIndexOf("/");
    const parent = slash === -1 ? ctx2.rootFolder : `${ctx2.rootFolder}/${path.slice(0, slash)}`;
    const name = path.slice(slash + 1);
    const { dateFormat, fields } = await resolveItemFields(ctx2, report, f, pass1Fields);
    const { values, errors } = toImportFormValues(f, fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: f.Id, ...e });
    const formValues = [{ FieldName: "Title", FieldValue: name }, ...values.filter((v) => v.FieldName !== "Title")];
    try {
      const { id } = await createItemWithRetry(ctx2, report, f.Id, {
        folderPath: parent,
        underlyingObjectType: 1,
        leafName: name,
        formValues
      });
      report.idMap[f.Id] = id;
      report.folders.created++;
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return false;
      }
      report.folders.failed++;
      report.warnings.push(`Folder "${path}" could not be created (${err.message || err}); its items will fail unless the folder exists.`);
    }
  }
  return true;
}
async function runItems(ctx2, data, report, pass1Fields, userIds, lookupIds, onStep) {
  const items = [...data.items].sort((a, b) => (a.Id ?? 0) - (b.Id ?? 0));
  let done = 0;
  for (const item2 of items) {
    if (ctx2.signal?.aborted) {
      report.aborted = "user";
      return false;
    }
    const { dateFormat, fields } = await resolveItemFields(ctx2, report, item2, pass1Fields);
    const { values, errors } = toImportFormValues(item2, fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: item2.Id, ...e });
    const folderPath = item2._dir ? `${ctx2.rootFolder}/${item2._dir}` : ctx2.rootFolder;
    try {
      const { id, droppedFields } = await createItemWithRetry(ctx2, report, item2.Id, {
        folderPath,
        underlyingObjectType: 0,
        formValues: values
      });
      report.idMap[item2.Id] = id;
      report.items.added++;
      if (droppedFields?.length) {
        report.warnings.push(`Item ${item2.Id} \u2192 ${id}: created without ${droppedFields.join(", ")} (see field errors).`);
      }
      if (Array.isArray(item2._attachments) && item2._attachments.length) {
        await runAttachments(ctx2, report, item2, id);
      }
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return false;
      }
      report.items.failed.push({ sourceId: item2.Id, error: err.message || String(err) });
    }
    done++;
    onStep?.({ phase: "items", index: done, total: items.length, label: `Copying item ${done} of ${items.length}\u2026` });
  }
  capItemFailures(report);
  return true;
}
async function runSelfLookups(ctx2, data, report, pass2, onStep) {
  if (!pass2.fields.length) return true;
  const allRows = [...data.folders, ...data.items];
  const updates = [];
  for (const item2 of allRows) {
    const newId2 = report.idMap[item2.Id];
    if (!newId2) continue;
    const { values, errors } = toImportFormValues(item2, pass2.fields, { idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: item2.Id, ...e });
    const set = values.filter((v) => v.FieldValue);
    if (set.length) updates.push({ item: item2, newId: newId2, formValues: set });
  }
  let i = 0;
  for (const { item: item2, newId: newId2, formValues } of updates) {
    if (ctx2.signal?.aborted) {
      report.aborted = "user";
      return false;
    }
    i++;
    onStep?.({ phase: "selfLookups", index: i, total: updates.length, label: `Linking self-references ${i} of ${updates.length}\u2026` });
    try {
      await ctx2.spWrite.validateUpdateListItem({ listId: ctx2.listId, itemId: newId2 }, formValues);
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return false;
      }
      if (err.fieldErrors) {
        for (const [field2, message] of Object.entries(err.fieldErrors)) report.fieldErrors.push({ sourceId: item2.Id, field: field2, message });
      } else {
        report.fieldErrors.push({ sourceId: item2.Id, field: pass2.fields.map((f) => f.name).join(","), message: err.message });
      }
    }
  }
  return true;
}
async function runAuthorship(ctx2, data, report, userIds, onStep) {
  const allRows = [...data.folders, ...data.items].filter((i2) => report.idMap[i2.Id]);
  let i = 0;
  for (const item2 of allRows) {
    if (ctx2.signal?.aborted) {
      report.aborted = "user";
      return false;
    }
    i++;
    onStep?.({ phase: "authorship", index: i, total: allRows.length, label: `Restoring authorship ${i} of ${allRows.length}\u2026` });
    let dateFormat = ctx2.baseFormat;
    let dropNames = [];
    const wanted = ["Created", "Modified"].filter((n) => item2?.[n] != null);
    if (wanted.length) {
      if (!ctx2.dateCalibrationOk) {
        dropNames = wanted;
      } else {
        dateFormat = await ctx2.dateFormatForFields(item2, wanted);
        dropNames = wanted.filter((n) => dateFormat.perField.get(n) === null);
      }
    }
    for (const n of dropNames) {
      if (!ctx2.warnedDateFields.has(n)) {
        ctx2.warnedDateFields.add(n);
        report.warnings.push(`\u201C${n}\u201D not written \u2014 the target web's date format could not be learned.`);
      }
    }
    let values = authorshipFormValues(item2, { dateFormat, userIds });
    if (dropNames.length) values = values.filter((v) => !dropNames.includes(v.FieldName));
    if (!values.length) continue;
    try {
      await ctx2.spWrite.validateUpdateListItem({ listId: ctx2.listId, itemId: report.idMap[item2.Id] }, values, { newDocumentUpdate: true });
      report.authorship.applied++;
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return false;
      }
      report.authorship.failed++;
      if (err.fieldErrors) {
        for (const [field2, message] of Object.entries(err.fieldErrors)) report.fieldErrors.push({ sourceId: item2.Id, field: field2, message });
      } else {
        report.warnings.push(`Item ${item2.Id}: authorship could not be restored (${err.message || err}).`);
      }
    }
  }
  return true;
}
async function applyListData({
  dataDoc,
  client: client2,
  spWrite,
  listId,
  options = {},
  onStep,
  signal
} = {}) {
  const { includeAttachments = false, preserveAuthorship = false, sourceClient = null } = options;
  const data = normalizeDataDoc(dataDoc);
  const report = newDataReport(data);
  if (signal?.aborted) {
    report.aborted = "user";
    return report;
  }
  let listRow;
  try {
    listRow = await readTargetListRow(client2, listId);
  } catch (err) {
    if (isExpiredSession(err)) {
      report.aborted = "auth";
      return report;
    }
    throw err;
  }
  if (Number(listRow.BaseType) === 1) {
    throw new SpFileError(
      "Item import into a document library arrives with stage 2.",
      { code: "library-items" }
    );
  }
  report.listTitle = listRow.Title || report.listTitle;
  const rootFolder = String(listRow.RootFolder?.ServerRelativeUrl || "").replace(/\/+$/, "");
  let targetFieldRows;
  try {
    targetFieldRows = await readTargetFields2(client2, listId);
  } catch (err) {
    if (isExpiredSession(err)) {
      report.aborted = "auth";
      return report;
    }
    throw err;
  }
  const writable = writableFields(data.fields, targetFieldRows);
  const isMock = spWrite.isMock();
  const dayOffsetCache = /* @__PURE__ */ new Map();
  let baseFormat = { order: "mdy", sep: "/", timeSep: ":" };
  let dateCalibrationOk = true;
  const dateFieldsAll = writable.filter((w) => w.typeAsString === "DateTime");
  if (!isMock && (dateFieldsAll.length || preserveAuthorship)) {
    try {
      baseFormat = await calibrateDateFormat(
        client2,
        spWrite,
        listId,
        rootFolder,
        dateFieldsAll[0] ? dateFieldsAll[0].name : "Created"
      );
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return report;
      }
      dateCalibrationOk = false;
    }
  }
  const warnedDateFields = /* @__PURE__ */ new Set();
  if (!dateCalibrationOk) {
    for (const w of dateFieldsAll) {
      warnedDateFields.add(w.name);
      report.warnings.push(`\u201C${w.name}\u201D not written \u2014 the target web's date format could not be learned.`);
    }
    if (preserveAuthorship) {
      for (const n of ["Created", "Modified"]) warnedDateFields.add(n);
      report.warnings.push("\u201CCreated\u201D/\u201CModified\u201D not written \u2014 the target web\u2019s date format could not be learned.");
    }
  }
  const effectiveWritable = dateCalibrationOk ? writable : writable.filter((w) => w.typeAsString !== "DateTime");
  const { pass1, pass2 } = partitionPasses(data.items, effectiveWritable);
  const offsetForValue = async (rawValue) => {
    if (isMock) return 0;
    const utc = rawValue ? new Date(rawValue) : /* @__PURE__ */ new Date();
    if (Number.isNaN(utc.getTime())) return 0;
    try {
      return await webLocalOffsetMinutes(dayOffsetCache, client2, utc);
    } catch {
      return null;
    }
  };
  const dateFormatForFields = async (item2, names) => {
    const perField = /* @__PURE__ */ new Map();
    for (const name of names) {
      const raw = item2?.[name];
      if (raw == null) continue;
      perField.set(name, await offsetForValue(raw));
    }
    return { ...baseFormat, perField };
  };
  const lookupIds = /* @__PURE__ */ new Map();
  for (const w of pass1.fields) {
    if (!LOOKUP_TYPES.has(w.typeAsString)) continue;
    const lookupListId = cleanGuid4(w.tf?.LookupList);
    if (!lookupListId) continue;
    try {
      const showField = w.tf?.LookupField || "Title";
      const { items: rows, partial } = await client2.getAll(
        guidPath3(lookupListId, "/items"),
        { select: ["Id", showField] },
        { cap: LOOKUP_INDEX_CAP, allowLargeCap: true }
      );
      const byValue = /* @__PURE__ */ new Map();
      for (const r of rows) {
        const key2 = String(r[showField] ?? "");
        if (!byValue.has(key2)) byValue.set(key2, []);
        byValue.get(key2).push(r.Id);
      }
      lookupIds.set(w.name, byValue);
      if (partial) {
        report.warnings.push(`Lookup "${w.name}": the target lookup list has more items than could be indexed at once \u2014 some values may be left unresolved.`);
      }
    } catch (err) {
      if (isExpiredSession(err)) {
        report.aborted = "auth";
        return report;
      }
      report.warnings.push(`Lookup "${w.name}": target lookup list could not be read (${err.message || err}); values will be left empty.`);
    }
  }
  const userIds = await buildUserIds(spWrite, data.users, report);
  if (report.aborted) return report;
  const ctx2 = {
    client: client2,
    spWrite,
    listId,
    rootFolder,
    baseFormat,
    dateCalibrationOk,
    dateFormatForFields,
    warnedDateFields,
    signal,
    includeAttachments,
    sourceClient
  };
  if (!await runFolders(ctx2, data, report, pass1.fields, userIds, lookupIds, onStep) || report.aborted) return report;
  if (!await runItems(ctx2, data, report, pass1.fields, userIds, lookupIds, onStep) || report.aborted) return report;
  if (!await runSelfLookups(ctx2, data, report, pass2, onStep) || report.aborted) return report;
  if (preserveAuthorship) await runAuthorship(ctx2, data, report, userIds, onStep);
  return report;
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

// ../src/workbench/grid.js
var el4 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function bindNewTab2(a) {
  a.target = "_blank";
  a.rel = "noopener";
  a.addEventListener("click", (e) => {
    e.stopPropagation();
    if (e.ctrlKey || e.metaKey || e.shiftKey || e.altKey || e.button === 1) return;
    e.preventDefault();
    window.open(a.href, "_blank", "noopener");
  });
  return a;
}
function createMenuButton2(label, title, items) {
  const wrap = el4("span", "wb-menu-wrap");
  const btn = el4("button", "btn btn-xs", label);
  btn.type = "button";
  btn.title = title;
  const menu = el4("div", "wb-menu");
  menu.hidden = true;
  for (const [itemLabel, run] of items) {
    const item2 = el4("button", "wb-menu-item", itemLabel);
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
  return wrap;
}

// ../src/workbench/list-schema-dialog.js
var el5 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var cleanId = (v) => String(v || "").replace(/[{}]/g, "").toLowerCase();
var canonUrl = (u) => String(u || "").replace(/\/+$/, "").toLowerCase();
var plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
var STATE_MAP = {
  planned: "pending",
  running: "running",
  done: "ok",
  skipped: "skip",
  failed: "failed",
  blocked: "blocked"
};
function hasBlockingFieldFailures(report) {
  return (report?.steps || []).some((s) => s.kind === "field.create" && (s.status === "blocked" || s.status === "failed" && !s.final));
}
function verbFor(step2) {
  switch (step2.kind) {
    case "list.create":
      return "Creating the list";
    case "list.adopt":
      return "Opening the existing list";
    case "list.settings":
      return "Applying list settings";
    case "ct.attach":
      return "Attaching content type";
    case "field.create":
      return "Creating field";
    case "field.merge":
      return "Updating field";
    case "field.base":
      return "Updating the Title column";
    case "view.upsert":
      return "Creating view";
    case "list.validation":
      return "Applying the validation formula";
    default:
      return "Working";
  }
}
function stepLine(step2) {
  return step2.error ? `${step2.label} \u2014 ${step2.error}` : step2.label;
}
function buildHeadline(report, { isMock }) {
  if (report.aborted === "auth" || report.itemsReport?.aborted === "auth") return EXPIRED_SESSION_NOTE;
  if (report.aborted === "probe") {
    const msg = report.warnings?.[report.warnings.length - 1] || "a read failed";
    return `The run stopped: ${msg} \u2014 Retry resumes from the list as it is now.`;
  }
  const failed = report.steps.filter((s) => s.status === "failed").length;
  const retryable = report.steps.filter((s) => s.status === "failed" && !s.final).length;
  let headline;
  if (report.adopted) {
    headline = `Updated \u2018${report.title}\u2019 \u2014 ${plural(report.fields.added, "field")} added, ${report.fields.skipped} already present, ${plural(report.views.added + report.views.updated, "view")} rebuilt.`;
  } else {
    const fieldsTotal = report.fields.added + report.fields.skipped + report.fields.failed.length;
    const ctPart = report.contentTypes.attached ? `, ${plural(report.contentTypes.attached, "content type")}` : "";
    const dest = report.targetWebUrl || "the target site";
    headline = `${report.created ? "Created" : "Could not create"} \u2018${report.title}\u2019 on ${dest} \u2014 ${report.fields.added} of ${fieldsTotal} fields, ${plural(report.views.added + report.views.updated, "view")}${ctPart}.`;
  }
  if (failed) {
    headline += retryable ? ` ${plural(failed, "step")} failed \u2014 see below, then Retry failed steps.` : ` ${plural(failed, "step")} could not be applied \u2014 see below.`;
  }
  if (isMock) headline += " (mock mode \u2014 the fixture web does not change).";
  return headline;
}
function openSchemaApplyDialog({
  doc,
  dataDoc = null,
  mode = "copy",
  client: client2,
  createClient,
  navigate,
  inspectSite: inspectSite2,
  mockWriter: mockWriter2
} = {}) {
  return new Promise((resolve) => {
    const d = normalizeSchemaDoc(doc);
    const isLibraryDoc = schemaBaseType(d) === 1;
    let targetClient = createClient();
    const isMockMode = !targetClient.context().live;
    const writerFor = (c) => createSpWriteClient({ client: c, mockWriter: isMockMode ? mockWriter2 : void 0 });
    let spWrite = writerFor(targetClient);
    let probe = { targetLists: [], existingList: null, existingFields: [], existingViews: [], existingContentTypeIds: [], availableContentTypes: null };
    let connected = false;
    let connectedInputValue = null;
    let existingPolicy = "new";
    let phase = "form";
    let listExists = false;
    let lastReport = null;
    let itemsPhaseRun = false;
    let liById = /* @__PURE__ */ new Map();
    let abortController = null;
    let titleTouched = false;
    let lastCollisionKey = null;
    const dialog = el5("dialog", "app-dialog sp-metadata-dialog wb-schema-dialog");
    const panel = el5("div", "app-dialog__panel");
    const head = el5("div", "app-dialog__head");
    head.append(el5("h2", "", mode === "copy" ? "Copy list to\u2026" : "New list from schema"));
    const closeBtn = el5("button", "btn btn-ghost btn-xs wb-schema-close", "\u2715");
    closeBtn.type = "button";
    closeBtn.setAttribute("aria-label", "Close");
    head.append(closeBtn);
    const context = el5("p", "app-dialog__context", mode === "copy" ? `Copy the schema of \u2018${d.list.title}\u2019 to a new or existing list.` : `Create a list from the imported schema for \u2018${d.list.title}\u2019.`);
    panel.append(head, context);
    const targetField = el5("div", "app-dialog__field wb-schema-target-row");
    targetField.append(el5("label", "", "Target site"));
    const targetRow = el5("div", "wb-schema-target-inputrow");
    const targetInput = el5("input", "wb-schema-target");
    targetInput.type = "text";
    targetInput.autocomplete = "off";
    targetInput.setAttribute("list", "wb-schema-target-list");
    const datalist = el5("datalist");
    datalist.id = "wb-schema-target-list";
    targetRow.append(targetInput, datalist);
    const connectBtn = el5("button", "btn btn-xs wb-schema-connect", "Connect");
    connectBtn.type = "button";
    targetRow.append(connectBtn);
    targetField.append(targetRow);
    const targetStatus = el5("div", "wb-schema-target-status");
    targetStatus.hidden = true;
    targetField.append(targetStatus);
    panel.append(targetField);
    function fillDatalist() {
      datalist.textContent = "";
      const host = el5("option", "", "This site (host web)");
      host.value = "";
      datalist.append(host);
      for (const fav of getFavorites()) {
        const opt = el5("option", "", fav.title || fav.url);
        opt.value = fav.url;
        datalist.append(opt);
      }
      for (const rec of getRecents()) {
        const opt = el5("option", "", rec.title || rec.url);
        opt.value = rec.url;
        datalist.append(opt);
      }
    }
    fillDatalist();
    const titleField = el5("div", "app-dialog__field wb-schema-title-row");
    titleField.append(el5("label", "", "Title"));
    const titleInput = el5("input", "wb-schema-title");
    titleInput.type = "text";
    titleInput.autocomplete = "off";
    titleField.append(titleInput);
    const titleStatus = el5("div", "wb-schema-title-status");
    titleField.append(titleStatus);
    panel.append(titleField);
    const descField = el5("div", "app-dialog__field wb-schema-description-row");
    descField.append(el5("label", "", "Description"));
    const descInput = el5("textarea", "wb-schema-description");
    descInput.rows = 2;
    descInput.value = d.list.description || "";
    descField.append(descInput);
    panel.append(descField);
    const lookupSection = el5("div", "wb-schema-lookups-section");
    lookupSection.hidden = true;
    lookupSection.append(el5("h3", "", "Lookup targets"));
    const lookupTable = el5("table", "wb-table wb-schema-lookups");
    const lookupHead = el5("thead");
    const headRow = el5("tr");
    headRow.append(el5("th", "", "Column"), el5("th", "", "Source list"), el5("th", "", "Target list"));
    lookupHead.append(headRow);
    const lookupBody = el5("tbody");
    lookupTable.append(lookupHead, lookupBody);
    lookupSection.append(lookupTable);
    const policyRow = el5("div", "wb-schema-lookup-policy");
    policyRow.hidden = true;
    const policyLabel = el5("div", "", "When a lookup target is missing on the target:");
    const skipLabel = el5("label", "wb-schema-lookup-policy-opt");
    const skipRadio = el5("input");
    skipRadio.type = "radio";
    skipRadio.name = "wb-schema-missing-lookup";
    skipRadio.value = "skip";
    skipRadio.checked = true;
    skipLabel.append(skipRadio, el5("span", "", "Skip the column"));
    const textLabel = el5("label", "wb-schema-lookup-policy-opt");
    const textRadio = el5("input");
    textRadio.type = "radio";
    textRadio.name = "wb-schema-missing-lookup";
    textRadio.value = "text";
    textLabel.append(textRadio, el5("span", "", "Create it as a single line of text"));
    policyRow.append(policyLabel, skipLabel, textLabel);
    lookupSection.append(policyRow);
    panel.append(lookupSection);
    const ctSection = el5("div", "wb-schema-cts-section");
    ctSection.hidden = true;
    ctSection.append(el5("h3", "", "Content types"));
    const ctList = el5("ul", "wb-schema-cts-list");
    ctSection.append(ctList);
    ctSection.append(el5("p", "wb-schema-note", "Missing content types are not created in this stage."));
    panel.append(ctSection);
    const itemsFieldset = el5("fieldset", "wb-schema-items");
    const legend = el5("legend", "", "Items");
    itemsFieldset.append(legend);
    const itemsRow = el5("div", "wb-schema-items-row");
    const includeItemsCb = el5("input");
    includeItemsCb.type = "checkbox";
    const includeItemsLabel = el5("label", "wb-schema-items-opt");
    includeItemsLabel.append(includeItemsCb, el5("span", "", "Include items"));
    const includeAttachmentsCb = el5("input");
    includeAttachmentsCb.type = "checkbox";
    const includeAttachmentsLabel = el5("label", "wb-schema-items-opt");
    includeAttachmentsLabel.append(includeAttachmentsCb, el5("span", "", "Include attachments"));
    const preserveAuthorshipCb = el5("input");
    preserveAuthorshipCb.type = "checkbox";
    const preserveAuthorshipLabel = el5("label", "wb-schema-items-opt");
    preserveAuthorshipLabel.append(preserveAuthorshipCb, el5("span", "", "Preserve authorship"));
    itemsRow.append(includeItemsLabel, includeAttachmentsLabel, preserveAuthorshipLabel);
    itemsFieldset.append(itemsRow);
    const itemsNote = el5("p", "wb-schema-note wb-schema-items-note");
    itemsFieldset.append(itemsNote);
    panel.append(itemsFieldset);
    function updateItemsAvailability() {
      if (isLibraryDoc) {
        itemsFieldset.disabled = true;
        itemsRow.hidden = true;
        itemsNote.textContent = "Files are not copied \u2014 a library copy is schema only.";
        return;
      }
      itemsRow.hidden = false;
      const available = mode === "copy" || Boolean(dataDoc);
      itemsFieldset.disabled = !available;
      itemsNote.textContent = available ? "" : "Choose a matching item-data file alongside the schema (New from schema\u2026) to import items.";
      includeAttachmentsCb.disabled = !available || !includeItemsCb.checked;
      preserveAuthorshipCb.disabled = !available || !includeItemsCb.checked;
    }
    includeItemsCb.addEventListener("change", updateItemsAvailability);
    updateItemsAvailability();
    const existingSection = el5("div", "wb-schema-existing sp-metadata-consent");
    existingSection.hidden = true;
    const newTitleLabel = el5("label", "sp-metadata-consent__row");
    const newTitleRadio = el5("input");
    newTitleRadio.type = "radio";
    newTitleRadio.name = "wb-schema-existing-policy";
    newTitleRadio.value = "new";
    newTitleRadio.checked = true;
    newTitleLabel.append(newTitleRadio, el5("span", "sp-metadata-consent__label", "Choose another title"));
    const resumeLabel = el5("label", "sp-metadata-consent__row");
    const resumeRadio = el5("input");
    resumeRadio.type = "radio";
    resumeRadio.name = "wb-schema-existing-policy";
    resumeRadio.value = "resume";
    resumeLabel.append(resumeRadio, el5("span", "sp-metadata-consent__label", "Add missing fields and views to the existing list"));
    existingSection.append(newTitleLabel, resumeLabel);
    const gateRow = el5("label", "sp-metadata-consent__row wb-schema-gate");
    gateRow.hidden = true;
    const gateBox = el5("input");
    gateBox.type = "checkbox";
    const gateLabel = el5("span", "sp-metadata-consent__label");
    gateRow.append(gateBox, gateLabel);
    existingSection.append(gateRow);
    panel.append(existingSection);
    const planPanel = el5("div", "wb-schema-plan");
    planPanel.hidden = true;
    const planHeading = el5("p", "wb-schema-plan-heading");
    planPanel.append(planHeading);
    const stepsList = el5("ol", "wb-schema-steps");
    planPanel.append(stepsList);
    const masterLine = el5("p", "wb-schema-master");
    masterLine.hidden = true;
    planPanel.append(masterLine);
    panel.append(planPanel);
    const reportPanel = el5("div", "wb-schema-report");
    reportPanel.hidden = true;
    const reportHeadline = el5("p", "wb-schema-report-headline");
    const reportCounts = el5("table", "wb-table wb-schema-report-counts");
    const reportFailed = el5("div", "wb-schema-report-failed");
    const reportWarnings = el5("div", "wb-schema-report-warnings");
    const reportLinks = el5("div", "wb-schema-report-links");
    reportPanel.append(reportHeadline, reportCounts, reportFailed, reportWarnings, reportLinks);
    panel.append(reportPanel);
    const error = el5("div", "sp-files-error");
    error.setAttribute("role", "alert");
    error.hidden = true;
    panel.append(error);
    const actions = el5("div", "app-dialog__actions sp-metadata-actions wb-schema-actions-row");
    const cancelBtn = el5("button", "btn btn-ghost wb-schema-cancel", "Cancel");
    cancelBtn.type = "button";
    const dryRunBtn = el5("button", "btn btn-xs wb-schema-dryrun", "Dry run");
    dryRunBtn.type = "button";
    const retryBtn = el5("button", "btn btn-xs wb-schema-retry", "Retry failed steps");
    retryBtn.type = "button";
    retryBtn.hidden = true;
    const downloadBtn = el5("button", "btn btn-xs wb-schema-report-download", "Download report .md");
    downloadBtn.type = "button";
    downloadBtn.hidden = true;
    const createBtn = el5("button", "btn btn-run wb-schema-create", "Create list");
    createBtn.type = "button";
    actions.append(cancelBtn, dryRunBtn, retryBtn, downloadBtn, createBtn);
    panel.append(actions);
    dialog.append(panel);
    document.body.append(dialog);
    const finish = (outcome) => {
      dialog.close();
      dialog.remove();
      resolve(outcome);
    };
    function showError(message) {
      error.textContent = message;
      error.hidden = !message;
    }
    function sameAsSource() {
      return mode === "copy" && probe.existingList && cleanId(probe.existingList.id) === cleanId(d.source?.listId);
    }
    function baseTypeMismatch() {
      return Boolean(probe.existingList) && probe.existingList.baseType != null && Number(probe.existingList.baseType) !== schemaBaseType(d);
    }
    function unsupportedTemplate() {
      const bt = schemaBaseType(d);
      const isGenericList = bt === 0 && Number(d.list.baseTemplate) === 100;
      return !(isGenericList || bt === 1);
    }
    function eligible() {
      return !unsupportedTemplate() && !baseTypeMismatch();
    }
    function updateTitleStatus() {
      if (!probe.existingList) {
        titleStatus.textContent = connected ? "Available" : "";
        titleStatus.classList.remove("wb-schema-title-taken");
      } else if (sameAsSource()) {
        titleStatus.textContent = "Choose a title different from the source list \u2014 this is the same list on the same site.";
        titleStatus.classList.add("wb-schema-title-taken");
      } else {
        titleStatus.textContent = `A list named \u2018${probe.existingList.title}\u2019 already exists on the target.`;
        titleStatus.classList.add("wb-schema-title-taken");
      }
    }
    function updateExistingSection() {
      const collisionKey = probe.existingList ? `${canonUrl(targetClient.webUrl())}::${cleanId(probe.existingList.id)}` : null;
      if (collisionKey !== lastCollisionKey) {
        lastCollisionKey = collisionKey;
        existingPolicy = "new";
        newTitleRadio.checked = true;
        gateBox.checked = false;
      }
      const show = Boolean(probe.existingList) && !sameAsSource();
      existingSection.hidden = !show;
      if (!show) {
        gateRow.hidden = true;
        return;
      }
      gateLabel.textContent = `I understand this changes \u2018${probe.existingList.title}\u2019: missing columns and views are added, views with the same title have their columns rebuilt, and attachments or folders are switched on if the source needs them. Nothing is deleted.`;
      gateRow.hidden = existingPolicy !== "resume";
    }
    function customLookups() {
      return d.fields.filter((f) => f.custom && LOOKUP_TYPES.has(f.type));
    }
    function buildLookupMap() {
      const map = {};
      for (const select of lookupBody.querySelectorAll("select[data-key]")) {
        if (select.value && select.dataset.key) map[select.dataset.key] = select.value;
      }
      return map;
    }
    function renderLookupRows() {
      const lookups = customLookups();
      lookupSection.hidden = lookups.length === 0;
      lookupBody.textContent = "";
      let anyMissing = false;
      for (const f of lookups) {
        const tr = el5("tr");
        tr.append(el5("td", "", f.displayName || f.internalName));
        tr.append(el5("td", "", f.lookupList || f.lookupListId || "\u2014"));
        const td = el5("td");
        if (f.isSelfLookup) {
          td.append(el5("span", "", "(this list)"));
        } else {
          const select = el5("select", "wb-schema-lookup-target");
          const emptyOpt = el5("option", "", "\u2014 missing \u2014");
          emptyOpt.value = "";
          select.append(emptyOpt);
          for (const l of probe.targetLists || []) {
            const opt = el5("option", "", l.title);
            opt.value = l.title;
            select.append(opt);
          }
          const wanted = String(f.lookupList || "").toLowerCase();
          const found = (probe.targetLists || []).find((l) => String(l.title).toLowerCase() === wanted);
          select.value = found ? found.title : "";
          if (!found) anyMissing = true;
          select.dataset.key = f.lookupList || f.lookupListId || "";
          td.append(select);
        }
        tr.append(td);
        lookupBody.append(tr);
      }
      policyRow.hidden = !anyMissing;
    }
    function renderContentTypes() {
      ctSection.hidden = !d.list.contentTypesEnabled;
      if (!d.list.contentTypesEnabled) return;
      ctList.textContent = "";
      const available = new Set((probe.availableContentTypes || []).map((c) => c.id));
      const already = new Set(probe.existingContentTypeIds || []);
      const seen = /* @__PURE__ */ new Set();
      for (const ct of d.contentTypes) {
        const parentId = ct.parentId || parentContentTypeId(ct.id);
        if (isBuiltinParent(parentId) || seen.has(parentId)) continue;
        seen.add(parentId);
        const status = already.has(parentId) ? "already on the target" : available.has(parentId) ? "available on the target" : "missing on the target; will be skipped";
        ctList.append(el5("li", "", `${ct.name} \u2014 ${status}`));
      }
    }
    function canDryRun() {
      return eligible() && connected && Boolean(titleInput.value.trim()) && !sameAsSource() && (!probe.existingList || existingPolicy === "resume");
    }
    function canCreate() {
      return canDryRun() && (!probe.existingList || gateBox.checked) && connectedInputValue !== null && canonUrl(connectedInputValue) === canonUrl(targetInput.value);
    }
    function refreshGate() {
      dryRunBtn.disabled = !canDryRun();
      createBtn.disabled = !canCreate();
      createBtn.textContent = probe.existingList && existingPolicy === "resume" ? "Add to existing list" : "Create list";
    }
    async function refreshProbeForTitle(title) {
      try {
        probe = await probeTarget(targetClient, { title, doc: d });
      } catch (err) {
        showError(err.message || String(err));
        return;
      }
      updateTitleStatus();
      updateExistingSection();
      renderLookupRows();
      renderContentTypes();
      if (baseTypeMismatch()) {
        showError(`\u2018${probe.existingList.title}\u2019 already exists on the target as a different type of list \u2014 it can\u2019t be used for this schema.`);
      } else if (unsupportedTemplate()) {
        showError("Only generic lists and document libraries can be created from a schema.");
      } else {
        showError("");
      }
      refreshGate();
    }
    let connectSeq = 0;
    async function connect(rawUrl) {
      const seq = ++connectSeq;
      const stale = () => seq !== connectSeq;
      let candidate = null;
      showError("");
      targetStatus.hidden = true;
      connectBtn.disabled = true;
      connected = false;
      connectedInputValue = null;
      refreshGate();
      try {
        candidate = createClient();
        await candidate.connectWeb(rawUrl);
      } catch (err) {
        if (stale()) return false;
        connectBtn.disabled = false;
        targetStatus.textContent = err.message || String(err);
        targetStatus.hidden = false;
        probe = { targetLists: [], existingList: null, existingFields: [], existingViews: [], existingContentTypeIds: [], availableContentTypes: null };
        updateTitleStatus();
        refreshGate();
        return false;
      }
      if (stale()) return false;
      let initial;
      try {
        initial = await probeTarget(candidate, { title: "", doc: d });
      } catch (err) {
        if (stale()) return false;
        connectBtn.disabled = false;
        showError(err.message || String(err));
        return false;
      }
      if (stale()) return false;
      targetClient = candidate;
      spWrite = writerFor(candidate);
      connectBtn.disabled = false;
      connected = true;
      connectedInputValue = rawUrl;
      probe = initial;
      if (!titleTouched) titleInput.value = defaultTargetTitle(d, probe.targetLists);
      await refreshProbeForTitle(titleInput.value.trim());
      return true;
    }
    connectBtn.addEventListener("click", () => connect(targetInput.value));
    targetInput.addEventListener("input", () => {
      connected = false;
      connectedInputValue = null;
      probe = { targetLists: [], existingList: null, existingFields: [], existingViews: [], existingContentTypeIds: [], availableContentTypes: null };
      targetStatus.textContent = "Connect to check this site.";
      targetStatus.hidden = false;
      updateTitleStatus();
      refreshGate();
    });
    let titleDebounce = null;
    titleInput.addEventListener("input", () => {
      titleTouched = true;
      clearTimeout(titleDebounce);
      titleDebounce = setTimeout(() => {
        if (connected) refreshProbeForTitle(titleInput.value.trim());
      }, 150);
    });
    newTitleRadio.addEventListener("change", () => {
      existingPolicy = "new";
      gateBox.checked = false;
      updateExistingSection();
      refreshGate();
    });
    resumeRadio.addEventListener("change", () => {
      existingPolicy = "resume";
      updateExistingSection();
      refreshGate();
    });
    gateBox.addEventListener("change", refreshGate);
    function buildStepsList(steps) {
      stepsList.textContent = "";
      liById = /* @__PURE__ */ new Map();
      for (const step2 of steps) {
        const li = el5("li", "", stepLine(step2));
        li.dataset.state = STATE_MAP[step2.status] || "pending";
        stepsList.append(li);
        liById.set(step2.id, li);
      }
    }
    function handleStep(step2, i, total) {
      const li = liById.get(step2.id);
      if (li) {
        li.dataset.state = STATE_MAP[step2.status] || "pending";
        li.textContent = stepLine(step2);
      }
      masterLine.hidden = false;
      masterLine.textContent = `${verbFor(step2)} ${i + 1} of ${total}\u2026`;
      if (currentCtx?.listId && !listExists) {
        listExists = true;
        closeBtn.hidden = true;
        cancelBtn.hidden = true;
      }
    }
    let currentCtx = null;
    async function runItemsPhase(targetListId) {
      if (!includeItemsCb.checked || itemsFieldset.disabled) return;
      let sourceData = dataDoc;
      if (mode === "copy") {
        masterLine.hidden = false;
        masterLine.textContent = "Reading source items\u2026";
        try {
          const capture = await captureListData(client2, d.source.listId, { schemaDoc: d });
          sourceData = capture.doc;
        } catch (err) {
          lastReport.itemsError = err.message || String(err);
          return;
        }
      }
      if (!sourceData) return;
      try {
        lastReport.itemsReport = await applyListData({
          dataDoc: sourceData,
          client: currentCtx.client,
          spWrite: currentCtx.spWrite,
          listId: targetListId,
          options: {
            includeAttachments: includeAttachmentsCb.checked,
            preserveAuthorship: preserveAuthorshipCb.checked,
            sourceClient: mode === "copy" ? client2 : null
          },
          onStep: (info) => {
            masterLine.hidden = false;
            masterLine.textContent = info.label;
          },
          signal: currentCtx.signal
        });
      } catch (err) {
        lastReport.itemsError = err.message || String(err);
      }
    }
    let savedItemsReport = null;
    let savedItemsError = "";
    async function maybeRunItemsPhase(report, targetListId) {
      if (!includeItemsCb.checked || itemsFieldset.disabled) return;
      if (itemsPhaseRun) {
        if (savedItemsReport) report.itemsReport = savedItemsReport;
        else if (savedItemsError) report.itemsError = savedItemsError;
        return;
      }
      if (hasBlockingFieldFailures(report)) {
        report.itemsHeld = "Items wait until the failed columns are created \u2014 Retry, then items import.";
        return;
      }
      report.itemsHeld = "";
      itemsPhaseRun = true;
      await runItemsPhase(targetListId);
      savedItemsReport = report.itemsReport || null;
      savedItemsError = report.itemsError || "";
    }
    async function runDryRun() {
      if (!canDryRun()) return;
      showError("");
      dryRunBtn.disabled = true;
      try {
        probe = await probeTarget(targetClient, { title: titleInput.value.trim(), doc: d });
        updateTitleStatus();
        updateExistingSection();
        const plan = buildApplyPlan(d, buildOptions(), probe);
        planPanel.hidden = false;
        planHeading.textContent = "Plan \u2014 nothing has been written.";
        masterLine.hidden = true;
        buildStepsList(plan.steps);
      } catch (err) {
        showError(err.message || String(err));
      } finally {
        refreshGate();
      }
    }
    function buildOptions() {
      return {
        title: titleInput.value.trim(),
        description: descInput.value,
        existing: probe.existingList && existingPolicy === "resume" ? "resume" : "fail",
        missingLookup: [...policyRow.querySelectorAll('input[name="wb-schema-missing-lookup"]')].find((r) => r.checked)?.value === "text" ? "text" : "skip",
        lookupMap: buildLookupMap(),
        targetWebUrl: targetClient.webUrl()
      };
    }
    function setPhase(next) {
      phase = next;
      const inForm = next === "form";
      const inRunning = next === "running";
      const inReport = next === "report";
      targetField.hidden = !inForm;
      titleField.hidden = !inForm;
      descField.hidden = !inForm;
      itemsFieldset.hidden = !inForm;
      if (inForm) {
        renderLookupRows();
        renderContentTypes();
        updateExistingSection();
      } else {
        lookupSection.hidden = true;
        ctSection.hidden = true;
        existingSection.hidden = true;
      }
      if (inRunning) planPanel.hidden = false;
      else if (inReport) planPanel.hidden = true;
      reportPanel.hidden = !inReport;
      dryRunBtn.hidden = !inForm;
      createBtn.hidden = !inForm;
      retryBtn.hidden = true;
      downloadBtn.hidden = !inReport;
      cancelBtn.hidden = false;
      cancelBtn.textContent = inReport ? "Close" : "Cancel";
      closeBtn.hidden = false;
    }
    async function runCreate() {
      if (!canCreate()) return;
      showError("");
      const client3 = targetClient;
      const writer = spWrite;
      const unlock = () => {
        targetInput.disabled = false;
        connectBtn.disabled = false;
        dryRunBtn.disabled = false;
        createBtn.disabled = false;
      };
      targetInput.disabled = true;
      connectBtn.disabled = true;
      dryRunBtn.disabled = true;
      createBtn.disabled = true;
      let plan;
      try {
        probe = await probeTarget(client3, { title: titleInput.value.trim(), doc: d });
        plan = buildApplyPlan(d, buildOptions(), probe);
      } catch (err) {
        showError(err.message || String(err));
        unlock();
        return;
      }
      targetInput.disabled = false;
      connectBtn.disabled = false;
      setPhase("running");
      planHeading.textContent = `Applying the plan to \u2018${client3.webUrl()}\u2019\u2026`;
      buildStepsList(plan.steps);
      listExists = false;
      abortController = new AbortController();
      currentCtx = { client: client3, spWrite: writer, onStep: handleStep, signal: abortController.signal };
      try {
        const report = await runPlan(plan, currentCtx);
        lastReport = report;
        if (report.listId && !report.aborted) await maybeRunItemsPhase(report, report.listId);
        renderReport(lastReport);
        setPhase("report");
      } catch (err) {
        setPhase("form");
        showError(err.message || String(err));
        dryRunBtn.disabled = false;
        createBtn.disabled = false;
      }
    }
    async function runRetry() {
      if (!lastReport) return;
      showError("");
      retryBtn.disabled = true;
      let plan;
      const client3 = currentCtx?.client || targetClient;
      const writer = currentCtx?.spWrite || spWrite;
      try {
        let resumeProbe = await probeTarget(client3, { title: lastReport.title, doc: d });
        if (lastReport.listId) {
          const wantedId = String(lastReport.listId).toLowerCase();
          const byId = resumeProbe.targetLists.find((l) => String(l.id).toLowerCase() === wantedId);
          if (!byId) {
            showError(`The list \u2018${lastReport.title}\u2019 could not be found on the target anymore \u2014 retry cannot continue.`);
            retryBtn.disabled = false;
            return;
          }
          if (!resumeProbe.existingList || String(resumeProbe.existingList.id).toLowerCase() !== wantedId) {
            resumeProbe = await probeTarget(client3, { title: byId.title, doc: d });
          }
        }
        if (resumeProbe.existingList) {
          plan = buildApplyPlan(
            d,
            { ...buildOptions(), title: resumeProbe.existingList.title, existing: "resume" },
            resumeProbe
          );
        } else if (lastReport.listId) {
          showError(`The list \u2018${lastReport.title}\u2019 could not be found on the target anymore \u2014 retry cannot continue.`);
          retryBtn.disabled = false;
          return;
        } else {
          plan = buildApplyPlan(d, { ...buildOptions(), title: lastReport.title }, resumeProbe);
        }
      } catch (err) {
        showError(err.message || String(err));
        retryBtn.disabled = false;
        return;
      }
      setPhase("running");
      planHeading.textContent = "Retrying \u2014 resuming from the list as it is now\u2026";
      buildStepsList(plan.steps);
      listExists = Boolean(plan.existingListId);
      abortController = new AbortController();
      currentCtx = { client: client3, spWrite: writer, onStep: handleStep, signal: abortController.signal };
      try {
        const report = await runPlan(plan, currentCtx);
        lastReport = report;
        if (report.listId && !report.aborted) await maybeRunItemsPhase(report, report.listId);
      } catch (err) {
        showError(err.message || String(err));
      } finally {
        setPhase("report");
        renderReport(lastReport);
        retryBtn.disabled = false;
      }
    }
    function renderReportCounts(report) {
      reportCounts.textContent = "";
      const tbody = el5("tbody");
      const rows = [
        ["Fields", `${report.fields.added} added \xB7 ${report.fields.skipped} skipped \xB7 ${report.fields.failed.length} failed`],
        ["Views", `${report.views.added} added \xB7 ${report.views.updated} updated \xB7 ${report.views.failed.length} failed`],
        ["Content types", `${report.contentTypes.attached} attached \xB7 ${report.contentTypes.skipped} skipped \xB7 ${report.contentTypes.failed} failed`],
        ["Validation", report.validation.applied ? "applied" : report.validation.error ? "failed" : "\u2014"]
      ];
      if (report.itemsReport) {
        const ir = report.itemsReport;
        const failedText = ir.items.failedTruncated ? `${ir.items.failed.length}+${ir.items.failedTruncated} failed` : `${ir.items.failed.length} failed`;
        rows.push(["Items", `${ir.items.added} added \xB7 ${failedText}`]);
        rows.push(["Folders", `${ir.folders.created} created \xB7 ${ir.folders.failed} failed`]);
        if (includeAttachmentsCb.checked) {
          rows.push(["Attachments", `${ir.attachments.added} added \xB7 ${ir.attachments.skipped} skipped \xB7 ${ir.attachments.failed} failed`]);
        }
        if (preserveAuthorshipCb.checked) {
          rows.push(["Authorship", `${ir.authorship.applied} applied \xB7 ${ir.authorship.failed} failed`]);
        }
      } else if (report.itemsError) {
        rows.push(["Items", `could not be imported \u2014 ${report.itemsError}`]);
      } else if (report.itemsHeld) {
        rows.push(["Items", report.itemsHeld]);
      }
      for (const [label, value] of rows) {
        const tr = el5("tr");
        tr.append(el5("td", "", label), el5("td", "", value));
        tbody.append(tr);
      }
      reportCounts.append(tbody);
    }
    function renderReportFailed(report) {
      reportFailed.textContent = "";
      const failed = report.steps.filter((s) => s.status === "failed");
      if (failed.length) {
        reportFailed.append(el5("h3", "", "Failed steps"));
        const table2 = el5("table", "wb-table");
        const tbody = el5("tbody");
        for (const s of failed) {
          const tr = el5("tr");
          tr.append(el5("td", "", s.label), el5("td", "", s.error || ""));
          tbody.append(tr);
        }
        table2.append(tbody);
        reportFailed.append(table2);
      }
      const failedItems = report.itemsReport?.items.failed;
      if (failedItems?.length) {
        reportFailed.append(el5("h3", "", "Failed items"));
        const itable = el5("table", "wb-table");
        const itbody = el5("tbody");
        for (const f of failedItems) {
          const tr = el5("tr");
          tr.append(el5("td", "", `Source id ${f.sourceId}`), el5("td", "", f.error || ""));
          itbody.append(tr);
        }
        itable.append(itbody);
        reportFailed.append(itable);
        if (report.itemsReport.items.failedTruncated) {
          reportFailed.append(el5("p", "wb-schema-note", `+${report.itemsReport.items.failedTruncated} more not shown.`));
        }
      }
      const itemFieldErrors = report.itemsReport?.fieldErrors;
      if (itemFieldErrors?.length) {
        reportFailed.append(el5("h3", "", "Item field errors"));
        const ftable = el5("table", "wb-table");
        const ftbody = el5("tbody");
        for (const fe of itemFieldErrors) {
          const tr = el5("tr");
          tr.append(el5("td", "", `Source id ${fe.sourceId} \u2014 ${fe.field}`), el5("td", "", fe.message || ""));
          ftbody.append(tr);
        }
        ftable.append(ftbody);
        reportFailed.append(ftable);
      }
    }
    function renderReportWarnings(report) {
      reportWarnings.textContent = "";
      const warnings = [...report.warnings || [], ...report.itemsReport?.warnings || []];
      if (!warnings.length) return;
      reportWarnings.append(el5("h3", "", "Warnings"));
      const list2 = el5("ul", "wb-grid-notice");
      for (const w of warnings) list2.append(el5("li", "", w));
      reportWarnings.append(list2);
    }
    function renderReportLinks(report) {
      reportLinks.textContent = "";
      if (!report.listId) return;
      const sameWeb = canonUrl(targetClient.webUrl()) === canonUrl(client2.webUrl());
      if (sameWeb) {
        const openBtn = el5("button", "btn btn-xs wb-schema-report-open", "Open the new list");
        openBtn.type = "button";
        openBtn.addEventListener("click", () => {
          navigate({ view: "lists", listId: report.listId, listTitle: report.title });
          finish("created");
        });
        reportLinks.append(openBtn);
      } else {
        const settingsLink = el5("a", "btn btn-xs wb-schema-report-settings", "Open list settings \u2197");
        settingsLink.href = linkUrl(targetClient.webUrl(), LIST_SETTINGS, { guid: report.listId });
        bindNewTab2(settingsLink);
        const inspectBtn = el5("button", "btn btn-xs wb-schema-report-inspect", "Inspect the target site");
        inspectBtn.type = "button";
        inspectBtn.addEventListener("click", async () => {
          await inspectSite2(targetClient.webUrl());
          navigate({ view: "lists", listId: report.listId, listTitle: report.title });
          finish("created");
        });
        reportLinks.append(settingsLink, inspectBtn);
      }
    }
    function renderReport(report) {
      reportHeadline.textContent = buildHeadline(report, { isMock: spWrite.isMock() });
      reportHeadline.classList.toggle("wb-schema-report-auth", report.aborted === "auth");
      renderReportCounts(report);
      renderReportFailed(report);
      renderReportWarnings(report);
      renderReportLinks(report);
      retryBtn.hidden = !(report.steps.some((s) => s.status === "failed" && !s.final) || report.aborted === "probe" && Boolean(report.listId));
      downloadBtn.hidden = false;
    }
    dryRunBtn.addEventListener("click", runDryRun);
    createBtn.addEventListener("click", runCreate);
    retryBtn.addEventListener("click", runRetry);
    downloadBtn.addEventListener("click", () => {
      if (!lastReport) return;
      const stem2 = String(lastReport.title || "list").toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "list";
      downloadText(`schema-report-${stem2}.md`, buildApplyReport({ report: lastReport, doc: d, targetWebUrl: lastReport.targetWebUrl }), "text/markdown");
    });
    cancelBtn.addEventListener("click", () => {
      if (phase === "running") {
        abortController?.abort();
        return;
      }
      if (phase === "report") {
        finish(lastReport?.listId ? "created" : "cancelled");
        return;
      }
      finish("cancelled");
    });
    closeBtn.addEventListener("click", () => {
      if (phase === "running") return;
      finish(phase === "report" && lastReport?.listId ? "created" : "cancelled");
    });
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      if (phase === "running") {
        if (!listExists) abortController?.abort();
        return;
      }
      finish(phase === "report" && lastReport?.listId ? "created" : "cancelled");
    });
    setPhase("form");
    dryRunBtn.disabled = true;
    createBtn.disabled = true;
    dialog.showModal();
    targetInput.value = client2.webUrl();
    connect(client2.webUrl());
  });
}

// ../src/workbench/list-schema-script.js
var TEMPLATE_NAMES = { 100: "GenericList", 101: "DocumentLibrary" };
function planFor(doc, opts = {}) {
  const lookupMap = opts.lookupMap || {};
  const assumed = [...new Set((doc?.fields || []).filter((f) => f.lookupList && !f.isSelfLookup).map((f) => lookupMapGet(lookupMap, f.lookupList) ?? lookupMapGet(lookupMap, f.lookupListId) ?? f.lookupList))].map((title) => ({ title }));
  const assumedContentTypes = [...new Set((doc?.contentTypes || []).map((ct) => ct.parentId || parentContentTypeId(ct.id)).filter((parentId) => !isBuiltinParent(parentId)))].map((id) => ({ id }));
  return buildApplyPlan(doc, {
    title: opts.title,
    description: opts.description,
    lookupMap,
    missingLookup: opts.missingLookup
  }, opts.probe || { targetLists: assumed, availableContentTypes: assumedContentTypes });
}
var isDone = (step2) => !["skipped", "failed", "blocked"].includes(step2.status);
var LIST_TOKEN = "{lookuplist}";
var PRIMARY_TOKEN = "{primaryfield}";
function fieldXml(step2) {
  const f = step2.payload.field;
  if (step2.payload.asText) return textFallbackXml(f);
  const refs = step2.refs || {};
  return scrubSchemaXml(f.schemaXml, {
    fieldType: f.type,
    lookupListId: refs.lookupListId ? "lookuplist" : null,
    primaryFieldId: refs.primaryFieldId ? "primaryfield" : null
  });
}
var FIELD_OPTIONS = 8;
var psEsc = (s) => String(s ?? "").replaceAll("'", "''");
var psLit = (s) => `'${psEsc(s)}'`;
function psCsomValue(v) {
  if (typeof v === "boolean") return v ? "$true" : "$false";
  if (typeof v === "number") return String(v);
  return psLit(v);
}
function psHashtable(obj) {
  const parts = Object.entries(obj).map(([k, v]) => {
    if (typeof v === "boolean") return `${k}=$${v}`;
    if (typeof v === "number") return `${k}=${v}`;
    return `${k}=${psLit(v)}`;
  });
  return `@{ ${parts.join("; ")} }`;
}
var psArray = (arr) => `@(${(arr || []).map((s) => psLit(s)).join(", ")})`;
function pushCsomBlock(lines, plan, props) {
  const entries = Object.entries(props).filter(([, v]) => v != null);
  if (!entries.length) return;
  lines.push(`$l = Get-PnPList -Identity ${psLit(plan.title)}`);
  for (const [k, v] of entries) lines.push(`$l.${k} = ${psCsomValue(v)}`);
  lines.push("$l.Update()");
  lines.push("Invoke-PnPQuery");
}
function toPnpPowerShellProvisioning(doc, opts = {}) {
  const plan = planFor(doc, opts);
  const connect = `Connect-PnPOnline -Url "${opts.targetWebUrl || "https://tenant.sharepoint.com/sites/yoursite"}" -Interactive`;
  const lines = ["# PnP.PowerShell provisioning", `# ${plan.title}`, connect, `$list = ${psLit(plan.title)}`, ""];
  for (const step2 of plan.steps) {
    if (!isDone(step2)) {
      lines.push(`# SKIP ${step2.label}${step2.error ? ` \u2014 ${step2.error}` : ""}`);
      continue;
    }
    switch (step2.kind) {
      case "list.create": {
        const template = TEMPLATE_NAMES[step2.payload.baseTemplate] || step2.payload.baseTemplate;
        lines.push(`New-PnPList -Title ${psLit(step2.payload.title)} -Template ${template}${step2.payload.contentTypesEnabled ? " -EnableContentTypes" : ""}`);
        break;
      }
      case "list.adopt":
        lines.push(`# Using existing list "${step2.payload.title}"`);
        break;
      case "list.settings":
        pushCsomBlock(lines, plan, { ...step2.payload.groupA, ...step2.payload.groupB });
        break;
      case "ct.attach":
        lines.push(`Add-PnPContentTypeToList -List ${psLit(plan.title)} -ContentType ${psLit(step2.payload.contentTypeId)}`);
        break;
      case "field.create": {
        let xmlExpr = psLit(fieldXml(step2));
        const refs = step2.refs || {};
        if (refs.lookupListId) {
          const listIdentity = refs.lookupListId.self ? plan.title : refs.lookupListId.id || refs.lookupListId.list;
          xmlExpr = `(${xmlExpr}).Replace('${LIST_TOKEN}', ('{' + (Get-PnPList -Identity ${psLit(listIdentity)}).Id + '}'))`;
        }
        if (refs.primaryFieldId) {
          xmlExpr = `(${xmlExpr}).Replace('${PRIMARY_TOKEN}', ('{' + (Get-PnPField -List ${psLit(plan.title)} -Identity ${psLit(refs.primaryFieldId.field)}).Id + '}'))`;
        }
        lines.push(`Add-PnPFieldFromXml -List ${psLit(plan.title)} -FieldXml ${xmlExpr}`);
        break;
      }
      case "field.merge":
        lines.push(`Set-PnPField -List ${psLit(plan.title)} -Identity ${psLit(step2.payload.internalName)} -Values ${psHashtable(step2.payload.merges)}`);
        break;
      case "field.base": {
        const values = { Title: step2.payload.displayName };
        if (step2.payload.internalName === "Title") values.Required = step2.payload.required;
        if (step2.payload.description) values.Description = step2.payload.description;
        lines.push(`Set-PnPField -List ${psLit(plan.title)} -Identity ${psLit(step2.payload.internalName)} -Values ${psHashtable(values)}`);
        break;
      }
      case "view.upsert": {
        const titleLit = psLit(step2.payload.title);
        lines.push(`$v = Get-PnPView -List $list -Identity ${titleLit} -ErrorAction SilentlyContinue`);
        if (step2.payload.defaultView) {
          lines.push("if (-not $v) { $v = Get-PnPView -List $list | Where-Object DefaultView }");
        }
        lines.push("if ($v) {");
        lines.push(`  Set-PnPView -List $list -Identity $v -Fields ${psArray(step2.payload.fields)} -Values ${psHashtable({ ViewQuery: step2.payload.viewQuery, RowLimit: step2.payload.rowLimit, Paged: step2.payload.paged })}`);
        lines.push("} else {");
        lines.push(`  Add-PnPView -List $list -Title ${titleLit} -Fields ${psArray(step2.payload.fields)} -Query ${psLit(step2.payload.viewQuery)} -RowLimit ${step2.payload.rowLimit}${step2.payload.paged ? " -Paged" : ""} -SetAsDefault:$${step2.payload.defaultView ? "true" : "false"}`);
        lines.push("}");
        break;
      }
      case "list.validation":
        pushCsomBlock(lines, plan, {
          ValidationFormula: step2.payload.validationFormula,
          ValidationMessage: step2.payload.validationMessage
        });
        break;
      default:
        break;
    }
  }
  return lines.join("\n");
}
function toPnpjs2Provisioning(doc, opts = {}) {
  const plan = planFor(doc, opts);
  const lines = [
    "// PnPjs 2.x \u2014 paste into the DCSPad JS pane (pnpjs2 framework enabled)",
    `// ${plan.title}`,
    ""
  ];
  for (const step2 of plan.steps) {
    if (!isDone(step2)) {
      lines.push(`// SKIP ${step2.label}${step2.error ? ` \u2014 ${step2.error}` : ""}`);
      continue;
    }
    switch (step2.kind) {
      case "list.create":
        lines.push(`const newList = (await sp.web.lists.add(${JSON.stringify(step2.payload.title)}, ${JSON.stringify(step2.payload.description || "")}, ${step2.payload.baseTemplate}, ${step2.payload.contentTypesEnabled})).list;`);
        lines.push('const newListId = (await newList.select("Id")()).Id;');
        break;
      case "list.adopt":
        lines.push(`const newList = sp.web.lists.getById(${JSON.stringify(step2.payload.listId)});`);
        lines.push(`const newListId = ${JSON.stringify(step2.payload.listId)};`);
        break;
      case "list.settings": {
        const settings = { ...step2.payload.groupA, ...step2.payload.groupB };
        if (Object.keys(settings).length) lines.push(`await newList.update(${JSON.stringify(settings, null, 2)});`);
        break;
      }
      case "ct.attach":
        lines.push(`await newList.contentTypes.addAvailableContentType(${JSON.stringify(step2.payload.contentTypeId)});`);
        break;
      case "field.create": {
        const refs = step2.refs || {};
        let xmlExpr = JSON.stringify(fieldXml(step2));
        const pre = [];
        if (refs.lookupListId) {
          if (refs.lookupListId.self) {
            pre.push("  const lookupId = newListId;");
          } else if (refs.lookupListId.id) {
            pre.push(`  const lookupId = ${JSON.stringify(refs.lookupListId.id)};`);
          } else {
            pre.push(`  const lookupId = (await sp.web.lists.getByTitle(${JSON.stringify(refs.lookupListId.list)}).select("Id")()).Id;`);
          }
          xmlExpr += `.replace(${JSON.stringify(LIST_TOKEN)}, \`{\${lookupId}}\`)`;
        }
        if (refs.primaryFieldId) {
          pre.push(`  const primaryId = (await newList.fields.getByInternalNameOrTitle(${JSON.stringify(refs.primaryFieldId.field)}).select("Id")()).Id;`);
          xmlExpr += `.replace(${JSON.stringify(PRIMARY_TOKEN)}, \`{\${primaryId}}\`)`;
        }
        lines.push(
          "{",
          ...pre,
          `  await newList.fields.createFieldAsXml({ SchemaXml: ${xmlExpr}, Options: ${step2.payload.options ?? FIELD_OPTIONS} });`,
          "}"
        );
        break;
      }
      case "field.merge":
        lines.push(`await newList.fields.getByInternalNameOrTitle(${JSON.stringify(step2.payload.internalName)}).update(${JSON.stringify(step2.payload.merges)});`);
        break;
      case "field.base": {
        const values = { Title: step2.payload.displayName };
        if (step2.payload.internalName === "Title") values.Required = step2.payload.required;
        if (step2.payload.description) values.Description = step2.payload.description;
        lines.push(`await newList.fields.getByInternalNameOrTitle(${JSON.stringify(step2.payload.internalName)}).update(${JSON.stringify(values)});`);
        break;
      }
      case "view.upsert": {
        const titleJson = JSON.stringify(step2.payload.title);
        const settingsJson = JSON.stringify({
          ViewQuery: step2.payload.viewQuery,
          RowLimit: step2.payload.rowLimit,
          Paged: step2.payload.paged
        });
        lines.push("{");
        lines.push("  let existing = null;");
        lines.push(`  try { existing = await newList.views.getByTitle(${titleJson})(); } catch { existing = null; }`);
        if (step2.payload.defaultView) {
          lines.push("  if (!existing) existing = await newList.defaultView();");
        }
        lines.push("  let view;");
        lines.push("  if (existing) {");
        lines.push("    view = newList.views.getByTitle(existing.Title);");
        lines.push(`    await view.update(${settingsJson});`);
        lines.push("  } else {");
        lines.push(`    view = (await newList.views.add(${titleJson}, false, ${settingsJson})).view;`);
        lines.push("  }");
        lines.push("  await view.fields.removeAll();");
        for (const name of step2.payload.fields) lines.push(`  await view.fields.add(${JSON.stringify(name)});`);
        if (step2.payload.defaultView) {
          lines.push("  await view.update({ DefaultView: true });");
        }
        lines.push("}");
        break;
      }
      case "list.validation":
        lines.push(`await newList.update(${JSON.stringify({
          ValidationFormula: step2.payload.validationFormula,
          ValidationMessage: step2.payload.validationMessage
        })});`);
        break;
      default:
        break;
    }
  }
  return lines.join("\n");
}

// ../src/workbench/list-data-import-dialog.js
var DEFAULT_CLOSE_ANYWAY_MS = 15e3;
var closeAnywayMs = DEFAULT_CLOSE_ANYWAY_MS;
var CLOSE_ANYWAY_NOTE = "The current request is still in flight \u2014 closing leaves it running; the Items tab will refresh.";
var el6 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var guidPath4 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
var FIELD_SELECT = [
  "Id",
  "InternalName",
  "TypeAsString",
  "ReadOnlyField",
  "LookupList",
  "LookupField",
  "DisplayFormat"
];
var stem = (s) => String(s || "list").toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "list";
var plural2 = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
function dropReason(name, meta, targetFieldRow) {
  if (!targetFieldRow) return "missing on this list";
  if (targetFieldRow.ReadOnlyField) return "read-only on this list";
  if (NEVER_WRITE.has(name) || NEVER_WRITE_TYPES.has(targetFieldRow.TypeAsString)) return "system column \u2014 not imported";
  if (!(meta?.custom || name === "Title")) return "system column \u2014 not imported";
  return "not writable";
}
var AUTHORSHIP_FIELD_NAMES = ["Author", "Editor", "Created", "Modified"];
function buildFieldTable(dataDoc, targetFieldRows, writable, preserveAuthorship) {
  const table2 = el6("table", "wb-table wb-import-fields");
  const thead = el6("thead");
  const headRow = el6("tr");
  headRow.append(el6("th", "", "Column"), el6("th", "", "Status"));
  thead.append(headRow);
  const tbody = el6("tbody");
  const byName = new Map(targetFieldRows.map((f) => [f.InternalName, f]));
  const writtenNames = new Set(writable.map((w) => w.name));
  for (const [name, meta] of Object.entries(dataDoc.fields || {})) {
    if (AUTHORSHIP_FIELD_NAMES.includes(name)) continue;
    const tr = el6("tr");
    tr.append(el6("td", "", name));
    const status = writtenNames.has(name) ? "written" : dropReason(name, meta, byName.get(name));
    const td = el6("td");
    td.className = writtenNames.has(name) ? "" : "wb-import-dropped";
    td.textContent = status;
    tr.append(td);
    tbody.append(tr);
  }
  for (const name of AUTHORSHIP_FIELD_NAMES) {
    const tr = el6("tr", "wb-import-authorship-row");
    tr.append(el6("td", "", name));
    const td = el6("td");
    td.className = preserveAuthorship ? "" : "wb-import-dropped";
    td.textContent = preserveAuthorship ? "written \u2014 preserve authorship" : "not written (authorship off)";
    tr.append(td);
    tbody.append(tr);
  }
  table2.append(thead, tbody);
  return table2;
}
function itemFailureSummary(items) {
  const total = (items.failed?.length || 0) + (items.failedTruncated || 0);
  return items.failedTruncated ? `${total} failed (${items.failedTruncated} not itemised below)` : `${plural2(total, "failure")}`;
}
function buildHeadline2(report, { listTitle, isMock }) {
  if (report.aborted === "auth") return EXPIRED_SESSION_NOTE;
  if (report.aborted === "user") return "Import cancelled.";
  let headline = `Imported into \u2018${listTitle}\u2019 \u2014 ${plural2(report.items.added, "item")} added, ${itemFailureSummary(report.items)}.`;
  if (isMock) headline += " (mock mode \u2014 the fixture web does not change).";
  return headline;
}
function buildReportMarkdown(report, { listTitle }) {
  const lines = [`# Import report \u2014 ${listTitle}`, "", buildHeadline2(report, { listTitle, isMock: false }), ""];
  lines.push("## Counts", "");
  lines.push(`- Items: ${report.items.added} added, ${itemFailureSummary(report.items)}`);
  lines.push(`- Folders: ${report.folders.created} created, ${report.folders.failed} failed`);
  lines.push(`- Attachments: ${report.attachments.added} added, ${report.attachments.skipped} skipped, ${report.attachments.failed} failed`);
  lines.push(`- Authorship: ${report.authorship.applied} applied, ${report.authorship.failed} failed`);
  lines.push("");
  if (report.items.failed.length) {
    lines.push("## Failed items", "");
    for (const f of report.items.failed) lines.push(`- Source id ${f.sourceId}: ${f.error}`);
    lines.push("");
  }
  if (report.fieldErrors.length) {
    lines.push("## Field errors", "");
    for (const fe of report.fieldErrors) lines.push(`- Source id ${fe.sourceId} \u2014 ${fe.field}: ${fe.message}`);
    lines.push("");
  }
  if (report.warnings.length) {
    lines.push("## Warnings", "");
    for (const w of report.warnings) lines.push(`- ${w}`);
    lines.push("");
  }
  return lines.join("\n");
}
function openImportDataDialog({
  dataDoc,
  client: client2,
  listId,
  listTitle,
  mockWriter: mockWriter2,
  invalidateItems
} = {}) {
  return new Promise((resolve) => {
    const isMock = !client2.context().live;
    const spWrite = createSpWriteClient({ client: client2, mockWriter: isMock ? mockWriter2 : void 0 });
    const dialog = el6("dialog", "app-dialog sp-metadata-dialog wb-schema-dialog wb-import-data-dialog");
    const panel = el6("div", "app-dialog__panel");
    const head = el6("div", "app-dialog__head");
    head.append(el6("h2", "", "Import data"));
    const closeBtn = el6("button", "btn btn-ghost btn-xs wb-schema-close", "\u2715");
    closeBtn.type = "button";
    closeBtn.setAttribute("aria-label", "Close");
    head.append(closeBtn);
    const context = el6(
      "p",
      "app-dialog__context",
      `Add items from \u2018${dataDoc.source?.listTitle || "a list"}\u2019${dataDoc.source?.siteUrl ? ` (${dataDoc.source.siteUrl})` : ""} to \u2018${listTitle}\u2019.`
    );
    panel.append(head, context);
    const countsLine = el6(
      "p",
      "wb-schema-note",
      `${plural2(dataDoc.items.length, "item")}, ${plural2(dataDoc.folders.length, "folder")} in the source file.`
    );
    panel.append(countsLine);
    const fieldsSection = el6("div", "wb-schema-section");
    fieldsSection.append(el6("h3", "", "Columns"));
    const fieldsStatus = el6("p", "wb-schema-note", "Reading this list\u2019s columns\u2026");
    fieldsSection.append(fieldsStatus);
    panel.append(fieldsSection);
    const optsFieldset = el6("fieldset", "wb-schema-items");
    optsFieldset.append(el6("legend", "", "Options"));
    const optsRow = el6("div", "wb-schema-items-row");
    const attachCb = el6("input");
    attachCb.type = "checkbox";
    const attachLabel = el6("label", "wb-schema-items-opt");
    attachLabel.append(attachCb, el6("span", "", "Include attachments"));
    const authorshipCb = el6("input");
    authorshipCb.type = "checkbox";
    const authorshipLabel = el6("label", "wb-schema-items-opt");
    authorshipLabel.append(authorshipCb, el6("span", "", "Preserve authorship"));
    optsRow.append(attachLabel, authorshipLabel);
    optsFieldset.append(optsRow);
    panel.append(optsFieldset);
    const consentSection = el6("div", "sp-metadata-consent");
    const consentLabel = el6("label", "sp-metadata-consent__row");
    const consentBox = el6("input");
    consentBox.type = "checkbox";
    const consentText = el6(
      "span",
      "sp-metadata-consent__label",
      `I understand this adds ${dataDoc.items.length} item${dataDoc.items.length === 1 ? "" : "s"} to \u2018${listTitle}\u2019. Existing items are not changed or de-duplicated.`
    );
    consentLabel.append(consentBox, consentText);
    consentSection.append(consentLabel);
    panel.append(consentSection);
    const planPanel = el6("div", "wb-schema-plan");
    planPanel.hidden = true;
    const masterLine = el6("p", "wb-schema-master");
    const closeAnywayNote = el6("p", "wb-schema-note wb-import-closeanyway-note", CLOSE_ANYWAY_NOTE);
    closeAnywayNote.hidden = true;
    const closeAnywayBtn = el6("button", "btn btn-xs wb-import-closeanyway", "Close anyway");
    closeAnywayBtn.type = "button";
    closeAnywayBtn.hidden = true;
    planPanel.append(masterLine, closeAnywayNote, closeAnywayBtn);
    panel.append(planPanel);
    const reportPanel = el6("div", "wb-schema-report");
    reportPanel.hidden = true;
    const reportHeadline = el6("p", "wb-schema-report-headline");
    const reportCounts = el6("table", "wb-table wb-schema-report-counts");
    const reportFailed = el6("div", "wb-schema-report-failed");
    const reportWarnings = el6("div", "wb-schema-report-warnings");
    reportPanel.append(reportHeadline, reportCounts, reportFailed, reportWarnings);
    panel.append(reportPanel);
    const error = el6("div", "sp-files-error");
    error.setAttribute("role", "alert");
    error.hidden = true;
    panel.append(error);
    const actions = el6("div", "app-dialog__actions sp-metadata-actions wb-schema-actions-row");
    const cancelBtn = el6("button", "btn btn-ghost wb-schema-cancel", "Cancel");
    cancelBtn.type = "button";
    const downloadBtn = el6("button", "btn btn-xs wb-schema-report-download", "Download report .md");
    downloadBtn.type = "button";
    downloadBtn.hidden = true;
    const importBtn = el6("button", "btn btn-run wb-import-run", "Import");
    importBtn.type = "button";
    importBtn.disabled = true;
    actions.append(cancelBtn, downloadBtn, importBtn);
    panel.append(actions);
    dialog.append(panel);
    document.body.append(dialog);
    let phase = "form";
    let lastReport = null;
    let abortController = null;
    let cancelRequested = false;
    let closeAnywayTimer = null;
    let invalidated = false;
    let previewOk = false;
    let fieldsTableEl = null;
    let cachedTargetFieldRows = null;
    let cachedWritable = null;
    const finish = (outcome) => {
      dialog.close();
      dialog.remove();
      resolve(outcome);
    };
    const mutated = (report) => Boolean(report && ((report.items?.added || 0) > 0 || (report.folders?.created || 0) > 0));
    let detached = false;
    function invalidateOnce() {
      invalidated = true;
      invalidateItems?.();
    }
    function hideCloseAnyway() {
      closeAnywayNote.hidden = true;
      closeAnywayBtn.hidden = true;
      if (closeAnywayTimer) {
        clearTimeout(closeAnywayTimer);
        closeAnywayTimer = null;
      }
    }
    function requestCancel() {
      if (phase !== "running" || cancelRequested) return;
      cancelRequested = true;
      abortController?.abort();
      closeAnywayTimer = setTimeout(() => {
        closeAnywayTimer = null;
        if (phase === "running") {
          closeAnywayNote.hidden = false;
          closeAnywayBtn.hidden = false;
        }
      }, closeAnywayMs);
    }
    function setPhase(next) {
      phase = next;
      const inForm = next === "form";
      const inReport = next === "report";
      countsLine.hidden = !inForm;
      fieldsSection.hidden = !inForm;
      optsFieldset.hidden = !inForm;
      consentSection.hidden = !inForm;
      planPanel.hidden = next !== "running";
      reportPanel.hidden = !inReport;
      importBtn.hidden = !inForm;
      downloadBtn.hidden = !inReport;
      cancelBtn.textContent = inReport ? "Close" : "Cancel";
      closeBtn.hidden = next === "running";
    }
    function updateImportEnabled() {
      importBtn.disabled = !(consentBox.checked && previewOk);
    }
    consentBox.addEventListener("change", updateImportEnabled);
    function rebuildFieldsTable() {
      if (!cachedTargetFieldRows) return;
      fieldsTableEl?.remove();
      fieldsTableEl = buildFieldTable(dataDoc, cachedTargetFieldRows, cachedWritable, authorshipCb.checked);
      fieldsSection.append(fieldsTableEl);
    }
    authorshipCb.addEventListener("change", rebuildFieldsTable);
    client2.getAll(guidPath4(listId, "/fields"), { select: FIELD_SELECT }).then(({ items: targetFieldRows }) => {
      fieldsStatus.remove();
      cachedTargetFieldRows = targetFieldRows;
      cachedWritable = writableFields(dataDoc.fields, targetFieldRows);
      rebuildFieldsTable();
      previewOk = true;
      updateImportEnabled();
    }).catch((err) => {
      previewOk = false;
      showFailure(fieldsStatus, err, "this list\u2019s columns");
      updateImportEnabled();
    });
    function renderReportCounts(report) {
      reportCounts.textContent = "";
      const tbody = el6("tbody");
      const rows = [
        ["Items", `${report.items.added} added \xB7 ${itemFailureSummary(report.items)}`],
        ["Folders", `${report.folders.created} created \xB7 ${report.folders.failed} failed`]
      ];
      if (attachCb.checked) rows.push(["Attachments", `${report.attachments.added} added \xB7 ${report.attachments.skipped} skipped \xB7 ${report.attachments.failed} failed`]);
      if (authorshipCb.checked) rows.push(["Authorship", `${report.authorship.applied} applied \xB7 ${report.authorship.failed} failed`]);
      for (const [label, value] of rows) {
        const tr = el6("tr");
        tr.append(el6("td", "", label), el6("td", "", value));
        tbody.append(tr);
      }
      reportCounts.append(tbody);
    }
    function renderReportFailed(report) {
      reportFailed.textContent = "";
      if (report.items.failed.length) {
        reportFailed.append(el6("h3", "", "Failed items"));
        const table2 = el6("table", "wb-table");
        const tbody = el6("tbody");
        for (const f of report.items.failed) {
          const tr = el6("tr");
          tr.append(el6("td", "", `Source id ${f.sourceId}`), el6("td", "", f.error || ""));
          tbody.append(tr);
        }
        table2.append(tbody);
        reportFailed.append(table2);
        if (report.items.failedTruncated) {
          reportFailed.append(el6("p", "wb-schema-note", `+${report.items.failedTruncated} more not shown.`));
        }
      }
      if (report.fieldErrors.length) {
        reportFailed.append(el6("h3", "", "Field errors"));
        const table2 = el6("table", "wb-table");
        const tbody = el6("tbody");
        for (const fe of report.fieldErrors) {
          const tr = el6("tr");
          tr.append(el6("td", "", `Source id ${fe.sourceId} \u2014 ${fe.field}`), el6("td", "", fe.message || ""));
          tbody.append(tr);
        }
        table2.append(tbody);
        reportFailed.append(table2);
      }
    }
    function renderReportWarnings(report) {
      reportWarnings.textContent = "";
      if (!report.warnings?.length) return;
      reportWarnings.append(el6("h3", "", "Warnings"));
      const list2 = el6("ul", "wb-grid-notice");
      for (const w of report.warnings) list2.append(el6("li", "", w));
      reportWarnings.append(list2);
    }
    function renderReport(report) {
      reportHeadline.textContent = buildHeadline2(report, { listTitle, isMock });
      reportHeadline.classList.toggle("wb-schema-report-auth", report.aborted === "auth");
      renderReportCounts(report);
      renderReportFailed(report);
      renderReportWarnings(report);
      downloadBtn.hidden = false;
    }
    importBtn.addEventListener("click", async () => {
      if (importBtn.disabled) return;
      error.hidden = true;
      importBtn.disabled = true;
      cancelRequested = false;
      hideCloseAnyway();
      setPhase("running");
      masterLine.textContent = "Importing\u2026";
      abortController = new AbortController();
      try {
        const report = await applyListData({
          dataDoc,
          client: client2,
          spWrite,
          listId,
          options: { includeAttachments: attachCb.checked, preserveAuthorship: authorshipCb.checked },
          onStep: (info) => {
            masterLine.textContent = info.label;
          },
          signal: abortController.signal
        });
        hideCloseAnyway();
        lastReport = report;
        setPhase("report");
        renderReport(report);
        if (mutated(report)) invalidateOnce();
      } catch (err) {
        if (detached) {
          invalidateOnce();
          return;
        }
        hideCloseAnyway();
        setPhase("form");
        importBtn.disabled = false;
        error.textContent = err?.message || String(err);
        error.hidden = false;
      }
    });
    downloadBtn.addEventListener("click", () => {
      if (!lastReport) return;
      downloadText(`import-report-${stem(listTitle)}.md`, buildReportMarkdown(lastReport, { listTitle }), "text/markdown");
    });
    closeAnywayBtn.addEventListener("click", () => {
      detached = true;
      invalidateOnce();
      finish("closed-during-run");
    });
    cancelBtn.addEventListener("click", () => {
      if (phase === "running") {
        requestCancel();
        return;
      }
      if (phase === "report") {
        finish(lastReport && !lastReport.aborted ? "imported" : "cancelled");
        return;
      }
      finish("cancelled");
    });
    closeBtn.addEventListener("click", () => {
      if (phase === "running") return;
      finish(phase === "report" && lastReport && !lastReport.aborted ? "imported" : "cancelled");
    });
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      if (phase === "running") {
        requestCancel();
        return;
      }
      finish(phase === "report" && lastReport && !lastReport.aborted ? "imported" : "cancelled");
    });
    setPhase("form");
    dialog.showModal();
  });
}

// ../src/workbench/list-tools.js
var el7 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var fileStem = (s) => String(s || "list").toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "list";
function copyEligible(doc) {
  const bt = schemaBaseType(doc);
  return bt === 0 && Number(doc.list.baseTemplate) === 100 || bt === 1;
}
function cardShell(id, title, description) {
  const node = el7("div", `wb-linkgroup wb-tool-card wb-tool-${id}`);
  node.append(el7("h3", "", title));
  node.append(el7("p", "wb-tool-desc", description));
  const body = el7("div", "wb-tool-body");
  node.append(body);
  return { node, body };
}
function disabledLine(reason) {
  const line = el7("p", "wb-tool-reason", reason);
  return line;
}
var copyTool = {
  id: "copy",
  title: "Copy this list\u2026",
  description: "Create a new list or library from this one\u2019s schema, on this site or another.",
  render(card, ctx2) {
    const { body } = card;
    const status = el7("p", "wb-tool-status", "Reading the list schema\u2026");
    body.append(status);
    ctx2.getSchema().then(({ doc }) => {
      status.remove();
      const summary = schemaSummary(doc);
      if (copyEligible(doc)) {
        const btn = el7("button", "btn btn-xs wb-tools-copy", "Copy to\u2026");
        btn.type = "button";
        btn.title = "Create a new list from this schema, on this site or another one.";
        btn.addEventListener("click", () => ctx2.openCopy(doc));
        body.append(btn);
      } else {
        const btn = el7("button", "btn btn-xs wb-tools-copy", "Copy to\u2026");
        btn.type = "button";
        btn.disabled = true;
        body.append(btn);
        body.append(disabledLine(
          `Only generic lists and document libraries can be copied \u2014 this is a ${summary.kind.toLowerCase()}.`
        ));
      }
    }).catch((err) => showFailure(status, err, "this list\u2019s schema"));
  }
};
var exportSchemaTool = {
  id: "export-schema",
  title: "Export schema",
  description: "Download or copy this list\u2019s schema, or generate a provisioning script.",
  render(card, ctx2) {
    const { body } = card;
    const status = el7("p", "wb-tool-status", "Reading the list schema\u2026");
    body.append(status);
    ctx2.getSchema().then(({ doc }) => {
      status.remove();
      const stem2 = fileStem(ctx2.listTitle);
      body.append(createMenuButton2("Export \u25BE", "Export this list\u2019s schema", [
        ["Download schema .json", () => downloadText(`schema-${stem2}.json`, JSON.stringify(doc, null, 2), "application/json")],
        ["Copy schema JSON", (btn) => copyText(JSON.stringify(doc, null, 2), btn)],
        ["Copy as PnP.PowerShell (provision)", (btn) => copyText(toPnpPowerShellProvisioning(doc, { targetWebUrl: ctx2.client.webUrl() }), btn)],
        ["Copy as PnPjs 2 (provision)", (btn) => copyText(toPnpjs2Provisioning(doc, {}), btn)]
      ]));
    }).catch((err) => showFailure(status, err, "this list\u2019s schema"));
  }
};
var exportDataTool = {
  id: "export-data",
  title: "Export data",
  description: "Read every item in this list (not just what\u2019s on screen elsewhere) and export it as JSON.",
  render(card, ctx2) {
    const { body } = card;
    const stem2 = fileStem(ctx2.listTitle);
    const status = el7("div", "wb-grid-status wb-tool-status");
    status.hidden = true;
    let busy = false;
    async function run(useDoc) {
      if (busy) return;
      busy = true;
      status.hidden = false;
      status.classList.remove("wb-error", "wb-denied");
      status.removeAttribute("title");
      let n = null;
      try {
        const { doc: schemaDoc } = await ctx2.getSchema();
        n = schemaDoc?.source?.itemCount;
        status.textContent = n != null ? `Reading ${n} item${n === 1 ? "" : "s"}\u2026` : "Reading items\u2026";
        const { doc } = await captureListData(ctx2.client, ctx2.listId, { schemaDoc });
        status.hidden = true;
        useDoc(doc);
      } catch (err) {
        showFailure(status, err, "this list\u2019s item data");
      } finally {
        busy = false;
      }
    }
    const downloadBtn = el7("button", "btn btn-xs", "Download data .json");
    downloadBtn.type = "button";
    downloadBtn.addEventListener("click", () => run((doc) => downloadText(`data-${stem2}.json`, JSON.stringify(doc, null, 2), "application/json")));
    const copyBtn = el7("button", "btn btn-xs", "Copy data JSON");
    copyBtn.type = "button";
    copyBtn.addEventListener("click", () => run((doc) => copyText(JSON.stringify(doc, null, 2), copyBtn)));
    body.append(downloadBtn, copyBtn, status);
  }
};
function showInlineNotice(host, message) {
  host.textContent = "";
  host.hidden = false;
  host.classList.add("wb-consent-error");
  host.append(el7("span", "wb-consent-text", message));
  const dismiss = el7("button", "btn btn-xs", "Dismiss");
  dismiss.type = "button";
  dismiss.addEventListener("click", () => {
    host.hidden = true;
  });
  host.append(dismiss);
}
var importDataTool = {
  id: "import-data",
  title: "Import data into this list",
  description: "Add items from a data .json exported by this Workbench, or by SPUtils.",
  render(card, ctx2) {
    const { body } = card;
    const status = el7("p", "wb-tool-status", "Reading the list schema\u2026");
    body.append(status);
    ctx2.getSchema().then(({ doc: schemaDoc }) => {
      status.remove();
      const isLibrary = schemaBaseType(schemaDoc) === 1;
      const btn = el7("button", "btn btn-xs wb-tools-import", "Choose data .json\u2026");
      btn.type = "button";
      const input = el7("input", "wb-tools-import-file");
      input.type = "file";
      input.accept = ".json";
      input.hidden = true;
      const notice = el7("div", "wb-consent wb-tool-import-notice");
      notice.hidden = true;
      if (isLibrary) {
        btn.disabled = true;
        body.append(btn, disabledLine(
          "Item import into a document library isn\u2019t supported \u2014 a library copy is schema only."
        ));
        return;
      }
      btn.addEventListener("click", () => input.click());
      input.addEventListener("change", async () => {
        const file = input.files?.[0];
        input.value = "";
        if (!file) return;
        notice.hidden = true;
        if (file.size > MAX_IMPORT_BYTES) {
          showInlineNotice(notice, `\u2018${file.name}\u2019 is ${(file.size / 1048576).toFixed(1)} MB \u2014 above the 5 MB import limit.`);
          return;
        }
        let json;
        try {
          json = JSON.parse(await file.text());
        } catch {
          showInlineNotice(notice, `\u2018${file.name}\u2019 isn\u2019t valid JSON.`);
          return;
        }
        let dataDoc;
        try {
          dataDoc = normalizeDataDoc(json);
        } catch {
          showInlineNotice(notice, `\u2018${file.name}\u2019 is not a list data document (expected kind ${DATA_KIND}).`);
          return;
        }
        const problem = validateDataDoc(json);
        if (problem) {
          showInlineNotice(notice, `\u2018${file.name}\u2019 is not a usable list data document (${problem}).`);
          return;
        }
        await openImportDataDialog({
          dataDoc,
          client: ctx2.client,
          listId: ctx2.listId,
          listTitle: ctx2.listTitle,
          mockWriter: ctx2.mockWriter,
          invalidateItems: ctx2.invalidateItems
        });
      });
      body.append(btn, input, notice);
    }).catch((err) => showFailure(status, err, "this list\u2019s schema"));
  }
};
var LIST_TOOLS = [copyTool, exportSchemaTool, exportDataTool, importDataTool];
function buildToolsPane(wrap, ctx2) {
  const grid = el7("div", "wb-tools");
  wrap.append(grid);
  for (const tool of LIST_TOOLS) {
    const shell2 = cardShell(tool.id, tool.title, tool.description);
    grid.append(shell2.node);
    tool.render(shell2, ctx2);
  }
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
  "IsCatalog",
  "Created",
  "LastItemModifiedDate",
  "EntityTypeName",
  "Description",
  "DefaultViewUrl",
  "RootFolder/ServerRelativeUrl"
];
var INTERNAL_TEMPLATES = /* @__PURE__ */ new Set([112, 113, 114, 116, 121, 122, 123, 124, 125, 175, 544]);
var INTERNAL_TITLES = /* @__PURE__ */ new Set([
  "TaxonomyHiddenList",
  "Style Library",
  "Form Templates",
  "Cache Profiles",
  "Device Channels",
  "Quick Deploy Items",
  "Reusable Content",
  "Content and Structure Reports",
  "Site Collection Documents",
  "Site Collection Images",
  "Suggested Content Browser Locations"
]);
function isInternalList(list2) {
  const path = String(list2?.RootFolder?.ServerRelativeUrl || "").toLowerCase();
  return Boolean(list2?.Hidden) || Boolean(list2?.IsCatalog) || INTERNAL_TEMPLATES.has(list2?.BaseTemplate) || path.includes("/_catalogs") || path.endsWith("/formservertemplates") || INTERNAL_TITLES.has(list2?.Title);
}
var FIELD_SELECT2 = [
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
var fileStem2 = (s) => String(s || "list").toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "") || "list";
var el8 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var guidPath5 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
function parseItemsQuery(text) {
  const out = { filter: "", orderby: "", error: "" };
  const raw = String(text || "").trim().replace(/^\?/, "");
  if (!raw) return out;
  const parts = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] === "'") {
      if (quoted && raw[i + 1] === "'") {
        i++;
        continue;
      }
      quoted = !quoted;
    } else if (raw[i] === "&" && !quoted) {
      parts.push(raw.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(raw.slice(start));
  const seen = /* @__PURE__ */ new Set();
  for (const part of parts) {
    const piece = part.trim();
    if (!piece) continue;
    const clause = /^\$?(filter|orderby)\s*=\s*(.+)$/i.exec(piece);
    if (clause) {
      const name = clause[1].toLowerCase();
      if (seen.has(name)) {
        out.error = `Duplicate $${name} clause \u2014 combine them into one.`;
        return out;
      }
      seen.add(name);
      out[name] = clause[2].trim();
    } else if (piece.startsWith("$")) {
      out.error = "Only $filter and $orderby are supported here \u2014 columns and max items have their own controls.";
      return out;
    } else {
      out.filter = out.filter ? `${out.filter} and ${piece}` : piece;
    }
  }
  return out;
}
function createListsView({
  client: client2,
  navigate,
  updateRoute,
  inspectSite: inspectSite2,
  createClient,
  mockWriter: mockWriter2
}) {
  const root = el8("section", "wb-view wb-view-lists");
  const webOrigin = () => {
    try {
      return new URL(client2.webUrl()).origin;
    } catch {
      return "";
    }
  };
  const absUrl = (rel) => rel ? `${webOrigin()}${encodeSpPath(rel)}` : "";
  const gridPane = el8("div", "wb-pane");
  const head = el8("div", "wb-view-head");
  head.innerHTML = '<h2>Lists &amp; libraries</h2><p class="wb-view-hint">Every list in this web \u2014 SharePoint-internal plumbing sits behind the expander below the grid. Click a row for fields, views, content types, and items.</p>';
  const grid = createGrid({
    columns: [
      { key: "Title", label: "Title" },
      { key: "BaseTemplate", label: "Template", format: (v) => BASE_TEMPLATE_NAMES[v] || String(v ?? "") },
      { key: "ItemCount", label: "Items" },
      { key: "Hidden", label: "Hidden" },
      { key: "Url", label: "Url", value: (row) => row.RootFolder?.ServerRelativeUrl || "", mono: true, copyable: true, link: absUrl },
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
          a.title = "Open list settings in a new tab";
          a.textContent = "\u2699";
          bindNewTab(a);
          return a;
        }
      }
    ],
    onOpen: (row) => navigate({ view: "lists", listId: row.Id, listTitle: row.Title }),
    emptyText: "No lists in this web.",
    subject: "the lists in this web",
    filterPlaceholder: "Filter lists\u2026",
    exportName: "sp-lists",
    descriptor: {
      path: "web/lists",
      options: { select: LIST_SELECT, expand: "RootFolder", orderby: "Title", top: 5e3 },
      webUrl: client2.webUrl()
    }
  });
  const moreBtn = el8("button", "btn btn-xs wb-lists-more");
  moreBtn.type = "button";
  moreBtn.hidden = true;
  const importNotice = el8("div", "wb-consent wb-schema-import-notice");
  importNotice.hidden = true;
  gridPane.append(head, importNotice, grid.el, moreBtn);
  const newFromSchemaBtn = el8("button", "btn btn-xs wb-schema-new", "New from schema\u2026");
  newFromSchemaBtn.type = "button";
  const schemaFileInput = el8("input", "wb-schema-file");
  schemaFileInput.type = "file";
  schemaFileInput.accept = ".json";
  schemaFileInput.multiple = true;
  schemaFileInput.hidden = true;
  schemaFileInput.setAttribute("aria-label", "Choose a list schema JSON file, and optionally a matching item-data JSON file");
  newFromSchemaBtn.addEventListener("click", () => schemaFileInput.click());
  function showImportNotice(message) {
    importNotice.textContent = "";
    importNotice.hidden = false;
    importNotice.classList.add("wb-consent-error");
    importNotice.append(el8("span", "wb-consent-text", message));
    const dismiss = el8("button", "btn btn-xs", "Dismiss");
    dismiss.type = "button";
    dismiss.addEventListener("click", () => {
      importNotice.hidden = true;
    });
    importNotice.append(dismiss);
  }
  schemaFileInput.addEventListener("change", async () => {
    const files = [...schemaFileInput.files || []];
    schemaFileInput.value = "";
    if (!files.length) return;
    importNotice.hidden = true;
    let doc = null;
    let dataDoc = null;
    for (const file of files) {
      if (file.size > MAX_IMPORT_BYTES) {
        showImportNotice(`\u2018${file.name}\u2019 is ${(file.size / 1048576).toFixed(1)} MB \u2014 above the 5 MB import limit.`);
        return;
      }
      let json;
      try {
        json = JSON.parse(await file.text());
      } catch {
        showImportNotice(`\u2018${file.name}\u2019 isn\u2019t valid JSON.`);
        return;
      }
      if (json?.kind === DATA_KIND) {
        try {
          dataDoc = normalizeDataDoc(json);
        } catch {
          showImportNotice(`\u2018${file.name}\u2019 is not a list data document (expected kind ${DATA_KIND}).`);
          return;
        }
      } else {
        try {
          doc = normalizeSchemaDoc(json);
        } catch {
          showImportNotice(`\u2018${file.name}\u2019 is not a list schema document (expected kind ${SCHEMA_KIND}).`);
          return;
        }
      }
    }
    if (!doc) {
      showImportNotice(`Choose a list schema document (expected kind ${SCHEMA_KIND}).`);
      return;
    }
    if (dataDoc && String(dataDoc.source?.listId || "").toLowerCase() !== String(doc.source?.listId || "").toLowerCase()) {
      showImportNotice(
        `The data file\u2019s source list (\u2018${dataDoc.source?.listTitle || "unknown"}\u2019) doesn\u2019t match the schema\u2019s (\u2018${doc.source?.listTitle || "unknown"}\u2019) \u2014 choose a matching pair.`
      );
      return;
    }
    const outcome = await openSchemaApplyDialog({
      doc,
      dataDoc,
      mode: "import",
      client: client2,
      createClient,
      navigate,
      inspectSite: inspectSite2,
      mockWriter: mockWriter2
    });
    if (outcome === "created") {
      listsLoaded = false;
      loadLists();
    }
  });
  grid.actionsEl.prepend(newFromSchemaBtn, schemaFileInput);
  const detailPane = el8("div", "wb-pane");
  detailPane.hidden = true;
  root.append(gridPane, detailPane);
  let listsLoaded = false;
  let allLists = [];
  let listsPartial = false;
  let showInternal = false;
  const tabCache = /* @__PURE__ */ new Map();
  function renderLists() {
    const internal = allLists.filter(isInternalList);
    grid.setRows(
      showInternal ? allLists : allLists.filter((l) => !isInternalList(l)),
      { partial: listsPartial }
    );
    moreBtn.hidden = internal.length === 0;
    const label = `${internal.length} internal list${internal.length === 1 ? "" : "s"}`;
    moreBtn.textContent = showInternal ? `Hide ${label} \u25B4` : `Show ${label} \u25BE`;
    moreBtn.title = "SharePoint-internal plumbing: hidden lists, galleries, the taxonomy cache";
  }
  moreBtn.addEventListener("click", () => {
    showInternal = !showInternal;
    renderLists();
  });
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
      allLists = items;
      listsPartial = partial;
      renderLists();
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
  function buildItemsPane(wrap, listId, listTitle) {
    const controls = el8("span", "wb-items-controls");
    const viewLabel = el8("label", "wb-items-label", "Columns ");
    const viewSel = el8("select", "wb-items-view");
    viewSel.setAttribute("aria-label", "Column source: a view or all columns");
    viewLabel.append(viewSel);
    const maxLabel = el8("label", "wb-items-label", "Max ");
    const maxIn = el8("input", "wb-items-max");
    maxIn.type = "number";
    maxIn.min = "1";
    maxIn.max = "5000";
    maxIn.value = "500";
    maxIn.setAttribute("aria-label", "Maximum items to fetch");
    maxLabel.append(maxIn);
    const queryToggle = el8("button", "wb-items-querytoggle", "Query \u25BE");
    queryToggle.type = "button";
    const QUERY_HINT = "Advanced: raw OData clauses for the item query \u2014 $filter=\u2026 and/or $orderby=\u2026, joined with &. A bare expression counts as $filter. Order defaults to ID desc.";
    queryToggle.title = QUERY_HINT;
    queryToggle.setAttribute("aria-expanded", "false");
    const queryWrap = el8("span", "wb-items-querywrap");
    queryWrap.hidden = true;
    const queryIn = el8("input", "wb-items-query");
    queryIn.type = "text";
    queryIn.placeholder = "$filter=Status eq 'Active'&$orderby=DueDate desc";
    queryIn.setAttribute("aria-label", "OData $filter and $orderby for the item query");
    queryIn.title = QUERY_HINT;
    const queryErr = el8("span", "wb-items-error");
    queryWrap.append(queryIn, queryErr);
    controls.append(viewLabel, maxLabel, queryToggle, queryWrap);
    const gridBox = el8("div", "wb-items-grid");
    wrap.append(gridBox);
    function setQueryOpen(open) {
      queryWrap.hidden = !open;
      queryToggle.textContent = open ? "Query \u25B4" : "Query \u25BE";
      queryToggle.setAttribute("aria-expanded", open ? "true" : "false");
      if (open) queryIn.focus();
    }
    queryToggle.addEventListener("click", () => {
      const opening = queryWrap.hidden;
      if (!opening && queryErr.textContent) {
        queryIn.value = "";
        queryErr.textContent = "";
        reload();
      }
      setQueryOpen(opening);
    });
    let itemsGrid = null;
    let current = null;
    let viewsFilled = false;
    let loadSeq = 0;
    let itemsCache = { key: "", items: null, partial: false };
    const clampMax = () => {
      const n = Math.floor(Number(maxIn.value));
      const max = Number.isFinite(n) ? Math.min(Math.max(n, 1), 5e3) : 500;
      maxIn.value = String(max);
      return max;
    };
    const exportDoc = () => itemsGrid && current ? buildItemsMarkdown({
      listTitle,
      webUrl: client2.webUrl(),
      viewTitle: current.viewTitle,
      items: itemsGrid.getExportRows(),
      // selected rows when any, else visible
      fields: current.fields,
      viewFieldNames: current.viewFieldNames,
      filter: current.filter,
      orderby: current.orderby
    }) : "";
    function markApplied() {
      const applied = Boolean(current && (current.filter || current.orderby));
      queryToggle.classList.toggle("wb-applied", applied);
      queryToggle.title = applied ? `Active query: ${queryIn.value.trim()}` : QUERY_HINT;
    }
    async function reload() {
      const seq = ++loadSeq;
      const parsed = parseItemsQuery(queryIn.value);
      queryErr.textContent = parsed.error;
      if (parsed.error) return;
      const max = clampMax();
      let status = null;
      if (itemsGrid) {
        itemsGrid.setLoading("Loading items\u2026");
      } else {
        gridBox.textContent = "";
        status = el8("div", "wb-grid-status", "Loading items\u2026");
        gridBox.append(status);
      }
      let fields = null;
      let viewFieldNames = null;
      let query = null;
      try {
        const [fieldsResult, { items: views }] = await Promise.all([
          cached2(listId, "fields", () => client2.getAll(guidPath5(listId, "/fields"), { select: FIELD_SELECT2 })),
          cached2(listId, "views", () => client2.getAll(guidPath5(listId, "/views"), { select: VIEW_SELECT }))
        ]);
        fields = fieldsResult.items;
        if (!viewsFilled) {
          viewsFilled = true;
          const none = el8("option", "", "All columns");
          none.value = "";
          viewSel.append(none);
          for (const view of views.filter((v) => !v.PersonalView)) {
            const opt = el8("option", "", view.DefaultView ? `${view.Title} (default)` : view.Title);
            opt.value = view.Id;
            viewSel.append(opt);
          }
        }
        let viewTitle = "";
        if (viewSel.value) {
          const vf = await cached2(listId, `viewfields::${viewSel.value}`, () => client2.get(guidPath5(listId, `/views(guid'${viewSel.value}')/viewfields`)));
          viewFieldNames = vf?.Items?.results || vf?.Items || [];
          viewTitle = views.find((v) => v.Id === viewSel.value)?.Title || "";
        }
        const hasAttachments = fields.some((f) => f.TypeAsString === "Attachments");
        const expand = ["FieldValuesAsText", ...hasAttachments ? ["AttachmentFiles"] : []];
        const query2 = {
          path: guidPath5(listId, "/items"),
          options: {
            select: ["*", ...expand],
            expand,
            ...parsed.filter ? { filter: parsed.filter } : {},
            orderby: parsed.orderby || "ID desc",
            top: max
          }
        };
        const dataKey = JSON.stringify([max, parsed.filter, parsed.orderby]);
        let items;
        let partial;
        if (itemsCache.items && itemsCache.key === dataKey) {
          ({ items, partial } = itemsCache);
        } else {
          ({ items, partial } = await client2.getAll(query2.path, query2.options, { cap: max }));
          if (seq !== loadSeq) return;
          if (!parsed.orderby) {
            items.sort((a, b) => (Number(b.ID ?? b.Id) || 0) - (Number(a.ID ?? a.Id) || 0));
          }
          itemsCache = { key: dataKey, items, partial };
        }
        if (seq !== loadSeq) return;
        mountItemsGrid(buildColumns(fields, viewFieldNames, items), query2);
        itemsGrid.setRows(items, { partial });
        current = {
          fields,
          viewFieldNames,
          viewTitle,
          filter: parsed.filter,
          orderby: parsed.orderby
        };
        markApplied();
      } catch (err) {
        if (seq !== loadSeq) return;
        const raw = err?.message || String(err);
        const message = /SPQueryThrottledException|list view threshold/i.test(raw) ? `SharePoint throttled this query \u2014 filter/order by an indexed column and keep the matched set under 5,000. (${raw})` : raw;
        if (!itemsGrid && fields) {
          mountItemsGrid(buildColumns(fields, viewFieldNames, []), query);
        }
        if (itemsGrid) {
          itemsGrid.setError({ message });
        } else if (status) {
          status.textContent = message;
          status.classList.add("wb-error");
        }
      }
    }
    function buildColumns(fields, viewFieldNames, items) {
      const content = viewFieldNames ? viewColumnFields(fields, viewFieldNames) : contentFields(fields);
      const filesOf = (row) => Array.isArray(row.AttachmentFiles) ? row.AttachmentFiles : row.AttachmentFiles?.results || [];
      const anyAttachments = items.some((row) => filesOf(row).length);
      return [
        { key: "Title", label: "Title", value: (row) => itemTitle(row) },
        { key: "ID", label: "ID", num: true, value: (row) => row.ID ?? row.Id },
        ...content.map((f) => ({
          key: f.InternalName,
          label: f.Title || f.InternalName,
          value: (row) => fieldText(row, f)
        })),
        ...anyAttachments ? [{
          key: "Attachments",
          label: "Attachments",
          value: (row) => filesOf(row).map((f) => f?.FileName || "").filter(Boolean).join(", ")
        }] : [],
        { key: "Created", label: "Created", format: fmtDate },
        { key: "CreatedBy", label: "Created By", value: (row) => personText(row, "Author") },
        { key: "Modified", label: "Modified", format: fmtDate },
        { key: "ModifiedBy", label: "Modified By", value: (row) => personText(row, "Editor") }
      ];
    }
    function mountItemsGrid(columns, query) {
      const newGrid = createGrid({
        columns,
        rowKey: "ID",
        emptyText: "No items in this list.",
        subject: "this list\u2019s items",
        filterPlaceholder: "Filter items\u2026",
        exportName: `items-${fileStem2(listTitle)}`,
        descriptor: query ? { ...query, webUrl: client2.webUrl() } : null,
        selectable: true,
        toolbarExtras: controls,
        exportExtras: [
          ["Download .md", () => {
            const md = exportDoc();
            if (md) downloadMarkdown(`items-${fileStem2(listTitle)}`, md);
          }],
          ["Copy .md", (btn) => {
            const md = exportDoc();
            if (md) copyText(md, btn);
          }]
        ]
      });
      const focused = controls.contains(document.activeElement) ? document.activeElement : null;
      gridBox.textContent = "";
      gridBox.append(newGrid.el);
      itemsGrid = newGrid;
      if (focused) focused.focus();
    }
    viewSel.addEventListener("change", reload);
    maxIn.addEventListener("change", reload);
    queryIn.addEventListener("change", reload);
    queryIn.addEventListener("keydown", (e) => {
      if (e.key === "Enter") reload();
    });
    reload();
  }
  function fieldNote(f, fields) {
    if (!f.custom) return "system column \u2014 not copied";
    if (TAXONOMY_TYPES.has(f.type)) return "managed metadata \u2014 not recreated";
    if (f.isDependentLookup) {
      const primary = fields.find((p) => p.id === f.primaryFieldId);
      return `dependent lookup of ${primary?.displayName || primary?.internalName || "a primary column"}`;
    }
    if (f.isSelfLookup) return "lookup to this list \u2014 rebinds to the copy";
    return "";
  }
  function schemaChip(cls, text, title) {
    const chip = el8("span", `wb-info-chip ${cls}`, text);
    if (title) chip.title = title;
    return chip;
  }
  function copyableCell(text) {
    const span = el8("span", "sp-copy", text);
    span.title = "Click to copy";
    span.addEventListener("click", () => copyText(text, span));
    return span;
  }
  function buildSettingsSection(doc) {
    const section = el8("div", "wb-schema-section");
    section.append(el8("h3", "", "Settings"));
    const table2 = el8("table", "wb-table wb-schema-settings");
    const tbody = el8("tbody");
    const rows = [
      ["Description", doc.list.description || "\u2014"],
      ["Base template", `${BASE_TEMPLATE_NAMES[doc.list.baseTemplate] || "Template"} (${doc.list.baseTemplate})`],
      ["Versioning", doc.list.enableVersioning ? `on (major limit ${doc.list.majorVersionLimit ?? "\u2014"})` : "off"],
      ["Content types", doc.list.contentTypesEnabled ? "on" : "off"],
      ["Attachments", doc.list.enableAttachments ? "on" : "off"],
      ["Folder creation", doc.list.enableFolderCreation ? "on" : "off"],
      ["Moderation", doc.list.enableModeration ? "on" : "off"],
      ["Force checkout", doc.list.forceCheckout ? "on" : "off"],
      ["Validation formula", doc.list.validationFormula || "\u2014"],
      ["On Quick Launch", doc.list.onQuickLaunch ? "yes" : "no"]
    ];
    for (const [label, value] of rows) {
      const tr = el8("tr");
      tr.append(el8("td", "", label));
      const td = el8("td");
      td.append(copyableCell(String(value)));
      tr.append(td);
      tbody.append(tr);
    }
    table2.append(tbody);
    section.append(table2);
    return section;
  }
  function buildFieldsSection(doc) {
    const section = el8("div", "wb-schema-section");
    section.append(el8("h3", "", "Fields"));
    const fieldsGrid = createGrid({
      columns: [
        { key: "displayName", label: "Title" },
        { key: "internalName", label: "Internal name", mono: true, copyable: true },
        { key: "type", label: "Type" },
        { key: "custom", label: "Custom" },
        { key: "required", label: "Required" },
        { key: "indexed", label: "Indexed" },
        { key: "lookupList", label: "Lookup target", value: (f) => f.lookupList || "" },
        { key: "note", label: "Note", value: (f) => fieldNote(f, doc.fields) }
      ],
      rowKey: "internalName",
      emptyText: "No fields.",
      filterPlaceholder: "Filter fields\u2026"
    });
    fieldsGrid.setRows(doc.fields);
    section.append(fieldsGrid.el);
    return section;
  }
  function buildViewsSection(doc) {
    const section = el8("div", "wb-schema-section");
    section.append(el8("h3", "", "Views"));
    const viewsGrid = createGrid({
      columns: [
        { key: "title", label: "Title" },
        { key: "defaultView", label: "Default" },
        { key: "columns", label: "Columns", value: (v) => (v.fields || []).length, num: true },
        { key: "rowLimit", label: "Row limit", num: true },
        { key: "paged", label: "Paged" },
        { key: "customFormatter", label: "Formatting", value: (v) => v.customFormatter ? "yes" : "" },
        { key: "viewQuery", label: "CAML query", mono: true, copyable: true }
      ],
      rowKey: "title",
      emptyText: "No views.",
      filterPlaceholder: "Filter views\u2026"
    });
    viewsGrid.setRows(doc.views);
    section.append(viewsGrid.el);
    return section;
  }
  function buildContentTypesSection(doc) {
    const section = el8("div", "wb-schema-section");
    section.append(el8("h3", "", "Content types"));
    if (!doc.list.contentTypesEnabled) {
      section.append(el8("p", "wb-schema-note", "Content types are exported for reference \u2014 they are not recreated unless enabled on the target."));
    }
    const ctGrid = createGrid({
      columns: [
        { key: "name", label: "Name" },
        { key: "id", label: "Id", mono: true, copyable: true },
        { key: "group", label: "Group" },
        { key: "sealed", label: "Sealed" }
      ],
      rowKey: "id",
      emptyText: "No content types.",
      filterPlaceholder: "Filter content types\u2026"
    });
    ctGrid.setRows(doc.contentTypes);
    section.append(ctGrid.el);
    return section;
  }
  function buildWarningsSection(doc) {
    const section = el8("div", "wb-schema-section");
    section.append(el8("h3", "", "Warnings"));
    const list2 = el8("ul", "wb-grid-notice");
    for (const w of doc.warnings) list2.append(el8("li", "", w));
    section.append(list2);
    return section;
  }
  async function openCopy(doc) {
    await openSchemaApplyDialog({
      doc,
      mode: "copy",
      client: client2,
      createClient,
      navigate,
      inspectSite: inspectSite2,
      mockWriter: mockWriter2
    });
  }
  function renderSchemaPane(doc) {
    const summary = schemaSummary(doc);
    const head2 = el8("div", "wb-schema-head");
    const chips = el8("div", "wb-schema-chips");
    chips.append(
      schemaChip("wb-schema-kind", summary.kind, `BaseTemplate ${doc.list.baseTemplate}`),
      schemaChip("wb-schema-fields", summary.fieldsText, `${doc.fields.length} total fields, ${doc.fields.filter((f) => f.custom).length} custom`),
      schemaChip("wb-schema-views", summary.viewsText),
      schemaChip("wb-schema-cts", summary.contentTypesText, doc.list.contentTypesEnabled ? "Content types are enabled on this list." : "Content types are not enabled on this list."),
      schemaChip("wb-schema-versioning", summary.versioningText)
    );
    if (summary.isLibrary) {
      chips.append(schemaChip(
        "wb-schema-libkind",
        "document library",
        "A library copy carries the schema only \u2014 files are not copied."
      ));
    }
    const hint = el8("p", "wb-schema-hint", "Export and copy this list from the Tools tab.");
    head2.append(chips, hint);
    const sections = el8("div", "wb-schema-sections");
    sections.append(buildSettingsSection(doc));
    sections.append(buildFieldsSection(doc));
    sections.append(buildViewsSection(doc));
    if (doc.contentTypes.length) sections.append(buildContentTypesSection(doc));
    if (doc.warnings.length) sections.append(buildWarningsSection(doc));
    const root2 = el8("div", "wb-schema");
    root2.append(head2, sections);
    return root2;
  }
  function buildSchemaPane(wrap, listId) {
    const status = el8("div", "wb-grid-status", "Reading the list schema\u2026");
    wrap.append(status);
    cached2(listId, "schema", () => captureListSchema(client2, listId)).then(({ doc }) => {
      status.remove();
      wrap.append(renderSchemaPane(doc));
    }).catch((err) => {
      showFailure(status, err, "this list\u2019s schema");
    });
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
        exportName: `fields-${fileStem2(title)}`,
        query: { path: guidPath5(listId, "/fields"), options: { select: FIELD_SELECT2 } }
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
          { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true, link: absUrl },
          { key: "ViewQuery", label: "CAML query", mono: true, copyable: true }
        ],
        exportName: `views-${fileStem2(title)}`,
        query: { path: guidPath5(listId, "/views"), options: { select: VIEW_SELECT } }
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
        exportName: `contenttypes-${fileStem2(title)}`,
        query: { path: guidPath5(listId, "/contenttypes"), options: { select: CT_SELECT } }
      })
    },
    { id: "schema", label: "Schema" },
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
        exportName: `permissions-${fileStem2(title)}`,
        query: {
          path: guidPath5(listId, "/roleassignments"),
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
    { id: "items", label: "Items" },
    { id: "tools", label: "Tools" },
    { id: "raw", label: "Raw" }
  ];
  function showDetail(route) {
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = "";
    const listId = route.listId;
    const back = el8("button", "btn btn-xs wb-back", "\u2190 All lists");
    back.type = "button";
    back.addEventListener("click", () => navigate({ view: "lists" }));
    const title = el8("h2", "", route.listTitle || "List");
    const sub = el8("span", "wb-detail-id sp-copy", listId);
    sub.title = "Click to copy the list id";
    sub.addEventListener("click", () => copyText(listId, sub));
    const settingsLink = el8("a", "btn btn-xs wb-detail-settings", "List settings \u2197");
    settingsLink.href = linkUrl(client2.webUrl(), LIST_SETTINGS, { guid: listId });
    bindNewTab(settingsLink);
    settingsLink.title = "Open this list\u2019s settings page in a new tab";
    const headRow = el8("div", "wb-detail-head");
    headRow.append(back, title, sub, settingsLink);
    const tabsBar = el8("div", "wb-tabs");
    tabsBar.setAttribute("role", "tablist");
    const body = el8("div", "wb-tab-body");
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
      updateRoute?.({ tab: tab.id });
    }
    function pane(tab) {
      if (panes.has(tab.id)) return panes.get(tab.id);
      const wrap = el8("div", "wb-tab-pane");
      panes.set(tab.id, wrap);
      if (tab.id === "items") {
        buildItemsPane(wrap, listId, route.listTitle || "List");
        return wrap;
      }
      if (tab.id === "schema") {
        buildSchemaPane(wrap, listId);
        return wrap;
      }
      if (tab.id === "tools") {
        buildToolsPane(wrap, {
          client: client2,
          listId,
          listTitle: route.listTitle || "List",
          getSchema: () => cached2(listId, "schema", () => captureListSchema(client2, listId)),
          openCopy,
          createClient,
          navigate,
          inspectSite: inspectSite2,
          mockWriter: mockWriter2,
          // Drops the Items tab's own built pane so it rebuilds — and
          // reloads its rows — the next time it's opened. Never touches
          // buildItemsPane's private itemsCache directly; deleting the
          // cached pane is the one seam this tab needs into a sibling.
          invalidateItems: () => {
            panes.delete("items");
          }
        });
        return wrap;
      }
      if (tab.id === "raw") {
        const status = el8("div", "wb-grid-status", "Loading raw list entity\u2026");
        wrap.append(status);
        cached2(listId, "raw", () => client2.get(guidPath5(listId))).then((json) => {
          status.remove();
          const node = toNode(json, 0, { maxDepth: 8, maxItems: 250 });
          const inspector = el8("div", "wb-raw");
          inspector.append(enhance(node) ?? renderValue(node));
          wrap.append(inspector);
        }).catch((err) => {
          showFailure(status, err, "this list\u2019s raw entity");
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
      const btn = el8("button", "wb-tab", tab.label);
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

// ../src/workbench/views/security.js
var el9 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var roleNames = (row) => (row.RoleDefinitionBindings?.results || row.RoleDefinitionBindings || []).map((r) => r.Name).filter(Boolean).join(", ");
function createSecurityView({ client: client2 }) {
  const root = el9("section", "wb-view wb-view-security");
  const spWrite = createSpWriteClient({ client: client2 });
  const head = el9("div", "wb-view-head");
  head.innerHTML = '<h2>Permissions</h2><p class="wb-view-hint">Site groups, membership, role definitions, and who holds what on this web. The inheritance scan is on-demand \u2014 it makes SharePoint evaluate security per list.</p>';
  const headLinks = el9("div", "wb-head-links");
  const permGroup = LINK_GROUPS.find((g) => g.title === "Permissions & people");
  for (const link of (permGroup?.links || []).filter((l) => l.label !== "Access requests")) {
    const a = el9("a", "btn btn-xs wb-head-link", `${link.label} \u2197`);
    bindNewTab(a);
    a.dataset.path = link.path;
    headLinks.append(a);
  }
  head.append(headLinks);
  const tabsBar = el9("div", "wb-tabs");
  const body = el9("div", "wb-tab-body");
  root.append(head, tabsBar, body);
  const panes = /* @__PURE__ */ new Map();
  function groupsPane() {
    const wrap = el9("div", "wb-tab-pane");
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
      subject: "this web\u2019s groups",
      filterPlaceholder: "Filter groups\u2026",
      exportName: "sp-groups",
      descriptor: { ...groupsQuery, webUrl: client2.webUrl() }
    });
    const membersBox = el9("div", "wb-subpanel");
    membersBox.hidden = true;
    const membersTitle = el9("h3", "wb-subpanel-title", "");
    const membersHost = el9("div", "wb-subpanel-body");
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
        subject: "the members of this web",
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
    const wrap = el9("div", "wb-tab-pane");
    const notice = el9("div", "wb-consent");
    notice.hidden = true;
    function showNotice(message, { isError = false, confirm = null } = {}) {
      notice.textContent = "";
      notice.hidden = false;
      notice.classList.toggle("wb-consent-error", isError);
      notice.append(el9("span", "wb-consent-text", message));
      if (confirm) {
        const yes = el9("button", "btn btn-xs", confirm.label);
        yes.type = "button";
        yes.addEventListener("click", () => {
          notice.hidden = true;
          confirm.run();
        });
        notice.append(yes);
      }
      const dismiss = el9("button", "btn btn-xs", confirm ? "Cancel" : "Dismiss");
      dismiss.type = "button";
      dismiss.addEventListener("click", () => {
        notice.hidden = true;
      });
      notice.append(dismiss);
    }
    const addBar = el9("div", "wb-members-add");
    const groupSelect = el9("select", "wb-members-group");
    groupSelect.setAttribute("aria-label", "Group to add the user to");
    const loginInput = el9("input", "wb-members-login");
    loginInput.type = "text";
    loginInput.placeholder = "user@tenant.com or i:0#.f|membership|\u2026";
    const addBtn = el9("button", "btn btn-xs", "Add to group");
    addBtn.type = "button";
    addBar.append(el9("span", "wb-qb-label", "Add user"), groupSelect, loginInput, addBtn);
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
      subject: "this group\u2019s members",
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
          const opt = el9("option", "", group.Title);
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
    const wrap = el9("div", "wb-tab-pane");
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
      subject: "this web\u2019s permission levels",
      filterPlaceholder: "Filter roles\u2026",
      exportName: "sp-roledefinitions",
      descriptor: {
        path: "web/roledefinitions",
        options: { select: ["Id", "Name", "Description", "RoleTypeKind", "Hidden", "BasePermissions"] },
        webUrl: client2.webUrl()
      }
    });
    const decodeBox = el9("div", "wb-subpanel");
    decodeBox.hidden = true;
    const decodeTitle = el9("h3", "wb-subpanel-title", "");
    const decodeBody = el9("div", "wb-subpanel-body wb-flags");
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
      for (const flag of d.flags) decodeBody.append(el9("span", "wb-flag", flag));
      if (d.isEmpty) decodeBody.append(el9("span", "wb-view-hint", "No permission bits set."));
    }
    return wrap;
  }
  function assignmentsPane() {
    const wrap = el9("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "PrincipalId",
      columns: [
        { key: "Member", label: "Principal", value: (row) => row.Member?.Title || "" },
        { key: "LoginName", label: "Login", value: (row) => row.Member?.LoginName || "", mono: true, copyable: true },
        { key: "PrincipalType", label: "Type", value: (row) => row.Member?.PrincipalType, format: principalTypeName },
        { key: "Roles", label: "Roles", value: roleNames }
      ],
      emptyText: "No role assignments.",
      subject: "this web\u2019s permission assignments",
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
    const wrap = el9("div", "wb-tab-pane");
    const bar = el9("div", "wb-scan-bar");
    const btn = el9("button", "btn", "Scan lists for unique permissions");
    btn.type = "button";
    const hint = el9(
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
    const btn = el9("button", "wb-tab", tab.label);
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
var el10 = (tag, cls, text) => {
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
  const root = el10("section", "wb-view wb-view-site");
  const absUrl = (rel) => {
    try {
      return rel ? `${new URL(client2.webUrl()).origin}${encodeSpPath(rel)}` : "";
    } catch {
      return "";
    }
  };
  const head = el10("div", "wb-view-head");
  head.innerHTML = '<h2>Site overview</h2><p class="wb-view-hint">Web and site collection properties, features, subwebs, and the property bag.</p>';
  const tabsBar = el10("div", "wb-tabs");
  const body = el10("div", "wb-tab-body");
  root.append(head, tabsBar, body);
  const panes = /* @__PURE__ */ new Map();
  function sheetPane(query, extraSections = []) {
    const wrap = el10("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "Property",
      columns: [
        { key: "Property", label: "Property", mono: true, copyable: true },
        { key: "Value", label: "Value", copyable: true }
      ],
      emptyText: "Nothing returned.",
      subject: "these properties",
      filterPlaceholder: "Filter properties\u2026",
      exportName: query.exportName,
      descriptor: { path: query.path, options: query.options, webUrl: client2.webUrl() }
    });
    wrap.append(grid.el);
    grid.setLoading("Loading\u2026");
    client2.get(query.path, query.options).then((entity2) => grid.setRows(entityRows(entity2))).catch((err) => grid.setError(err));
    for (const extra of extraSections) {
      const box = el10("div", "wb-subpanel");
      box.hidden = false;
      box.append(el10("h3", "wb-subpanel-title", extra.title));
      const hostEl = el10("div", "wb-subpanel-body");
      box.append(hostEl);
      const extraGrid = createGrid({
        rowKey: "Property",
        columns: [
          { key: "Property", label: "Property", mono: true, copyable: true },
          { key: "Value", label: "Value", copyable: true }
        ],
        emptyText: "Nothing returned.",
        subject: "these properties",
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
    const wrap = el10("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "DefinitionId",
      columns: [
        { key: "Scope", label: "Scope" },
        { key: "DisplayName", label: "Feature", format: (v) => v || "(no display name)" },
        { key: "DefinitionId", label: "Definition id", mono: true, copyable: true }
      ],
      emptyText: "No activated features.",
      subject: "the features on this web",
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
    const wrap = el10("div", "wb-tab-pane");
    const query = {
      path: "web/webs",
      options: { select: ["Id", "Title", "ServerRelativeUrl", "WebTemplate", "Created", "Language"] }
    };
    const grid = createGrid({
      columns: [
        { key: "Title", label: "Title" },
        { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true, link: absUrl },
        { key: "WebTemplate", label: "Template" },
        { key: "Language", label: "Language" },
        { key: "Created", label: "Created", format: (v) => v ? String(v).slice(0, 10) : "" },
        { key: "Id", label: "Id", mono: true, copyable: true }
      ],
      emptyText: "No subwebs.",
      subject: "subwebs",
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
    const wrap = el10("div", "wb-tab-pane");
    const grid = createGrid({
      rowKey: "RawKey",
      columns: [
        { key: "Key", label: "Key (decoded)", mono: true },
        { key: "RawKey", label: "Raw key", mono: true, copyable: true },
        { key: "Value", label: "Value", copyable: true }
      ],
      emptyText: "Empty property bag.",
      subject: "this web\u2019s property bag",
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
    const btn = el10("button", "wb-tab", tab.label);
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
var el11 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
var fmtDate2 = (v) => v ? String(v).slice(0, 10) : "";
function createSiteHomeView({ client: client2, navigate, inspectSite: inspectSite2 }) {
  const root = el11("section", "wb-view wb-view-sitehome");
  const absUrl = (rel) => {
    try {
      return rel ? `${new URL(client2.webUrl()).origin}${encodeSpPath(rel)}` : "";
    } catch {
      return "";
    }
  };
  const head = el11("div", "wb-view-head");
  head.innerHTML = '<h2>Site</h2><p class="wb-view-hint">The inspected web at a glance. Full property sheets are under Advanced.</p>';
  const cards = el11("div", "wb-home-cards");
  const webCard = el11("div", "wb-home-card");
  const userCard = el11("div", "wb-home-card");
  cards.append(webCard, userCard);
  const subwebsBox = el11("div", "wb-home-subwebs");
  root.append(head, cards, subwebsBox);
  let loadedForWeb = "";
  const failureRow = (err, subject) => showFailure(el11("div", "wb-grid-status"), err, subject);
  function factRow(label, value, { copyFull = "" } = {}) {
    const row = el11("div", "wb-home-fact");
    row.append(el11("span", "wb-home-fact-label", label));
    const v = el11("span", copyFull ? "wb-home-fact-value sp-copy" : "wb-home-fact-value", value || "\u2014");
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
    webCard.append(el11("h3", "wb-home-card-title", "This web"));
    userCard.textContent = "";
    userCard.append(el11("h3", "wb-home-card-title", "You"));
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
      webCard.append(failureRow(err, "this web\u2019s details"));
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
      const roleRow = el11("div", "wb-home-fact");
      roleRow.append(el11("span", "wb-home-fact-label", "Role"));
      roleRow.append(el11(
        "span",
        user2.IsSiteAdmin ? "wb-role-chip wb-role-admin" : "wb-role-chip wb-role-user",
        user2.IsSiteAdmin ? "Site admin" : "Site user"
      ));
      userCard.append(roleRow);
    } catch (err) {
      userCard.append(failureRow(err, "your own account"));
    }
    const grid = createGrid({
      columns: [
        { key: "Title", label: "Subweb" },
        { key: "ServerRelativeUrl", label: "Url", mono: true, copyable: true, link: absUrl },
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
      // Classic webs routinely refuse web/webs to anyone without rights on
      // the child webs — the one denial this landing page hits by default.
      subject: "subwebs",
      filterPlaceholder: "Filter subwebs\u2026",
      exportName: "sp-subwebs",
      descriptor: {
        path: "web/webs",
        options: { select: ["Id", "Title", "ServerRelativeUrl", "Created", "WebTemplate"] },
        webUrl: client2.webUrl()
      }
    });
    subwebsBox.append(el11("h3", "wb-home-card-title", "Subwebs"), grid.el);
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
var el12 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function createLinksView({ client: client2 }) {
  const root = el12("section", "wb-view wb-view-links");
  const head = el12("div", "wb-view-head");
  head.innerHTML = '<h2>Panels</h2><p class="wb-view-hint">Quick jumps to the SharePoint configuration panels you actually reach for. Links open in a new tab; hover for the underlying page.</p>';
  const body = el12("div", "wb-links");
  root.append(head, body);
  function load2() {
    const webUrl = client2.webUrl();
    body.textContent = "";
    for (const group of LINK_GROUPS) {
      const card = el12("div", "wb-linkgroup");
      card.append(el12("h3", "", group.title));
      for (const link of group.links) {
        const row = el12("a", "wb-link");
        row.href = linkUrl(webUrl, link);
        bindNewTab(row);
        row.append(el12("span", "wb-link-label", link.label));
        row.append(el12("span", "wb-link-go", "\u2197"));
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
var FIELD_SELECT3 = [
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
var el13 = (tag, cls, text) => {
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
  const root = el13("section", "wb-view wb-view-query");
  const head = el13("div", "wb-view-head");
  head.innerHTML = '<h2>Query builder</h2><p class="wb-view-hint">Compose an OData query against any list \u2014 or any /_api endpoint \u2014 and run it. \u201CCopy as\u201D turns the query into a script.</p>';
  const composer = el13("div", "wb-qb");
  const targetRow = el13("div", "wb-qb-row");
  const listSelect = el13("select", "wb-qb-list");
  listSelect.setAttribute("aria-label", "Query target list");
  const endpointInput = el13("input", "wb-qb-endpoint");
  endpointInput.type = "text";
  endpointInput.placeholder = "web/currentuser \u2014 path after /_api/";
  endpointInput.hidden = true;
  targetRow.append(el13("span", "wb-qb-label", "Target"), listSelect, endpointInput);
  const fieldsBox = el13("div", "wb-qb-fields");
  const fieldsList = el13("div", "wb-qb-fieldlist");
  fieldsBox.append(el13("span", "wb-qb-label", "$select"), fieldsList);
  const filtersBox = el13("div", "wb-qb-filters");
  const filterRows = el13("div", "wb-qb-filterrows");
  const addFilter = el13("button", "btn btn-xs", "+ Filter");
  addFilter.type = "button";
  filtersBox.append(el13("span", "wb-qb-label", "$filter"), filterRows, addFilter);
  const optionsRow = el13("div", "wb-qb-row");
  const orderSelect = el13("select", "wb-qb-order");
  const orderDir = el13("select", "wb-qb-orderdir");
  for (const [v, label] of [["asc", "ascending"], ["desc", "descending"]]) {
    const opt = el13("option", "", label);
    opt.value = v;
    orderDir.append(opt);
  }
  const topInput = el13("input", "wb-qb-top");
  topInput.type = "number";
  topInput.min = "1";
  topInput.max = String(MAX_TOP);
  topInput.value = String(DEFAULT_TOP);
  const expandInput = el13("input", "wb-qb-expand");
  expandInput.type = "text";
  expandInput.placeholder = "extra $expand (comma-separated)";
  optionsRow.append(
    el13("span", "wb-qb-label", "$orderby"),
    orderSelect,
    orderDir,
    el13("span", "wb-qb-label", "$top"),
    topInput,
    el13("span", "wb-qb-label", "$expand"),
    expandInput
  );
  const rawRow = el13("div", "wb-qb-rawrow");
  const rawArea = el13("textarea", "wb-qb-raw");
  rawArea.spellcheck = false;
  rawArea.setAttribute("aria-label", "Raw query");
  const rawNote = el13("span", "wb-qb-rawnote", "Editing the raw query overrides the builder.");
  rawNote.hidden = true;
  const runBtn = el13("button", "btn btn-xs wb-qb-run wb-primary", "Run \u25B6");
  runBtn.type = "button";
  const backToBuilder = el13("button", "btn btn-xs", "Back to builder");
  backToBuilder.type = "button";
  backToBuilder.hidden = true;
  rawRow.append(rawArea, rawNote, runBtn, backToBuilder);
  composer.append(targetRow, fieldsBox, filtersBox, optionsRow, rawRow);
  const results = el13("div", "wb-qb-results");
  root.append(head, composer, results);
  let lists = [];
  let fields = [];
  let rawMode = false;
  let loadedForWeb = "";
  const fieldByName = (name) => fields.find((f) => f.InternalName === name);
  const guidPath7 = (listId) => `web/lists(guid'${listId}')/items`;
  function pickedListId() {
    return listSelect.value === "::endpoint" ? "" : listSelect.value;
  }
  function addFilterRow(saved = {}) {
    const row = el13("div", "wb-qb-filterrow");
    const join2 = el13("select", "wb-qb-join");
    for (const [v, label] of [["and", "AND"], ["or", "OR"]]) {
      const opt = el13("option", "", label);
      opt.value = v;
      join2.append(opt);
    }
    join2.value = saved.join || "and";
    if (!filterRows.childElementCount) join2.classList.add("wb-qb-join-first");
    const fieldSel = el13("select", "wb-qb-field");
    for (const f of fields) {
      const opt = el13("option", "", f.InternalName);
      opt.value = f.InternalName;
      fieldSel.append(opt);
    }
    if (saved.field) fieldSel.value = saved.field;
    const opSel = el13("select", "wb-qb-op");
    for (const [v, label] of OPERATORS) {
      const opt = el13("option", "", label);
      opt.value = v;
      opSel.append(opt);
    }
    if (saved.op) opSel.value = saved.op;
    const valueInput = el13("input", "wb-qb-value");
    valueInput.type = "text";
    valueInput.placeholder = "value";
    valueInput.value = saved.value || "";
    const remove = el13("button", "btn btn-xs", "\xD7");
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
    const path = listId ? guidPath7(listId) : String(endpointInput.value || "").trim().replace(/^\/+/, "");
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
    const blank = el13("option", "", "(no ordering)");
    blank.value = "";
    orderSelect.append(blank);
    const wanted = new Set(savedSelect || ["Id", "Title"]);
    for (const f of fields) {
      const entry = EXPANDABLE_TYPES.has(f.TypeAsString) ? `${f.InternalName}/Title` : f.InternalName;
      const label = el13("label", "wb-qb-fieldopt");
      const box = el13("input");
      box.type = "checkbox";
      box.value = entry;
      box.checked = wanted.has(entry);
      box.addEventListener("change", onBuilderChange);
      label.append(box, document.createTextNode(entry));
      label.append(el13("span", "wb-qb-fieldtype", f.TypeAsString));
      fieldsList.append(label);
      const opt = el13("option", "", f.InternalName);
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
    fieldsList.append(el13("div", "wb-qb-loading", "Loading fields\u2026"));
    try {
      const { items } = await client2.getAll(`web/lists(guid'${listId}')/fields`, {
        select: FIELD_SELECT3
      });
      fields = items.filter((f) => !f.Hidden);
      renderFieldList(saved?.select);
      for (const savedRow of saved?.filters || []) addFilterRow(savedRow);
      if (saved?.orderby) orderSelect.value = saved.orderby;
      if (saved?.orderdir) orderDir.value = saved.orderdir;
      onBuilderChange();
    } catch (err) {
      fieldsList.textContent = "";
      fieldsList.append(showFailure(el13("div", "wb-qb-loading"), err, "this list\u2019s fields"));
    }
  }
  function renderListPicker(savedListId = "") {
    listSelect.textContent = "";
    for (const list2 of lists) {
      const opt = el13("option", "", list2.Hidden ? `${list2.Title} (hidden)` : list2.Title);
      opt.value = list2.Id;
      listSelect.append(opt);
    }
    const endpoint = el13("option", "", "\u2014 arbitrary endpoint \u2014");
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
      subject: "what this query asked for",
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
      results.append(showFailure(el13("div", "wb-grid-status"), err, "the lists in this web"));
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
var el14 = (tag, cls, text) => {
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
  const row = el14("div", "wb-editor-row");
  row.dataset.internal = field2.InternalName || "";
  const label = el14("label", "wb-editor-label", field2.Title || field2.InternalName);
  const typeBadge = el14("span", "wb-editor-type", type);
  label.append(typeBadge);
  const control = el14("div", "wb-editor-control");
  const error = el14("div", "wb-editor-error");
  error.hidden = true;
  row.append(label, control, error);
  let getValue = () => "";
  const textInput = (tag, value) => {
    const input = el14(tag === "textarea" ? "textarea" : "input");
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
      const select = el14("select");
      const options = choicesOf(field2);
      const blank = el14("option", "", "\u2014");
      blank.value = "";
      select.append(blank);
      for (const choice of options) {
        const opt = el14("option", "", choice);
        opt.value = choice;
        select.append(opt);
      }
      if (initial && !options.includes(initial)) {
        const opt = el14("option", "", `${initial} (current)`);
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
      const listBox = el14("div", "wb-editor-choices");
      const initialSet = new Set(Array.isArray(initial) ? initial : []);
      const boxes = [];
      for (const choice of choicesOf(field2)) {
        const lab = el14("label", "wb-editor-choice");
        const box = el14("input");
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
      const box = el14("input");
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
  const row = el14("div", "wb-editor-row wb-editor-readonly");
  row.dataset.internal = field2.InternalName || "";
  const label = el14("label", "wb-editor-label", field2.Title || field2.InternalName);
  label.append(el14("span", "wb-editor-type", String(field2.TypeAsString || "")));
  const value = el14("div", "wb-editor-static", displayText || "");
  value.title = field2.ReadOnlyField ? "Read-only field" : "Not editable in the workbench";
  row.append(label, value);
  return row;
}
function createFieldEditorForm({ fields, item: item2 = {}, itemAsText = {}, onSave }) {
  const root = el14("div", "wb-editor-form");
  const rows = el14("div", "wb-editor-rows");
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
  const bar = el14("div", "wb-editor-bar");
  const save = el14("button", "btn btn-xs", "Save metadata");
  save.type = "button";
  const status = el14("span", "wb-editor-status");
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
function textPartIsEmpty(html) {
  if (!html) return true;
  if (/<img\b/i.test(html)) return false;
  return !textOfControl({ kind: "text", innerHTML: html });
}
function contentParts(controls) {
  const parts = [];
  let unreadable = 0;
  for (const control of controls || []) {
    if (control.kind === "text") {
      const html = String(control.innerHTML || "").trim();
      if (!textPartIsEmpty(html)) parts.push({ kind: "text", label: "Text", html, lines: [] });
    } else if (control.kind === "webpart") {
      const lines = Object.values(
        control.webPartData?.serverProcessedContent?.searchablePlainTexts || {}
      ).filter((v) => typeof v === "string" && v.trim()).map((v) => v.trim());
      if (!lines.length) continue;
      const title = String(control.webPartData?.title || "").trim();
      parts.push({
        kind: "webpart",
        label: title || webPartName(control.webPartId),
        html: "",
        lines
      });
    } else if (control.kind === "unknown") {
      unreadable += 1;
    }
  }
  const counts = /* @__PURE__ */ new Map();
  for (const part of parts) counts.set(part.label, (counts.get(part.label) || 0) + 1);
  const seen = /* @__PURE__ */ new Map();
  for (const part of parts) {
    if (counts.get(part.label) > 1) {
      const n = (seen.get(part.label) || 0) + 1;
      seen.set(part.label, n);
      part.label = `${part.label} ${n}`;
    }
  }
  return { parts, unreadable };
}
function contentBlocks(controls, override) {
  const { parts, unreadable } = override ? { parts: override, unreadable: 0 } : contentParts(controls);
  const blocks = [];
  for (const part of parts) {
    blocks.push(`## ${part.label}`);
    blocks.push(part.kind === "text" ? sanitizeHtml(part.html) : part.lines.map((t) => `- ${t}`).join("\n"));
  }
  if (unreadable) {
    blocks.push(`*[${unreadable} part${unreadable === 1 ? "" : "s"} could not be read \u2014 see the raw export.]*`);
  }
  return blocks;
}
var METADATA_SKIP = /* @__PURE__ */ new Set([
  "CanvasContent1",
  "LayoutWebpartsContent",
  "FieldValuesAsText",
  "PublishingPageContent",
  "WikiField",
  "Author",
  "Editor"
  // flattened into Created/Modified lines
]);
function buildContentExport({
  item: item2 = {},
  controls = [],
  parts = null,
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
    contentBlocks(controls, parts).join("\n\n"),
    "",
    "---",
    "",
    ...meta,
    ""
  ].join("\n");
}
function buildRawExport({ item: item2 = {}, controls = [], webParts = [] }) {
  const payload = { item: item2, controls };
  if (webParts.length) payload.webParts = webParts;
  return JSON.stringify(payload, null, 2);
}
var slug = (s) => String(s).toLowerCase().replace(/[^a-z0-9-_]+/g, "-").replace(/^-+|-+$/g, "");
function exportFileStem(item2) {
  const name = String(item2.FileLeafRef || item2.Title || "page").replace(/\.aspx$/i, "");
  return slug(name) || "page";
}
function bundleEntryName(item2, libraryRootPath) {
  const dir = String(item2?.FileDirRef || "");
  const root = String(libraryRootPath || "").replace(/\/+$/, "");
  let folder = "";
  if (root && dir.toLowerCase().startsWith(root.toLowerCase())) {
    folder = dir.slice(root.length).replace(/^\/+/, "");
  }
  const segments = folder.split("/").map(slug).filter(Boolean);
  segments.push(`${exportFileStem(item2 || {})}-content.md`);
  return segments.join("/");
}
function dedupeEntryNames(names) {
  const seen = /* @__PURE__ */ new Set();
  return (names || []).map((raw) => {
    const name = String(raw);
    if (!seen.has(name)) {
      seen.add(name);
      return name;
    }
    const dot = name.lastIndexOf(".");
    const stem2 = dot > 0 ? name.slice(0, dot) : name;
    const ext = dot > 0 ? name.slice(dot) : "";
    let n = 2;
    while (seen.has(`${stem2}-${n}${ext}`)) n += 1;
    const unique = `${stem2}-${n}${ext}`;
    seen.add(unique);
    return unique;
  });
}
function buildExportReport({ total = 0, exported = 0, failures = [] }) {
  const lines = ["# Export report", "", `${exported} of ${total} pages exported.`, ""];
  if (failures.length) {
    lines.push("Not exported:", "");
    for (const failure of failures) {
      const reason = String(failure.reason || "").replace(/\s+/g, " ").trim();
      lines.push(`- ${failure.name}${reason ? ` \u2014 ${reason}` : ""}`);
    }
    lines.push("");
  }
  return lines.join("\n");
}

// ../src/workbench/zip.js
var LOCAL_SIG = 67324752;
var CENTRAL_SIG = 33639248;
var EOCD_SIG = 101010256;
var VERSION = 20;
var FLAG_UTF8 = 2048;
var METHOD_STORE = 0;
var MAX_ENTRIES = 65535;
var MAX_NAME_BYTES = 65535;
var MAX_UINT32 = 4294967295;
var CONTROL_CHARS = /[\u0000-\u001f\u007f]/;
var crcTable = null;
function table() {
  if (crcTable) return crcTable;
  crcTable = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 3988292384 ^ c >>> 1 : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  return crcTable;
}
function crc32(bytes) {
  const t = table();
  let c = 4294967295;
  for (let i = 0; i < bytes.length; i += 1) c = t[(c ^ bytes[i]) & 255] ^ c >>> 8;
  return (c ^ 4294967295) >>> 0;
}
function dosStamp(date) {
  const year = Math.max(1980, date.getFullYear());
  return {
    time: date.getHours() << 11 | date.getMinutes() << 5 | date.getSeconds() >> 1,
    date: year - 1980 << 9 | date.getMonth() + 1 << 5 | date.getDate()
  };
}
function safeEntryName(name) {
  return String(name || "").replace(/\\/g, "/").split("/").filter((seg) => seg && seg !== "." && seg !== "..").join("/");
}
function buildZip(entries, { date = /* @__PURE__ */ new Date() } = {}) {
  const enc = new TextEncoder();
  const stamp = dosStamp(date);
  const list2 = entries || [];
  if (!Array.isArray(list2)) throw new TypeError("Zip entries must be an array.");
  if (list2.length > MAX_ENTRIES) {
    throw new RangeError(`A zip cannot hold more than ${MAX_ENTRIES} entries (got ${list2.length}).`);
  }
  const files = list2.map((entry) => {
    const safe = safeEntryName(entry.name);
    if (!safe) {
      throw new RangeError(`Zip entry name ${JSON.stringify(String(entry.name ?? ""))} is empty once it is made relative \u2014 an entry with no name cannot be extracted.`);
    }
    if (CONTROL_CHARS.test(safe)) {
      throw new RangeError(`Zip entry name ${JSON.stringify(safe)} carries a control character \u2014 extractors truncate the name there, so it would unpack under a different name than it was written under.`);
    }
    const name = enc.encode(safe);
    if (name.length > MAX_NAME_BYTES) {
      throw new RangeError(`Zip entry name is ${name.length} bytes; the limit is ${MAX_NAME_BYTES}.`);
    }
    const body = enc.encode(String(entry.text ?? ""));
    if (body.length > MAX_UINT32) {
      throw new RangeError(`Zip entry ${JSON.stringify(safe)} is ${body.length} bytes; anything over 4 GB needs Zip64, which this writer does not implement.`);
    }
    return { name, body, crc: crc32(body) };
  });
  const localSize = files.reduce((n, f) => n + 30 + f.name.length + f.body.length, 0);
  const centralSize = files.reduce((n, f) => n + 46 + f.name.length, 0);
  if (localSize > MAX_UINT32 || centralSize > MAX_UINT32 || localSize + centralSize > MAX_UINT32) {
    throw new RangeError(`This archive would be ${localSize + centralSize} bytes; anything over 4 GB needs Zip64, which this writer does not implement.`);
  }
  const out = new Uint8Array(localSize + centralSize + 22);
  const view = new DataView(out.buffer);
  let at = 0;
  const u16 = (v) => {
    view.setUint16(at, v, true);
    at += 2;
  };
  const u32 = (v) => {
    view.setUint32(at, v >>> 0, true);
    at += 4;
  };
  const raw = (bytes) => {
    out.set(bytes, at);
    at += bytes.length;
  };
  for (const file of files) {
    file.offset = at;
    u32(LOCAL_SIG);
    u16(VERSION);
    u16(FLAG_UTF8);
    u16(METHOD_STORE);
    u16(stamp.time);
    u16(stamp.date);
    u32(file.crc);
    u32(file.body.length);
    u32(file.body.length);
    u16(file.name.length);
    u16(0);
    raw(file.name);
    raw(file.body);
  }
  const centralAt = at;
  for (const file of files) {
    u32(CENTRAL_SIG);
    u16(VERSION);
    u16(VERSION);
    u16(FLAG_UTF8);
    u16(METHOD_STORE);
    u16(stamp.time);
    u16(stamp.date);
    u32(file.crc);
    u32(file.body.length);
    u32(file.body.length);
    u16(file.name.length);
    u16(0);
    u16(0);
    u16(0);
    u16(0);
    u32(0);
    u32(file.offset);
    raw(file.name);
  }
  const centralEnd = at;
  u32(EOCD_SIG);
  u16(0);
  u16(0);
  u16(files.length);
  u16(files.length);
  u32(centralEnd - centralAt);
  u32(centralAt);
  u16(0);
  return out;
}

// ../src/workbench/classic-page.js
var SITE_PAGES_BASE_TEMPLATE = 119;
var PUBLISHING_PAGES_BASE_TEMPLATE = 850;
var PUBLISHING_BODY_FIELD = "PublishingPageContent";
var WIKI_BODY_FIELD = "WikiField";
function libraryKindOf(list2) {
  if (!list2) return null;
  if (list2.BaseTemplate === SITE_PAGES_BASE_TEMPLATE) return "modern";
  if (list2.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE) return "publishing";
  return "generic";
}
var libraryKindLabel = (kind) => ({
  modern: "modern Site Pages library",
  publishing: "classic publishing Pages library",
  generic: "pages library"
})[kind] || "pages library";
function htmlHasContent(html) {
  const raw = String(html ?? "").trim();
  if (!raw) return false;
  if (/<img\b/i.test(raw)) return true;
  const doc = new DOMParser().parseFromString(raw, "text/html");
  return Boolean((doc.body?.textContent || "").trim());
}
function pageContentKindOf(item2) {
  const it = item2 || {};
  if (String(it.CanvasContent1 ?? "").trim()) return "canvas";
  if (htmlHasContent(it[PUBLISHING_BODY_FIELD])) return "publishing";
  if (htmlHasContent(it[WIKI_BODY_FIELD])) return "wiki";
  return "empty";
}
var pageContentKindLabel = (kind) => ({
  canvas: "modern canvas page",
  publishing: "classic publishing page",
  wiki: "classic wiki page",
  webparts: "classic web-part page",
  empty: "page with no readable body"
})[kind] || "page";
function unwrapCdata(value) {
  const raw = String(value ?? "");
  const m = /^\s*<!\[CDATA\[([\s\S]*)\]\]>\s*$/.exec(raw);
  return m ? m[1] : raw;
}
function normalizeWebPart(entry, index = 0) {
  const wp = entry?.WebPart || {};
  const props = wp.Properties || {};
  const content = unwrapCdata(props.Content).trim();
  const contentLink = String(props.ContentLink ?? "").trim();
  return {
    id: entry?.Id || wp.Id || `wp-${index}`,
    title: String(wp.Title ?? "").trim(),
    zoneIndex: Number.isFinite(wp.ZoneIndex) ? wp.ZoneIndex : index,
    order: index,
    hidden: Boolean(wp.Hidden),
    closed: Boolean(wp.IsClosed),
    hasHtml: Boolean(content) || Boolean(contentLink),
    content,
    contentLink,
    properties: props
  };
}
function classicWebParts(entries) {
  return (entries || []).map((entry, i) => normalizeWebPart(entry, i)).sort((a, b) => a.zoneIndex - b.zoneIndex || a.order - b.order);
}
function webPartPart(wp) {
  const label = wp.title || "Embedded content";
  if (htmlHasContent(wp.content)) {
    return { kind: "text", label, html: wp.content, lines: [] };
  }
  if (wp.contentLink) {
    return {
      kind: "webpart",
      label,
      html: "",
      lines: [`Content linked from ${wp.contentLink}`]
    };
  }
  return null;
}
function guidOf(value) {
  const text = String(value ?? "").replace(/_/g, "-");
  const m = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i.exec(text);
  return m ? m[0].toLowerCase() : "";
}
function bodyParts(bodyHtml, bodyLabel, webParts, used) {
  const raw = String(bodyHtml ?? "").trim();
  const parts = [];
  const pushText = (html) => {
    if (htmlHasContent(html)) parts.push({ kind: "text", label: bodyLabel, html: html.trim(), lines: [] });
  };
  if (!raw) return parts;
  const doc = new DOMParser().parseFromString(raw, "text/html");
  const boxes = [...doc.querySelectorAll(".ms-rte-wpbox")];
  if (!boxes.length) {
    pushText(raw);
    return parts;
  }
  const byGuid = /* @__PURE__ */ new Map();
  for (const wp of webParts) {
    const guid = guidOf(wp.id);
    if (guid && !byGuid.has(guid)) byGuid.set(guid, wp);
  }
  boxes.forEach((box, i) => {
    const marker = doc.createElement("dcspad-wp");
    marker.setAttribute("data-i", String(i));
    box.replaceWith(marker);
  });
  const segments = doc.body.innerHTML.split(/<dcspad-wp data-i="(\d+)"><\/dcspad-wp>/);
  for (let s = 0; s < segments.length; s += 1) {
    if (s % 2 === 0) {
      pushText(segments[s]);
      continue;
    }
    const box = boxes[Number(segments[s])];
    const wp = byGuid.get(guidOf(box.outerHTML));
    if (wp) {
      used.add(wp);
      const part = webPartPart(wp);
      if (part) parts.push(part);
    }
  }
  return parts;
}
function classicContentParts({ item: item2 = {}, webParts = [], contentKind = null } = {}) {
  const kind = contentKind || pageContentKindOf(item2);
  const bodyField = kind === "wiki" ? WIKI_BODY_FIELD : PUBLISHING_BODY_FIELD;
  const bodyLabel = kind === "wiki" ? "Wiki content" : "Page content";
  const used = /* @__PURE__ */ new Set();
  const parts = bodyParts(item2[bodyField], bodyLabel, webParts, used);
  for (const wp of webParts) {
    if (used.has(wp) || !wp.hasHtml) continue;
    const part = webPartPart(wp);
    if (part) parts.push(part);
  }
  const counts = /* @__PURE__ */ new Map();
  for (const part of parts) counts.set(part.label, (counts.get(part.label) || 0) + 1);
  const seen = /* @__PURE__ */ new Map();
  for (const part of parts) {
    if (counts.get(part.label) > 1) {
      const n = (seen.get(part.label) || 0) + 1;
      seen.set(part.label, n);
      part.label = `${part.label} ${n}`;
    }
  }
  return { parts, unreadable: 0 };
}

// ../src/workbench/views/pages.js?v=2
var PAGE_SELECT_BASE = [
  "Id",
  "Title",
  "FileLeafRef",
  "FileRef",
  "FileDirRef",
  "Modified",
  "UniqueId",
  "Editor/Title"
];
var PAGE_SELECT_MODERN = [...PAGE_SELECT_BASE, "PromotedState"];
var MAX_BULK_PAGES = 200;
var BULK_CONCURRENCY = 4;
function pageQueryPlan(fieldInternalNames, kind) {
  const names = fieldInternalNames ? new Set(fieldInternalNames) : null;
  const hasField = (f) => names ? names.has(f) : kind === "modern";
  const showPromoted = hasField("PromotedState");
  const showTitle = names ? names.has("Title") : true;
  const modern = showPromoted && hasField("CanvasContent1");
  const gridSelect = (showPromoted ? PAGE_SELECT_MODERN : PAGE_SELECT_BASE).filter((f) => f !== "Title" || showTitle);
  return {
    showPromoted,
    showTitle,
    gridSelect,
    // Ladders, not single shapes — see queryLadder(). Rung 0 is the query we
    // want; every rung below it gives something up to stay answerable.
    gridShapes: [
      { options: { select: gridSelect, expand: "Editor" } },
      // No lookup projection, so no expand to satisfy: the Editor column
      // goes blank and everything else still lists.
      {
        options: { select: gridSelect.filter((f) => !f.includes("/")) },
        lost: "the Editor column"
      },
      // Nothing named at all. SPO returns the item's own fields, which is
      // every column this grid reads except the expanded Editor.
      { options: {}, lost: "the Editor column" }
    ],
    detailShapes: [
      {
        options: modern ? { select: DETAIL_SELECT, expand: DETAIL_EXPAND } : { select: CLASSIC_DETAIL_SELECT, expand: DETAIL_EXPAND }
      },
      // '*' still carries every content field the drilldown reads
      // (CanvasContent1, PublishingPageContent, WikiField); only the two
      // expanded people fields are out of reach, leaving their raw ids.
      { options: { select: ["*"] }, lost: "the author and editor names" },
      { options: {}, lost: "the author and editor names" }
    ]
  };
}
async function queryLadder(shapes, attempt, startAt = 0) {
  let firstError;
  const from = Math.min(Math.max(startAt, 0), shapes.length - 1);
  for (let index = from; index < shapes.length; index += 1) {
    const shape = shapes[index];
    try {
      return {
        value: await attempt(shape.options),
        lost: shape.lost || "",
        index,
        reason: firstError?.message || ""
      };
    } catch (err) {
      if (err?.status !== 400) throw err;
      firstError ||= err;
    }
  }
  throw firstError;
}
var DETAIL_EXPAND = ["Author", "Editor"];
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
var CLASSIC_DETAIL_SELECT = ["*", "Author/Title", "Editor/Title"];
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
var SITE_PAGES_BASE_TEMPLATE2 = 119;
var PUBLISHING_PAGES_BASE_TEMPLATE2 = 850;
var PAGES_LIBRARY_RANKS = [
  (l) => !l.Hidden && l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE2,
  (l) => !l.Hidden && l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE2,
  (l) => !l.Hidden && String(l.Title).toLowerCase() === "pages",
  (l) => l.BaseTemplate === SITE_PAGES_BASE_TEMPLATE2,
  (l) => l.BaseTemplate === PUBLISHING_PAGES_BASE_TEMPLATE2
];
function pagesLibraryCandidates(items) {
  const seen = /* @__PURE__ */ new Set();
  const found = [];
  for (const matches of PAGES_LIBRARY_RANKS) {
    for (const list2 of items || []) {
      if (seen.has(list2.Id) || !matches(list2)) continue;
      seen.add(list2.Id);
      found.push(list2);
    }
  }
  return found;
}
var promotedLabel = (v) => ({ 0: "", 1: "News (pending)", 2: "News" })[v] ?? String(v ?? "");
var fmtDate4 = (v) => v ? String(v).slice(0, 10) : "";
var el15 = (tag, cls, text) => {
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
var guidPath6 = (listId, sub = "") => `web/lists(guid'${listId}')${sub}`;
function reducedChip(lost, where, because = "") {
  const chip = el15("span", "wb-info-chip wb-reduced-chip", "some fields unavailable");
  chip.title = `SharePoint rejected part of this query${where ? ` for ${where}` : ""}, so ${lost} could not be read.` + (because ? `

SharePoint said: ${because}` : "");
  return chip;
}
function createPagesView({ client: client2, navigate, updateRoute }) {
  const root = el15("section", "wb-view wb-view-pages");
  const spWrite = createSpWriteClient({ client: client2 });
  const gridPane = el15("div", "wb-pane");
  const head = el15("div", "wb-view-head");
  head.innerHTML = '<h2>Pages</h2><p class="wb-view-hint">Every page in this web\u2019s pages library, subfolders included. Click a row to inspect content, metadata, and structure.</p>';
  const strip = el15("div", "wb-lib-strip");
  strip.append(el15("span", "wb-lib-wait", "locating library\u2026"));
  head.append(strip);
  const libraryLink = el15("a", "btn btn-xs wb-head-link", "Open \u2197");
  bindNewTab(libraryLink);
  libraryLink.hidden = true;
  strip.append(libraryLink);
  function renderLibraryStrip() {
    if (!current) {
      strip.hidden = true;
      return;
    }
    strip.hidden = false;
    strip.textContent = "";
    if (libraries.length > 1) {
      const select = el15("select", "wb-lib-select wb-lib-picker");
      select.setAttribute("aria-label", "Pages library to inspect");
      select.title = `This web has ${libraries.length} pages libraries \u2014 pick which one to inspect.`;
      for (const lib of libraries) {
        const opt = el15("option", "", `${lib.title} \xB7 ${lib.baseTemplate}${lib.hidden ? " \xB7 hidden" : ""}`);
        opt.value = lib.listId;
        if (lib.listId === current.listId) opt.selected = true;
        select.append(opt);
      }
      select.addEventListener("change", () => {
        switchLibrary(libraries.find((l) => l.listId === select.value));
      });
      strip.append(select);
    } else {
      const name = el15("span", "wb-lib-name sp-copy", current.title);
      if (current.rootPath) {
        name.title = `Click to copy the library path
${current.rootPath}`;
        name.addEventListener("click", () => copyText(current.rootPath, name));
      }
      strip.append(name);
    }
    const kind = el15(
      "span",
      `wb-info-chip wb-lib-kind wb-lib-${current.kind}`,
      libraryKindLabel(current.kind)
    );
    kind.title = `BaseTemplate ${current.baseTemplate}` + (current.hidden ? " \xB7 hidden library" : "") + (current.rootPath ? `
${current.rootPath}` : "");
    strip.append(kind);
    if (current.viewUrl) {
      libraryLink.href = current.viewUrl;
      libraryLink.title = `Open ${current.title} in a new tab`;
      libraryLink.hidden = false;
    } else {
      libraryLink.hidden = true;
    }
    strip.append(libraryLink);
  }
  const masterStatus = el15("div", "wb-grid-status");
  masterStatus.hidden = true;
  gridPane.append(head, masterStatus);
  const detailPane = el15("div", "wb-pane");
  detailPane.hidden = true;
  root.append(gridPane, detailPane);
  let librariesPromise = null;
  let libraries = [];
  let current = null;
  let grid = null;
  let pagesLoaded = false;
  const detailCache = /* @__PURE__ */ new Map();
  let detailRung = 0;
  let fieldsPromise = null;
  let detailRun = 0;
  let loadRun = 0;
  const toLibrary = (list2) => ({
    listId: list2.Id,
    title: list2.Title,
    kind: libraryKindOf(list2),
    baseTemplate: list2.BaseTemplate,
    hidden: Boolean(list2.Hidden),
    rootPath: list2.RootFolder?.ServerRelativeUrl || "",
    viewUrl: list2.DefaultViewUrl || list2.RootFolder?.ServerRelativeUrl || ""
  });
  function pagesLibraries() {
    if (!librariesPromise) {
      librariesPromise = client2.getAll("web/lists", {
        select: ["Id", "Title", "BaseTemplate", "Hidden", "DefaultViewUrl", "RootFolder/ServerRelativeUrl"],
        expand: "RootFolder",
        top: 5e3
      }).then(({ items }) => {
        libraries = pagesLibraryCandidates(items).map(toLibrary);
        if (!current) current = libraries[0] || null;
        return libraries;
      }).catch((err) => {
        librariesPromise = null;
        throw err;
      });
    }
    return librariesPromise;
  }
  function adoptLibrary(next) {
    current = next;
    detailCache.clear();
    webPartCache.clear();
    fieldsPromise = null;
    planPromise = null;
    detailRung = 0;
    if (grid) {
      grid.el.remove();
      grid = null;
    }
    pagesLoaded = false;
  }
  function switchLibrary(next) {
    if (!next || next.listId === current?.listId) return;
    adoptLibrary(next);
    navigate({ view: "pages", libId: next.listId });
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
  async function contentMarkdownFor(item2, sitePages, parts = null) {
    const parsed = parseCanvasContent(item2.CanvasContent1);
    let readingParts = parts;
    if (!readingParts) {
      const contentKind = pageContentKindOf(item2);
      if (contentKind === "canvas") {
        readingParts = contentParts(parsed.controls).parts;
      } else {
        const fetched = await classicWebPartsOf(item2.FileRef);
        readingParts = classicContentParts({
          item: item2,
          webParts: fetched.parts,
          contentKind
        }).parts;
      }
    }
    const web = await webIdentity();
    return buildContentExport({
      item: item2,
      controls: parsed.controls,
      parts: readingParts,
      siteTitle: web.Title || "",
      webUrl: web.Url || client2.webUrl(),
      libraryTitle: sitePages.title,
      libraryRootPath: sitePages.rootPath
    });
  }
  let exporting = false;
  async function exportContentZip() {
    if (exporting || !grid || !current) return;
    const sitePages = current;
    const rows = grid.getExportRows();
    if (!rows.length) return;
    if (rows.length > MAX_BULK_PAGES) {
      masterStatus.textContent = `${rows.length} pages selected \u2014 this export is capped at ${MAX_BULK_PAGES}. Narrow the selection or filter the grid first.`;
      masterStatus.classList.add("wb-error");
      masterStatus.hidden = false;
      return;
    }
    exporting = true;
    masterStatus.classList.remove("wb-error");
    masterStatus.hidden = false;
    let done = 0;
    const progress = () => {
      masterStatus.textContent = `Exporting ${done} of ${rows.length} page${rows.length === 1 ? "" : "s"}\u2026`;
    };
    progress();
    const results = new Array(rows.length);
    try {
      const plan = await queryPlan(sitePages);
      let next = 0;
      const worker = async () => {
        for (let i = next++; i < rows.length; i = next++) {
          const row = rows[i];
          try {
            const { item: item2 } = await pageItem(sitePages.listId, row.Id, plan.detailShapes);
            results[i] = { item: item2, text: await contentMarkdownFor(item2, sitePages) };
          } catch (err) {
            results[i] = {
              failure: {
                name: row.FileLeafRef || row.Title || `Page ${row.Id}`,
                reason: err?.message || String(err)
              }
            };
          }
          done += 1;
          progress();
        }
      };
      await Promise.all(
        Array.from({ length: Math.min(BULK_CONCURRENCY, rows.length) }, worker)
      );
      if (current !== sitePages) {
        masterStatus.hidden = true;
        return;
      }
      const ok = results.filter((r) => r && r.text !== void 0);
      const failures = results.filter((r) => r && r.failure).map((r) => r.failure);
      const names = dedupeEntryNames(
        ok.map((r) => bundleEntryName(r.item, sitePages.rootPath))
      );
      const entries = ok.map((r, i) => ({ name: names[i], text: r.text }));
      if (failures.length) {
        entries.push({
          name: "_export-report.md",
          text: buildExportReport({ total: rows.length, exported: ok.length, failures })
        });
      }
      if (!entries.length) {
        masterStatus.textContent = "No pages could be read, so there was nothing to export.";
        masterStatus.hidden = false;
        return;
      }
      downloadBytes("sp-pages-content.zip", buildZip(entries), "application/zip");
      masterStatus.hidden = true;
    } catch (err) {
      showFailure(masterStatus, err, `the pages in ${sitePages.title}`);
    } finally {
      exporting = false;
    }
  }
  async function loadPages() {
    if (pagesLoaded) return;
    const run = ++loadRun;
    masterStatus.hidden = true;
    try {
      await pagesLibraries();
      if (!current) {
        strip.hidden = true;
        masterStatus.textContent = "This web has no pages library \u2014 looked for modern Site Pages (BaseTemplate 119), classic publishing Pages (850), and any library titled \u201CPages\u201D.";
        masterStatus.hidden = false;
        return;
      }
      const sitePages = current;
      renderLibraryStrip();
      if (!grid) {
        const plan = await queryPlan(sitePages);
        const paging = { orderby: "FileLeafRef", top: 5e3 };
        const query = { path: guidPath6(sitePages.listId, "/items") };
        const descriptor = {
          ...query,
          options: { ...plan.gridShapes[0].options, ...paging },
          webUrl: client2.webUrl()
        };
        grid = createGrid({
          columns: [
            { key: "FileLeafRef", label: "Name", mono: true },
            // Dropped with the field itself: a library without Title would
            // otherwise carry a column that can only ever be blank.
            ...plan.showTitle ? [{ key: "Title", label: "Title" }] : [],
            {
              key: "Folder",
              label: "Folder",
              value: (row) => folderOf(row.FileDirRef, sitePages.rootPath),
              format: (v) => v ? `/${v}` : ""
            },
            ...plan.showPromoted ? [{ key: "PromotedState", label: "Promoted", format: promotedLabel }] : [],
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
                a.title = "Open the page in a new tab";
                a.textContent = "\u2197";
                bindNewTab(a);
                return a;
              }
            }
          ],
          onOpen: (row) => navigate({
            view: "pages",
            pageId: row.Id,
            pageName: row.FileLeafRef || row.Title,
            // Without this a reload resolves the saved id against the ranked
            // default library — same id, different page, and the Metadata tab
            // would then write to the wrong item.
            libId: sitePages.listId
          }),
          emptyText: `No pages in ${sitePages.title}.`,
          subject: `the pages in ${sitePages.title}`,
          filterPlaceholder: "Filter pages\u2026",
          toolbarExtras: strip,
          exportName: "sp-pages",
          // Ticking rows scopes every export, the zip included. The checkbox
          // cell owns its own clicks (see grid.js), so selecting a page and
          // opening one stay distinct gestures on the same row.
          selectable: true,
          exportExtras: [
            ["Download content .zip", () => exportContentZip()]
          ],
          // The same object the ladder rewrites below, on purpose: the
          // "Copy as…" menu reads it at click time, so a script copied out of
          // a degraded grid reproduces the query that actually worked rather
          // than the one SharePoint rejected.
          descriptor
        });
        gridPane.append(grid.el);
        grid.setLoading("Loading pages\u2026");
        const { value, lost, reason } = await queryLadder(plan.gridShapes, (options) => {
          descriptor.options = { ...options, ...paging };
          return client2.getAll(query.path, descriptor.options);
        });
        const { items, partial } = value;
        if (run !== loadRun) return;
        if (lost) strip.insertBefore(reducedChip(lost, sitePages.title, reason), libraryLink);
        grid.setRows(items, { partial });
        pagesLoaded = true;
      }
    } catch (err) {
      if (run !== loadRun) return;
      if (strip.querySelector(".wb-lib-wait")) strip.hidden = true;
      if (grid) grid.setError(err);
      else showFailure(masterStatus, err, "this web\u2019s pages");
    }
  }
  let planPromise = null;
  function queryPlan(sitePages) {
    if (!planPromise) {
      planPromise = listFields(sitePages.listId).then((fields) => pageQueryPlan(fields.map((f) => f.InternalName), sitePages.kind)).catch(() => pageQueryPlan(null, sitePages.kind));
    }
    return planPromise;
  }
  function pageItem(listId, pageId, shapes) {
    const key2 = `${listId}:${pageId}`;
    const path = guidPath6(listId, `/items(${pageId})`);
    if (!detailCache.has(key2)) {
      detailCache.set(
        key2,
        queryLadder(shapes, (options) => client2.get(path, options), detailRung).then(({ value, lost, index, reason }) => {
          detailRung = index;
          return { item: value, lost, reason };
        }).catch((err) => {
          detailCache.delete(key2);
          throw err;
        })
      );
    }
    return detailCache.get(key2);
  }
  const webPartCache = /* @__PURE__ */ new Map();
  function classicWebPartsOf(fileRef) {
    const key2 = String(fileRef || "");
    if (!key2) return Promise.resolve({ parts: [], error: null });
    if (!webPartCache.has(key2)) {
      const path = `web/getfilebyserverrelativepath(decodedurl='${odataPathLiteral(key2)}')/getlimitedwebpartmanager(scope=1)/webparts`;
      webPartCache.set(key2, client2.getAll(path, { expand: "WebPart/Properties" }).then(({ items }) => ({ parts: classicWebParts(items), error: null })).catch((err) => {
        webPartCache.delete(key2);
        return { parts: [], error: err };
      }));
    }
    return webPartCache.get(key2);
  }
  function listFields(listId) {
    if (!fieldsPromise) {
      fieldsPromise = client2.getAll(guidPath6(listId, "/fields"), { select: FIELD_SELECT4 }).then(({ items }) => items).catch((err) => {
        fieldsPromise = null;
        throw err;
      });
    }
    return fieldsPromise;
  }
  function structurePane(parsed) {
    const wrap = el15("div", "wb-tab-pane");
    const tree = el15("div", "wb-canvas-tree");
    const { sections, unplaced } = buildSectionTree(parsed.controls);
    if (!sections.length && !unplaced.length) {
      tree.append(el15("div", "wb-grid-status", "No canvas sections on this page."));
    }
    sections.forEach((section, i) => {
      const bits = [`${section.columns.length} column${section.columns.length === 1 ? "" : "s"}`];
      if (section.emphasis) bits.push(`emphasis ${section.emphasis}`);
      if (section.vertical) bits.push("vertical");
      if (section.collapsible) bits.push("collapsible");
      tree.append(el15("div", "wb-canvas-section", `Section ${i + 1} \u2014 ${bits.join(", ")}`));
      for (const column of section.columns) {
        const row = el15("div", "wb-canvas-column");
        const width = typeof column.sectionFactor === "number" ? `${column.sectionFactor}/12` : "auto";
        row.append(el15("span", "wb-canvas-width", width));
        if (!column.controls.length) row.append(el15("span", "wb-canvas-chip wb-canvas-empty", "empty"));
        for (const control of column.controls) {
          const chipLabel = control.kind === "text" ? "Text" : control.kind === "webpart" ? control.webPartData.title || webPartName(control.webPartId) : control.kind;
          const chip = el15("span", "wb-canvas-chip", chipLabel);
          chip.title = control.kind === "webpart" ? `${webPartName(control.webPartId)} \xB7 ${control.webPartId}` : textOfControl(control).slice(0, 200);
          row.append(chip);
        }
        tree.append(row);
      }
    });
    if (unplaced.length) {
      tree.append(el15("div", "wb-canvas-section", `Unplaced entries (${unplaced.length})`));
      for (const control of unplaced) {
        const row = el15("div", "wb-canvas-column");
        row.append(el15("span", "wb-canvas-chip", control.kind));
        tree.append(row);
      }
    }
    wrap.append(tree);
    return wrap;
  }
  function webPartsPane(parsed) {
    const wrap = el15("div", "wb-tab-pane");
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
  function classicWebPartsPane(webParts, error) {
    const wrap = el15("div", "wb-tab-pane");
    if (error) {
      const notice = el15(
        "div",
        "wb-grid-notice",
        "\u26A0 The page\u2019s web parts could not be read \u2014 " + (error.message || String(error))
      );
      wrap.append(notice);
    }
    const rows = webParts.map((wp) => ({
      Title: wp.title,
      Zone: wp.zoneIndex,
      Content: wp.content ? "inline" : wp.contentLink ? "linked" : "",
      ContentLink: wp.contentLink,
      Hidden: wp.hidden ? "yes" : "",
      Closed: wp.closed ? "yes" : "",
      Id: wp.id
    }));
    const partsGrid = createGrid({
      columns: [
        { key: "Title", label: "Title" },
        { key: "Zone", label: "Zone index" },
        { key: "Content", label: "Content" },
        { key: "ContentLink", label: "Content link", mono: true, copyable: true },
        { key: "Hidden", label: "Hidden" },
        { key: "Closed", label: "Closed" },
        { key: "Id", label: "Web part id", mono: true, copyable: true }
      ],
      emptyText: "No web parts on this page.",
      filterPlaceholder: "Filter web parts\u2026",
      exportName: "sp-page-webparts"
    });
    wrap.append(partsGrid.el);
    partsGrid.setRows(rows);
    return wrap;
  }
  function textPane(parts, notice) {
    const wrap = el15("div", "wb-tab-pane wb-text-pane");
    if (notice) wrap.append(notice);
    if (!parts.length) {
      wrap.append(el15("div", "wb-grid-status", "No readable content on this page."));
      return wrap;
    }
    const contentBlock = el15("div", "wb-text-block");
    contentBlock.append(el15("div", "wb-subpanel-title", "Content"));
    const rendered = el15("div", "wb-text-rendered");
    for (const part of parts) {
      rendered.append(el15("h3", "wb-text-part", part.label));
      if (part.kind === "text") {
        const body = el15("div", "wb-text-body");
        body.innerHTML = sanitizeHtml(part.html);
        rendered.append(body);
      } else {
        const list2 = el15("ul", "wb-text-lines");
        for (const line of part.lines) list2.append(el15("li", "", line));
        rendered.append(list2);
      }
    }
    contentBlock.append(rendered);
    wrap.append(contentBlock);
    const withHtml = parts.filter((p) => p.kind === "text");
    if (withHtml.length) {
      const htmlBlock = el15("div", "wb-text-block");
      htmlBlock.append(el15("div", "wb-subpanel-title", "HTML"));
      htmlBlock.append(el15(
        "pre",
        "wb-text-raw",
        withHtml.map((p) => `<!-- ${p.label} -->
${p.html}`).join("\n\n")
      ));
      wrap.append(htmlBlock);
    }
    return wrap;
  }
  function metadataPane(listId, pageId) {
    const wrap = el15("div", "wb-tab-pane");
    const status = el15("div", "wb-grid-status", "Loading metadata\u2026");
    wrap.append(status);
    (async () => {
      const fields = await listFields(listId);
      let item2;
      let itemAsText = {};
      try {
        item2 = await client2.get(guidPath6(listId, `/items(${pageId})`), {
          expand: "FieldValuesAsText"
        });
        itemAsText = item2.FieldValuesAsText || {};
      } catch {
        item2 = await client2.get(guidPath6(listId, `/items(${pageId})`));
        try {
          itemAsText = await client2.get(guidPath6(listId, `/items(${pageId})/FieldValuesAsText`));
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
      showFailure(status, err, "this page\u2019s metadata");
    });
    return wrap;
  }
  function rawPane(item2, parsed, webParts = []) {
    const wrap = el15("div", "wb-tab-pane");
    const payload = { item: item2, parsedCanvas: parsed.controls };
    if (webParts.length) payload.webParts = webParts;
    const node = toNode(payload, 0, { maxDepth: 10, maxItems: 400 });
    const inspector = el15("div", "wb-raw");
    inspector.append(enhance(node) ?? renderValue(node));
    wrap.append(inspector);
    return wrap;
  }
  async function showDetail(route) {
    const run = ++detailRun;
    gridPane.hidden = true;
    detailPane.hidden = false;
    detailPane.textContent = "";
    const back = el15("button", "btn btn-xs wb-back", "\u2190 All pages");
    back.type = "button";
    back.addEventListener("click", () => navigate({ view: "pages", libId: current?.listId }));
    const title = el15("h2", "", route.pageName || `Page ${route.pageId}`);
    const headRow = el15("div", "wb-detail-head");
    headRow.append(back, title);
    detailPane.append(headRow);
    const status = el15("div", "wb-grid-status", "Loading page\u2026");
    detailPane.append(status);
    let sitePages;
    let item2;
    let lostFields = "";
    let lostReason = "";
    try {
      await pagesLibraries();
      sitePages = current;
      if (!sitePages) throw new Error("This web has no pages library.");
      const plan = await queryPlan(sitePages);
      ({ item: item2, lost: lostFields, reason: lostReason } = await pageItem(
        sitePages.listId,
        route.pageId,
        plan.detailShapes
      ));
    } catch (err) {
      if (run !== detailRun) return;
      showFailure(status, err, "this page");
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
      const frag = el15("span", "wb-detail-id sp-copy", item2.FileRef);
      frag.title = `Click to copy the full URL
${fullUrl}`;
      frag.addEventListener("click", () => copyText(fullUrl, frag));
      headRow.append(frag);
    }
    const parsed = parseCanvasContent(item2.CanvasContent1);
    const contentKind = pageContentKindOf(item2);
    const isCanvas = contentKind === "canvas";
    let classicParts = [];
    let webParts = [];
    let webPartError = null;
    if (!isCanvas) {
      const fetched = await classicWebPartsOf(item2.FileRef);
      if (run !== detailRun) return;
      webParts = fetched.parts;
      webPartError = fetched.error;
      classicParts = classicContentParts({ item: item2, webParts, contentKind }).parts;
    }
    const readingParts = isCanvas ? contentParts(parsed.controls).parts : classicParts;
    const displayKind = !isCanvas && contentKind === "empty" && readingParts.length ? "webparts" : contentKind;
    const kindChip = el15("span", "wb-info-chip wb-detail-kind", pageContentKindLabel(displayKind));
    kindChip.title = isCanvas ? "Modern canvas page \u2014 Structure shows its sections and columns." : `${pageContentKindLabel(displayKind)} \u2014 no canvas sections or columns, so the Structure tab does not apply. Content Editor and Script Editor web-part content is merged into Extract.`;
    headRow.append(kindChip);
    if (lostFields) headRow.append(reducedChip(lostFields, "this page", lostReason));
    const actions = el15("span", "wb-detail-actions");
    const exportContent = el15("button", "btn btn-xs", "Export content");
    exportContent.type = "button";
    exportContent.title = "One human-readable file: metadata, merged web-part content, full metadata";
    const exportRaw = el15("button", "btn btn-xs", "Export raw");
    exportRaw.type = "button";
    exportRaw.title = "Item + parsed canvas controls as JSON, for scripts";
    actions.append(exportContent, exportRaw);
    if (item2.FileRef) {
      const open = el15("a", "btn btn-xs", "Open page \u2197");
      open.href = item2.FileRef;
      bindNewTab(open);
      actions.append(open);
    }
    headRow.append(actions);
    exportContent.addEventListener("click", async () => {
      downloadText(
        `${exportFileStem(item2)}-content.md`,
        await contentMarkdownFor(item2, sitePages, readingParts),
        "text/markdown;charset=utf-8"
      );
    });
    exportRaw.addEventListener("click", () => {
      downloadText(
        `${exportFileStem(item2)}-raw.json`,
        buildRawExport({ item: item2, controls: parsed.controls, webParts }),
        "application/json"
      );
    });
    if (isCanvas && parsed.errors.length) {
      const notice = el15(
        "div",
        "wb-grid-notice",
        `\u26A0 ${parsed.errors.length} canvas entr${parsed.errors.length === 1 ? "y" : "ies"} could not be fully parsed \u2014 shown raw where possible.`
      );
      notice.title = parsed.errors.join("\n");
      detailPane.append(notice);
    }
    const tabsBar = el15("div", "wb-tabs");
    tabsBar.setAttribute("role", "tablist");
    const body = el15("div", "wb-tab-body");
    const panes = /* @__PURE__ */ new Map();
    const TABS = [
      {
        id: "text",
        label: "Extract",
        build: () => textPane(readingParts, webPartError ? el15(
          "div",
          "wb-grid-notice",
          `\u26A0 This page\u2019s web parts could not be read, so embedded content may be missing \u2014 ${webPartError.message || String(webPartError)}`
        ) : null)
      },
      { id: "metadata", label: "Metadata", build: () => metadataPane(sitePages.listId, route.pageId) },
      ...isCanvas ? [{ id: "structure", label: "Structure", build: () => structurePane(parsed) }] : [],
      {
        id: "webparts",
        label: "Web parts",
        build: () => isCanvas ? webPartsPane(parsed) : classicWebPartsPane(webParts, webPartError)
      },
      { id: "raw", label: "Raw", build: () => rawPane(item2, parsed, webParts) }
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
      const btn = el15("button", "wb-tab", tab.label);
      btn.type = "button";
      btn.dataset.tab = tab.id;
      btn.setAttribute("role", "tab");
      btn.addEventListener("click", () => activate(tab));
      tabsBar.append(btn);
    }
    detailPane.append(tabsBar, body);
    activate(TABS.find((t) => t.id === route.tab) || TABS[0]);
  }
  async function applyRouteLibrary(route) {
    if (!route?.libId) return true;
    await pagesLibraries();
    if (route.libId === current?.listId) return true;
    const wanted = libraries.find((l) => l.listId === route.libId);
    if (!wanted) return false;
    adoptLibrary(wanted);
    return true;
  }
  function rememberLibrary() {
    if (current && libraries.length > 1) updateRoute?.({ libId: current.listId });
  }
  function showMissingLibrary() {
    detailPane.hidden = true;
    gridPane.hidden = false;
    strip.hidden = true;
    if (grid) {
      grid.el.remove();
      grid = null;
    }
    pagesLoaded = false;
    masterStatus.textContent = "The pages library this page was opened from is no longer available on this web \u2014 it may have been deleted or renamed, or your access may have changed. Pick a library to continue.";
    masterStatus.classList.add("wb-error");
    masterStatus.hidden = false;
  }
  function load2(route) {
    if (route?.pageId) {
      detailRun += 1;
      const run = detailRun;
      applyRouteLibrary(route).catch(() => false).then((ok) => {
        if (run !== detailRun) return;
        if (!ok) {
          showMissingLibrary();
          return;
        }
        rememberLibrary();
        showDetail(route);
      });
    } else {
      detailRun += 1;
      detailPane.hidden = true;
      gridPane.hidden = false;
      applyRouteLibrary(route).catch(() => false).then((ok) => {
        masterStatus.classList.remove("wb-error");
        if (!ok) {
          showMissingLibrary();
          return;
        }
        rememberLibrary();
        loadPages();
      });
    }
  }
  return { el: root, load: load2 };
}

// ../src/workbench/upload-metadata.js
function metadataFieldStates(libraryFields) {
  const fields = Array.isArray(libraryFields) ? libraryFields : [];
  const states = {};
  for (const spec of FILE_METADATA_SPECS) {
    const match = fields.find((f) => String(f.InternalName || "").toLowerCase() === spec.internalName.toLowerCase());
    let reason = "";
    if (!match) reason = `${spec.internalName} is not available in this library.`;
    else if (match.ReadOnlyField) reason = `${spec.internalName} is read-only.`;
    else if (match.Hidden) reason = `${spec.internalName} is hidden in this library.`;
    else if (!spec.types.includes(String(match.TypeAsString || ""))) {
      reason = `${spec.internalName} is not a supported text field.`;
    }
    states[spec.key] = {
      key: spec.key,
      label: spec.label,
      internalName: match?.InternalName || spec.internalName,
      entityPropertyName: match?.EntityPropertyName || match?.InternalName || spec.internalName,
      available: !reason,
      reason
    };
  }
  return states;
}
var anyMetadataAvailable = (states) => Object.values(states || {}).some((s) => s.available);
var el16 = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function openUploadMetadataDialog({
  fileName,
  overwrite = false,
  states,
  values = {},
  doUpload,
  doMetadata
}) {
  return new Promise((resolve, reject) => {
    const dialog = el16("dialog", "app-dialog sp-metadata-dialog wb-upload-metadata");
    const panel = el16("div", "app-dialog__panel");
    const head = el16("div", "app-dialog__head");
    head.append(el16("h2", "", "File metadata"));
    const closeBtn = el16("button", "btn btn-ghost btn-xs", "\u2715");
    closeBtn.type = "button";
    closeBtn.setAttribute("aria-label", "Close");
    head.append(closeBtn);
    const context = el16("p", "app-dialog__context", overwrite ? `Review metadata before replacing ${fileName}.` : `Add metadata before uploading ${fileName}.`);
    panel.append(head, context);
    const inputs = {};
    for (const state2 of Object.values(states)) {
      const label = el16(
        "label",
        `app-dialog__field sp-metadata-field ${state2.available ? "available" : "unavailable"}`
      );
      const headRow = el16("span", "sp-metadata-field__head");
      headRow.append(el16("span", "", state2.label));
      headRow.append(el16(
        "span",
        "sp-metadata-field__state",
        state2.available ? "Available" : "Unavailable"
      ));
      const input = state2.key === "description" ? el16("textarea") : el16("input");
      if (state2.key === "description") input.rows = 4;
      else {
        input.type = "text";
        input.maxLength = 255;
        input.autocomplete = "off";
      }
      input.className = `wb-upload-meta-${state2.key}`;
      input.disabled = !state2.available;
      input.value = state2.available ? String(values[state2.key] ?? "") : "";
      const hint = el16("span", "sp-metadata-field__hint", state2.available ? `Writes to the ${state2.internalName} field.` : state2.reason);
      label.append(headRow, input, hint);
      panel.append(label);
      inputs[state2.key] = input;
    }
    const error = el16("div", "sp-files-error");
    error.setAttribute("role", "alert");
    error.hidden = true;
    const actions = el16("div", "app-dialog__actions sp-metadata-actions");
    const cancel = el16("button", "btn btn-ghost", "Cancel");
    cancel.type = "button";
    const keep = el16("button", "btn wb-upload-meta-keep", "Keep file without metadata");
    keep.type = "button";
    keep.hidden = true;
    const primary = el16("button", "btn btn-run wb-upload-meta-go", "Upload file");
    primary.type = "button";
    actions.append(cancel, keep, primary);
    panel.append(error, actions);
    dialog.append(panel);
    document.body.append(dialog);
    const readValues = () => Object.fromEntries(
      Object.entries(inputs).map(([key2, input]) => [key2, input.value])
    );
    const finish = (outcome) => {
      dialog.close();
      dialog.remove();
      resolve(outcome);
    };
    const fail = (err) => {
      dialog.close();
      dialog.remove();
      reject(err);
    };
    let uploaded = false;
    let busy = false;
    async function run() {
      if (busy) return;
      busy = true;
      error.hidden = true;
      primary.disabled = true;
      cancel.disabled = true;
      keep.disabled = true;
      if (!uploaded) {
        primary.textContent = "Uploading\u2026";
        try {
          await doUpload();
          uploaded = true;
        } catch (err) {
          if (err && typeof err === "object") err.uploadMetadataValues = readValues();
          fail(err);
          return;
        }
      }
      primary.textContent = "Saving metadata\u2026";
      try {
        await doMetadata(readValues());
        finish("saved");
      } catch (err) {
        busy = false;
        cancel.hidden = true;
        closeBtn.hidden = true;
        keep.hidden = false;
        keep.disabled = false;
        primary.textContent = "Retry metadata";
        primary.disabled = false;
        error.textContent = `The file was uploaded, but its metadata could not be saved: ${err?.message || err}`;
        error.hidden = false;
      }
    }
    primary.addEventListener("click", run);
    keep.addEventListener("click", () => {
      if (!busy) finish("kept");
    });
    const dismiss = () => {
      if (busy) return;
      if (uploaded) return;
      finish("cancelled");
    };
    cancel.addEventListener("click", dismiss);
    closeBtn.addEventListener("click", dismiss);
    dialog.addEventListener("cancel", (e) => {
      e.preventDefault();
      dismiss();
    });
    dialog.showModal();
    Object.values(inputs).find((input) => !input.disabled)?.focus();
  });
}

// ../src/workbench/views/browser.js?v=4
var FIELD_SELECT5 = [
  "Id",
  "Title",
  "InternalName",
  "EntityPropertyName",
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
var CHECK_IN_COMMENT2 = "Uploaded from SP Workbench";
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
var el17 = (tag, cls, text) => {
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
  const root = el17("section", "wb-view wb-view-files");
  const spWrite = createSpWriteClient({ client: client2 });
  const head = el17("div", "wb-view-head");
  head.innerHTML = '<h2>Files</h2><p class="wb-view-hint">Browse any library or folder of this web \u2014 every file type, with download, binary upload, folder creation, and full metadata editing.</p>';
  const bar = el17("div", "wb-crumbs-bar");
  const librarySelect = el17("select", "wb-lib-select");
  librarySelect.setAttribute("aria-label", "Jump to a document library");
  const crumbs = el17("div", "wb-crumbs");
  bar.append(librarySelect, crumbs);
  const consent = el17("div", "wb-consent");
  consent.hidden = true;
  const gridWrap = el17("div", "wb-files-grid");
  const metaPanel = el17("div", "wb-subpanel wb-file-meta");
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
    const rootBtn = el17("button", "wb-crumb", rootPath === "/" ? "/" : rootPath);
    rootBtn.type = "button";
    rootBtn.addEventListener("click", () => navigate({ view: "files", path: rootPath }));
    let started = rootPath === "/";
    if (started) crumbs.append(rootBtn);
    for (const segment of segments) {
      acc += `/${segment}`;
      if (!started) {
        if (normalizedPath2(acc) === rootPath) {
          started = true;
          const btn2 = el17("button", "wb-crumb", rootPath);
          btn2.type = "button";
          btn2.addEventListener("click", () => navigate({ view: "files", path: rootPath }));
          crumbs.append(btn2);
        }
        continue;
      }
      crumbs.append(el17("span", "wb-crumb-sep", "/"));
      const target = acc;
      const btn = el17("button", "wb-crumb", segment);
      btn.type = "button";
      btn.addEventListener("click", () => navigate({ view: "files", path: target }));
      crumbs.append(btn);
    }
    syncCrumbOverflow();
  }
  function syncCrumbOverflow() {
    crumbs.scrollLeft = crumbs.scrollWidth;
    crumbs.classList.toggle("is-clipped", crumbs.scrollWidth > crumbs.clientWidth + 1);
  }
  const crumbObserver = typeof ResizeObserver === "function" ? new ResizeObserver(() => syncCrumbOverflow()) : null;
  crumbObserver?.observe(crumbs);
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
    const blank = el17("option", "", "Libraries\u2026");
    blank.value = "";
    librarySelect.append(blank);
    for (const lib of libraries) {
      const url = lib.RootFolder?.ServerRelativeUrl;
      if (!url) continue;
      const opt = el17("option", "", lib.Title);
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
            const wrap = el17("span", `wb-file-name wb-node-${fileRole(row)}`);
            const glyph = icon(row.kind === "folder" ? "folder" : "file");
            glyph.classList.add("wb-node");
            wrap.append(glyph, el17("span", "wb-file-name-text", name));
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
      subject: "this folder",
      filterPlaceholder: "Filter files\u2026",
      exportName: "sp-files",
      toolbarExtras: bar
    });
    const uploadBtn = el17("button", "btn btn-xs wb-primary", "Upload\u2026");
    uploadBtn.type = "button";
    const fileInput = el17("input");
    fileInput.type = "file";
    fileInput.hidden = true;
    fileInput.setAttribute("aria-label", "Choose a file to upload");
    uploadBtn.addEventListener("click", () => fileInput.click());
    fileInput.addEventListener("change", () => {
      if (fileInput.files?.length) startUpload(fileInput.files[0]);
      fileInput.value = "";
    });
    const newFolderBtn = el17("button", "btn btn-xs wb-newfolder", "New folder\u2026");
    newFolderBtn.type = "button";
    newFolderBtn.addEventListener("click", promptNewFolder);
    const refreshBtn = el17("button", "btn btn-xs", "Refresh");
    refreshBtn.type = "button";
    refreshBtn.addEventListener("click", () => listFolder(currentPath, { force: true }));
    grid.actionsEl.prepend(uploadBtn, fileInput, newFolderBtn, refreshBtn);
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
  function showConsent(message, onConfirm, { gate = "" } = {}) {
    consent.textContent = "";
    consent.hidden = false;
    consent.append(el17("span", "wb-consent-text", message));
    const replace = el17("button", "btn btn-xs", "Replace");
    replace.type = "button";
    replace.addEventListener("click", () => {
      consent.hidden = true;
      onConfirm();
    });
    const cancel = el17("button", "btn btn-xs", "Cancel");
    cancel.type = "button";
    cancel.addEventListener("click", () => {
      consent.hidden = true;
    });
    if (gate) {
      replace.disabled = true;
      const row = el17("label", "sp-metadata-consent__row wb-consent-gate");
      const box = el17("input");
      box.type = "checkbox";
      box.className = "wb-consent-checkout";
      box.addEventListener("change", () => {
        replace.disabled = !box.checked;
      });
      row.append(box, el17("span", "sp-metadata-consent__label", gate));
      consent.append(row);
    }
    consent.append(replace, cancel);
  }
  function uploadNotice(message, isError = false) {
    consent.textContent = "";
    consent.hidden = false;
    consent.classList.toggle("wb-consent-error", isError);
    consent.append(el17("span", "wb-consent-text", message));
    const dismiss = el17("button", "btn btn-xs", "Dismiss");
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
      await confirmReplace(file, folderPath, { existing });
      return;
    }
    await runUpload(file, { overwrite: false, folderPath });
  }
  async function libraryForcesCheckout(folderPath) {
    const listId = await parentListId(folderPath);
    const list2 = await client2.get(`web/lists(guid'${listId}')`, { select: "ForceCheckout" });
    return Boolean(list2?.ForceCheckout);
  }
  async function readCheckOut(filePath, existing) {
    if (existing && !isCheckedOut(existing.CheckOutType)) return { checkedOut: false };
    const file = await client2.get(fileApi(filePath, ""), {
      select: "CheckOutType,CheckedOutByUser/Id,CheckedOutByUser/Title,CheckedOutByUser/LoginName,CheckedOutByUser/Email",
      expand: "CheckedOutByUser"
    });
    return {
      checkedOut: isCheckedOut(file?.CheckOutType),
      user: file?.CheckedOutByUser || null
    };
  }
  async function checkOutState(folderPath, fileName, existing) {
    const state2 = {
      required: false,
      checkedOut: false,
      checkedOutByCurrentUser: false,
      checkedOutBy: ""
    };
    try {
      state2.required = await libraryForcesCheckout(folderPath);
      const { checkedOut, user: user2 } = await readCheckOut(`${folderPath}/${fileName}`, existing);
      state2.checkedOut = checkedOut;
      if (!checkedOut) return state2;
      state2.checkedOutBy = String(user2?.Title || user2?.LoginName || "").trim();
      state2.checkedOutByCurrentUser = isCheckedOutByCurrentUser(
        user2,
        client2.context()?.pageContext,
        { sameWeb: client2.webUrl() === client2.hostWebUrl() }
      );
    } catch {
    }
    return state2;
  }
  async function confirmReplace(file, folderPath, { existing = null, carriedValues = null } = {}) {
    const checkout = await checkOutState(folderPath, file.name, existing);
    if (checkout.checkedOut && !checkout.checkedOutByCurrentUser) {
      uploadNotice(
        `\u201C${file.name}\u201D is checked out to ${checkout.checkedOutBy || "another user"}. It cannot be replaced until they check it back in.`,
        true
      );
      return;
    }
    let gate = "";
    if (checkout.checkedOutByCurrentUser) {
      gate = `${checkout.required ? "This library requires check-out \u2014 " : ""}\u201C${file.name}\u201D is already checked out to you. Replace it and check it back in.`;
    } else if (checkout.required) {
      gate = `This library requires check-out \u2014 check \u201C${file.name}\u201D out before replacing it, then back in.`;
    }
    showConsent(
      `\u201C${file.name}\u201D already exists in this folder. Replace it?`,
      () => runUpload(file, { overwrite: true, folderPath, carriedValues, checkout }),
      { gate }
    );
  }
  async function uploadMetadataStates(folderPath) {
    try {
      const listId = await parentListId(folderPath);
      const fields = await listFields(listId);
      const states = metadataFieldStates(fields);
      return anyMetadataAvailable(states) ? states : null;
    } catch {
      return null;
    }
  }
  async function prefillUploadValues(states, folderPath, fileName) {
    const values = { title: "", description: "", docVersion: "" };
    const available = Object.values(states).filter((s) => s.available);
    try {
      const item2 = await client2.get(
        fileApi(`${folderPath}/${fileName}`, "/ListItemAllFields"),
        { select: available.map((s) => s.entityPropertyName) }
      );
      for (const s of available) {
        values[s.key] = String(item2?.[s.entityPropertyName] ?? item2?.[s.internalName] ?? "");
      }
    } catch {
    }
    return values;
  }
  async function runUpload(file, {
    overwrite,
    folderPath,
    carriedValues = null,
    checkout = null
  }) {
    let data;
    try {
      data = await file.arrayBuffer();
    } catch (err) {
      uploadNotice(`Could not read the file: ${err?.message || err}`, true);
      return;
    }
    let uploaded = null;
    let checkedOutHere = false;
    const bareUpload = async () => {
      if (overwrite && checkout?.required && !checkout.checkedOutByCurrentUser) {
        await spWrite.checkOutFile(`${folderPath}/${file.name}`);
        checkout.checkedOutByCurrentUser = true;
        checkedOutHere = true;
      }
      uploaded = await spWrite.uploadFile(folderPath, file.name, data, { overwrite });
      return uploaded;
    };
    const settle = async (message) => {
      const held = checkout?.checkedOutByCurrentUser || isCheckedOut(uploaded?.checkOutType);
      if (!held) return finishUpload(file, folderPath, message);
      try {
        await spWrite.checkInFile(`${folderPath}/${file.name}`, { comment: CHECK_IN_COMMENT2 });
        return finishUpload(file, folderPath, `${message} Checked in.`);
      } catch (err) {
        return finishUpload(
          file,
          folderPath,
          `${message} It could not be checked in and is still checked out to you: ${err?.message || err}`,
          true
        );
      }
    };
    const fail = async (err, carried) => {
      if (checkedOutHere && !uploaded) {
        if (currentPath === folderPath) await listFolder(folderPath, { force: true });
        uploadNotice(
          `Upload failed: ${err?.message || err} \u201C${file.name}\u201D is still checked out to you.`,
          true
        );
        return;
      }
      handleUploadError(err, file, folderPath, overwrite, carried);
    };
    const states = await uploadMetadataStates(folderPath);
    if (!states) {
      uploadNotice(`Uploading \u201C${file.name}\u201D\u2026`);
      try {
        await bareUpload();
        await settle(`Uploaded \u201C${file.name}\u201D \u2713`);
      } catch (err) {
        await fail(err, null);
      }
      return;
    }
    const baseline = overwrite ? await prefillUploadValues(states, folderPath, file.name) : { title: "", description: "", docVersion: "" };
    const values = carriedValues || baseline;
    const initial = { ...baseline };
    const filePath = `${folderPath}/${file.name}`;
    try {
      const outcome = await openUploadMetadataDialog({
        fileName: file.name,
        overwrite,
        states,
        values,
        doUpload: bareUpload,
        // Write only what the user changed against the prefill (a cleared
        // prefill still writes ''); untouched values cost no request.
        doMetadata: async (entered) => {
          const formValues = Object.values(states).filter((s) => s.available && String(entered[s.key] ?? "") !== String(initial[s.key] ?? "")).map((s) => ({ FieldName: s.internalName, FieldValue: String(entered[s.key] ?? "") }));
          if (!formValues.length) return;
          await spWrite.validateUpdateListItem(
            { fileServerRelativeUrl: filePath },
            formValues,
            { newDocumentUpdate: true, checkInComment: CHECK_IN_COMMENT2 }
          );
        }
      });
      if (outcome === "cancelled") return;
      await settle(outcome === "saved" ? `Uploaded \u201C${file.name}\u201D \u2713` : `Uploaded \u201C${file.name}\u201D \u2713 (kept without metadata)`);
    } catch (err) {
      await fail(err, err?.uploadMetadataValues || null);
    }
  }
  async function finishUpload(file, folderPath, message, isError = false) {
    if (currentPath === folderPath) await listFolder(folderPath, { force: true });
    uploadNotice(message, isError);
  }
  function handleUploadError(err, file, folderPath, overwrite, carriedValues) {
    if (err?.code === "conflict" && !overwrite) {
      void confirmReplace(file, folderPath, { carriedValues, existing: null });
      return;
    }
    uploadNotice(`Upload failed: ${err?.message || err}`, true);
  }
  function promptNewFolder() {
    const folderPath = currentPath;
    consent.classList.remove("wb-consent-error");
    consent.textContent = "";
    consent.hidden = false;
    consent.append(el17("span", "wb-consent-text", `New folder in ${folderPath}:`));
    const nameIn = el17("input", "wb-folder-name");
    nameIn.type = "text";
    nameIn.placeholder = "Folder name";
    nameIn.setAttribute("aria-label", "New folder name");
    const create = el17("button", "btn btn-xs wb-primary", "Create");
    create.type = "button";
    const cancel = el17("button", "btn btn-xs", "Cancel");
    cancel.type = "button";
    cancel.addEventListener("click", () => {
      consent.hidden = true;
    });
    const submit = async () => {
      const name = nameIn.value.trim();
      if (!name) {
        nameIn.focus();
        return;
      }
      create.disabled = true;
      try {
        await spWrite.createFolder(folderPath, name);
        await listFolder(folderPath, { force: true });
        uploadNotice(`Created folder \u201C${name}\u201D.`);
      } catch (err) {
        uploadNotice(`Could not create the folder: ${err?.message || err}`, true);
      }
    };
    create.addEventListener("click", submit);
    nameIn.addEventListener("keydown", (e) => {
      if (e.key === "Enter") submit();
    });
    consent.append(nameIn, create, cancel);
    nameIn.focus();
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
        select: FIELD_SELECT5
      }).then(({ items }) => items).catch((err) => {
        fieldsCache.delete(listId);
        throw err;
      }));
    }
    return fieldsCache.get(listId);
  }
  async function openMetadata(row) {
    metaPanel.hidden = false;
    metaPanel.textContent = "";
    const titleRow = el17("div", "wb-file-meta-head");
    titleRow.append(el17("h3", "wb-subpanel-title", `Metadata for ${row.Name}`));
    const close = el17("button", "btn btn-xs", "Close");
    close.type = "button";
    close.addEventListener("click", () => {
      metaPanel.hidden = true;
    });
    titleRow.append(close);
    metaPanel.append(titleRow);
    const body = el17("div", "wb-subpanel-body");
    metaPanel.append(body);
    const status = el17("div", "wb-grid-status", "Loading metadata\u2026");
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
      showFailure(status, err, "this file\u2019s metadata");
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
  function destroy() {
    crumbObserver?.disconnect();
  }
  return { el: root, load: load2, destroy };
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
  statusCtx2.textContent = ctx2.live ? `SP: ${ctx2.label}${inspectingNote} \xB7 Build #${APP_BUILD_INFO.build}` : `SP: mock data (deploy to SharePoint for live inspection)${inspectingNote} \xB7 Build #${APP_BUILD_INFO.build}`;
  statusCtx2.title = buildTooltipFor("SP Workbench");
}
var ctx = getSpContext();
applyWorkbenchContext(ctx);
var client = createSpRestClient({
  mockResolver: ctx.live ? null : mockResolver
});
var shell = createShell({
  mount: document.getElementById("wb-main"),
  // inspectSite is a hoisted declaration below; views get a late-bound ref.
  // createClient hands a view its own independent REST client (a *second*
  // web connection, e.g. the Schema tab's "Copy to…" target) without
  // disturbing the shell's own client or importing mock-data.js directly.
  deps: {
    client,
    inspectSite: (url) => inspectSite(url),
    createClient: () => createSpRestClient({ mockResolver: ctx.live ? null : mockResolver }),
    // Mock-mode-only write fixture (see mock-data.js) — the Schema tab's
    // apply dialog is the sole consumer; views never import mock-data.js
    // directly (invariant: only main.js touches it).
    mockWriter: ctx.live ? void 0 : mockWriter
  },
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
applyWorkbenchBuildMarker();
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
