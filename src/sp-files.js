// SharePoint document-library text transfer.
//
// This module has no DOM or storage ownership. It operates only on HTML, CSS,
// and JavaScript files in a selected same-tenant SharePoint web and keeps all
// paths inside that web's server-relative boundary.

import { getSpContext } from './bridge/sp-context.js';
import { MAX_IMPORT_BYTES, paneForFileName } from './io.js?v=2';

const ACCEPT_JSON = 'application/json;odata=nometadata';
const DIGEST_SAFETY_MS = 60_000;

export class SpFileError extends Error {
  constructor(message, { code = 'sharepoint', status = 0, cause } = {}) {
    super(message, { cause });
    this.name = 'SpFileError';
    this.code = code;
    this.status = status;
  }
}

function normalizedPath(value) {
  let path = String(value || '').trim().replaceAll('\\', '/');
  if (!path.startsWith('/')) path = `/${path}`;
  path = path.replace(/\/{2,}/g, '/');
  if (path.length > 1) path = path.replace(/\/+$/, '');
  return path;
}

function pathFromWebUrl(webUrl) {
  try {
    return normalizedPath(decodeURIComponent(new URL(webUrl).pathname));
  } catch {
    return '/';
  }
}

function odataPathLiteral(value) {
  // Encode URL-significant characters such as # and %, but retain OData's
  // doubled-apostrophe escaping inside the surrounding string literal.
  return encodeURIComponent(String(value)).replaceAll("'", "''");
}

function resultArray(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.results)) return value.results;
  return [];
}

function unwrapJson(data) {
  return data?.d?.GetContextWebInformation
    || data?.GetContextWebInformation
    || data?.d
    || data;
}

async function responseMessage(response) {
  try {
    const body = await response.clone().json();
    return body?.error?.message?.value
      || body?.error?.message
      || body?.['odata.error']?.message?.value
      || '';
  } catch {
    try { return (await response.text()).trim(); }
    catch { return ''; }
  }
}

async function requireOk(response, fallback, code) {
  if (response.ok) return response;
  const detail = await responseMessage(response);
  let message = detail || `${fallback} (HTTP ${response.status})`;
  let normalizedCode = code;
  if (response.status === 401 || response.status === 403) {
    message = detail
      || 'SharePoint denied this request. Check library permissions and try again.';
    normalizedCode = 'permission';
  } else if (response.status === 404) {
    message = detail || 'The SharePoint file or folder was not found.';
    normalizedCode = 'not-found';
  } else if (response.status === 409) {
    message = detail || 'A SharePoint file with that name already exists.';
    normalizedCode = 'conflict';
  }
  throw new SpFileError(message, {
    code: normalizedCode,
    status: response.status,
  });
}

export function createSpFilesClient({
  fetchImpl = (...args) => fetch(...args),
  getContext = getSpContext,
} = {}) {
  const digestCache = new Map();

  function context({ refresh = false } = {}) {
    const ctx = getContext({ refresh });
    if (!ctx?.live || !ctx.pageContext?.webAbsoluteUrl) {
      throw new SpFileError(
        'SharePoint file transfer requires an SP: Live context.',
        { code: 'not-live' },
      );
    }
    return ctx;
  }

  function webInfo(targetWebUrl = '') {
    const ctx = context({ refresh: true });
    const hostWebUrl = ctx.pageContext.webAbsoluteUrl.replace(/\/+$/, '');
    let webUrl = hostWebUrl;
    if (targetWebUrl) {
      try {
        const candidate = new URL(String(targetWebUrl).trim(), hostWebUrl);
        if (!/^https?:$/.test(candidate.protocol) || candidate.origin !== new URL(hostWebUrl).origin) {
          throw new Error('origin');
        }
        candidate.hash = '';
        candidate.search = '';
        webUrl = candidate.href.replace(/\/+$/, '');
      } catch {
        throw new SpFileError(
          'Enter a SharePoint site URL on this tenant, such as /sites/ProjectName.',
          { code: 'invalid-web-url' },
        );
      }
    }
    const rootPath = normalizedPath(
      webUrl === hostWebUrl && ctx.pageContext.webServerRelativeUrl
        ? ctx.pageContext.webServerRelativeUrl
        : pathFromWebUrl(webUrl),
    );
    return { ctx, webUrl, rootPath, hostWebUrl };
  }

  function checkedPath(path, rootPath) {
    const normalized = normalizedPath(path || rootPath);
    if (rootPath !== '/'
        && normalized !== rootPath
        && !normalized.startsWith(`${rootPath}/`)) {
      throw new SpFileError(
        'That path is outside the current SharePoint web.',
        { code: 'outside-web' },
      );
    }
    return normalized;
  }

  async function request(url, options = {}) {
    try {
      return await fetchImpl(url, {
        credentials: 'same-origin',
        ...options,
      });
    } catch (cause) {
      throw new SpFileError(
        `Could not reach SharePoint (${cause.message || cause}).`,
        { code: 'network', cause },
      );
    }
  }

  async function fetchContextInfo(targetWebUrl = '') {
    const requested = webInfo(targetWebUrl);
    const { webUrl } = requested;
    const response = await request(`${webUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { Accept: ACCEPT_JSON },
    });
    await requireOk(response, 'Could not obtain SharePoint request context', 'context');
    const info = unwrapJson(await response.json()) || {};
    const value = info.FormDigestValue || info.formDigestValue;
    if (!value) {
      throw new SpFileError(
        'SharePoint contextinfo did not return a request digest.',
        { code: 'context' },
      );
    }
    const timeoutSeconds =
      Number(info.FormDigestTimeoutSeconds || info.formDigestTimeoutSeconds) || 1800;
    const canonicalWebUrl = webInfo(
      info.WebFullUrl || info.webFullUrl || webUrl,
    ).webUrl;
    const cached = {
      value,
      expiresAt: Date.now() + (timeoutSeconds * 1000),
      webFullUrl: canonicalWebUrl,
      siteFullUrl: info.SiteFullUrl || info.siteFullUrl || '',
    };
    digestCache.set(webUrl.toLowerCase(), cached);
    digestCache.set(canonicalWebUrl.toLowerCase(), cached);
    return {
      ...cached,
      webUrl: canonicalWebUrl,
      rootPath: pathFromWebUrl(canonicalWebUrl),
    };
  }

  async function connectWeb(targetWebUrl = '') {
    const info = await fetchContextInfo(targetWebUrl);
    return {
      webUrl: info.webUrl,
      rootPath: info.rootPath,
      siteFullUrl: info.siteFullUrl,
    };
  }

  async function getDigest({ force = false, webUrl: targetWebUrl = '' } = {}) {
    const target = webInfo(targetWebUrl);
    const cacheKey = target.webUrl.toLowerCase();
    const cached = digestCache.get(cacheKey);
    if (!force && cached?.expiresAt - DIGEST_SAFETY_MS > Date.now()) {
      return cached.value;
    }

    // A page digest is only a candidate for the web that supplied the page
    // context. Other tenant sites always receive their own /contextinfo call.
    if (!force && !cached && target.webUrl === target.hostWebUrl) {
      const ctx = context({ refresh: true });
      const value = ctx.pageContext.formDigestValue;
      const timeoutSeconds = Number(ctx.pageContext.formDigestTimeoutSeconds) || 0;
      if (value && !ctx.pageContext.isDcsPadMock && timeoutSeconds > 0) {
        const pageDigest = {
          value,
          expiresAt: (ctx.capturedAt || Date.now()) + timeoutSeconds * 1000,
          webFullUrl: ctx.pageContext.webAbsoluteUrl,
          siteFullUrl: ctx.pageContext.siteAbsoluteUrl || '',
        };
        digestCache.set(cacheKey, pageDigest);
        if (pageDigest.expiresAt - DIGEST_SAFETY_MS > Date.now()) return value;
      }
    }

    return (await fetchContextInfo(target.webUrl)).value;
  }

  async function listFolder(serverRelativePath, { webUrl: targetWebUrl = '' } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')`
      + '?$select=Name,ServerRelativeUrl'
      + '&$expand=Folders($select=Name,ServerRelativeUrl),'
      + 'Files($select=Name,ServerRelativeUrl,Length,TimeLastModified)';
    const response = await request(endpoint, {
      headers: { Accept: ACCEPT_JSON },
    });
    await requireOk(response, 'Could not list the SharePoint folder', 'list');
    const data = unwrapJson(await response.json()) || {};
    const folders = resultArray(data.Folders)
      .map((item) => ({
        kind: 'folder',
        name: String(item.Name || ''),
        serverRelativeUrl: checkedPath(item.ServerRelativeUrl, rootPath),
      }))
      .filter((item) => item.name)
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
    const files = resultArray(data.Files)
      .map((item) => ({
        kind: 'file',
        name: String(item.Name || ''),
        pane: paneForFileName(item.Name),
        serverRelativeUrl: checkedPath(item.ServerRelativeUrl, rootPath),
        length: Number(item.Length) || 0,
        modified: item.TimeLastModified || '',
      }))
      .filter((item) => item.name && item.pane)
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
    return {
      path: checkedPath(data.ServerRelativeUrl || path, rootPath),
      rootPath,
      folders,
      files,
    };
  }

  async function readTextFile(serverRelativePath, { webUrl: targetWebUrl = '' } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const pane = paneForFileName(path);
    if (!pane) {
      throw new SpFileError(
        'Only HTML, CSS, and JavaScript files can be imported.',
        { code: 'unsupported-file' },
      );
    }
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')/$value`;
    const response = await request(endpoint);
    await requireOk(response, 'Could not download the SharePoint file', 'read');
    const length = Number(response.headers.get('content-length')) || 0;
    if (length > MAX_IMPORT_BYTES) {
      throw new SpFileError(
        'The selected SharePoint file is larger than the 5 MB import limit.',
        { code: 'too-large' },
      );
    }
    const text = await response.text();
    if (new Blob([text]).size > MAX_IMPORT_BYTES) {
      throw new SpFileError(
        'The selected SharePoint file is larger than the 5 MB import limit.',
        { code: 'too-large' },
      );
    }
    return {
      fileName: path.slice(path.lastIndexOf('/') + 1),
      pane,
      text,
      serverRelativeUrl: path,
    };
  }

  async function writeTextFile(
    folderPath,
    fileName,
    text,
    { overwrite = false, webUrl: targetWebUrl = '' } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const folder = checkedPath(folderPath, rootPath);
    const safeName = String(fileName || '').trim();
    if (!/^[a-z0-9][a-z0-9._-]*\.(?:html?|css|js)$/i.test(safeName)) {
      throw new SpFileError(
        'Use a safe HTML, CSS, or JS file name containing letters, numbers, dots, hyphens, or underscores.',
        { code: 'invalid-name' },
      );
    }

    const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(folder)}')/Files/AddUsingPath(`
      + `decodedUrl='${odataPathLiteral(safeName)}',overwrite=${overwrite ? 'true' : 'false'})`;

    const upload = async (forceDigest) => {
      const digest = await getDigest({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: 'POST',
        headers: {
          Accept: ACCEPT_JSON,
          'Content-Type': 'text/plain; charset=utf-8',
          'X-RequestDigest': digest,
        },
        body: text,
      });
    };

    let response = await upload(false);
    if (response.status === 403) response = await upload(true);
    await requireOk(response, 'Could not upload the SharePoint file', 'write');
    let result = {};
    try { result = unwrapJson(await response.json()) || {}; }
    catch { /* A successful upload may return no JSON body on older servers. */ }
    return {
      fileName: safeName,
      serverRelativeUrl:
        result.ServerRelativeUrl || `${folder.replace(/\/$/, '')}/${safeName}`,
    };
  }

  return {
    webInfo,
    connectWeb,
    getDigest,
    listFolder,
    readTextFile,
    writeTextFile,
  };
}

const defaultClient = createSpFilesClient();

export const getSpWebInfo = (webUrl) => defaultClient.webInfo(webUrl);
export const connectSpWeb = (webUrl) => defaultClient.connectWeb(webUrl);
export const getDigest = (options) => defaultClient.getDigest(options);
export const listFolder = (path, options) => defaultClient.listFolder(path, options);
export const readTextFile = (path, options) => defaultClient.readTextFile(path, options);
export const writeTextFile = (folder, name, text, options) =>
  defaultClient.writeTextFile(folder, name, text, options);
