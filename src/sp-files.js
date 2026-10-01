// SharePoint document-library text transfer and Browser resource discovery.
//
// This module has no DOM or storage ownership. It operates only on HTML, CSS,
// JavaScript, JSON, CSV, Markdown, and text files in a selected same-tenant SharePoint web
// and keeps all paths inside that web's server-relative boundary.

import { getSpContext } from './bridge/sp-context.js';
import { MAX_IMPORT_BYTES, fileTypeForFileName } from './io.js?v=2';
import {
  ACCEPT_JSON, SpFileError, odataPathLiteral, resultArray, unwrapJson, requireOk,
} from './sp-odata.js';

const DIGEST_SAFETY_MS = 60_000;
const LIBRARY_GUID = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i;
export const CHECK_IN_COMMENT = 'Saved from DCSPad';
// SP.CheckOutType: 0 = checked out online, 1 = checked out offline, 2 = none.
export const CHECK_OUT_TYPE_NONE = 2;

export function isCheckedOut(checkOutType) {
  const type = Number(checkOutType ?? CHECK_OUT_TYPE_NONE);
  return Number.isFinite(type) && type !== CHECK_OUT_TYPE_NONE;
}

// "Is this file checked out to me?" Login claims are the only identifier that
// survives a cross-site write: a site user Id is issued per site collection,
// so it is only trusted for the web that supplied the page context.
// Exported: the SP Workbench Files uploader shares this rule, so the two
// forced-check-out implementations can't drift on what "to me" means.
export function isCheckedOutByCurrentUser(user, pageContext = {}, { sameWeb = true } = {}) {
  if (!user) return false;
  // Page context carries the login as a bare UPN on SharePoint Online
  // ("joe@tenant.com") while SP.User.LoginName is the full claim
  // ("i:0#.f|membership|joe@tenant.com"), so a match is either form. A
  // mismatch is not an answer: fall through to the other identifiers.
  const login = String(pageContext?.userLoginName || '').trim().toLowerCase();
  const claim = String(user.LoginName || '').trim().toLowerCase();
  if (login && claim
      && (login === claim || claim.endsWith(`|${login}`) || login.endsWith(`|${claim}`))) {
    return true;
  }
  const email = String(pageContext?.userEmail || '').trim().toLowerCase();
  const userEmail = String(user.Email || user.UserPrincipalName || '').trim().toLowerCase();
  if (email && userEmail) return email === userEmail;
  const id = Number(pageContext?.userId);
  if (sameWeb && Number.isFinite(id) && id > 0) return id === Number(user.Id);
  return false;
}
// Exported: the SP Workbench upload dialog shares these specs so the two
// implementations can't drift on availability rules or internal names.
export const FILE_METADATA_SPECS = Object.freeze([
  { key: 'title', label: 'Title', internalName: 'Title', types: ['Text'] },
  { key: 'description', label: 'Description', internalName: '_ExtendedDescription', types: ['Note', 'Text'] },
  { key: 'docVersion', label: 'DocVersion', internalName: 'DocVersion', types: ['Text'] },
]);

export { SpFileError };

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

function browserTypeForFileName(fileName) {
  const name = String(fileName || '');
  if (/\.html?$/i.test(name)) return 'html';
  if (/\.(?:md|markdown)$/i.test(name)) return 'markdown';
  if (/\.css$/i.test(name)) return 'css';
  if (/\.js$/i.test(name)) return 'javascript';
  if (/\.json$/i.test(name)) return 'json';
  if (/\.csv$/i.test(name)) return 'csv';
  if (/\.txt$/i.test(name)) return 'text';
  return '';
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

  async function listFolder(
    serverRelativePath,
    { webUrl: targetWebUrl = '', purpose = 'code', additionalTypes = [] } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const endpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')`
      + '?$select=Name,ServerRelativeUrl,'
      + 'Folders/Name,Folders/ServerRelativeUrl,'
      + 'Files/Name,Files/ServerRelativeUrl,Files/Length,Files/TimeLastModified'
      + '&$expand=Folders,Files';
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
      .map((item) => {
        const fileType = fileTypeForFileName(item.Name, additionalTypes);
        return {
          kind: 'file',
          name: String(item.Name || ''),
          pane: fileType?.pane || '',
          fileType,
          browserType: browserTypeForFileName(item.Name),
          serverRelativeUrl: checkedPath(item.ServerRelativeUrl, rootPath),
          length: Number(item.Length) || 0,
          modified: item.TimeLastModified || '',
        };
      })
      .filter((item) =>
        item.name && (purpose === 'browser' ? item.browserType : item.fileType))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
    return {
      path: checkedPath(data.ServerRelativeUrl || path, rootPath),
      rootPath,
      folders,
      files,
    };
  }

  async function readTextFile(
    serverRelativePath,
    { webUrl: targetWebUrl = '', additionalTypes = [] } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const fileType = fileTypeForFileName(path, additionalTypes);
    if (!fileType) {
      throw new SpFileError(
        'That file type is not supported for SharePoint import.',
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
      pane: fileType.pane,
      fileType,
      text,
      serverRelativeUrl: path,
    };
  }

  // Raw binary read, no file-type restriction (readTextFile's cap and type
  // gate are for the code-import path; this is for arbitrary asset bytes —
  // e.g. a page-copy banner or web-part image). content-length is checked
  // before the body is read so an oversized file never has to be downloaded
  // to be rejected; the arrayBuffer length is checked again in case the
  // header was missing or wrong.
  async function readFileBytes(
    serverRelativePath,
    { webUrl: targetWebUrl = '', maxBytes = 50 * 1024 * 1024 } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')/$value`;
    const response = await request(endpoint);
    await requireOk(response, 'Could not download the SharePoint file', 'read');
    const declaredLength = Number(response.headers.get('content-length')) || 0;
    if (declaredLength > maxBytes) {
      throw new SpFileError(
        'The selected SharePoint file is larger than the transfer limit.',
        { code: 'too-large' },
      );
    }
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength > maxBytes) {
      throw new SpFileError(
        'The selected SharePoint file is larger than the transfer limit.',
        { code: 'too-large' },
      );
    }
    return {
      bytes,
      length: bytes.byteLength,
      contentType: response.headers.get('content-type') || '',
      serverRelativeUrl: path,
    };
  }

  // Check-out policy of the destination library, plus the destination file's
  // current check-out state. The file state is read whether or not the
  // library forces a check-out: a file held by someone else refuses the write
  // in any library, and one the operator holds stays checked out after it.
  // Never throws: a failed probe reports known:false and leaves SharePoint the
  // authority, so an unrelated probe failure cannot cost the operator their
  // export.
  async function checkOutState({ webUrl, hostWebUrl, rootPath, libraryId, filePath, ctx }) {
    const state = {
      required: false,
      known: false,
      checkedOut: false,
      checkedOutByCurrentUser: false,
      checkedOutBy: '',
      reason: '',
    };
    try {
      const policyResponse = await request(
        `${webUrl}/_api/web/lists(guid'${libraryId}')?$select=ForceCheckout`,
        { headers: { Accept: ACCEPT_JSON } },
      );
      await requireOk(
        policyResponse,
        'Could not read the destination library check-out policy',
        'checkout-policy',
      );
      const list = unwrapJson(await policyResponse.json()) || {};
      state.required = Boolean(list.ForceCheckout ?? list.forceCheckout);
      state.known = true;
      if (!filePath) return state;

      const path = checkedPath(filePath, rootPath);
      const fileResponse = await request(
        `${webUrl}/_api/web/GetFileByServerRelativePath(`
        + `decodedUrl='${odataPathLiteral(path)}')`
        + '?$select=CheckOutType,CheckedOutByUser/Id,CheckedOutByUser/Title,'
        + 'CheckedOutByUser/LoginName,CheckedOutByUser/Email'
        + '&$expand=CheckedOutByUser',
        { headers: { Accept: ACCEPT_JSON } },
      );
      await requireOk(
        fileResponse,
        'Could not read the destination file check-out state',
        'checkout-state',
      );
      const file = unwrapJson(await fileResponse.json()) || {};
      state.checkedOut = isCheckedOut(file.CheckOutType ?? file.checkOutType);
      if (state.checkedOut) {
        const user = file.CheckedOutByUser || file.checkedOutByUser || null;
        state.checkedOutBy = String(user?.Title || user?.LoginName || '').trim();
        state.checkedOutByCurrentUser = isCheckedOutByCurrentUser(
          user, ctx?.pageContext, { sameWeb: webUrl === hostWebUrl },
        );
      }
    } catch (error) {
      // The library policy may already be known; only the state read failed.
      state.known = false;
      state.reason = String(error?.message || error);
    }
    return state;
  }

  // The three check-out verbs share one shape: a digest-bearing POST to a
  // method on the file, with the same 403 force-refresh retry as every other
  // write here.
  async function postFileMethod(webUrl, path, method, fallback, code) {
    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')/${method}`;
    const attempt = async (forceDigest) => {
      const digest = await getDigest({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: 'POST',
        headers: { Accept: ACCEPT_JSON, 'X-RequestDigest': digest },
      });
    };
    let response = await attempt(false);
    if (response.status === 403) response = await attempt(true);
    await requireOk(response, fallback, code);
  }

  // Check a file out ahead of an overwrite. Required by libraries that set
  // ForceCheckout. The overwriting upload does not end the check-out — see
  // checkInFile(), which the caller runs once the file and its metadata land.
  async function checkOutFile(serverRelativePath, { webUrl: targetWebUrl = '' } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    await postFileMethod(
      webUrl, path, 'CheckOut()', 'Could not check out the SharePoint file', 'checkout',
    );
    return { serverRelativeUrl: path };
  }

  // Check a file back in, but only if SharePoint still reports it checked out.
  // Whether an overwriting upload leaves the check-out standing is the
  // server's business (and a new file in a ForceCheckout library is born
  // checked out), so the state is read rather than assumed: CheckIn() on a
  // file that is not checked out is an error, and skipping it on one that is
  // leaves the operator's work invisible to everyone else.
  async function checkInFile(
    serverRelativePath,
    { comment = '', checkInType = 0, webUrl: targetWebUrl = '' } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const stateResponse = await request(
      `${webUrl}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')?$select=CheckOutType`,
      { headers: { Accept: ACCEPT_JSON } },
    );
    await requireOk(
      stateResponse, 'Could not read the file check-out state', 'checkout-state',
    );
    const file = unwrapJson(await stateResponse.json()) || {};
    if (!isCheckedOut(file.CheckOutType ?? file.checkOutType)) {
      return { serverRelativeUrl: path, checkedIn: false };
    }
    // SP.CheckinType: 0 = minor, 1 = major, 2 = overwrite.
    const type = [0, 1, 2].includes(Number(checkInType)) ? Number(checkInType) : 0;
    const safeComment = String(comment || '').slice(0, 1023);
    await postFileMethod(
      webUrl,
      path,
      `CheckIn(comment='${odataPathLiteral(safeComment)}',checkintype=${type})`,
      'Could not check in the SharePoint file',
      'checkin',
    );
    return { serverRelativeUrl: path, checkedIn: true };
  }

  // Discard a check-out. SharePoint reverts the file to its last checked-in
  // version, so this is only offered before anything has been uploaded.
  async function undoCheckOutFile(serverRelativePath, { webUrl: targetWebUrl = '' } = {}) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    await postFileMethod(
      webUrl, path, 'UndoCheckOut()', 'Could not discard the check-out', 'checkout-undo',
    );
    return { serverRelativeUrl: path };
  }

  async function inspectFileMetadata(
    folderPath,
    { filePath = '', webUrl: targetWebUrl = '' } = {},
  ) {
    const { ctx, webUrl, rootPath, hostWebUrl } = webInfo(targetWebUrl);
    const folder = checkedPath(folderPath, rootPath);
    const libraryEndpoint = `${webUrl}/_api/web/GetFolderByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(folder)}')`
      + '?$select=ListItemAllFields/ParentList/Id'
      + '&$expand=ListItemAllFields,ListItemAllFields/ParentList';
    const libraryResponse = await request(libraryEndpoint, {
      headers: { Accept: ACCEPT_JSON },
    });
    const libraryData = libraryResponse.ok
      ? (unwrapJson(await libraryResponse.json()) || {})
      : {};
    let libraryId = String(
      libraryData.ListItemAllFields?.ParentList?.Id
      || libraryData.ListItemAllFields?.ParentList?.ID
      || '',
    ).replace(/[{}]/g, '').trim();

    // A document library's root folder has no corresponding list item, so
    // ListItemAllFields.ParentList can be empty even though files may be saved
    // there. Resolve the list directly from its root URL in that case.
    if (!LIBRARY_GUID.test(libraryId)) {
      const rootLibraryEndpoint = `${webUrl}/_api/web/GetList(@listUrl)`
        + `?@listUrl='${odataPathLiteral(folder)}'&$select=Id`;
      const rootLibraryResponse = await request(rootLibraryEndpoint, {
        headers: { Accept: ACCEPT_JSON },
      });
      await requireOk(
        rootLibraryResponse,
        'Could not resolve the destination SharePoint library',
        'metadata-library',
      );
      const rootLibraryData = unwrapJson(await rootLibraryResponse.json()) || {};
      libraryId = String(rootLibraryData.Id || rootLibraryData.ID || '')
        .replace(/[{}]/g, '')
        .trim();
    }
    if (!LIBRARY_GUID.test(libraryId)) {
      throw new SpFileError(
        'SharePoint did not identify the destination document library.',
        { code: 'metadata-library' },
      );
    }

    // Probed before the fields read so that a library whose columns cannot be
    // inspected still reports its check-out policy: the export can proceed
    // without metadata, but never without the check-out the library forces.
    const checkout = await checkOutState({
      webUrl, hostWebUrl, rootPath, libraryId, filePath, ctx,
    });

    try {
      return {
        fields: await inspectFields(webUrl, rootPath, libraryId, filePath),
        checkout,
      };
    } catch (error) {
      if (error && typeof error === 'object') error.checkout = checkout;
      throw error;
    }
  }

  async function inspectFields(webUrl, rootPath, libraryId, filePath) {
    const fieldsEndpoint = `${webUrl}/_api/web/lists(guid'${libraryId}')/Fields`
      + '?$select=InternalName,EntityPropertyName,Title,TypeAsString,ReadOnlyField,Hidden';
    const fieldsResponse = await request(fieldsEndpoint, {
      headers: { Accept: ACCEPT_JSON },
    });
    await requireOk(
      fieldsResponse,
      'Could not inspect the destination library metadata fields',
      'metadata-fields',
    );
    const fieldsData = unwrapJson(await fieldsResponse.json()) || {};
    const libraryFields = resultArray(fieldsData.value || fieldsData);

    const fields = {};
    for (const spec of FILE_METADATA_SPECS) {
      const match = libraryFields.find((field) =>
        String(field.InternalName || '').toLowerCase()
        === spec.internalName.toLowerCase());
      let reason = '';
      if (!match) reason = `${spec.internalName} is not available in this library.`;
      else if (match.ReadOnlyField) reason = `${spec.internalName} is read-only.`;
      else if (match.Hidden) reason = `${spec.internalName} is hidden in this library.`;
      else if (!spec.types.includes(String(match.TypeAsString || ''))) {
        reason = `${spec.internalName} is not a supported text field.`;
      }
      fields[spec.key] = {
        label: spec.label,
        internalName: match?.InternalName || spec.internalName,
        entityPropertyName:
          match?.EntityPropertyName || match?.InternalName || spec.internalName,
        available: !reason,
        reason,
        value: '',
      };
    }

    if (filePath) {
      const path = checkedPath(filePath, rootPath);
      const selected = Object.values(fields)
        .filter((field) => field.available)
        .map((field) => field.entityPropertyName);
      if (selected.length) {
        const valuesEndpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(`
          + `decodedUrl='${odataPathLiteral(path)}')/ListItemAllFields`
          + `?$select=${selected.map(encodeURIComponent).join(',')}`;
        const valuesResponse = await request(valuesEndpoint, {
          headers: { Accept: ACCEPT_JSON },
        });
        await requireOk(
          valuesResponse,
          'Could not read the destination file metadata',
          'metadata-read',
        );
        const values = unwrapJson(await valuesResponse.json()) || {};
        for (const field of Object.values(fields)) {
          if (field.available) {
            field.value = String(
              values[field.entityPropertyName] ?? values[field.internalName] ?? '',
            );
          }
        }
      }
    }

    return fields;
  }

  async function writeFileMetadata(
    serverRelativePath,
    fields,
    values,
    { webUrl: targetWebUrl = '' } = {},
  ) {
    const { webUrl, rootPath } = webInfo(targetWebUrl);
    const path = checkedPath(serverRelativePath, rootPath);
    const formValues = Object.entries(fields || {})
      .filter(([key, field]) => field?.available && Object.hasOwn(values || {}, key))
      .map(([key, field]) => ({
        FieldName: field.internalName,
        FieldValue: String(values[key] ?? ''),
      }));
    if (!formValues.length) return { updated: [] };

    const endpoint = `${webUrl}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(path)}')`
      + '/ListItemAllFields/ValidateUpdateListItem';
    const update = async (forceDigest) => {
      const digest = await getDigest({ force: forceDigest, webUrl });
      return request(endpoint, {
        method: 'POST',
        headers: {
          Accept: ACCEPT_JSON,
          'Content-Type': 'application/json;odata=nometadata',
          'X-RequestDigest': digest,
        },
        // bNewDocumentUpdate makes this the tail of the upload rather than a
        // new version — and on a checked-out file SharePoint checks it in as
        // part of the update, recording checkInComment.
        body: JSON.stringify({
          formValues,
          bNewDocumentUpdate: true,
          checkInComment: CHECK_IN_COMMENT,
        }),
      });
    };

    let response = await update(false);
    if (response.status === 403) response = await update(true);
    await requireOk(response, 'Could not save the SharePoint file metadata', 'metadata-write');
    const data = unwrapJson(await response.json()) || {};
    const results = resultArray(data.value || data.ValidateUpdateListItem || data);
    const failures = results.filter((result) =>
      result.HasException || String(result.ErrorMessage || '').trim());
    if (failures.length) {
      const detail = failures
        .map((result) =>
          `${result.FieldName || 'Field'}: ${result.ErrorMessage || 'SharePoint rejected the value.'}`)
        .join(' ');
      throw new SpFileError(
        `SharePoint rejected the file metadata. ${detail}`,
        { code: 'metadata-write' },
      );
    }
    return { updated: formValues.map((value) => value.FieldName) };
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
    if (!safeName || safeName === '.' || safeName === '..' || /[\\/]/.test(safeName)) {
      throw new SpFileError(
        'Enter a file name without folder separators.',
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
      // SP.File as returned by the upload: a new file in a ForceCheckout
      // library is born checked out. Undefined when the server doesn't say.
      checkOutType: result.CheckOutType,
    };
  }

  return {
    webInfo,
    connectWeb,
    getDigest,
    listFolder,
    readTextFile,
    readFileBytes,
    checkOutFile,
    checkInFile,
    undoCheckOutFile,
    inspectFileMetadata,
    writeFileMetadata,
    writeTextFile,
  };
}

const defaultClient = createSpFilesClient();

export const getSpWebInfo = (webUrl) => defaultClient.webInfo(webUrl);
export const connectSpWeb = (webUrl) => defaultClient.connectWeb(webUrl);
export const getDigest = (options) => defaultClient.getDigest(options);
export const listFolder = (path, options) => defaultClient.listFolder(path, options);
export const readTextFile = (path, options) => defaultClient.readTextFile(path, options);
export const readFileBytes = (path, options) => defaultClient.readFileBytes(path, options);
export const checkOutFile = (path, options) => defaultClient.checkOutFile(path, options);
export const checkInFile = (path, options) => defaultClient.checkInFile(path, options);
export const undoCheckOutFile = (path, options) =>
  defaultClient.undoCheckOutFile(path, options);
export const inspectFileMetadata = (folder, options) =>
  defaultClient.inspectFileMetadata(folder, options);
export const writeFileMetadata = (path, fields, values, options) =>
  defaultClient.writeFileMetadata(path, fields, values, options);
export const writeTextFile = (folder, name, text, options) =>
  defaultClient.writeTextFile(folder, name, text, options);
