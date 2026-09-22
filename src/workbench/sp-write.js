// Workbench write client — the POST-side companion to sp-rest.js.
//
// sp-rest.js is GET-only by design; this module layers the two write
// operations Tier 2 needs (ValidateUpdateListItem, binary AddUsingPath)
// on the digest cache that sp-files.js already owns. It never modifies
// sp-files.js — it only imports the exported getDigest, so the pad and the
// workbench share one digest per web without sharing any other state.
//
// Off SharePoint every write routes to a mock writer instead, so the
// editors stay exercisable (and testable) with zero network. Mock writes
// are recorded on globalThis.__DCSPAD_WB_WRITES__ for tests.

import { getDigest, isCheckedOut } from '../sp-files.js';
import {
  ACCEPT_JSON, SpFileError, odataPathLiteral, resultArray, unwrapJson, requireOk,
} from '../sp-odata.js';

// Single-request AddUsingPath ceiling. SharePoint accepts far larger files
// through chunked uploads; that is deliberately out of scope for v1.
export const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

// Sequential ids for the mock's AddValidateUpdateItemUsingPath and
// ensureuser answers — module-level so ids stay unique across every mock
// writer call in one page load (a fresh page/test reloads and resets it).
let mockNewItemId = 1000;
let mockEnsuredUserId = 9000;

// Built-in mock writer: record the call, answer with the minimal success
// shape each endpoint's caller parses. Exported so mock-data.js's stateful
// writer can fall back to it for URLs it doesn't itself register.
export function defaultMockWriter(url, body, contentType, headers) {
  const writes = (globalThis.__DCSPAD_WB_WRITES__ ||= []);
  writes.push({ url, body, contentType, headers });
  const lower = String(url).toLowerCase();
  // Stage 1b-b: list item DATA import writes. Checked ahead of the
  // 'validateupdatelistitem' substring test below — 'addvalidateupdateitem
  // usingpath' never matches it (no "list" between "validate" and "update"),
  // but keeping the more specific endpoints first avoids any future drift.
  if (lower.includes('addvalidateupdateitemusingpath')) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const formValues = Array.isArray(data?.formValues) ? data.formValues : [];
    const id = String(mockNewItemId++);
    return {
      value: [
        ...formValues.map((fv) => ({
          FieldName: fv.FieldName, FieldValue: fv.FieldValue, HasException: false, ErrorMessage: null,
        })),
        { FieldName: 'Id', FieldValue: id, HasException: false, ErrorMessage: null },
      ],
    };
  }
  if (lower.includes('/ensureuser')) {
    let data = {};
    try { data = JSON.parse(body); } catch { /* keep {} */ }
    const logon = String(data?.logonName || '');
    const email = logon.includes('@') ? logon : `${logon.replace(/[^a-z0-9.]+/gi, '.')}@mock.local`;
    return {
      Id: mockEnsuredUserId++,
      Title: logon,
      LoginName: `i:0#.f|membership|${email.toLowerCase()}`,
      Email: email,
    };
  }
  if (lower.includes('attachmentfiles/add(')) {
    const name = /attachmentfiles\/add\(filename='([^']*)'\)/.exec(lower)?.[1] || 'file';
    let decoded = name;
    try { decoded = decodeURIComponent(name); } catch { /* keep raw */ }
    return { FileName: decoded, ServerRelativeUrl: `/mock/attachments/${decoded}` };
  }
  if (lower.includes('validateupdatelistitem')) {
    let formValues = [];
    try { formValues = JSON.parse(body)?.formValues || []; } catch { /* keep [] */ }
    return {
      value: formValues.map((fv) => ({
        FieldName: fv.FieldName, HasException: false, ErrorMessage: null,
      })),
    };
  }
  if (lower.includes('addusingpath')) {
    const name = /addusingpath\(decodedurl='([^']*)'/.exec(lower)?.[1] || 'file';
    const folder = /getfolderbyserverrelativepath\(decodedurl='([^']*)'/.exec(lower)?.[1] || '';
    return { ServerRelativeUrl: `${decodeURIComponent(folder)}/${decodeURIComponent(name)}` };
  }
  return { ok: true };
}

export function createSpWriteClient({
  client,                                  // the workbench sp-rest client
  fetchImpl = (...args) => fetch(...args),
  mockWriter = null,
} = {}) {
  const isMock = () => !client.context().live;

  // headers merges over the base set (Accept/Content-Type/digest) so a
  // caller can add X-HTTP-Method/IF-MATCH (mergeJson) without repeating them.
  async function post(url, { body, contentType = 'application/json;odata=nometadata', headers = {} } = {}, {
    fallback = 'SharePoint write failed', code = 'write',
  } = {}) {
    if (isMock()) {
      return structuredClone((mockWriter || defaultMockWriter)(url, body, contentType, headers));
    }
    const attempt = async (forceDigest) => {
      const digest = await getDigest({ force: forceDigest, webUrl: client.webUrl() });
      try {
        return await fetchImpl(url, {
          method: 'POST',
          credentials: 'same-origin',
          headers: {
            Accept: ACCEPT_JSON,
            'Content-Type': contentType,
            'X-RequestDigest': digest,
            ...headers,
          },
          body,
        });
      } catch (cause) {
        throw new SpFileError(
          `Could not reach SharePoint (${cause.message || cause}).`,
          { code: 'network', cause },
        );
      }
    };
    let response = await attempt(false);
    if (response.status === 403) response = await attempt(true);
    // Throttling: one retry honouring Retry-After, capped like sp-rest.js's
    // read-side retry — a schema apply can throw dozens of writes at a list
    // in a row and SPO's list-write throttle is real.
    if (response.status === 429 || response.status === 503) {
      const after = Number(response.headers.get('Retry-After')) || 2;
      await new Promise((r) => setTimeout(r, Math.min(after, 30) * 1000));
      response = await attempt(false);
    }
    await requireOk(response, fallback, code);
    try { return unwrapJson(await response.json()) || {}; }
    catch { return {}; }   // a successful write may return no JSON body
  }

  // ValidateUpdateListItem against a list item, addressed either by
  // { listId, itemId } or by { fileServerRelativeUrl }. formValues is the
  // [{ FieldName, FieldValue }] array SharePoint expects (strings only —
  // see field-editor.js toFormValue for the per-type conventions).
  // Throws SpFileError('metadata-write') carrying err.fieldErrors
  // ({ FieldName: message }) so forms can map failures onto editors.
  async function validateUpdateListItem(
    pathKind, formValues, { newDocumentUpdate = false, checkInComment = '' } = {},
  ) {
    if (!Array.isArray(formValues) || !formValues.length) return { updated: [] };
    const base = `${client.webUrl()}/_api/web`;
    const endpoint = pathKind.fileServerRelativeUrl
      ? `${base}/GetFileByServerRelativePath(`
        + `decodedUrl='${odataPathLiteral(pathKind.fileServerRelativeUrl)}')`
        + '/ListItemAllFields/ValidateUpdateListItem'
      : `${base}/lists(guid'${pathKind.listId}')/items(${Number(pathKind.itemId)})`
        + '/ValidateUpdateListItem';
    const data = await post(endpoint, {
      // With bNewDocumentUpdate SharePoint checks a checked-out file in as
      // part of the update; checkInComment is what it records when it does.
      body: JSON.stringify({
        formValues,
        bNewDocumentUpdate: Boolean(newDocumentUpdate),
        ...(newDocumentUpdate && checkInComment ? { checkInComment } : {}),
      }),
    }, { fallback: 'Could not save the item metadata', code: 'metadata-write' });

    const results = resultArray(data.value || data.ValidateUpdateListItem || data);
    const failures = results.filter((result) =>
      result.HasException || String(result.ErrorMessage || '').trim());
    if (failures.length) {
      const fieldErrors = {};
      for (const failure of failures) {
        fieldErrors[failure.FieldName || ''] =
          failure.ErrorMessage || 'SharePoint rejected the value.';
      }
      const detail = failures
        .map((f) => `${f.FieldName || 'Field'}: ${f.ErrorMessage || 'SharePoint rejected the value.'}`)
        .join(' ');
      const err = new SpFileError(
        `SharePoint rejected the metadata. ${detail}`,
        { code: 'metadata-write' },
      );
      err.fieldErrors = fieldErrors;
      throw err;
    }
    return { updated: formValues.map((fv) => fv.FieldName) };
  }

  // AddValidateUpdateItemUsingPath — create a list item (or, with
  // underlyingObjectType 1, a folder) under folderPath, in one call that both
  // creates the row and sets its fields. Unlike ValidateUpdateListItem
  // (an update against an item that already exists), a rejected field here
  // aborts the WHOLE create — SharePoint never returns an Id alongside a
  // HasException row — so every failure throws the same SpFileError
  // ('metadata-write', err.fieldErrors) validateUpdateListItem throws; the
  // caller (list-data-apply.js) retries once without the rejected fields,
  // the same way importListData does.
  async function addValidateUpdateItem(listId, {
    folderPath, formValues, underlyingObjectType = 0, leafName,
  } = {}) {
    const endpoint = `${client.webUrl()}/_api/web/lists(guid'${listId}')/AddValidateUpdateItemUsingPath`;
    const data = await post(endpoint, {
      body: JSON.stringify({
        listItemCreateInfo: {
          FolderPath: { DecodedUrl: folderPath },
          UnderlyingObjectType: underlyingObjectType,
          ...(leafName ? { LeafName: { DecodedUrl: leafName } } : {}),
        },
        formValues,
        bNewDocumentUpdate: false,
      }),
    }, { fallback: 'Could not create the item', code: 'metadata-write' });

    const results = resultArray(data.value || data.AddValidateUpdateItemUsingPath || data);
    const failures = results.filter((result) => result.HasException || String(result.ErrorMessage || '').trim());
    if (failures.length) {
      const fieldErrors = {};
      for (const failure of failures) {
        fieldErrors[failure.FieldName || ''] = failure.ErrorMessage || 'SharePoint rejected the value.';
      }
      const detail = failures
        .map((f) => `${f.FieldName || 'Field'}: ${f.ErrorMessage || 'SharePoint rejected the value.'}`)
        .join(' ');
      const err = new SpFileError(`SharePoint rejected the metadata. ${detail}`, { code: 'metadata-write' });
      err.fieldErrors = fieldErrors;
      throw err;
    }
    const idRow = results.find((r) => /^id$/i.test(r.FieldName || ''));
    const id = idRow ? Number(idRow.FieldValue) : null;
    if (!id) {
      throw new SpFileError('SharePoint did not return the new item’s id.', { code: 'metadata-write' });
    }
    return { id };
  }

  // web/ensureuser — resolve an email or login to a target-web principal,
  // cached per login for this client's lifetime (one spWrite instance per
  // target connection, so the cache never survives a Connect to a different
  // web). Returns null for an empty logon; throws on a genuine write failure
  // — the caller (list-data-apply.js) decides what to do with that (SPUtils'
  // fallback to a synthetic claims key).
  const ensureUserCache = new Map();
  async function ensureUser(logonName) {
    const key = String(logonName || '').toLowerCase();
    if (!key) return null;
    if (ensureUserCache.has(key)) return ensureUserCache.get(key);
    const promise = (async () => {
      const data = await post(`${client.webUrl()}/_api/web/ensureuser`, {
        body: JSON.stringify({ logonName }),
      }, { fallback: 'Could not resolve the user', code: 'write' });
      return { id: data.Id, loginName: data.LoginName || '', email: data.Email || '', title: data.Title || '' };
    })();
    ensureUserCache.set(key, promise);
    try {
      return await promise;
    } catch (err) {
      ensureUserCache.delete(key);
      throw err;
    }
  }

  // Attach a file to an existing item. `bytes` is whatever fetch accepts as
  // a body (ArrayBuffer/Blob/typed array) — the caller (list-data-apply.js)
  // reads it from the source attachment's URL first; this only uploads it.
  async function addAttachment(listId, itemId, fileName, bytes) {
    const endpoint = `${client.webUrl()}/_api/web/lists(guid'${listId}')/items(${Number(itemId)})`
      + `/AttachmentFiles/add(FileName='${odataPathLiteral(fileName)}')`;
    const data = await post(endpoint, { body: bytes, contentType: 'application/octet-stream' },
      { fallback: 'Could not add the attachment', code: 'write' });
    return { fileName, serverRelativeUrl: data.ServerRelativeUrl || '' };
  }

  // Binary (or text) upload into a folder. data may be an ArrayBuffer,
  // typed array, Blob, or string — whatever fetch accepts as a body.
  async function uploadFile(folderServerRelativeUrl, fileName, data, { overwrite = false } = {}) {
    const safeName = String(fileName || '').trim();
    if (!safeName || safeName === '.' || safeName === '..' || /[\\/]/.test(safeName)) {
      throw new SpFileError(
        'Enter a file name without folder separators.',
        { code: 'invalid-name' },
      );
    }
    const size = data?.byteLength ?? data?.size ?? (typeof data === 'string' ? data.length : 0);
    if (size > MAX_UPLOAD_BYTES) {
      throw new SpFileError(
        `The file is larger than the ${Math.round(MAX_UPLOAD_BYTES / (1024 * 1024))} MB upload limit.`,
        { code: 'too-large' },
      );
    }
    const folder = String(folderServerRelativeUrl || '/').replace(/\/+$/, '') || '/';
    const endpoint = `${client.webUrl()}/_api/web/GetFolderByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(folder)}')/Files/AddUsingPath(`
      + `decodedUrl='${odataPathLiteral(safeName)}',overwrite=${overwrite ? 'true' : 'false'})`;
    const result = await post(endpoint, {
      body: data,
      contentType: 'application/octet-stream',
    }, { fallback: 'Could not upload the file', code: 'write' });
    return {
      fileName: safeName,
      serverRelativeUrl: result.ServerRelativeUrl || `${folder === '/' ? '' : folder}/${safeName}`,
      // SP.File as returned by the upload: a new file in a ForceCheckout
      // library is born checked out, and an overwrite leaves a check-out
      // standing. Undefined when the server (or the mock) doesn't say.
      checkOutType: result.CheckOutType,
    };
  }

  // Check a file out ahead of an overwrite, for libraries that set
  // ForceCheckout. The overwriting upload does not end the check-out, so the
  // caller pairs this with checkInFile() — the same contract as the pad's
  // export (sp-files.js checkOutFile / checkInFile).
  async function checkOutFile(fileServerRelativeUrl) {
    const endpoint = `${client.webUrl()}/_api/web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(fileServerRelativeUrl)}')/CheckOut()`;
    await post(endpoint, { body: '' },
      { fallback: 'Could not check out the file', code: 'checkout' });
    return { serverRelativeUrl: fileServerRelativeUrl };
  }

  // Check a file back in, but only if SharePoint still reports it checked
  // out: CheckIn() on a file that isn't is an error, and whether an upload
  // left one standing is the server's business, so the state is read rather
  // than assumed. The mock has no such state; it records the call.
  async function checkInFile(fileServerRelativeUrl, { comment = '' } = {}) {
    const file = `web/GetFileByServerRelativePath(`
      + `decodedUrl='${odataPathLiteral(fileServerRelativeUrl)}')`;
    if (!isMock()) {
      const state = await client.get(file, { select: 'CheckOutType' });
      if (!isCheckedOut(state?.CheckOutType)) {
        return { serverRelativeUrl: fileServerRelativeUrl, checkedIn: false };
      }
    }
    const safeComment = String(comment || '').slice(0, 1023);
    // SP.CheckinType 0 = minor version, matching the pad's export.
    await post(
      `${client.webUrl()}/_api/${file}/CheckIn(`
        + `comment='${odataPathLiteral(safeComment)}',checkintype=0)`,
      { body: '' },
      { fallback: 'Could not check in the file', code: 'checkin' },
    );
    return { serverRelativeUrl: fileServerRelativeUrl, checkedIn: true };
  }

  // Create a subfolder. '#' and '%' are legal in modern SPO names (hence
  // ResourcePath addressing); the rejected set is what SharePoint itself
  // refuses: " * : < > ? / \ | plus leading/trailing dots.
  async function createFolder(parentServerRelativeUrl, name) {
    const clean = String(name || '').trim();
    if (!clean || /["*:<>?/\\|]/.test(clean) || clean.startsWith('.') || clean.endsWith('.')) {
      throw new SpFileError(
        'Folder names cannot contain " * : < > ? / \\ | or start or end with a dot.',
        { code: 'invalid-name' },
      );
    }
    // Names SharePoint reserves outright (legacy device names, _vti_).
    if (/^(CON|PRN|AUX|NUL|COM\d|LPT\d)(\..*)?$/i.test(clean) || /_vti_/i.test(clean)) {
      throw new SpFileError(
        'That folder name is reserved by SharePoint.',
        { code: 'invalid-name' },
      );
    }
    const parent = String(parentServerRelativeUrl || '/').replace(/\/+$/, '') || '';
    const path = `${parent}/${clean}`;
    const endpoint = `${client.webUrl()}/_api/web/Folders/AddUsingPath(`
      + `decodedUrl='${odataPathLiteral(path)}')`;
    await post(endpoint, { body: '' },
      { fallback: 'Could not create the folder', code: 'write' });
    return { name: clean, serverRelativeUrl: path };
  }

  // Generic JSON POST against a /_api-relative path (group membership ops
  // and other small writes). Returns the parsed response body. `contentType`
  // is an escape hatch for the schema executor's verbose-odata fallback
  // (list-schema-apply.js) — every other caller leaves it unset and gets
  // post()'s own nometadata default.
  async function postJson(path, body = {}, {
    fallback = 'SharePoint write failed', code = 'write', headers = {}, contentType,
  } = {}) {
    const url = `${client.webUrl()}/_api/${String(path).replace(/^\/+/, '')}`;
    return post(url, {
      body: JSON.stringify(body), headers, ...(contentType ? { contentType } : {}),
    }, { fallback, code });
  }

  // SharePoint's REST MERGE: a POST carrying X-HTTP-Method: MERGE and
  // IF-MATCH: * (unconditional — the caller isn't tracking an etag). Used
  // for every partial-property update (list settings, field flags, view
  // properties) so a plan step never has to know which properties want a
  // full PUT versus a merge; on SharePoint everything here is a merge.
  async function mergeJson(path, body = {}, {
    fallback = 'SharePoint write failed', code = 'write', headers = {}, contentType,
  } = {}) {
    const url = `${client.webUrl()}/_api/${String(path).replace(/^\/+/, '')}`;
    return post(url, {
      body: JSON.stringify(body),
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*', ...headers },
      ...(contentType ? { contentType } : {}),
    }, { fallback, code });
  }

  return {
    validateUpdateListItem, addValidateUpdateItem, ensureUser, addAttachment,
    uploadFile, checkOutFile, checkInFile, createFolder,
    postJson, mergeJson, isMock,
  };
}
