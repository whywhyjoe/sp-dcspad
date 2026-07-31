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

import { getDigest } from '../sp-files.js';
import {
  ACCEPT_JSON, SpFileError, odataPathLiteral, resultArray, unwrapJson, requireOk,
} from '../sp-odata.js';

// Single-request AddUsingPath ceiling. SharePoint accepts far larger files
// through chunked uploads; that is deliberately out of scope for v1.
export const MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

// Built-in mock writer: record the call, answer with the minimal success
// shape each endpoint's caller parses.
function defaultMockWriter(url, body, contentType) {
  const writes = (globalThis.__DCSPAD_WB_WRITES__ ||= []);
  writes.push({ url, body, contentType });
  const lower = String(url).toLowerCase();
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

  async function post(url, { body, contentType = 'application/json;odata=nometadata' } = {}, {
    fallback = 'SharePoint write failed', code = 'write',
  } = {}) {
    if (isMock()) {
      return structuredClone((mockWriter || defaultMockWriter)(url, body, contentType));
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
  async function validateUpdateListItem(pathKind, formValues, { newDocumentUpdate = false } = {}) {
    if (!Array.isArray(formValues) || !formValues.length) return { updated: [] };
    const base = `${client.webUrl()}/_api/web`;
    const endpoint = pathKind.fileServerRelativeUrl
      ? `${base}/GetFileByServerRelativePath(`
        + `decodedUrl='${odataPathLiteral(pathKind.fileServerRelativeUrl)}')`
        + '/ListItemAllFields/ValidateUpdateListItem'
      : `${base}/lists(guid'${pathKind.listId}')/items(${Number(pathKind.itemId)})`
        + '/ValidateUpdateListItem';
    const data = await post(endpoint, {
      body: JSON.stringify({ formValues, bNewDocumentUpdate: Boolean(newDocumentUpdate) }),
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
    };
  }

  // Generic JSON POST against a /_api-relative path (group membership ops
  // and other small writes). Returns the parsed response body.
  async function postJson(path, body = {}, { fallback = 'SharePoint write failed', code = 'write' } = {}) {
    const url = `${client.webUrl()}/_api/${String(path).replace(/^\/+/, '')}`;
    return post(url, { body: JSON.stringify(body) }, { fallback, code });
  }

  return { validateUpdateListItem, uploadFile, postJson, isMock };
}
