// Shared SharePoint OData plumbing: the accept header, error type, and
// response-shape helpers used by every module that talks to /_api endpoints
// (sp-files.js for document transfer, src/workbench/ for the inspector).
// No DOM, no storage, no context — pure request/response helpers.

export const ACCEPT_JSON = 'application/json;odata=nometadata';

export class SpFileError extends Error {
  constructor(message, { code = 'sharepoint', status = 0, cause } = {}) {
    super(message, { cause });
    this.name = 'SpFileError';
    this.code = code;
    this.status = status;
  }
}

export function odataPathLiteral(value) {
  // Encode URL-significant characters such as # and %, but retain OData's
  // doubled-apostrophe escaping inside the surrounding string literal.
  return encodeURIComponent(String(value)).replaceAll("'", "''");
}

export function resultArray(value) {
  if (Array.isArray(value)) return value;
  if (Array.isArray(value?.results)) return value.results;
  return [];
}

export function unwrapJson(data) {
  return data?.d?.GetContextWebInformation
    || data?.GetContextWebInformation
    || data?.d
    || data;
}

export async function responseMessage(response) {
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

export async function requireOk(response, fallback, code) {
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
