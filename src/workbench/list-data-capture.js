// List item DATA — capture (source client). Stage 1b-a.
//
// Port of SPUtils' exportListData (utilities/dcspad-sp-utilities.js:2441-2534)
// into the Workbench's own I/O layer: reads items (with folders and
// attachment metadata alongside them, in one pass), resolves the person and
// lookup fields the way exportListData does, and hands the raw rows to
// list-data.js's buildDataDoc() to shape the document. Nothing here decides
// per-type FieldValue strings — that's list-data.js's job, for the (future)
// import/apply module.
//
// Attachment BYTES are embedded unconditionally (unlike SPUtils' opt-in
// includeAttachments), matching v1's `{ name, base64 }` shape, under a
// per-file (10 MB) and running total (50 MB) cap — over either, the entry
// degrades to a `{ name, url }` link (the `url` key is additive/v2) plus one
// warning naming the file. See captureAttachment()/fetchAttachmentBytes().

import {
  LOOKUP_TYPES, USER_TYPES, buildDataDoc,
} from './list-data.js';
import { SCHEMA_KIND } from './list-schema.js';
import { captureListSchema } from './list-schema-capture.js';
import { APP_BUILD_INFO } from '../build-info.js';
import { isExpiredSession } from './denied.js';
import { SpFileError } from '../sp-odata.js';

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;

// Default page-read ceiling for a data export — the opt-in cap sp-rest.js
// reserves for this one caller (getAll's default stays 5000 for everyone
// else). A caller may pass a smaller `maxItems`; it can never exceed this.
const DATA_CAP = 100000;

// Attachment embedding caps — per file and across the whole export. Over
// either cap the attachment degrades to its v2 `{ name, url }` link-only
// shape (see fetchAttachmentBytes/captureAttachment below) instead of
// growing the document without bound.
const ATTACHMENT_FILE_CAP = 10 * 1024 * 1024;
const ATTACHMENT_TOTAL_CAP = 50 * 1024 * 1024;

const toIdList = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]).filter((x) => x != null);

// Reads one attachment's bytes over plain fetch (same tenant, same-origin
// cookies) — used here to embed export bytes, and by list-data-apply.js's
// copy-mode fallback when an item's `_attachments` entry has no base64 (an
// imported-from-file document, or one that degraded to url-only over a cap).
export async function fetchAttachmentBytes(client, serverRelativeUrl) {
  let origin = '';
  try { origin = new URL(client.webUrl()).origin; } catch { /* keep '' */ }
  const abs = /^https?:/i.test(serverRelativeUrl) ? serverRelativeUrl : `${origin}${serverRelativeUrl}`;
  let res;
  try {
    res = await fetch(abs, { credentials: 'same-origin' });
  } catch (cause) {
    throw new SpFileError(`Could not reach the source site (${cause.message || cause}).`, { code: 'network', cause });
  }
  if (!res.ok) {
    throw new SpFileError(`The source attachment could not be read (HTTP ${res.status}).`, {
      code: res.status === 401 ? 'auth' : 'network', status: res.status,
    });
  }
  return res.arrayBuffer();
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

// SPUtils' v1 shape is `{ name, base64 }` — that's what a small-enough
// attachment gets here. Over either cap it degrades to `{ name, url }`
// (the `url` key is additive/v2; a v1 reader never looked for it) plus one
// warning naming the file, per finding #3.
async function captureAttachment(client, rawFile, warnings, totalState) {
  const name = rawFile?.FileName || '';
  const url = rawFile?.ServerRelativeUrl || '';
  if (!name) return null;
  if (totalState.bytes >= ATTACHMENT_TOTAL_CAP) {
    warnings.push(`Attachment "${name}" was not embedded — the ${ATTACHMENT_TOTAL_CAP / (1024 * 1024)} MB total attachment cap was reached; exported as a link only.`);
    return { name, url };
  }
  let bytes;
  try {
    bytes = await fetchAttachmentBytes(client, url);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    warnings.push(`Attachment "${name}" could not be read (${err.message || err}); exported as a link only.`);
    return { name, url };
  }
  if (bytes.byteLength > ATTACHMENT_FILE_CAP) {
    warnings.push(`Attachment "${name}" was not embedded — it is larger than the ${ATTACHMENT_FILE_CAP / (1024 * 1024)} MB per-file cap; exported as a link only.`);
    return { name, url };
  }
  if (totalState.bytes + bytes.byteLength > ATTACHMENT_TOTAL_CAP) {
    warnings.push(`Attachment "${name}" was not embedded — the ${ATTACHMENT_TOTAL_CAP / (1024 * 1024)} MB total attachment cap was reached; exported as a link only.`);
    return { name, url };
  }
  totalState.bytes += bytes.byteLength;
  return { name, base64: arrayBufferToBase64(bytes) };
}

export async function captureListData(client, listId, { schemaDoc = null, maxItems = null } = {}) {
  const warnings = [];
  const schema = schemaDoc && schemaDoc.kind === SCHEMA_KIND
    ? schemaDoc
    : (await captureListSchema(client, listId)).doc;

  const rootFolder = String(schema.source?.rootFolder || '').replace(/\/+$/, '');
  const relPath = (serverRelative) => {
    const p = String(serverRelative || '');
    if (!rootFolder || !p.toLowerCase().startsWith(rootFolder.toLowerCase())) return '';
    return p.slice(rootFolder.length).replace(/^\/+/, '');
  };

  const cap = maxItems ? Math.max(1, Math.min(Number(maxItems) || DATA_CAP, DATA_CAP)) : DATA_CAP;
  const hasAttachments = schema.fields.some((f) => f.type === 'Attachments');
  const expand = ['FieldValuesAsText', ...(hasAttachments ? ['AttachmentFiles'] : [])];
  // $select=* deliberately — the Items-tab gotcha applies here too (views/
  // lists.js buildItemsPane): a bare internal-name projection on a User or
  // Lookup field 400s ("$expand must contain X"). FSObjType/FileDirRef/
  // FileRef ride alongside '*' the way SPUtils asks for them
  // (dcspad-sp-utilities.js:2447), so folders and their paths come back in
  // the same read as item content — no separate folder query.
  const { items: rawRows, partial } = await client.getAll(
    guidPath(listId, '/items'),
    { select: ['*', 'FSObjType', 'FileDirRef', 'FileRef', ...expand], expand, orderby: 'ID asc' },
    { cap, allowLargeCap: true },
  );
  if (partial) {
    warnings.push(`Only the first ${rawRows.length} item(s) were read — the list has more than the ${cap}-item cap.`);
  }

  const userFields = schema.fields.filter((f) => USER_TYPES.has(f.type)
    && (f.custom || f.internalName === 'Author' || f.internalName === 'Editor'));
  const lookupFields = schema.fields.filter((f) => LOOKUP_TYPES.has(f.type) && f.custom);

  // Lookup value maps: id -> shown value, one read per distinct lookup
  // target (exportListData, dcspad-sp-utilities.js:2460-2470).
  const lookupValues = new Map();   // internalName -> Map(id -> value)
  for (const f of lookupFields) {
    if (!f.lookupListId) continue;
    try {
      const { items: targetRows, partial } = await client.getAll(
        guidPath(f.lookupListId, '/items'),
        { select: ['Id', f.lookupField || 'Title'] },
        { cap: DATA_CAP, allowLargeCap: true },
      );
      lookupValues.set(f.internalName, new Map(targetRows.map((r) => [r.Id, r[f.lookupField || 'Title']])));
      if (partial) {
        warnings.push(`Lookup “${f.internalName}”: the target list has more items than could be indexed at once — values for ids beyond that are looked up individually.`);
      }
    } catch (err) {
      // A denied read (403) or any other read failure is a fact about this
      // one lookup list, not a reason to abort the whole export — only an
      // expired session (401) means nothing further can be trusted.
      if (isExpiredSession(err)) throw err;
      warnings.push(`Lookup “${f.internalName}”: target list could not be read (${err?.message || err}); ids exported without values.`);
    }
  }

  // Site users referenced by a User/UserMulti field, resolved once (bulk
  // web/siteusers) and cached per id — the same resolveUser() shape SPUtils
  // uses (dcspad-sp-utilities.js:2404-2422): {Id, Email, LoginName, Title}.
  let siteUsersById = null;
  const ensureSiteUsers = async () => {
    if (siteUsersById) return siteUsersById;
    siteUsersById = new Map();
    try {
      const { items } = await client.getAll('web/siteusers', { select: ['Id', 'Email', 'LoginName', 'Title'] });
      for (const u of items) siteUsersById.set(u.Id, u);
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      warnings.push(`Site users could not be read (${err?.message || err}); person fields resolve one id at a time instead.`);
    }
    return siteUsersById;
  };
  const referencedUserIds = new Set();
  const resolveUser = async (id) => {
    const users = await ensureSiteUsers();
    referencedUserIds.add(id);
    if (!users.has(id)) {
      try {
        const u = await client.get(`web/siteusers/getbyid(${Number(id)})`, { select: ['Id', 'Email', 'LoginName', 'Title'] });
        users.set(id, u);
      } catch (err) {
        if (isExpiredSession(err)) throw err;
        users.set(id, null);
      }
    }
    const u = users.get(id);
    return u
      ? { Id: u.Id, Email: u.Email || '', LoginName: u.LoginName || '', Title: u.Title || '' }
      : { Id: id, Email: '', LoginName: '', Title: '' };
  };

  // A lookup id the index doesn't carry (the target list paged past the
  // cap) is fetched directly rather than exported unresolved — one request
  // per missing id, cached back onto the same map for any later row that
  // references it too.
  const resolveLookupValue = async (f, id) => {
    const map = lookupValues.get(f.internalName);
    if (!map) return null;
    if (map.has(id)) return map.get(id) ?? null;
    try {
      const row = await client.get(guidPath(f.lookupListId, `/items(${id})`), { select: ['Id', f.lookupField || 'Title'] });
      const value = row ? (row[f.lookupField || 'Title'] ?? null) : null;
      map.set(id, value);
      return value;
    } catch (err) {
      // Degrade a missing row, never an expired sign-in.
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
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) resolved[f.internalName] = await Promise.all(ids.map(resolveUser));
    }
    for (const f of lookupFields) {
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) {
        resolved[f.internalName] = await Promise.all(
          ids.map(async (id) => ({ Id: id, value: await resolveLookupValue(f, id) })),
        );
      }
    }
    // Attachment metadata rides the same items() read (AttachmentFiles was
    // expanded above) — bytes are embedded (base64, capped) or the entry
    // degrades to a link; see captureAttachment().
    const rawFiles = Array.isArray(raw.AttachmentFiles) ? raw.AttachmentFiles : (raw.AttachmentFiles?.results || []);
    const attachments = [];
    for (const f of rawFiles) {
      const a = await captureAttachment(client, f, warnings, attachmentTotal);
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
      type: f.type, custom: f.custom, readOnly: f.readOnly,
      lookupList: f.lookupList, lookupListId: f.lookupListId, lookupField: f.lookupField,
      isSelfLookup: f.isSelfLookup, allowMultipleValues: f.allowMultipleValues,
    };
  }

  const users = siteUsersById
    ? [...referencedUserIds].map((id) => siteUsersById.get(id)).filter(Boolean)
      .map((u) => ({ Id: u.Id, Email: u.Email || '', LoginName: u.LoginName || '', Title: u.Title || '' }))
    : [];

  const source = {
    siteUrl: client.webUrl(), listTitle: schema.source.listTitle, listId: schema.source.listId,
    rootFolder: schema.source.rootFolder ?? null, itemCount: schema.source.itemCount ?? null,
  };

  const doc = buildDataDoc({
    source, fields: fieldMap, items, folders, users, warnings,
    generatorBuild: APP_BUILD_INFO.build,
  });

  return { doc, raw: { items: rawRows } };
}
