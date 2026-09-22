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
// Attachment BYTES are never embedded (unlike SPUtils' opt-in includeAttachments):
// this capture only ever records each attachment's name and server-relative
// URL. A schema-only "how many bytes would this be" concern doesn't apply to
// metadata, so there's no size gate here — the gate belongs to whatever
// (future) module re-uploads them.

import {
  LOOKUP_TYPES, USER_TYPES, buildDataDoc,
} from './list-data.js';
import { SCHEMA_KIND } from './list-schema.js';
import { captureListSchema } from './list-schema-capture.js';
import { APP_BUILD_INFO } from '../build-info.js';
import { isDeniedRead, isExpiredSession } from './denied.js';

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;

// Default page-read ceiling for a data export — the opt-in cap sp-rest.js
// reserves for this one caller (getAll's default stays 5000 for everyone
// else). A caller may pass a smaller `maxItems`; it can never exceed this.
const DATA_CAP = 100000;

const toIdList = (v) => (v == null ? [] : Array.isArray(v) ? v : [v]).filter((x) => x != null);

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
      const { items: targetRows } = await client.getAll(
        guidPath(f.lookupListId, '/items'),
        { select: ['Id', f.lookupField || 'Title'] },
      );
      lookupValues.set(f.internalName, new Map(targetRows.map((r) => [r.Id, r[f.lookupField || 'Title']])));
    } catch (err) {
      if (isDeniedRead(err) || isExpiredSession(err)) throw err;
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
      if (isDeniedRead(err) || isExpiredSession(err)) throw err;
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
      } catch { users.set(id, null); }
    }
    const u = users.get(id);
    return u
      ? { Id: u.Id, Email: u.Email || '', LoginName: u.LoginName || '', Title: u.Title || '' }
      : { Id: id, Email: '', LoginName: '', Title: '' };
  };

  const items = [];
  const folders = [];
  for (const raw of rawRows) {
    const resolved = {};
    for (const f of userFields) {
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) resolved[f.internalName] = await Promise.all(ids.map(resolveUser));
    }
    for (const f of lookupFields) {
      const ids = toIdList(raw[`${f.internalName}Id`]);
      if (ids.length) {
        const map = lookupValues.get(f.internalName);
        resolved[f.internalName] = ids.map((id) => ({ Id: id, value: map ? (map.get(id) ?? null) : null }));
      }
    }
    // Attachment metadata rides the same items() read (AttachmentFiles was
    // expanded above); normalize it to {name, url} and drop the raw
    // expansion object from the row — bytes are never fetched or embedded.
    const rawFiles = Array.isArray(raw.AttachmentFiles) ? raw.AttachmentFiles : (raw.AttachmentFiles?.results || []);
    const attachments = rawFiles
      .map((f) => ({ name: f?.FileName || '', url: f?.ServerRelativeUrl || '' }))
      .filter((f) => f.name);
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
