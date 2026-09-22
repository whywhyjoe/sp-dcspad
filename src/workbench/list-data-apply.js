// List item DATA — apply executor (target client + spWrite). Stage 1b-b.
//
// Port of SPUtils' importListData (utilities/dcspad-sp-utilities.js:2812-3146)
// into the Workbench, the same way list-schema-apply.js ports importListData's
// sibling createListFromSchema. list-data.js already carries the PURE per-type
// FieldValue conventions (toImportFormValues), the writable-field selection
// (writableFields) and the three-pass split (partitionPasses) — this module
// is the I/O the pure core can't do: reading the target's live fields/lookup
// lists, learning the web's date format, resolving users, and turning each
// pass into AddValidateUpdateItemUsingPath / ValidateUpdateListItem /
// AttachmentFiles/add calls through sp-write.js.
//
// Matches SPUtils' toWebDateString(), which calls the server's
// webLocalDate() once per date VALUE, not once per item: list-data.js's
// toWebDateString() is pure and takes one flat offsetMinutes per call, so
// this executor precomputes a PER-FIELD offset for each item
// (dateFormatForFields()/resolveItemFields()) before calling
// toImportFormValues() — a Created stamp and a custom DueOn value on the
// same item can straddle a DST transition differently, so one offset for
// the whole item would get one of them wrong.

import {
  normalizeDataDoc, writableFields, partitionPasses, toImportFormValues,
  authorshipFormValues, folderOrder, LOOKUP_TYPES,
} from './list-data.js';
import { fetchAttachmentBytes } from './list-data-capture.js';
import { SpFileError } from '../sp-odata.js';
import { isExpiredSession } from './denied.js';

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;
const FAILED_CAP = 50;
// Same opt-in cap list-data-capture.js reserves for a lookup-target index.
const LOOKUP_INDEX_CAP = 100000;

function cleanGuid(v) {
  const s = String(v ?? '').replace(/[{}]/g, '').trim();
  return s ? s.toLowerCase() : null;
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

// ---- report -----------------------------------------------------------

function newDataReport(data) {
  return {
    listTitle: data.source?.listTitle || '',
    attempted: data.items.length,
    items: { added: 0, failed: [] },
    folders: { created: 0, failed: 0 },
    attachments: { added: 0, skipped: 0, failed: 0 },
    authorship: { applied: 0, failed: 0 },
    fieldErrors: [],
    idMap: {},
    warnings: [],
    aborted: '',
  };
}

function capItemFailures(report) {
  if (report.items.failed.length > FAILED_CAP) {
    report.items.failedTruncated = report.items.failed.length - FAILED_CAP;
    report.items.failed = report.items.failed.slice(0, FAILED_CAP);
  }
}

// ---- target reads -------------------------------------------------------

async function readTargetListRow(client, listId) {
  return client.get(guidPath(listId), {
    select: ['BaseType', 'Title', 'RootFolder/ServerRelativeUrl'],
    expand: 'RootFolder',
  });
}

async function readTargetFields(client, listId) {
  const { items } = await client.getAll(guidPath(listId, '/fields'), {
    select: ['Id', 'InternalName', 'TypeAsString', 'ReadOnlyField', 'LookupList', 'LookupField', 'DisplayFormat'],
  });
  return items;
}

// ---- date calibration (SPUtils calibrateDateFormat / webLocalDate) ------

async function getRegionalSettings(client) {
  return client.get('web/RegionalSettings', {
    select: ['DateFormat', 'DateSeparator', 'TimeSeparator', 'Time24', 'AM', 'PM', 'LocaleId'],
  });
}

function formatFromSettings(s) {
  return {
    order: s?.DateFormat === 2 ? 'ymd' : s?.DateFormat === 1 ? 'dmy' : 'mdy',
    sep: s?.DateSeparator || '/',
    timeSep: s?.TimeSeparator || ':',
  };
}

// Parses the server's own sample out of a rejected-date ErrorMessage, e.g.
// "Enter a date and time like this: 23/02/2012 02:25 PM".
function formatFromSample(sample) {
  const m = /(\d{1,4})([^\d\s])(\d{1,2})\2(\d{1,4})/.exec(String(sample || ''));
  if (!m) return null;
  const parts = [m[1], m[3], m[4]];
  const order = parts.map((p) => (p.length === 4 ? 'y' : Number(p) === 23 ? 'd' : 'm')).join('');
  if (!/^(mdy|dmy|ymd)$/.test(order)) return null;
  const rest = sample.slice(m.index + m[0].length);
  const tm = /\d{1,2}([^\d\s])\d{2}/.exec(rest);
  return { order, sep: m[2], timeSep: tm ? tm[1] : ':' };
}

// Provokes the target for its own date sample with a deliberately invalid
// value on a DateTime field — the create it rides on is aborted by that same
// invalid value (see addValidateUpdateItem's header note), so nothing is
// written. Falls back to RegionalSettings' own DateFormat/DateSeparator when
// the sample can't be parsed (or the probe field itself is missing/hidden).
async function calibrateDateFormat(client, spWrite, listId, rootFolder, probeField) {
  const settings = await getRegionalSettings(client);
  const fallback = formatFromSettings(settings);
  try {
    await spWrite.addValidateUpdateItem(listId, {
      folderPath: rootFolder,
      underlyingObjectType: 0,
      formValues: [{ FieldName: probeField, FieldValue: 'not-a-date' }],
    });
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    const sample = err?.fieldErrors?.[probeField];
    if (sample) {
      const parsed = formatFromSample(sample.split(':').slice(1).join(':'));
      if (parsed) return parsed;
    }
  }
  return fallback;
}

async function utcOffsetAtMs(client, isoInstant) {
  const data = await client.get(`web/RegionalSettings/TimeZone/utcToLocalTime(@d)?@d='${isoInstant}'`);
  const local = data?.value;
  if (!local) {
    throw new SpFileError('Web time zone could not be read; date not written.', { code: 'write' });
  }
  return new Date(`${local}Z`).getTime() - new Date(isoInstant).getTime();
}

// Per-UTC-day offset, cached — mirrors SPUtils' tzOffsetByDay/tzOffsetExact:
// a day whose noon offset disagrees with its neighbours is a DST transition,
// so that one day gets an exact per-instant lookup instead of the cached noon
// value.
async function webLocalOffsetMinutes(cache, client, utc) {
  const dayOf = (d) => d.toISOString().slice(0, 10);
  const offsetForDay = async (day) => {
    if (!cache.has(day)) cache.set(day, await utcOffsetAtMs(client, `${day}T12:00:00Z`));
    return cache.get(day);
  };
  const day = dayOf(utc);
  const here = await offsetForDay(day);
  const prev = await offsetForDay(dayOf(new Date(utc.getTime() - 86400000)));
  const next = await offsetForDay(dayOf(new Date(utc.getTime() + 86400000)));
  let offsetMs = here;
  if (here !== prev || here !== next) {
    offsetMs = await utcOffsetAtMs(client, utc.toISOString());
  }
  return offsetMs / 60000;
}

// ---- users (SPUtils loginFor/ensureUser) --------------------------------

// Resolves every distinct user list-data-capture.js already collected
// (dataDoc.users) into the target's own login names, once, up front —
// toImportFormValues is pure and synchronous, so every ensureUser() the
// import needs happens here, before any item is built.
async function buildUserIds(spWrite, users, report) {
  const userIds = new Map();
  for (const u of users || []) {
    const key = String(u?.Email || u?.LoginName || '').toLowerCase();
    if (!key) continue;
    // App/system principals ("SharePoint App", SYSTEM) have no email and
    // can't be matched by ValidateUpdateListItem; leave the field empty.
    if (!u.Email && !/^i:0#\.f\|membership\|/i.test(u.LoginName || '')) continue;
    try {
      const ensured = await spWrite.ensureUser(u.Email || u.LoginName);
      if (ensured?.loginName) userIds.set(key, ensured.loginName);
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return userIds; }
      if (u.Email) userIds.set(key, `i:0#.f|membership|${u.Email.toLowerCase()}`);
    }
  }
  return userIds;
}

// ---- attachments ----------------------------------------------------------

// Attachment BYTES travel as base64 on the item (v1's own shape, or an
// export that stayed under the size caps) — used in EITHER mode, import
// included, since that's just data on the document, not a network fetch.
// Only when base64 is absent (a document that degraded to `{ name, url }`
// over a cap) does copy mode fall back to fetching the source's bytes live;
// an imported-from-file document with no base64 can only skip and say why.
async function runAttachments(ctx, report, item, newItemId) {
  for (const a of item._attachments) {
    if (!ctx.includeAttachments) { report.attachments.skipped++; continue; }
    try {
      let bytes;
      // A zero-byte file is a valid v1 attachment: presence of the string,
      // not its length, says the bytes travelled in the document.
      if (typeof a.base64 === 'string') {
        bytes = base64ToArrayBuffer(a.base64);
      } else if (ctx.sourceClient && ctx.sourceClient.context().live && a.url) {
        bytes = await fetchAttachmentBytes(ctx.sourceClient, a.url);
      } else {
        report.attachments.skipped++;
        if (!ctx.warnedNoSourceClient) {
          ctx.warnedNoSourceClient = true;
          report.warnings.push(
            'Some attachments were not copied — the imported data document has no attachment bytes for them; '
            + 'use Copy to… from the source site to bring them across.',
          );
        }
        continue;
      }
      await ctx.spWrite.addAttachment(ctx.listId, newItemId, a.name, bytes);
      report.attachments.added++;
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.attachments.failed++;
      report.warnings.push(`Item ${item.Id} → ${newItemId}: attachment "${a.name}" failed (${err.message || err}).`);
    }
  }
}

// ---- item/folder create, with the retry-without-rejected-fields ---------

async function createItemWithRetry(ctx, report, sourceId, spec) {
  try {
    return await ctx.spWrite.addValidateUpdateItem(ctx.listId, spec);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    if (err?.code === 'metadata-write' && err.fieldErrors && Object.keys(err.fieldErrors).length) {
      for (const [field, message] of Object.entries(err.fieldErrors)) {
        report.fieldErrors.push({ sourceId, field, message });
      }
      const drop = new Set(Object.keys(err.fieldErrors));
      const retryValues = spec.formValues.filter((v) => !drop.has(v.FieldName));
      const retried = await ctx.spWrite.addValidateUpdateItem(ctx.listId, { ...spec, formValues: retryValues });
      return { ...retried, droppedFields: [...drop] };
    }
    throw err;
  }
}

// ---- per-item date resolution ---------------------------------------------

// Computes this item's date format with one offset PER FIELD (a Created
// stamp and a custom DueOn value can straddle a DST transition differently —
// see list-data.js's fieldDateFormat) and drops any DateTime field whose
// offset could not be learned, warning once per field NAME across the whole
// run rather than once per item, and never falling back to a guessed
// mdy/UTC value for it (finding #4/#10).
async function resolveItemFields(ctx, report, item, fields) {
  const dateNames = fields.filter((w) => w.typeAsString === 'DateTime').map((w) => w.name);
  if (!dateNames.length) return { dateFormat: ctx.baseFormat, fields };
  const dateFormat = await ctx.dateFormatForFields(item, dateNames);
  const failed = dateNames.filter((n) => dateFormat.perField.get(n) === null);
  for (const n of failed) {
    if (!ctx.warnedDateFields.has(n)) {
      ctx.warnedDateFields.add(n);
      report.warnings.push(`“${n}” not written — the target web's date format could not be learned.`);
    }
  }
  return { dateFormat, fields: failed.length ? fields.filter((w) => !failed.includes(w.name)) : fields };
}

// ---- pass 0: folders ------------------------------------------------------

async function runFolders(ctx, data, report, pass1Fields, userIds, lookupIds, onStep) {
  const ordered = folderOrder(data.folders);
  let i = 0;
  for (const f of ordered) {
    if (ctx.signal?.aborted) { report.aborted = 'user'; return false; }
    i++;
    onStep?.({ phase: 'folders', index: i, total: ordered.length, label: `Creating folder ${i} of ${ordered.length}…` });
    const path = String(f._folderPath || f.folderPath || '');
    const slash = path.lastIndexOf('/');
    const parent = slash === -1 ? ctx.rootFolder : `${ctx.rootFolder}/${path.slice(0, slash)}`;
    const name = path.slice(slash + 1);
    const { dateFormat, fields } = await resolveItemFields(ctx, report, f, pass1Fields);
    const { values, errors } = toImportFormValues(f, fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: f.Id, ...e });
    const formValues = [{ FieldName: 'Title', FieldValue: name }, ...values.filter((v) => v.FieldName !== 'Title')];
    try {
      const { id } = await createItemWithRetry(ctx, report, f.Id, {
        folderPath: parent, underlyingObjectType: 1, leafName: name, formValues,
      });
      report.idMap[f.Id] = id;
      report.folders.created++;
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return false; }
      report.folders.failed++;
      report.warnings.push(`Folder "${path}" could not be created (${err.message || err}); its items will fail unless the folder exists.`);
    }
  }
  return true;
}

// ---- pass 1: items ----------------------------------------------------

// Sequential, in ascending SOURCE id order (folders already went first, by
// depth) — a fresh target list then gets item ids in the same sequence the
// source had them. The earlier concurrent pool dequeued in that order too,
// but completion order (and so id ASSIGNMENT order) wasn't guaranteed by
// network timing; a live run once produced ids 1,2,3 for source items
// 2,3,1. Only pass 1's CREATE matters for this — pass 2/3 are updates
// against ids already assigned, with no ordering promise to keep.
async function runItems(ctx, data, report, pass1Fields, userIds, lookupIds, onStep) {
  const items = [...data.items].sort((a, b) => (a.Id ?? 0) - (b.Id ?? 0));
  let done = 0;
  for (const item of items) {
    if (ctx.signal?.aborted) { report.aborted = 'user'; return false; }
    const { dateFormat, fields } = await resolveItemFields(ctx, report, item, pass1Fields);
    const { values, errors } = toImportFormValues(item, fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: item.Id, ...e });
    const folderPath = item._dir ? `${ctx.rootFolder}/${item._dir}` : ctx.rootFolder;
    try {
      const { id, droppedFields } = await createItemWithRetry(ctx, report, item.Id, {
        folderPath, underlyingObjectType: 0, formValues: values,
      });
      report.idMap[item.Id] = id;
      report.items.added++;
      if (droppedFields?.length) {
        report.warnings.push(`Item ${item.Id} → ${id}: created without ${droppedFields.join(', ')} (see field errors).`);
      }
      if (Array.isArray(item._attachments) && item._attachments.length) {
        await runAttachments(ctx, report, item, id);
      }
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return false; }
      report.items.failed.push({ sourceId: item.Id, error: err.message || String(err) });
    }
    done++;
    onStep?.({ phase: 'items', index: done, total: items.length, label: `Copying item ${done} of ${items.length}…` });
  }
  capItemFailures(report);
  return true;
}

// ---- pass 2: self-referencing lookups ------------------------------------

async function runSelfLookups(ctx, data, report, pass2, onStep) {
  if (!pass2.fields.length) return true;
  const allRows = [...data.folders, ...data.items];
  const updates = [];
  for (const item of allRows) {
    const newId = report.idMap[item.Id];
    if (!newId) continue;
    const { values, errors } = toImportFormValues(item, pass2.fields, { idMap: report.idMap });
    for (const e of errors) report.fieldErrors.push({ sourceId: item.Id, ...e });
    const set = values.filter((v) => v.FieldValue);
    if (set.length) updates.push({ item, newId, formValues: set });
  }
  let i = 0;
  for (const { item, newId, formValues } of updates) {
    if (ctx.signal?.aborted) { report.aborted = 'user'; return false; }
    i++;
    onStep?.({ phase: 'selfLookups', index: i, total: updates.length, label: `Linking self-references ${i} of ${updates.length}…` });
    try {
      await ctx.spWrite.validateUpdateListItem({ listId: ctx.listId, itemId: newId }, formValues);
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return false; }
      if (err.fieldErrors) {
        for (const [field, message] of Object.entries(err.fieldErrors)) report.fieldErrors.push({ sourceId: item.Id, field, message });
      } else {
        report.fieldErrors.push({ sourceId: item.Id, field: pass2.fields.map((f) => f.name).join(','), message: err.message });
      }
    }
  }
  return true;
}

// ---- pass 3: authorship ---------------------------------------------------

async function runAuthorship(ctx, data, report, userIds, onStep) {
  const allRows = [...data.folders, ...data.items].filter((i) => report.idMap[i.Id]);
  let i = 0;
  for (const item of allRows) {
    if (ctx.signal?.aborted) { report.aborted = 'user'; return false; }
    i++;
    onStep?.({ phase: 'authorship', index: i, total: allRows.length, label: `Restoring authorship ${i} of ${allRows.length}…` });
    let dateFormat = ctx.baseFormat;
    let dropNames = [];
    const wanted = ['Created', 'Modified'].filter((n) => item?.[n] != null);
    if (wanted.length) {
      if (!ctx.dateCalibrationOk) {
        dropNames = wanted;
      } else {
        dateFormat = await ctx.dateFormatForFields(item, wanted);
        dropNames = wanted.filter((n) => dateFormat.perField.get(n) === null);
      }
    }
    for (const n of dropNames) {
      if (!ctx.warnedDateFields.has(n)) {
        ctx.warnedDateFields.add(n);
        report.warnings.push(`“${n}” not written — the target web's date format could not be learned.`);
      }
    }
    let values = authorshipFormValues(item, { dateFormat, userIds });
    if (dropNames.length) values = values.filter((v) => !dropNames.includes(v.FieldName));
    if (!values.length) continue;
    try {
      await ctx.spWrite.validateUpdateListItem({ listId: ctx.listId, itemId: report.idMap[item.Id] }, values, { newDocumentUpdate: true });
      report.authorship.applied++;
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return false; }
      report.authorship.failed++;
      if (err.fieldErrors) {
        for (const [field, message] of Object.entries(err.fieldErrors)) report.fieldErrors.push({ sourceId: item.Id, field, message });
      } else {
        report.warnings.push(`Item ${item.Id}: authorship could not be restored (${err.message || err}).`);
      }
    }
  }
  return true;
}

// ---- entry point ------------------------------------------------------

// { dataDoc, schemaDoc, client, spWrite, listId, options: { includeAttachments,
//   preserveAuthorship, sourceClient }, onStep, signal } → report.
//
// `client`/`spWrite` are the TARGET's — the same pair the schema executor
// just wrote the list with. `sourceClient`, when given (copy mode only), is
// the SOURCE site's own client, used only to fetch attachment bytes: same
// tenant, read-only, never used for anything else here.
export async function applyListData({
  dataDoc, client, spWrite, listId, options = {}, onStep, signal,
} = {}) {
  const { includeAttachments = false, preserveAuthorship = false, sourceClient = null } = options;
  const data = normalizeDataDoc(dataDoc);
  const report = newDataReport(data);
  if (signal?.aborted) { report.aborted = 'user'; return report; }

  let listRow;
  try {
    listRow = await readTargetListRow(client, listId);
  } catch (err) {
    if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
    throw err;
  }
  if (Number(listRow.BaseType) === 1) {
    throw new SpFileError(
      'Item import into a document library arrives with stage 2.',
      { code: 'library-items' },
    );
  }
  report.listTitle = listRow.Title || report.listTitle;
  const rootFolder = String(listRow.RootFolder?.ServerRelativeUrl || '').replace(/\/+$/, '');

  let targetFieldRows;
  try {
    targetFieldRows = await readTargetFields(client, listId);
  } catch (err) {
    if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
    throw err;
  }
  const writable = writableFields(data.fields, targetFieldRows);

  // Date calibration must never guess: a genuine failure to learn the
  // target's regional format (RegionalSettings itself unreadable — the
  // deliberately-invalid-date probe failing is fine, calibrateDateFormat
  // itself falls back to RegionalSettings for that) means every DateTime
  // field is DROPPED from every item rather than written as a guessed
  // mdy/UTC value (finding #4). Computed before partitionPasses so a
  // dropped field never even reaches pass 1/pass 3.
  const isMock = spWrite.isMock();
  const dayOffsetCache = new Map();
  let baseFormat = { order: 'mdy', sep: '/', timeSep: ':' };
  let dateCalibrationOk = true;
  const dateFieldsAll = writable.filter((w) => w.typeAsString === 'DateTime');
  if (!isMock && (dateFieldsAll.length || preserveAuthorship)) {
    try {
      baseFormat = await calibrateDateFormat(
        client, spWrite, listId, rootFolder, dateFieldsAll[0] ? dateFieldsAll[0].name : 'Created',
      );
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
      dateCalibrationOk = false;
    }
  }
  const warnedDateFields = new Set();
  if (!dateCalibrationOk) {
    for (const w of dateFieldsAll) {
      warnedDateFields.add(w.name);
      report.warnings.push(`“${w.name}” not written — the target web's date format could not be learned.`);
    }
    if (preserveAuthorship) {
      for (const n of ['Created', 'Modified']) warnedDateFields.add(n);
      report.warnings.push('“Created”/“Modified” not written — the target web’s date format could not be learned.');
    }
  }
  const effectiveWritable = dateCalibrationOk ? writable : writable.filter((w) => w.typeAsString !== 'DateTime');
  const { pass1, pass2 } = partitionPasses(data.items, effectiveWritable);

  // Per-VALUE offset (not per item) — see resolveItemFields()/list-data.js's
  // fieldDateFormat. A timezone lookup that fails for one specific value is
  // the same "never guess" rule: that field is dropped for that item, never
  // written with an offset of 0/UTC.
  const offsetForValue = async (rawValue) => {
    if (isMock) return 0;
    const utc = rawValue ? new Date(rawValue) : new Date();
    if (Number.isNaN(utc.getTime())) return 0;
    try { return await webLocalOffsetMinutes(dayOffsetCache, client, utc); }
    catch { return null; }
  };
  const dateFormatForFields = async (item, names) => {
    const perField = new Map();
    for (const name of names) {
      const raw = item?.[name];
      if (raw == null) continue;
      perField.set(name, await offsetForValue(raw));
    }
    return { ...baseFormat, perField };
  };

  const lookupIds = new Map();
  for (const w of pass1.fields) {
    if (!LOOKUP_TYPES.has(w.typeAsString)) continue;
    // SharePoint's own REST responses are inconsistent about braces on a
    // GUID-valued property — cleanGuid() the same way every other GUID this
    // codebase reads off the wire is handled (list-schema.js normalizeField,
    // scrubSchemaXml…). Left raw, a braced value built a malformed
    // `guid'{…}'` OData path, the read below failed, and the index for this
    // field was silently never built — every cross-list lookup value then
    // came back unresolved while a self-lookup (which never touches this
    // index — it resolves through idMap instead) kept working.
    const lookupListId = cleanGuid(w.tf?.LookupList);
    if (!lookupListId) continue;
    try {
      const showField = w.tf?.LookupField || 'Title';
      const { items: rows, partial } = await client.getAll(
        guidPath(lookupListId, '/items'), { select: ['Id', showField] },
        { cap: LOOKUP_INDEX_CAP, allowLargeCap: true },
      );
      const byValue = new Map();
      for (const r of rows) {
        const key = String(r[showField] ?? '');
        if (!byValue.has(key)) byValue.set(key, []);
        byValue.get(key).push(r.Id);
      }
      lookupIds.set(w.name, byValue);
      if (partial) {
        report.warnings.push(`Lookup "${w.name}": the target lookup list has more items than could be indexed at once — some values may be left unresolved.`);
      }
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
      report.warnings.push(`Lookup "${w.name}": target lookup list could not be read (${err.message || err}); values will be left empty.`);
    }
  }

  const userIds = await buildUserIds(spWrite, data.users, report);
  if (report.aborted) return report;

  const ctx = {
    client, spWrite, listId, rootFolder, baseFormat, dateCalibrationOk,
    dateFormatForFields, warnedDateFields, signal,
    includeAttachments, sourceClient,
  };

  if (!(await runFolders(ctx, data, report, pass1.fields, userIds, lookupIds, onStep)) || report.aborted) return report;
  if (!(await runItems(ctx, data, report, pass1.fields, userIds, lookupIds, onStep)) || report.aborted) return report;
  if (!(await runSelfLookups(ctx, data, report, pass2, onStep)) || report.aborted) return report;
  if (preserveAuthorship) await runAuthorship(ctx, data, report, userIds, onStep);
  return report;
}
