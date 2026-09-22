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
// Deviation from SPUtils, called out because it only exists here: SPUtils'
// toWebDateString() calls the server's webLocalDate() once per date VALUE
// (async, so every field gets its own exact offset). list-data.js's
// toWebDateString() is pure and takes one flat offsetMinutes per call, and
// toImportFormValues() applies ONE dateFormat to every field of one item —
// so this executor computes ONE offset per ITEM (from that item's Created
// date, falling back to "now"), not per value. That is coarser than SPUtils
// only on the rare item whose OWN custom DateTime value and Created date
// straddle a DST transition on the very same import; list-data.js's own
// header comment names this as the future apply module's call to make.

import {
  normalizeDataDoc, writableFields, partitionPasses, toImportFormValues,
  authorshipFormValues, folderOrder, LOOKUP_TYPES,
} from './list-data.js';
import { SpFileError } from '../sp-odata.js';
import { isExpiredSession } from './denied.js';

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;
const FAILED_CAP = 50;
const CONCURRENCY = 3;

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

// ---- small promise pool (SPUtils runPool) -------------------------------

async function runPool(items, limit, worker) {
  const queue = [...items];
  const n = Math.max(1, Math.min(limit, queue.length || 1));
  await Promise.all(new Array(n).fill(0).map(async () => {
    while (queue.length) { await worker(queue.shift()); }
  }));
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

async function fetchAttachmentBytes(sourceClient, serverRelativeUrl) {
  let origin = '';
  try { origin = new URL(sourceClient.webUrl()).origin; } catch { /* keep '' */ }
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

// Attachment BYTES only ever travel in copy mode, straight from the source
// site (same tenant) — a data document imported from a FILE never carries
// them (list-data-capture.js records name/url metadata only), so import mode
// can only skip and say why.
async function runAttachments(ctx, report, item, newItemId) {
  for (const a of item._attachments) {
    if (!ctx.includeAttachments) { report.attachments.skipped++; continue; }
    if (!ctx.sourceClient) {
      report.attachments.skipped++;
      if (!ctx.warnedNoSourceClient) {
        ctx.warnedNoSourceClient = true;
        report.warnings.push(
          'Attachments were not copied — the imported data document has no attachment bytes; '
          + 'use Copy to… from the source site to bring them across.',
        );
      }
      continue;
    }
    if (!ctx.sourceClient.context().live) { report.attachments.skipped++; continue; }
    try {
      const bytes = await fetchAttachmentBytes(ctx.sourceClient, a.url);
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
    const offsetMinutes = await ctx.offsetFor(f);
    const dateFormat = { ...ctx.baseFormat, offsetMinutes };
    const { values, errors } = toImportFormValues(f, pass1Fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
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

async function runItems(ctx, data, report, pass1Fields, userIds, lookupIds, onStep) {
  const items = [...data.items].sort((a, b) => (a.Id ?? 0) - (b.Id ?? 0));
  let done = 0;
  let aborted = false;
  await runPool(items, CONCURRENCY, async (item) => {
    if (aborted || ctx.signal?.aborted) { aborted = true; return; }
    const offsetMinutes = await ctx.offsetFor(item);
    const dateFormat = { ...ctx.baseFormat, offsetMinutes };
    const { values, errors } = toImportFormValues(item, pass1Fields, { dateFormat, userIds, lookupIds, idMap: report.idMap });
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
      if (isExpiredSession(err)) { aborted = true; report.aborted = 'auth'; return; }
      report.items.failed.push({ sourceId: item.Id, error: err.message || String(err) });
    }
    done++;
    onStep?.({ phase: 'items', index: done, total: items.length, label: `Copying item ${done} of ${items.length}…` });
  });
  capItemFailures(report);
  return !aborted;
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
    const offsetMinutes = await ctx.offsetFor(item);
    const dateFormat = { ...ctx.baseFormat, offsetMinutes };
    const values = authorshipFormValues(item, { dateFormat, userIds });
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
  const { pass1, pass2 } = partitionPasses(data.items, writable);

  const isMock = spWrite.isMock();
  const dayOffsetCache = new Map();
  let baseFormat = { order: 'mdy', sep: '/', timeSep: ':' };
  const dateField = writable.find((w) => w.typeAsString === 'DateTime');
  if (!isMock && (dateField || preserveAuthorship)) {
    try {
      baseFormat = await calibrateDateFormat(client, spWrite, listId, rootFolder, dateField ? dateField.name : 'Created');
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
      report.warnings.push(`Date format could not be learned from the target (${err.message || err}); using ${baseFormat.order} "${baseFormat.sep}" as a fallback.`);
    }
  }
  const offsetFor = async (item) => {
    if (isMock) return 0;
    const raw = item?.Created;
    const utc = raw ? new Date(raw) : new Date();
    if (Number.isNaN(utc.getTime())) return 0;
    try { return await webLocalOffsetMinutes(dayOffsetCache, client, utc); }
    catch { return 0; }
  };

  const lookupIds = new Map();
  for (const w of pass1.fields) {
    if (!LOOKUP_TYPES.has(w.typeAsString)) continue;
    const lookupListId = w.tf?.LookupList;
    if (!lookupListId) continue;
    try {
      const showField = w.tf?.LookupField || 'Title';
      const { items: rows } = await client.getAll(guidPath(lookupListId, '/items'), { select: ['Id', showField] });
      const byValue = new Map();
      for (const r of rows) {
        const key = String(r[showField] ?? '');
        if (!byValue.has(key)) byValue.set(key, []);
        byValue.get(key).push(r.Id);
      }
      lookupIds.set(w.name, byValue);
    } catch (err) {
      if (isExpiredSession(err)) { report.aborted = 'auth'; return report; }
      report.warnings.push(`Lookup "${w.name}": target lookup list could not be read (${err.message || err}); values will be left empty.`);
    }
  }

  const userIds = await buildUserIds(spWrite, data.users, report);
  if (report.aborted) return report;

  const ctx = {
    client, spWrite, listId, rootFolder, baseFormat, offsetFor, signal,
    includeAttachments, sourceClient,
  };

  if (!(await runFolders(ctx, data, report, pass1.fields, userIds, lookupIds, onStep)) || report.aborted) return report;
  if (!(await runItems(ctx, data, report, pass1.fields, userIds, lookupIds, onStep)) || report.aborted) return report;
  if (!(await runSelfLookups(ctx, data, report, pass2, onStep)) || report.aborted) return report;
  if (preserveAuthorship) await runAuthorship(ctx, data, report, userIds, onStep);
  return report;
}
