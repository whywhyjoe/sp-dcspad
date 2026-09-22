// List schema — apply executor (target client + spWrite).
//
// Walks the pure Plan buildApplyPlan (list-schema.js) produces and turns each
// Step into a write against the target web. Nothing here decides WHAT to do
// — that is the planner's job, done before a single byte is written, so a
// dry-run preview and a live run render/execute the exact same Plan. This
// module only knows HOW to do each step kind, and the bookkeeping (fieldIdMap,
// blocked-by-dependency cascading, the report counters) that turning a
// static plan into a sequence of live writes requires.
//
// Bodies are nometadata JSON (no __metadata) — see writeJson()'s one-time
// odata=verbose retry for the tenants that still demand it on a 400.

import {
  buildApplyPlan, newReport, scrubSchemaXml, textFallbackXml,
} from './list-schema.js';
import { probeTarget } from './list-schema-capture.js';
import { SpFileError, odataPathLiteral } from '../sp-odata.js';
import { isExpiredSession } from './denied.js';

const listPath = (id, sub = '') => `web/lists(guid'${id}')${sub}`;

// ---- verbose-odata fallback -------------------------------------------------
// Some tenants still 400 a nometadata write that omits __metadata. One
// helper, used by every POST/MERGE in this module, so the retry logic never
// drifts between step kinds.
function needsVerboseRetry(err) {
  if (err?.status !== 400) return false;
  const msg = String(err?.message || '');
  return /__metadata/i.test(msg) || /\bSP\.[A-Za-z]+(\.[A-Za-z]+)*\b/.test(msg) || /entity ?type/i.test(msg);
}

async function writeJson(spWrite, method, path, body, entityType, opts = {}) {
  try {
    return await spWrite[method](path, body, opts);
  } catch (err) {
    if (!needsVerboseRetry(err)) throw err;
    // SP.XmlSchemaFieldCreationInformation is the one creation-info entity
    // in this module: its __metadata belongs INSIDE body.parameters (that's
    // where SharePoint's verbose odata reader looks for a creation-info
    // type), not at the body's own top level like every other entity here.
    const verboseBody = entityType === 'SP.XmlSchemaFieldCreationInformation'
      ? { ...body, parameters: { __metadata: { type: entityType }, ...(body.parameters || {}) } }
      : { __metadata: { type: entityType }, ...body };
    return spWrite[method](path, verboseBody, { ...opts, contentType: 'application/json;odata=verbose' });
  }
}

// ---- between-step reads ------------------------------------------------
// The plan is built from a probe taken before the list existed (or before
// any of this run's writes landed); the executor re-reads the target's own
// fields/lists at the points the plan depends on that being current — see
// the module doc and CLAUDE.md-adjacent brief for exactly when.

async function readTargetFields(ctx) {
  const { items } = await ctx.client.getAll(listPath(ctx.listId, '/fields'), {
    select: ['Id', 'InternalName', 'Title', 'TypeAsString', 'Hidden', 'ReadOnlyField', 'Indexed', 'EnforceUniqueValues'],
  });
  return items.map((f) => ({
    id: f.Id, internalName: f.InternalName, title: f.Title, typeAsString: f.TypeAsString,
    hidden: !!f.Hidden, readOnly: !!f.ReadOnlyField, indexed: !!f.Indexed,
    enforceUniqueValues: !!f.EnforceUniqueValues,
  }));
}

// Deviation from the plan's per-title `$filter`: every mock resolver in this
// codebase ignores $filter, so this reads every list once (nometadata is
// already unprojected-friendly) and matches titles case-insensitively on the
// client — cheaper than N filtered reads on a real tenant too.
async function readTargetListsByTitle(ctx) {
  const { items } = await ctx.client.getAll('web/lists', { select: ['Id', 'Title'] });
  const map = new Map();
  for (const l of items) map.set(String(l.Title || '').toLowerCase(), l.Id);
  return map;
}

async function readTargetViews(ctx) {
  const { items } = await ctx.client.getAll(listPath(ctx.listId, '/views'), {
    select: ['Id', 'Title', 'DefaultView'],
  });
  return items.map((v) => ({ id: v.Id, title: v.Title, defaultView: !!v.DefaultView }));
}

function seedFieldIdMap(ctx) {
  for (const f of ctx.targetFields || []) {
    if (!(f.internalName in ctx.fieldIdMap)) ctx.fieldIdMap[f.internalName] = f.id;
  }
}

// The full between-step re-read: fields, lists-by-title, and views. Used at
// every point the executor needs the target's live ground truth — priming a
// retry/resume, right after the list is created or adopted, and again right
// before the first view step.
async function readTargetState(ctx) {
  ctx.targetFields = await readTargetFields(ctx);
  seedFieldIdMap(ctx);
  ctx.targetListsByTitle = await readTargetListsByTitle(ctx);
  ctx.targetViews = await readTargetViews(ctx);
}

// Reads the executor takes BETWEEN steps are never a write, but a failure
// here is still a run-stopping fact: a 401 aborts cleanly (report.aborted =
// 'auth'); anything else aborts as a probe failure (report.aborted =
// 'probe', the message recorded as a warning) rather than guessing from
// stale data. Returns true when `fn` succeeded; on false the caller must
// stop and return the report immediately — every step it hasn't reached
// yet stays 'planned'.
async function safeRead(ctx, report, fn) {
  try {
    await fn();
    return true;
  } catch (err) {
    if (isExpiredSession(err)) { report.aborted = 'auth'; return false; }
    report.aborted = 'probe';
    report.warnings.push(err?.message || String(err));
    return false;
  }
}

function resolveListIdByTitle(ctx, title) {
  const id = ctx.targetListsByTitle?.get(String(title || '').toLowerCase());
  if (!id) {
    throw new SpFileError(
      `Lookup target list ‘${title}’ was not found on the target.`,
      { code: 'not-found' },
    );
  }
  return id;
}

// ---- step runners ------------------------------------------------------

async function runListCreate(step, ctx, report) {
  const { title, description, baseTemplate, contentTypesEnabled, urlName } = step.payload;
  const proposedUrlName = urlName || String(title || '').replace(/\s+/g, '');

  // The mock's generic 'web' echo fallback would answer this GET with the
  // web entity (wrong Title), misreporting every mock create as a URL
  // collision — the precheck is only meaningful against a real tenant.
  if (!ctx.spWrite.isMock()) {
    let webRel = '';
    try { webRel = new URL(ctx.client.webUrl()).pathname.replace(/\/+$/, ''); } catch { /* keep '' */ }
    const listUrl = `${webRel}/Lists/${proposedUrlName}`;
    try {
      const existing = await ctx.client.get(
        `web/GetList(@listUrl)?@listUrl='${odataPathLiteral(listUrl)}'&$select=Id,Title`,
      );
      if (existing?.Title && existing.Title !== title) {
        throw new SpFileError(
          `The list URL ‘${listUrl}’ is already taken by ‘${existing.Title}’ — it may be sitting `
          + 'in the recycle bin. Choose a different title, or empty the recycle bin first.',
          { code: 'url-taken' },
        );
      }
    } catch (err) {
      if (err?.status === 404 || err?.code === 'not-found') { /* the URL is free */ }
      else throw err;
    }
  }

  const body = {
    Title: urlName || title, Description: description || '', BaseTemplate: baseTemplate,
    ContentTypesEnabled: !!contentTypesEnabled, AllowContentTypes: true,
  };
  const data = await writeJson(ctx.spWrite, 'postJson', 'web/lists', body, 'SP.List',
    { fallback: 'Could not create the list', code: 'write' });
  ctx.listId = data.Id;
  report.listId = data.Id;
  report.created = true;
  report.rootFolder = data.RootFolder?.ServerRelativeUrl || '';
  // The list step is DONE the moment POST web/lists answers with an Id —
  // everything past this point is a separate concern. Renaming off the
  // URL-safe title back to the one the user chose is a nicety, not part of
  // list creation: a rejected MERGE here is a warning (the list exists,
  // just under its URL name), never a reason to fail the step that already
  // succeeded.
  if (urlName) {
    try {
      await writeJson(ctx.spWrite, 'mergeJson', listPath(data.Id), { Title: title }, 'SP.List', {});
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.warnings.push(
        `The list was created under the URL name ‘${urlName}’ — renaming its title to ‘${title}’ failed (${err.message || err}).`,
      );
    }
  }
  return data;
}

async function runListAdopt(step) {
  const { listId, title } = step.payload;
  return { id: listId, title };
}

async function mergeSettingsGroup(ctx, report, group) {
  const keys = Object.keys(group || {});
  if (!keys.length) return;   // empty groups send nothing
  try {
    await writeJson(ctx.spWrite, 'mergeJson', listPath(ctx.listId), group, 'SP.List',
      { fallback: 'Could not apply list settings', code: 'write' });
    report.settings.applied.push(...keys);
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    // One rejected key should not lose the whole group — retry one property
    // at a time and keep whatever the target accepts.
    for (const key of keys) {
      try {
        await writeJson(ctx.spWrite, 'mergeJson', listPath(ctx.listId), { [key]: group[key] }, 'SP.List',
          { fallback: `Could not apply ${key}`, code: 'write' });
        report.settings.applied.push(key);
      } catch (err2) {
        if (isExpiredSession(err2)) throw err2;
        report.settings.failed.push(key);
        report.warnings.push(`Setting ‘${key}’ was rejected on the target — ${err2.message}`);
      }
    }
  }
}

async function runListSettings(step, ctx, report) {
  const { groupA, groupB } = step.payload;
  await mergeSettingsGroup(ctx, report, groupA);
  await mergeSettingsGroup(ctx, report, groupB);
}

async function runCtAttach(step, ctx) {
  const { contentTypeId } = step.payload;
  return writeJson(ctx.spWrite, 'postJson', `${listPath(ctx.listId)}/contenttypes/addAvailableContentType`,
    { contentTypeId }, 'SP.List', { fallback: 'Could not attach the content type', code: 'write' });
}

function resolveFieldRefs(step, ctx) {
  const refs = step.refs || {};
  let lookupListId = null;
  if (refs.lookupListId) {
    if (refs.lookupListId.self) lookupListId = ctx.listId;
    else if (refs.lookupListId.id) lookupListId = refs.lookupListId.id;
    else if (refs.lookupListId.list) lookupListId = resolveListIdByTitle(ctx, refs.lookupListId.list);
  }
  let primaryFieldId = null;
  if (refs.primaryFieldId?.field) {
    primaryFieldId = ctx.fieldIdMap[refs.primaryFieldId.field] || null;
    if (!primaryFieldId) {
      throw new SpFileError(
        `Its primary lookup column ‘${refs.primaryFieldId.field}’ was not created on the target.`,
        { code: 'blocked' },
      );
    }
  }
  return { lookupListId, primaryFieldId };
}

async function runFieldCreate(step, ctx) {
  const { field: f, asText, options } = step.payload;
  const { lookupListId, primaryFieldId } = resolveFieldRefs(step, ctx);
  const xml = asText
    ? textFallbackXml(f)
    : scrubSchemaXml(f.schemaXml, { lookupListId, primaryFieldId, fieldType: f.type });
  const body = { parameters: { SchemaXml: xml, Options: options ?? 8 } };
  const data = await writeJson(
    ctx.spWrite, 'postJson', `${listPath(ctx.listId)}/fields/createfieldasxml`, body,
    'SP.XmlSchemaFieldCreationInformation',
    { fallback: `Could not create the column ‘${f.displayName}’`, code: 'write' },
  );
  if (data.InternalName && data.InternalName !== f.internalName) {
    throw new SpFileError(
      `Created as ‘${data.InternalName}’, expected ‘${f.internalName}’.`,
      { code: 'name-mismatch' },
    );
  }
  ctx.fieldIdMap[f.internalName] = data.Id;
  return data;
}

const FIELD_MERGE_ORDER = ['Indexed', 'EnforceUniqueValues', 'CustomFormatter'];

async function runFieldMerge(step, ctx, report) {
  const { internalName, merges } = step.payload;
  const fid = ctx.fieldIdMap[internalName];
  if (!fid) {
    throw new SpFileError(`Column ‘${internalName}’ has no id on the target to merge onto.`, { code: 'write' });
  }
  const keys = FIELD_MERGE_ORDER.filter((k) => k in merges);
  const failures = [];
  for (const key of keys) {
    try {
      await writeJson(ctx.spWrite, 'mergeJson', `${listPath(ctx.listId)}/fields(guid'${fid}')`,
        { [key]: merges[key] }, 'SP.Field', { fallback: `Could not apply ${key}`, code: 'write' });
      report.fieldMerges.applied++;
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.fieldMerges.failed++;
      failures.push({ key, message: err.message });
    }
  }
  // The step is failed only if every property in it was rejected — a
  // partial success still leaves the column usable.
  if (keys.length && failures.length === keys.length) {
    throw new SpFileError(failures.map((f) => f.message).join(' '), { code: 'write' });
  }
  for (const f of failures) {
    report.warnings.push(`‘${internalName}’: ${f.key} was not applied — ${f.message}`);
  }
  return { failed: failures.map((f) => f.key) };
}

async function runFieldBase(step, ctx) {
  const { displayName, required } = step.payload;
  return writeJson(ctx.spWrite, 'mergeJson', `${listPath(ctx.listId)}/fields/getbyinternalnameortitle('Title')`,
    { Title: displayName, Required: !!required }, 'SP.Field',
    { fallback: 'Could not update the Title column', code: 'write' });
}

async function runViewUpsert(step, ctx, report) {
  const {
    title, viewId: matchedViewId, fields: wanted, viewQuery, rowLimit, paged, defaultView,
    customFormatter, jsLink, viewTypeKind, scope, aggregations, aggregationsStatus,
    tabularView, mobileView, mobileDefaultView, includeRootFolder, viewData, viewJoins,
  } = step.payload;
  let viewId = matchedViewId;
  // The plan-time probe cannot see a view SharePoint creates as part of
  // list.create itself (its own default "All Items", born after the probe
  // ran) — match the live re-read (ctx.targetViews) by title, falling back
  // to the target's own default view when the source view is itself the
  // default, so this upserts that view instead of posting a duplicate.
  if (!viewId) {
    const targetViews = ctx.targetViews || [];
    const found = targetViews.find((tv) => String(tv.title).toLowerCase() === String(title).toLowerCase())
      || (defaultView ? targetViews.find((tv) => tv.defaultView) : null);
    if (found) viewId = found.id;
  }
  let created = false;

  if (viewId) {
    await writeJson(ctx.spWrite, 'mergeJson', `${listPath(ctx.listId)}/views(guid'${viewId}')`,
      // No Title: a view matched as the target's (localized) default view
      // keeps its own name — the consent only promises rebuilt columns.
      { ViewQuery: viewQuery || '', RowLimit: rowLimit ?? 30, Paged: paged !== false },
      'SP.View', { fallback: `Could not update the view ‘${title}’`, code: 'write' });
  } else {
    const body = {
      Title: title, PersonalView: false, ViewQuery: viewQuery || '', RowLimit: rowLimit ?? 30,
      Paged: paged !== false, DefaultView: false,
    };
    // ViewTypeKind only when it says something HTML (1) doesn't already —
    // some tenants reject an explicit 1, and null means "not captured".
    if (viewTypeKind != null && viewTypeKind !== 1) body.ViewTypeKind = viewTypeKind;
    const data = await writeJson(ctx.spWrite, 'postJson', `${listPath(ctx.listId)}/views`, body,
      'SP.View', { fallback: `Could not create the view ‘${title}’`, code: 'write' });
    viewId = data.Id;
    created = true;
  }

  // Re-check the wanted columns against the live (re-read) field set — a
  // plan-time guess about which base columns a fresh list carries can be
  // wrong; what's actually there now is the ground truth.
  const available = new Set((ctx.targetFields || []).map((f) => f.internalName));
  const toAdd = [];
  for (const name of wanted || []) {
    if (available.has(name)) toAdd.push(name);
    else report.warnings.push(`View ‘${title}’: column ‘${name}’ is not on the target — left out.`);
  }

  let previous = [];
  try {
    const vf = await ctx.client.get(`${listPath(ctx.listId)}/views(guid'${viewId}')/viewfields`);
    previous = vf?.Items?.results ?? vf?.Items ?? [];
  } catch { /* nothing to restore to if this read itself fails */ }

  await writeJson(ctx.spWrite, 'postJson', `${listPath(ctx.listId)}/views(guid'${viewId}')/viewfields/removeallviewfields`,
    {}, 'SP.View', { fallback: `Could not clear the columns on ‘${title}’`, code: 'write' });

  try {
    for (const name of toAdd) {
      await writeJson(ctx.spWrite, 'postJson',
        `${listPath(ctx.listId)}/views(guid'${viewId}')/viewfields/addviewfield('${odataPathLiteral(name)}')`,
        {}, 'SP.View', { fallback: `Could not add column ‘${name}’ to ‘${title}’`, code: 'write' });
    }
  } catch (err) {
    if (isExpiredSession(err)) throw err;
    // Restore whatever the view carried before this step touched it —
    // best-effort; the step still fails either way.
    try {
      await ctx.spWrite.postJson(`${listPath(ctx.listId)}/views(guid'${viewId}')/viewfields/removeallviewfields`, {});
      for (const name of previous) {
        await ctx.spWrite.postJson(
          `${listPath(ctx.listId)}/views(guid'${viewId}')/viewfields/addviewfield('${odataPathLiteral(name)}')`, {},
        );
      }
    } catch { /* best effort restore */ }
    throw err;
  }

  // One optional MERGE for the non-default view properties, sent after the
  // columns are settled so a rejection here never costs the view its
  // fields — only properties that differ from SharePoint's own default are
  // sent, and a rejection is a warning, never a failed step.
  const propsMerge = {};
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
      await writeJson(ctx.spWrite, 'mergeJson', `${listPath(ctx.listId)}/views(guid'${viewId}')`, propsMerge, 'SP.View', {});
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.warnings.push(`View ‘${title}’: some view properties were not applied — ${err.message}`);
    }
  }

  const optionalMerge = async (body, label) => {
    try {
      await writeJson(ctx.spWrite, 'mergeJson', `${listPath(ctx.listId)}/views(guid'${viewId}')`, body, 'SP.View', {});
    } catch (err) {
      if (isExpiredSession(err)) throw err;
      report.warnings.push(`View ‘${title}’: ${label} — ${err.message}`);
    }
  };
  if (customFormatter) await optionalMerge({ CustomFormatter: customFormatter }, 'the column-formatting JSON was not applied');
  if (jsLink) await optionalMerge({ JSLink: jsLink }, 'JSLink was not applied');
  if (defaultView) await optionalMerge({ DefaultView: true }, 'could not be set as the default view');

  return { id: viewId, created };
}

async function runListValidation(step, ctx) {
  const { validationFormula, validationMessage } = step.payload;
  return writeJson(ctx.spWrite, 'mergeJson', listPath(ctx.listId),
    { ValidationFormula: validationFormula, ValidationMessage: validationMessage || '' }, 'SP.List',
    { fallback: 'Could not apply the validation formula', code: 'write' });
}

export const STEP_RUNNERS = {
  'list.create': runListCreate,
  'list.adopt': runListAdopt,
  'list.settings': runListSettings,
  'ct.attach': runCtAttach,
  'field.create': runFieldCreate,
  'field.merge': runFieldMerge,
  'field.base': runFieldBase,
  'view.upsert': runViewUpsert,
  'list.validation': runListValidation,
};

// ---- report tallying -----------------------------------------------------
// One place that reads a finalized step (any status) and folds it into the
// report's summary counters — so success (inside a runner, via early
// returns above) and failure (caught generically below) never disagree
// about what counts as "added" vs "failed".

function recordSkippedFieldId(step, ctx) {
  if (step.kind !== 'field.create' || step.status !== 'skipped') return;
  const name = step.payload?.field?.internalName;
  if (!name) return;
  const found = (ctx.targetFields || []).find((f) => f.internalName === name);
  if (found) ctx.fieldIdMap[name] = found.id;
}

function tallyStep(step, report) {
  switch (step.kind) {
    case 'field.create': {
      const name = step.payload?.field?.internalName;
      if (step.status === 'done') report.fields.added++;
      else if (step.status === 'skipped') report.fields.skipped++;
      else if (step.status === 'failed') report.fields.failed.push({ internalName: name, error: step.error });
      break;
    }
    case 'view.upsert': {
      if (step.status === 'done') {
        if (step.result?.created) report.views.added++;
        else report.views.updated++;
      } else if (step.status === 'failed') {
        report.views.failed.push({ title: step.payload?.title, error: step.error });
      }
      break;
    }
    case 'ct.attach': {
      if (step.status === 'done') report.contentTypes.attached++;
      else if (step.status === 'failed') report.contentTypes.failed++;
      else if (step.status === 'skipped') report.contentTypes.skipped++;
      break;
    }
    case 'list.validation': {
      report.validation.applied = step.status === 'done';
      if (step.status === 'failed') report.validation.error = step.error;
      break;
    }
    default: break;
  }
}

// ---- runner --------------------------------------------------------------

// ctx: { client, spWrite, onStep, signal, listId?, fieldIdMap?, targetFields?,
//        targetListsByTitle? } — everything but client/spWrite is optional;
// runPlan fills it in as it goes (and, for a retry plan, from
// plan.existingListId — see the brief's retry contract).
export async function runPlan(plan, ctx = {}) {
  if (plan.existingListId && !ctx.listId) ctx.listId = plan.existingListId;
  ctx.fieldIdMap = ctx.fieldIdMap || {};
  ctx.targetFields = ctx.targetFields || [];
  ctx.targetListsByTitle = ctx.targetListsByTitle || null;
  ctx.targetViews = ctx.targetViews || [];

  const report = newReport(plan);
  const steps = report.steps;
  const total = steps.length;
  const byId = new Map(steps.map((s) => [s.id, s]));

  // A retry (or a resumed/adopted list) already has a listId before the
  // loop starts — prime the reads once here instead of waiting for a
  // list.create/list.adopt step that may not exist in this plan. A read
  // failure here (a 401, or anything else) means nothing that follows can
  // be trusted, so the run stops before touching a single step.
  if (ctx.listId && !ctx.targetFields.length) {
    const ok = await safeRead(ctx, report, () => readTargetState(ctx));
    if (!ok) return report;
  }
  let fieldsRereadForViews = false;

  const isBlockedByDeps = (s) => (s.dependsOn || []).some((depId) => {
    const dep = byId.get(depId);
    return dep && !dep.optional && (dep.status === 'failed' || dep.status === 'blocked');
  });

  for (let i = 0; i < steps.length; i++) {
    const s = steps[i];

    if (s.status === 'skipped' || s.status === 'failed' || s.status === 'blocked') {
      recordSkippedFieldId(s, ctx);
      tallyStep(s, report);
      ctx.onStep?.(s, i, total);
      continue;
    }

    if (ctx.signal?.aborted) { report.aborted = 'user'; break; }

    if (isBlockedByDeps(s)) {
      s.status = 'blocked';
      tallyStep(s, report);
      ctx.onStep?.(s, i, total);
      continue;
    }

    // The pre-view field/view re-read must succeed before any view step
    // runs — a stale field set could add a view column that doesn't
    // actually exist on the target (or miss one that does), so a failure
    // here stops the run rather than falling back to what the plan-time
    // probe guessed. `s` is still 'planned' at this point, same as
    // everything after it.
    if (s.kind === 'view.upsert' && !fieldsRereadForViews) {
      fieldsRereadForViews = true;
      const ok = await safeRead(ctx, report, async () => {
        ctx.targetFields = await readTargetFields(ctx);
        ctx.targetViews = await readTargetViews(ctx);
      });
      if (!ok) return report;
    }

    s.status = 'running';
    ctx.onStep?.(s, i, total);

    let aborted = false;
    try {
      const runner = STEP_RUNNERS[s.kind];
      if (!runner) throw new SpFileError(`No executor for step kind ‘${s.kind}’.`, { code: 'internal' });
      const result = await runner(s, ctx, report);
      s.status = 'done';
      s.result = result ?? null;
    } catch (err) {
      // A dependency that is missing only at run time (a dependent lookup's
      // primary has no id) is the same fact the planner calls 'blocked'.
      s.status = err?.code === 'blocked' ? 'blocked' : 'failed';
      s.error = err?.message || String(err);
      if (isExpiredSession(err)) {
        report.aborted = 'auth';
        aborted = true;
      } else if (s.optional) {
        report.warnings.push(`${s.label}: ${s.error}`);
      }
    }

    recordSkippedFieldId(s, ctx);
    tallyStep(s, report);
    ctx.onStep?.(s, i, total);
    if (aborted) break;

    // The list now exists (or was adopted) — everything after this depends
    // on what's really on the target, not what the plan-time probe guessed.
    // Re-read it OUTSIDE the write's own try/catch, so a read failure here
    // aborts the run cleanly instead of being blamed on the list.create/
    // list.adopt step that actually succeeded (see the commit-boundary note
    // in runListCreate — the same principle, for reads instead of writes).
    if (s.status === 'done' && (s.kind === 'list.create' || s.kind === 'list.adopt')) {
      const ok = await safeRead(ctx, report, () => readTargetState(ctx));
      if (!ok) return report;
    }
  }

  return report;
}

// probe (probeTarget) → plan (buildApplyPlan) → run (runPlan) — dryRun stops
// after the plan, so the dialog's preview and the live run share every read
// and every planning decision, and only the run performs a write.
export async function applyListSchema({
  doc, client, spWrite, options = {}, onStep, signal,
} = {}) {
  const probe = await probeTarget(client, { title: options.title, doc });
  const plan = buildApplyPlan(doc, { ...options, targetWebUrl: options.targetWebUrl || client.webUrl() }, probe);
  if (options.dryRun) {
    const report = newReport(plan);
    report.dryRun = true;
    return report;
  }
  return runPlan(plan, { client, spWrite, onStep, signal });
}
