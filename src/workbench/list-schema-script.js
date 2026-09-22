// List schema — provisioning script emitters.
//
// Each emitter walks the SAME Plan buildApplyPlan() produces (an empty probe
// stands in for "nothing exists yet on the target" when there's no live
// target to probe — the common case: this is a "copy as script" action off
// the Schema tab's Export menu, not the Copy-to dialog's dry run) so the
// generated script and the future executor can never drift onto different
// step orders. Mirrors scriptgen.js's three-dialect shape.

import {
  buildApplyPlan, scrubSchemaXml, textFallbackXml, parentContentTypeId, isBuiltinParent, lookupMapGet,
} from './list-schema.js';

const TEMPLATE_NAMES = { 100: 'GenericList', 101: 'DocumentLibrary' };

// Without a real probe the script assumes a target where every lookup list
// the source points at exists by the same (or mapped) title — the script
// resolves their ids when it runs, and fails loudly there if one is absent,
// rather than silently emitting every lookup column as skipped. Content
// types get the same treatment: every non-builtin parent id the source
// carries is assumed available, so a ct.attach step is planned (not failed)
// with no live probe to check against.
function planFor(doc, opts = {}) {
  const lookupMap = opts.lookupMap || {};
  const assumed = [...new Set((doc?.fields || [])
    .filter((f) => f.lookupList && !f.isSelfLookup)
    .map((f) => lookupMapGet(lookupMap, f.lookupList) ?? lookupMapGet(lookupMap, f.lookupListId) ?? f.lookupList))]
    .map((title) => ({ title }));
  const assumedContentTypes = [...new Set((doc?.contentTypes || [])
    .map((ct) => ct.parentId || parentContentTypeId(ct.id))
    .filter((parentId) => !isBuiltinParent(parentId)))]
    .map((id) => ({ id }));
  return buildApplyPlan(doc, {
    title: opts.title, description: opts.description,
    lookupMap, missingLookup: opts.missingLookup,
  }, opts.probe || { targetLists: assumed, availableContentTypes: assumedContentTypes });
}

const isDone = (step) => !['skipped', 'failed', 'blocked'].includes(step.status);

// The field's own scrubbed XML. A script has no probe, so lookup bindings are
// left as placeholder tokens that the emitted code swaps for ids it reads on
// the target at run time: the lookup list by title (or the new list itself
// for a self-lookup), and a dependent lookup's primary column by name.
const LIST_TOKEN = '{lookuplist}';
const PRIMARY_TOKEN = '{primaryfield}';

function fieldXml(step) {
  const f = step.payload.field;
  if (step.payload.asText) return textFallbackXml(f);
  const refs = step.refs || {};
  return scrubSchemaXml(f.schemaXml, {
    fieldType: f.type,
    lookupListId: refs.lookupListId ? 'lookuplist' : null,
    primaryFieldId: refs.primaryFieldId ? 'primaryfield' : null,
  });
}

// SP.AddFieldOptions.AddFieldInternalNameHint — without it SharePoint may
// derive the internal name from DisplayName (see list-schema.js). Fallback
// only: every field.create step carries its own resolved `options` value
// (list-schema.js buildApplyPlan) — that payload is the source of truth.
const FIELD_OPTIONS = 8;

// The one encoder for every document-derived string in the PS emitter —
// titles, internal names, XML, CAML, JSON formatter values, hashtable
// values, field arrays. Single-quoted PowerShell strings never interpolate,
// so a formatter's `$schema` (or any other `$`/backtick content) survives
// verbatim; only a literal single quote needs escaping (doubled).
const psEsc = (s) => String(s ?? '').replaceAll("'", "''");
const psLit = (s) => `'${psEsc(s)}'`;

function psCsomValue(v) {
  if (typeof v === 'boolean') return v ? '$true' : '$false';
  if (typeof v === 'number') return String(v);
  return psLit(v);
}

function psHashtable(obj) {
  const parts = Object.entries(obj).map(([k, v]) => {
    if (typeof v === 'boolean') return `${k}=$${v}`;
    if (typeof v === 'number') return `${k}=${v}`;
    return `${k}=${psLit(v)}`;
  });
  return `@{ ${parts.join('; ')} }`;
}
const psArray = (arr) => `@(${(arr || []).map((s) => psLit(s)).join(', ')})`;

// Settings and validation go through CSOM ($l = Get-PnPList; $l.Prop = …;
// $l.Update(); Invoke-PnPQuery) instead of Set-PnPList switches: the switch
// surface doesn't cover every captured property (ListExperienceOptions,
// ReadSecurity/WriteSecurity, DraftVersionVisibility…), CSOM does, uniformly.
function pushCsomBlock(lines, plan, props) {
  const entries = Object.entries(props).filter(([, v]) => v != null);
  if (!entries.length) return;
  lines.push(`$l = Get-PnPList -Identity ${psLit(plan.title)}`);
  for (const [k, v] of entries) lines.push(`$l.${k} = ${psCsomValue(v)}`);
  lines.push('$l.Update()');
  lines.push('Invoke-PnPQuery');
}

export function toPnpPowerShellProvisioning(doc, opts = {}) {
  const plan = planFor(doc, opts);
  const connect = `Connect-PnPOnline -Url "${opts.targetWebUrl || 'https://tenant.sharepoint.com/sites/yoursite'}" -Interactive`;
  const lines = ['# PnP.PowerShell provisioning', `# ${plan.title}`, connect, `$list = ${psLit(plan.title)}`, ''];
  for (const step of plan.steps) {
    if (!isDone(step)) {
      lines.push(`# SKIP ${step.label}${step.error ? ` — ${step.error}` : ''}`);
      continue;
    }
    switch (step.kind) {
      case 'list.create': {
        const template = TEMPLATE_NAMES[step.payload.baseTemplate] || step.payload.baseTemplate;
        lines.push(`New-PnPList -Title ${psLit(step.payload.title)} -Template ${template}`
          + `${step.payload.contentTypesEnabled ? ' -EnableContentTypes' : ''}`);
        break;
      }
      case 'list.adopt':
        lines.push(`# Using existing list "${step.payload.title}"`);
        break;
      case 'list.settings':
        pushCsomBlock(lines, plan, { ...step.payload.groupA, ...step.payload.groupB });
        break;
      case 'ct.attach':
        lines.push(`Add-PnPContentTypeToList -List ${psLit(plan.title)} -ContentType ${psLit(step.payload.contentTypeId)}`);
        break;
      case 'field.create': {
        let xmlExpr = psLit(fieldXml(step));
        const refs = step.refs || {};
        if (refs.lookupListId) {
          const listIdentity = refs.lookupListId.self ? plan.title : (refs.lookupListId.id || refs.lookupListId.list);
          xmlExpr = `(${xmlExpr}).Replace('${LIST_TOKEN}', ('{' + (Get-PnPList -Identity ${psLit(listIdentity)}).Id + '}'))`;
        }
        if (refs.primaryFieldId) {
          xmlExpr = `(${xmlExpr}).Replace('${PRIMARY_TOKEN}', ('{' + (Get-PnPField -List ${psLit(plan.title)} -Identity ${psLit(refs.primaryFieldId.field)}).Id + '}'))`;
        }
        lines.push(`Add-PnPFieldFromXml -List ${psLit(plan.title)} -FieldXml ${xmlExpr}`);
        break;
      }
      case 'field.merge':
        lines.push(`Set-PnPField -List ${psLit(plan.title)} -Identity ${psLit(step.payload.internalName)} -Values ${psHashtable(step.payload.merges)}`);
        break;
      case 'field.base':
        lines.push(`Set-PnPField -List ${psLit(plan.title)} -Identity 'Title' -Values ${psHashtable({ Title: step.payload.displayName, Required: step.payload.required })}`);
        break;
      case 'view.upsert': {
        const titleLit = psLit(step.payload.title);
        lines.push(`$v = Get-PnPView -List $list -Identity ${titleLit} -ErrorAction SilentlyContinue`);
        if (step.payload.defaultView) {
          lines.push('if (-not $v) { $v = Get-PnPView -List $list | Where-Object DefaultView }');
        }
        lines.push('if ($v) {');
        lines.push(`  Set-PnPView -List $list -Identity $v -Fields ${psArray(step.payload.fields)} -Values ${psHashtable({ ViewQuery: step.payload.viewQuery, RowLimit: step.payload.rowLimit, Paged: step.payload.paged })}`);
        lines.push('} else {');
        lines.push(`  Add-PnPView -List $list -Title ${titleLit} -Fields ${psArray(step.payload.fields)} -Query ${psLit(step.payload.viewQuery)} -RowLimit ${step.payload.rowLimit}${step.payload.paged ? ' -Paged' : ''} -SetAsDefault:$${step.payload.defaultView ? 'true' : 'false'}`);
        lines.push('}');
        break;
      }
      case 'list.validation':
        pushCsomBlock(lines, plan, {
          ValidationFormula: step.payload.validationFormula,
          ValidationMessage: step.payload.validationMessage,
        });
        break;
      default:
        break;
    }
  }
  return lines.join('\n');
}

export function toPnpjs2Provisioning(doc, opts = {}) {
  const plan = planFor(doc, opts);
  const lines = [
    '// PnPjs 2.x — paste into the DCSPad JS pane (pnpjs2 framework enabled)',
    `// ${plan.title}`, '',
  ];
  for (const step of plan.steps) {
    if (!isDone(step)) {
      lines.push(`// SKIP ${step.label}${step.error ? ` — ${step.error}` : ''}`);
      continue;
    }
    switch (step.kind) {
      case 'list.create':
        lines.push(`const newList = (await sp.web.lists.add(${JSON.stringify(step.payload.title)}, `
          + `${JSON.stringify(step.payload.description || '')}, ${step.payload.baseTemplate}, `
          + `${step.payload.contentTypesEnabled})).list;`);
        lines.push('const newListId = (await newList.select("Id")()).Id;');
        break;
      case 'list.adopt':
        lines.push(`const newList = sp.web.lists.getById(${JSON.stringify(step.payload.listId)});`);
        lines.push(`const newListId = ${JSON.stringify(step.payload.listId)};`);
        break;
      case 'list.settings': {
        const settings = { ...step.payload.groupA, ...step.payload.groupB };
        if (Object.keys(settings).length) lines.push(`await newList.update(${JSON.stringify(settings, null, 2)});`);
        break;
      }
      case 'ct.attach':
        lines.push(`await newList.contentTypes.addAvailableContentType(${JSON.stringify(step.payload.contentTypeId)});`);
        break;
      case 'field.create': {
        // A block per column: the lookup/primary ids are block-scoped consts.
        const refs = step.refs || {};
        let xmlExpr = JSON.stringify(fieldXml(step));
        const pre = [];
        if (refs.lookupListId) {
          if (refs.lookupListId.self) {
            pre.push('  const lookupId = newListId;');
          } else if (refs.lookupListId.id) {
            // A GUID lookupMap binding — no read needed, the id is literal.
            pre.push(`  const lookupId = ${JSON.stringify(refs.lookupListId.id)};`);
          } else {
            pre.push(`  const lookupId = (await sp.web.lists.getByTitle(${JSON.stringify(refs.lookupListId.list)}).select("Id")()).Id;`);
          }
          xmlExpr += `.replace(${JSON.stringify(LIST_TOKEN)}, \`{\${lookupId}}\`)`;
        }
        if (refs.primaryFieldId) {
          pre.push(`  const primaryId = (await newList.fields.getByInternalNameOrTitle(${JSON.stringify(refs.primaryFieldId.field)}).select("Id")()).Id;`);
          xmlExpr += `.replace(${JSON.stringify(PRIMARY_TOKEN)}, \`{\${primaryId}}\`)`;
        }
        lines.push('{', ...pre,
          `  await newList.fields.createFieldAsXml({ SchemaXml: ${xmlExpr}, Options: ${step.payload.options ?? FIELD_OPTIONS} });`, '}');
        break;
      }
      case 'field.merge':
        lines.push(`await newList.fields.getByInternalNameOrTitle(${JSON.stringify(step.payload.internalName)}).update(${JSON.stringify(step.payload.merges)});`);
        break;
      case 'field.base':
        lines.push(`await newList.fields.getByInternalNameOrTitle("Title").update(${JSON.stringify({ Title: step.payload.displayName, Required: step.payload.required })});`);
        break;
      case 'view.upsert': {
        // Block-scoped: every view gets its own `view` binding. Upsert by
        // title, falling back to the list's own default view when the
        // source view is the default (its title may be localized on the
        // target) — then reconcile fields and, when the source is default,
        // the target's DefaultView flag too.
        const titleJson = JSON.stringify(step.payload.title);
        const settingsJson = JSON.stringify({
          ViewQuery: step.payload.viewQuery, RowLimit: step.payload.rowLimit, Paged: step.payload.paged,
        });
        lines.push('{');
        lines.push('  let existing = null;');
        lines.push(`  try { existing = await newList.views.getByTitle(${titleJson})(); } catch { existing = null; }`);
        if (step.payload.defaultView) {
          lines.push('  if (!existing) existing = await newList.defaultView();');
        }
        lines.push('  let view;');
        lines.push('  if (existing) {');
        lines.push('    view = newList.views.getByTitle(existing.Title);');
        lines.push(`    await view.update(${settingsJson});`);
        lines.push('  } else {');
        lines.push(`    view = (await newList.views.add(${titleJson}, false, ${settingsJson})).view;`);
        lines.push('  }');
        lines.push('  await view.fields.removeAll();');
        for (const name of step.payload.fields) lines.push(`  await view.fields.add(${JSON.stringify(name)});`);
        if (step.payload.defaultView) {
          lines.push('  await view.update({ DefaultView: true });');
        }
        lines.push('}');
        break;
      }
      case 'list.validation':
        lines.push(`await newList.update(${JSON.stringify({
          ValidationFormula: step.payload.validationFormula, ValidationMessage: step.payload.validationMessage,
        })});`);
        break;
      default:
        break;
    }
  }
  return lines.join('\n');
}

// A one-liner around the console tool this feature ports — the interchange
// document works both ways, so the fastest path for someone already in the
// SPUtils console is just handing the exported doc back to it.
export function toSpUtilsCall(doc, opts = {}) {
  const args = { title: opts.title || doc?.list?.title, dryRun: !!opts.dryRun };
  if (opts.lookupMap && Object.keys(opts.lookupMap).length) args.lookupMap = opts.lookupMap;
  return [
    '// Paste into a console where DCSPad SP Utilities (SPUtils) is loaded.',
    `const schema = ${JSON.stringify(doc, null, 2)};`,
    `await SPUtils.createListFromSchema(schema, ${JSON.stringify(args, null, 2)});`,
  ].join('\n');
}

export const PROVISION_FORMATS = [
  ['Copy as PnP.PowerShell (provision)', toPnpPowerShellProvisioning],
  ['Copy as PnPjs 2 (provision)', toPnpjs2Provisioning],
];
