// SP Workbench — List Schema (stage 1a, slice 1: plumbing + read side).
//
// Own suite, own mock webs (/sites/schema source, /sites/target apply
// target — see mock-data.js), and its own stubbed live endpoints; layout
// mirrors workbench-edit.mjs. Slice 1 ships capture + the Schema tab (read
// only — "Copy to…" is a gated no-op until slice 2's dialog/executor land),
// so this suite covers: the pure list-schema.js/list-schema-script.js core,
// the Schema tab in mock mode, and the stubbed-live capture reads. Dialog/
// apply checks belong to slice 2.

import { readFileSync } from 'fs';
import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const WB_URL = process.env.DCSPAD_WORKBENCH_URL
  || APP_URL.replace(/index\.html.*$/, 'workbench.html');

const browser = await launchBrowser();

// ---- Pure (13): list-schema.js + list-schema-script.js, no page state ----

const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await page.goto(WB_URL);
await page.waitForSelector('.wb-home-cards');

await check('pure: a v1 document normalizes to v2 with defaults; the wrong kind is refused', () =>
  page.evaluate(async () => {
    const { normalizeSchemaDoc, SCHEMA_KIND } = await import('/src/workbench/list-schema.js');
    const v1 = {
      kind: SCHEMA_KIND, version: 1,
      source: { listTitle: 'Requests' },
      list: { title: 'Requests', baseTemplate: 100, enableVersioning: true },
      fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false }],
      views: [{ title: 'All Items', fields: ['Title'] }],
      contentTypes: [], warnings: [],
    };
    const v2 = normalizeSchemaDoc(v1);
    let refused = false;
    try { normalizeSchemaDoc({ kind: 'not-a-schema' }); } catch (e) { refused = e?.code === 'bad-schema'; }
    return v2.version === 2 && v2.list.baseTemplate === 100 && v2.list.onQuickLaunch === false
      && v2.fields[0].internalName === 'Title' && v2.fields[0].hidden === false
      && v2.views[0].fields[0] === 'Title' && refused;
  }));

await check('pure: isCustomField reads FromBaseType, CanBeDeleted, Hidden, and Computed', () =>
  page.evaluate(async () => {
    const { isCustomField } = await import('/src/workbench/list-schema.js');
    const base = { FromBaseType: false, CanBeDeleted: true, Hidden: false, TypeAsString: 'Text' };
    return isCustomField(base) === true
      && isCustomField({ ...base, FromBaseType: true }) === false
      && isCustomField({ ...base, CanBeDeleted: false }) === false
      && isCustomField({ ...base, Hidden: true }) === false
      && isCustomField({ ...base, TypeAsString: 'Computed' }) === false;
  }));

await check('pure: scrubSchemaXml strips source attrs, keeps User List=UserInfo, rebinds a lookup, strips FieldRef ids, throws on bad xml', () =>
  page.evaluate(async () => {
    const { scrubSchemaXml } = await import('/src/workbench/list-schema.js');
    const lookupXml = '<Field ID="{a}" SourceID="{b}" ColName="c" RowOrdinal="0" Version="1" WebId="{d}"'
      + ' List="{old}" Sealed="TRUE" Customization="x" Name="Client" Type="Lookup" DisplayName="Client"'
      + ' Indexed="TRUE" EnforceUniqueValues="TRUE" />';
    const out = scrubSchemaXml(lookupXml, { lookupListId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', fieldType: 'Lookup' });
    const userXml = '<Field ID="{a}" Name="Owner" Type="User" List="UserInfo" />';
    const outUser = scrubSchemaXml(userXml, { fieldType: 'User' });
    const calcXml = '<Field ID="{a}" Name="Total" Type="Calculated"><FieldRefs>'
      + '<FieldRef ID="{sourceguid}" Name="Budget"/></FieldRefs></Field>';
    const outCalc = scrubSchemaXml(calcXml, { fieldType: 'Calculated' });
    let threw = false;
    try { scrubSchemaXml('<Field this is not xml', {}); } catch (e) { threw = e?.code === 'bad-xml'; }
    return !out.includes('ID=') && !out.includes('SourceID') && !out.includes('RowOrdinal')
      && !out.includes('Sealed=') && !out.includes('Customization') && !out.includes('WebId')
      && out.includes('List="{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}"')
      && !out.includes('Indexed=') && !out.includes('EnforceUniqueValues=')
      && outUser.includes('List="UserInfo"')
      && !outCalc.includes('sourceguid') && outCalc.includes('Name="Budget"')
      && threw;
  }));

await check('pure: field creation order is stable within a tier (plain → lookup → dependent → calculated)', () =>
  page.evaluate(async () => {
    const { orderFields } = await import('/src/workbench/list-schema.js');
    const fields = [
      { internalName: 'Calc1', type: 'Calculated' },
      { internalName: 'Plain1', type: 'Text' },
      { internalName: 'Dep1', type: 'Lookup', isDependentLookup: true },
      { internalName: 'Look1', type: 'Lookup' },
      { internalName: 'Plain2', type: 'Number' },
    ];
    return orderFields(fields).map((f) => f.internalName).join(',') === 'Plain1,Plain2,Look1,Dep1,Calc1';
  }));

await check('pure: buildApplyPlan skips a same-type existing column and fails a mismatched one', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'Same', displayName: 'Same', type: 'Text', custom: true, schemaXml: '<Field Name="Same" Type="Text"/>' },
        { internalName: 'Diff', displayName: 'Diff', type: 'Number', custom: true, schemaXml: '<Field Name="Diff" Type="Number"/>' },
      ],
      views: [],
    });
    const probe = {
      existingList: null,
      existingFields: [
        { internalName: 'Same', title: 'Same', typeAsString: 'Text' },
        { internalName: 'Diff', title: 'Diff', typeAsString: 'Text' },
      ],
    };
    const plan = buildApplyPlan(doc, { title: 'Requests' }, probe);
    const same = plan.steps.find((s) => s.id === 'field:Same');
    const diff = plan.steps.find((s) => s.id === 'field:Diff');
    return same.status === 'skipped' && diff.status === 'failed' && diff.error.includes('Number');
  }));

await check('pure: an existing-title collision refuses by default and adopts on resume', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({ list: { title: 'Requests', baseTemplate: 100 }, fields: [], views: [] });
    const probe = { existingList: { id: 'xyz', title: 'Requests', baseTemplate: 100, contentTypesEnabled: false } };
    let threwCode = null;
    try { buildApplyPlan(doc, { title: 'Requests', existing: 'fail' }, probe); } catch (e) { threwCode = e.code; }
    const resumed = buildApplyPlan(doc, { title: 'Requests', existing: 'resume' }, probe);
    return threwCode === 'exists' && resumed.steps[0].kind === 'list.adopt' && resumed.existingListId === 'xyz';
  }));

await check('pure: a lookupMap entry rebinds a lookup column to its mapped target list', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [{
        internalName: 'Client', displayName: 'Client', type: 'Lookup', custom: true,
        lookupList: 'Clients', schemaXml: '<Field Name="Client" Type="Lookup"/>',
      }],
      views: [],
    });
    const probe = { existingList: null, existingFields: [], targetLists: [{ id: 'target-clients', title: 'Customers' }] };
    const plan = buildApplyPlan(doc, { title: 'Requests', lookupMap: { Clients: 'Customers' } }, probe);
    const step = plan.steps.find((s) => s.id === 'field:Client');
    return step.refs.lookupListId?.list === 'Customers' && step.status === 'planned';
  }));

await check('pure: a missing lookup target is skipped by default, made text under the text policy; a dependent of the skipped primary is blocked, not independently skipped', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        {
          id: 'p1', internalName: 'Region', displayName: 'Region', type: 'Lookup', custom: true,
          lookupList: 'Regions', schemaXml: '<Field Name="Region" Type="Lookup"/>',
        },
        {
          internalName: 'RegionCode', displayName: 'Region code', type: 'Lookup', custom: true,
          isDependentLookup: true, primaryFieldId: 'p1', lookupList: 'Regions',
          schemaXml: '<Field Name="RegionCode" Type="Lookup"/>',
        },
      ],
      views: [],
    });
    const probe = { existingList: null, existingFields: [], targetLists: [] };
    const skipped = buildApplyPlan(doc, { title: 'Requests' }, probe);
    const skipStep = skipped.steps.find((s) => s.id === 'field:Region');
    const dependentStep = skipped.steps.find((s) => s.id === 'field:RegionCode');
    const texted = buildApplyPlan(doc, { title: 'Requests', missingLookup: 'text' }, probe);
    const textStep = texted.steps.find((s) => s.id === 'field:Region');
    return skipStep.status === 'skipped' && skipStep.skipReason === 'lookup-target-missing'
      && dependentStep.status === 'blocked'
      && textStep.status === 'planned' && textStep.payload.asText === true
      && texted.warnings.some((w) => w.includes('Region'));
  }));

await check('pure: a content type attaches only when the source enables them and the target has it available', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc, parentContentTypeId } = await import('/src/workbench/list-schema.js');
    const parent = '0x0100442912F2B6C7409A8FF25CE5504F1FD';
    const listCtId = `${parent}00${'a'.repeat(32)}`;
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100, contentTypesEnabled: true },
      fields: [], views: [],
      contentTypes: [{ id: listCtId, name: 'Request', parentId: parentContentTypeId(listCtId) }],
    });
    const disabled = { ...doc, list: { ...doc.list, contentTypesEnabled: false } };
    const notEnabled = buildApplyPlan(disabled, { title: 'Requests' }, { existingList: null, existingFields: [] });
    const noCtStep = !notEnabled.steps.some((s) => s.kind === 'ct.attach');

    const unavailable = buildApplyPlan(doc, { title: 'Requests' }, {
      existingList: null, existingFields: [], availableContentTypes: [],
    });
    const failedStep = unavailable.steps.find((s) => s.kind === 'ct.attach');

    const available = buildApplyPlan(doc, { title: 'Requests' }, {
      existingList: null, existingFields: [], availableContentTypes: [{ id: parent, name: 'Request' }],
    });
    const plannedStep = available.steps.find((s) => s.kind === 'ct.attach');

    return noCtStep && failedStep?.status === 'failed' && plannedStep?.status === 'planned';
  }));

await check('pure: apply-plan step order is list → settings → content types → fields → merges → title → views → validation', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc, parentContentTypeId } = await import('/src/workbench/list-schema.js');
    const parent = '0x0100442912F2B6C7409A8FF25CE5504F1FD';
    const listCtId = `${parent}00${'b'.repeat(32)}`;
    const doc = buildSchemaDoc({
      list: {
        title: 'Requests', baseTemplate: 100, contentTypesEnabled: true,
        validationFormula: '=[Budget]>0', validationMessage: 'Must be positive',
      },
      fields: [
        { internalName: 'Title', displayName: 'Request Title', type: 'Text', custom: false, required: true, schemaXml: '<Field Name="Title" Type="Text"/>' },
        { internalName: 'Budget', displayName: 'Budget', type: 'Number', custom: true, indexed: true, schemaXml: '<Field Name="Budget" Type="Number" Indexed="TRUE"/>' },
      ],
      views: [{ title: 'All Items', hidden: false, fields: ['Budget'], rowLimit: 30, paged: true, viewQuery: '' }],
      contentTypes: [{ id: listCtId, name: 'Request', parentId: parentContentTypeId(listCtId) }],
    });
    const plan = buildApplyPlan(doc, { title: 'Requests' }, {
      existingList: null, existingFields: [], availableContentTypes: [{ id: parent, name: 'Request' }],
    });
    return plan.steps.map((s) => s.kind).join(',')
      === 'list.create,list.settings,ct.attach,field.create,field.merge,field.base,view.upsert,list.validation';
  }));

await check('pure: buildApplyReport renders fields/views/content-types/warnings/failed-steps sections', () =>
  page.evaluate(async () => {
    const { buildApplyReport } = await import('/src/workbench/list-schema.js');
    const report = {
      title: 'Requests Copy', created: true,
      fields: { added: 3, skipped: 1, failed: [{ internalName: 'X' }] },
      views: { added: 1, updated: 0, failed: [] },
      contentTypes: { attached: 1, skipped: 0, failed: 0 },
      warnings: ['A managed metadata column was skipped.'],
      steps: [{ id: 'field:X', label: 'Add Text column X', status: 'failed', error: 'boom' }],
    };
    const md = buildApplyReport({ report, targetWebUrl: 'https://tenant/sites/target' });
    return md.startsWith('# List schema report')
      && md.includes('Created ‘Requests Copy’')
      && md.includes('## Fields') && md.includes('- Failed: 1')
      && md.includes('## Views') && md.includes('## Content types')
      && md.includes('## Warnings') && md.includes('managed metadata')
      && md.includes('## Failed steps') && md.includes('boom');
  }));

await check('pure: script emitters pin New-PnPList/Add-PnPFieldFromXml and sp.web.lists.add/createFieldAsXml/views.add, with no ID= in the field xml', () =>
  page.evaluate(async () => {
    const { toPnpPowerShellProvisioning, toPnpjs2Provisioning } = await import('/src/workbench/list-schema-script.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [{
        internalName: 'Status', displayName: 'Status', type: 'Choice', custom: true,
        schemaXml: '<Field ID="{x}" Name="Status" Type="Choice" DisplayName="Status"/>',
      }],
      views: [{ title: 'All Items', hidden: false, fields: ['Status'], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const ps = toPnpPowerShellProvisioning(doc, {});
    const js = toPnpjs2Provisioning(doc, {});
    return ps.includes('New-PnPList') && ps.includes('Add-PnPFieldFromXml') && !ps.includes('ID=')
      && js.includes('sp.web.lists.add(') && js.includes('createFieldAsXml(') && js.includes('views.add(');
  }));

await check('pure: the PnPjs script compiles with two views and lookups, and binds lookups at run time instead of skipping them', () =>
  page.evaluate(async () => {
    const { toPnpPowerShellProvisioning, toPnpjs2Provisioning } = await import('/src/workbench/list-schema-script.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        {
          id: 'p1', internalName: 'Client', displayName: 'Client', type: 'Lookup', custom: true,
          lookupList: 'Clients', schemaXml: '<Field Name="Client" Type="Lookup" DisplayName="Client" List="{aa}" ShowField="Title"/>',
        },
        {
          internalName: 'ClientCode', displayName: 'Client code', type: 'Lookup', custom: true,
          isDependentLookup: true, primaryFieldId: 'p1', lookupList: 'Clients',
          schemaXml: '<Field Name="ClientCode" Type="Lookup" DisplayName="Client code" List="{aa}" FieldRef="{p1}" ShowField="Code"/>',
        },
      ],
      views: [
        { title: 'All Items', hidden: false, fields: ['Client'], rowLimit: 30, paged: true, viewQuery: '' },
        { title: 'By client', hidden: false, fields: ['Client', 'ClientCode'], rowLimit: 30, paged: true, viewQuery: '' },
      ],
    });
    const js = toPnpjs2Provisioning(doc, {});
    const ps = toPnpPowerShellProvisioning(doc, {});
    let compiles = true;
    try {
      // eslint-disable-next-line no-new-func
      new (Object.getPrototypeOf(async () => {}).constructor)('sp', js);
    } catch { compiles = false; }
    return compiles
      && !js.includes('// SKIP') && !ps.includes('# SKIP')
      && js.includes('getByTitle("Clients")') && js.includes('Options: 8')
      && js.includes('getByInternalNameOrTitle("Client")')
      && ps.includes("Get-PnPList -Identity 'Clients'") && ps.includes("Get-PnPField -List 'Requests' -Identity 'Client'");
  }));

await check('pure: settingsPayload never sends Description — list.create already supplies it and adopt must not change it', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100, description: 'Has a description' },
      fields: [], views: [],
    });
    const fresh = buildApplyPlan(doc, { title: 'Requests' }, { existingList: null, existingFields: [] });
    const freshSettings = fresh.steps.find((s) => s.id === 'settings').payload.groupA;
    const adopted = buildApplyPlan(doc, { title: 'Requests', existing: 'resume' },
      { existingList: { id: 'x', title: 'Requests' }, targetLists: [{ title: 'Requests' }] });
    const adoptSettings = adopted.steps.find((s) => s.id === 'settings').payload.groupA;
    return !('Description' in freshSettings) && !('Description' in adoptSettings);
  }));

await check('pure: buildApplyPlan refuses an unsupported template (unsupported-template — generic lists and libraries are fine, a survey is not) and a same-title BASE TYPE mismatch (base-type-mismatch, opt-in on probe.existingList.baseType)', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    // 101 (document library) is stage 2 — supported. 102 (Survey) never was.
    const libDoc = buildSchemaDoc({ list: { title: 'Documents', baseTemplate: 101 }, fields: [], views: [] });
    const libPlan = buildApplyPlan(libDoc, { title: 'Documents' }, { targetLists: [] });
    const surveyDoc = buildSchemaDoc({ list: { title: 'Poll', baseTemplate: 102 }, fields: [], views: [] });
    let surveyCode = null;
    try { buildApplyPlan(surveyDoc, { title: 'Poll' }, {}); } catch (e) { surveyCode = e.code; }

    const genericDoc = buildSchemaDoc({ list: { title: 'Requests', baseTemplate: 100 }, fields: [], views: [] });
    let mismatchCode = null;
    try {
      buildApplyPlan(genericDoc, { title: 'Requests' }, {
        existingList: { id: 'x', title: 'Requests', baseTemplate: 101, baseType: 1 },
      });
    } catch (e) { mismatchCode = e.code; }
    // The reverse direction: a library doc refused onto an existing generic list.
    let reverseMismatchCode = null;
    try {
      buildApplyPlan(libDoc, { title: 'Documents' }, {
        existingList: { id: 'y', title: 'Documents', baseTemplate: 100, baseType: 0 },
      });
    } catch (e) { reverseMismatchCode = e.code; }
    // A hand-built probe that never carries BaseType at all (as several
    // other pure fixtures in this file do) must not be treated as a
    // mismatch — the check is opt-in on the probe actually saying so. Even
    // though this probe's baseTemplate (101) would have mismatched under
    // the old template comparison — the check now looks at baseType only.
    const noTemplateInProbe = buildApplyPlan(genericDoc, { title: 'Requests', existing: 'resume' },
      { existingList: { id: 'x', title: 'Requests', baseTemplate: 101 }, targetLists: [{ title: 'Requests' }] });
    // A non-101 BaseType-1 template (a picture library, 109) is eligible —
    // recreated as a standard document library, with a warning naming the
    // source template — and a same-baseType collision against it (also
    // BaseType 1, different template) is NOT a mismatch: only the
    // list/library split matters now, not the exact template number.
    const pictureDoc = buildSchemaDoc({ list: { title: 'Photos', baseTemplate: 109 }, fields: [], views: [] });
    const picturePlan = buildApplyPlan(pictureDoc, { title: 'Photos' }, { targetLists: [] });
    const pictureCreate = picturePlan.steps.find((s) => s.id === 'list');
    const pictureAdopt = buildApplyPlan(pictureDoc, { title: 'Photos', existing: 'resume' },
      { existingList: { id: 'z', title: 'Photos', baseTemplate: 101, baseType: 1 }, targetLists: [{ title: 'Photos' }] });

    return libPlan.steps[0].kind === 'list.create' && surveyCode === 'unsupported-template'
      && mismatchCode === 'base-type-mismatch' && reverseMismatchCode === 'base-type-mismatch'
      && noTemplateInProbe.steps[0].kind === 'list.adopt'
      && pictureCreate.kind === 'list.create' && pictureCreate.payload.baseTemplate === 101
      && picturePlan.warnings.some((w) => w.includes('Recreated as a standard document library (source template 109)'))
      && pictureAdopt.steps[0].kind === 'list.adopt';
  }));

await check('pure: adopting reconciles only attachments/folders, nulls never reach a MERGE, and a dependent lookup on a failed primary is blocked', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: {
        title: 'Requests', baseTemplate: 100, enableVersioning: true, enableAttachments: true,
        enableFolderCreation: true, readSecurity: null, hidden: false,
      },
      fields: [
        { id: 'p1', internalName: 'Region', displayName: 'Region', type: 'Lookup', custom: true, lookupList: 'Regions', schemaXml: '<Field Name="Region" Type="Lookup"/>' },
        { internalName: 'RegionCode', displayName: 'Region code', type: 'Lookup', custom: true, isDependentLookup: true, primaryFieldId: 'p1', lookupList: 'Regions', schemaXml: '<Field Name="RegionCode" Type="Lookup"/>' },
      ],
      views: [],
    });
    // Regions exists on the target, but the target already has a Region
    // column of another type — the primary fails, so the dependent is blocked.
    const fresh = buildApplyPlan(doc, { title: 'Requests Copy' }, {
      targetLists: [{ title: 'Regions' }],
      existingFields: [{ internalName: 'Region', title: 'Region', typeAsString: 'Text' }],
    });
    const freshSettings = fresh.steps.find((s) => s.id === 'settings').payload.groupA;
    const adopted = buildApplyPlan(doc, { title: 'Requests', existing: 'resume' },
      { existingList: { id: 'x', title: 'Requests' }, targetLists: [{ title: 'Requests' }] });
    const adoptSettings = adopted.steps.find((s) => s.id === 'settings').payload.groupA;
    const dependent = fresh.steps.find((s) => s.id === 'field:RegionCode');
    return !('ReadSecurity' in freshSettings) && freshSettings.EnableVersioning === true
      && JSON.stringify(adoptSettings) === JSON.stringify({ EnableAttachments: true, EnableFolderCreation: true })
      && fresh.steps.find((s) => s.id === 'field:Region').status === 'failed'
      && dependent.status === 'blocked';
  }));

await check('pure: re-review fixes — Copy suffix skips a taken copy, a GUID-keyed lookupMap binds, a site column is not assumed in a fresh view, PS keeps -Paged', () =>
  page.evaluate(async () => {
    const { defaultTargetTitle, buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { toPnpPowerShellProvisioning } = await import('/src/workbench/list-schema-script.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'LinkTitle', type: 'Computed', custom: false, fromBaseType: true },
        { internalName: 'SiteCol', type: 'Text', custom: false, fromBaseType: false },
        { internalName: 'Client', displayName: 'Client', type: 'Lookup', custom: true, lookupList: null,
          lookupListId: '11111111-2222-3333-4444-555555555555', schemaXml: '<Field Name="Client" Type="Lookup"/>' },
      ],
      views: [{ title: 'All Items', hidden: false, fields: ['LinkTitle', 'SiteCol'], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const title = defaultTargetTitle(doc, [{ title: 'Requests' }, { title: 'requests copy' }]);
    const plan = buildApplyPlan(doc, { title: 'X', lookupMap: { '11111111-2222-3333-4444-555555555555': 'Customers' } },
      { targetLists: [{ id: 't1', title: 'customers' }] });
    const client = plan.steps.find((s) => s.id === 'field:Client');
    const view = plan.steps.find((s) => s.id === 'view:All Items');
    const ps = toPnpPowerShellProvisioning(doc, { lookupMap: { clients: 'Customers' } });
    return title === 'Requests Copy 2'
      && client.status === 'planned' && client.refs.lookupListId?.list?.toLowerCase() === 'customers'
      && view.payload.fields.join(',') === 'LinkTitle'
      && ps.includes(' -Paged');
  }));

await check('pure: executor — no views read without view steps, a 401 on the rename keeps the list done, default-view fallback never double-claims a target view', () =>
  page.evaluate(async () => {
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const { SpFileError } = await import('/src/sp-odata.js');
    const mkClient = (views, reads) => ({
      webUrl: () => 'https://t/sites/x',
      get: async (path) => { reads.push(path); return {}; },
      getAll: async (path) => {
        reads.push(path);
        if (path.endsWith('/views')) return { items: views };
        if (path.endsWith('/fields')) return { items: [{ Id: 'f1', InternalName: 'LinkTitle', Title: 'Title', TypeAsString: 'Computed' }] };
        return { items: [] };
      },
    });
    const mkWrite = (posts, { fail401On = '' } = {}) => {
      const go = async (path, body, opts) => {
        posts.push({ path, body });
        if (fail401On && path.includes(fail401On) && body?.Title) throw new SpFileError('expired', { code: 'auth', status: 401 });
        if (path === 'web/lists') return { Id: 'L1', RootFolder: { ServerRelativeUrl: '/x' } };
        return { Id: `v${posts.length}` };
      };
      return { postJson: go, mergeJson: go, isMock: () => true };
    };
    const listStep = (urlName = '') => ({ id: 'list', kind: 'list.create', label: 'l', dependsOn: [], status: 'planned',
      payload: { title: 'Req', description: '', baseTemplate: 100, contentTypesEnabled: false, urlName }, refs: {}, optional: false });
    // (a) no view steps → no /views read
    const readsA = [];
    await runPlan({ title: 'Req', steps: [listStep()] }, { client: mkClient([], readsA), spWrite: mkWrite([]) });
    const noViewsRead = !readsA.some((p) => p.endsWith('/views'));
    // (b) rename 401 → list done, aborted auth
    const reportB = await runPlan({ title: 'Req', steps: [listStep('ReqUrl')] },
      { client: mkClient([], []), spWrite: mkWrite([], { fail401On: "lists(guid'L1')" }) });
    const listDone = reportB.steps[0].status === 'done' && reportB.aborted === 'auth' && reportB.listId === 'L1';
    // (c) source default 'Alle Elemente' + source 'All Items' vs target default 'All Items'
    const vstep = (id, title, defaultView) => ({ id, kind: 'view.upsert', label: id, dependsOn: ['list'], status: 'planned', refs: {}, optional: false,
      payload: { title, viewId: null, fields: ['LinkTitle'], viewQuery: '', rowLimit: 30, paged: true, defaultView } });
    const postsC = [];
    const reportC = await runPlan({ title: 'Req', steps: [listStep(), vstep('view:Alle', 'Alle Elemente', true), vstep('view:All', 'All Items', false)] },
      { client: mkClient([{ Id: 'T1', Title: 'All Items', DefaultView: true }], []), spWrite: mkWrite(postsC) });
    const viewPosts = postsC.filter((p) => /\/views$/.test(p.path));
    const oneToOne = reportC.views.updated === 1 && reportC.views.added === 1 && viewPosts.length === 1
      && viewPosts[0].body.Title === 'Alle Elemente';
    return noViewsRead && listDone && oneToOne;
  }));

// Live dev-tenant finding: a document library's own _ExtendedDescription
// reports CanBeDeleted:true / FromBaseType:false, so capture treats it as a
// custom column and the plan schedules a create for it — but SharePoint
// already provisioned that exact name as part of the library's base type
// the instant the list was created. Posting createfieldasxml for it renamed
// the new field to '_ExtendedDescription0' and the step failed. Fixed in
// runFieldCreate: check the post-create re-read (ctx.targetFields) before
// posting.
await check('pure: executor — a field.create whose name is already on the target (post-create re-read) is skipped instead of posted, feeds its id into fieldIdMap for a following field.merge, and a same-name/different-type collision fails naming both types', () =>
  page.evaluate(async () => {
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const mkClient = (fields) => ({
      webUrl: () => 'https://t/sites/x',
      get: async () => ({}),
      getAll: async (path) => {
        if (path.endsWith('/fields')) return { items: fields };
        return { items: [] };
      },
    });
    const mkWrite = (posts) => {
      const go = async (path, body) => {
        posts.push({ path, body });
        if (path === 'web/lists') return { Id: 'L1', RootFolder: { ServerRelativeUrl: '/x' } };
        return { Id: `v${posts.length}`, InternalName: 'ignored' };
      };
      return { postJson: go, mergeJson: go, isMock: () => true };
    };
    const listStep = () => ({ id: 'list', kind: 'list.create', label: 'l', dependsOn: [], status: 'planned',
      payload: { title: 'Docs', description: '', baseTemplate: 101, contentTypesEnabled: false, urlName: '' }, refs: {}, optional: false });
    const fieldStep = (name, type) => ({
      id: `field:${name}`, kind: 'field.create', label: name, dependsOn: ['list'], status: 'planned', refs: {}, optional: false,
      payload: { field: { internalName: name, displayName: name, type, schemaXml: `<Field Name="${name}" Type="${type}"/>` }, options: 8 },
    });
    const mergeStep = (name) => ({
      id: `merge:${name}`, kind: 'field.merge', label: `merge ${name}`, dependsOn: [`field:${name}`], status: 'planned', refs: {}, optional: true,
      payload: { internalName: name, merges: { Indexed: true } },
    });

    // (a) same name, same type already on the target → skipped, nothing
    // posted to createfieldasxml, and its dependent merge still runs against
    // the EXISTING field's id.
    const postsA = [];
    const reportA = await runPlan(
      { title: 'Docs', steps: [listStep(), fieldStep('_ExtendedDescription', 'Note'), mergeStep('_ExtendedDescription')] },
      { client: mkClient([{ Id: 'existing-id', InternalName: '_ExtendedDescription', Title: 'Description', TypeAsString: 'Note' }]), spWrite: mkWrite(postsA) },
    );
    const stepA = reportA.steps.find((s) => s.id === 'field:_ExtendedDescription');
    const mergeA = reportA.steps.find((s) => s.id === 'merge:_ExtendedDescription');
    const noCreatePosted = !postsA.some((p) => p.path.includes('createfieldasxml'));
    const mergePosted = postsA.some((p) => p.path.includes("fields(guid'existing-id')"));

    // (b) same name, a DIFFERENT type already on the target → failed, naming
    // both types; still nothing posted to createfieldasxml.
    const postsB = [];
    const reportB = await runPlan(
      { title: 'Docs', steps: [listStep(), fieldStep('Budget', 'Number')] },
      { client: mkClient([{ Id: 'existing-id', InternalName: 'Budget', Title: 'Budget', TypeAsString: 'Text' }]), spWrite: mkWrite(postsB) },
    );
    const stepB = reportB.steps.find((s) => s.id === 'field:Budget');
    const noCreatePostedB = !postsB.some((p) => p.path.includes('createfieldasxml'));

    return stepA.status === 'skipped' && stepA.skipReason === 'already on the target' && noCreatePosted
      && reportA.fields.skipped === 1 && reportA.fields.added === 0
      && mergeA?.status === 'done' && mergePosted
      && stepB.status === 'failed' && stepB.error.includes('Text') && stepB.error.includes('Number') && noCreatePostedB;
  }));

await check('pure: a v1 publishing Pages (850) schema plans as a library, and a library _ExtendedDescription rename is a base-field update, never a create', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, SCHEMA_KIND } = await import('/src/workbench/list-schema.js');
    const v1 = {
      kind: SCHEMA_KIND, version: 1, source: { listTitle: 'Pages' },
      list: { title: 'Pages', baseTemplate: 850 },
      fields: [{ internalName: '_ExtendedDescription', displayName: 'Summary', type: 'Note', custom: true,
        description: 'What this page is about', schemaXml: '<Field Name="_ExtendedDescription" Type="Note"/>' }],
      views: [], contentTypes: [], warnings: [],
    };
    const plan = buildApplyPlan(v1, { title: 'Pages copy' }, { targetLists: [] });
    const create = plan.steps.find((s) => s.kind === 'list.create');
    const base = plan.steps.find((s) => s.kind === 'field.base' && s.payload.internalName === '_ExtendedDescription');
    return create.payload.baseTemplate === 101
      && !plan.steps.some((s) => s.id === 'field:_ExtendedDescription')
      && base && base.payload.displayName === 'Summary' && base.payload.description === 'What this page is about';
  }));

await check('pure: defaultTargetTitle keeps the source title when free, appends Copy when taken', () =>
  page.evaluate(async () => {
    const { defaultTargetTitle, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({ list: { title: 'Requests', baseTemplate: 100 }, fields: [], views: [] });
    const free = defaultTargetTitle(doc, [{ title: 'Other' }]);
    const taken = defaultTargetTitle(doc, [{ title: 'requests' }]);
    return free === 'Requests' && taken === 'Requests Copy';
  }));

await check('pure: a fresh-plan view keeps LinkTitle — a non-custom base column a fresh list already carries', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true },
        { internalName: 'LinkTitle', displayName: 'Title (linked)', type: 'Computed', custom: false, fromBaseType: true },
      ],
      views: [{ title: 'All Items', hidden: false, fields: ['Title', 'LinkTitle'], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const plan = buildApplyPlan(doc, { title: 'Requests' }, { existingList: null, existingFields: [] });
    const view = plan.steps.find((s) => s.id === 'view:All Items');
    return view.payload.fields.includes('Title') && view.payload.fields.includes('LinkTitle');
  }));

await check('pure: a resumed plan still merges Indexed onto an existing same-type column', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [{
        internalName: 'Budget', displayName: 'Budget', type: 'Number', custom: true,
        indexed: true, schemaXml: '<Field Name="Budget" Type="Number" Indexed="TRUE"/>',
      }],
      views: [],
    });
    const probe = {
      existingList: null,
      existingFields: [{ internalName: 'Budget', title: 'Budget', typeAsString: 'Number' }],
    };
    const plan = buildApplyPlan(doc, { title: 'Requests' }, probe);
    const create = plan.steps.find((s) => s.id === 'field:Budget');
    const merge = plan.steps.find((s) => s.id === 'merge:Budget');
    return create.status === 'skipped' && !!merge && merge.payload.merges.Indexed === true
      && merge.dependsOn.includes('field:Budget');
  }));

await check('pure: a GUID lookupMap value binds directly; target-title matching is case-insensitive', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const guid = 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE';
    const docGuid = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [{
        internalName: 'Client', displayName: 'Client', type: 'Lookup', custom: true,
        lookupList: 'Clients', schemaXml: '<Field Name="Client" Type="Lookup"/>',
      }],
      views: [],
    });
    const planGuid = buildApplyPlan(docGuid, { title: 'Requests', lookupMap: { Clients: `{${guid}}` } },
      { existingList: null, existingFields: [], targetLists: [] });
    const guidStep = planGuid.steps.find((s) => s.id === 'field:Client');

    const docLower = buildSchemaDoc({
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [{
        internalName: 'Region', displayName: 'Region', type: 'Lookup', custom: true,
        lookupList: 'Regions', schemaXml: '<Field Name="Region" Type="Lookup"/>',
      }],
      views: [],
    });
    const planLower = buildApplyPlan(docLower, { title: 'Requests' },
      { existingList: null, existingFields: [], targetLists: [{ title: 'regions' }] });
    const lowerStep = planLower.steps.find((s) => s.id === 'field:Region');

    return guidStep.status === 'planned' && guidStep.refs.lookupListId?.id === guid.toLowerCase()
      && lowerStep.status === 'planned' && lowerStep.refs.lookupListId?.list === 'regions';
  }));

await check('pure: PS emits an UPSERT view (Get-PnPView/$l.Update()/Invoke-PnPQuery, no -ValidationFormula, formatter JSON single-quoted); PnPjs still compiles, Options: 12 under content types, and upserts the default view', () =>
  page.evaluate(async () => {
    const { toPnpPowerShellProvisioning, toPnpjs2Provisioning } = await import('/src/workbench/list-schema-script.js');
    const { buildSchemaDoc, parentContentTypeId } = await import('/src/workbench/list-schema.js');
    const parent = '0x0100442912F2B6C7409A8FF25CE5504F1FD';
    const listCtId = `${parent}00${'c'.repeat(32)}`;
    const doc = buildSchemaDoc({
      list: {
        title: 'Requests', baseTemplate: 100, contentTypesEnabled: true,
        validationFormula: '=[Budget]>0', validationMessage: 'Must be positive',
      },
      fields: [{
        internalName: 'Notes', displayName: 'Notes', type: 'Note', custom: true,
        customFormatter: '{"$schema":"https://developer.microsoft.com/json-schemas/sp/column-formatting.schema.json"}',
        schemaXml: '<Field Name="Notes" Type="Note"/>',
      }],
      views: [{ title: 'All Items', hidden: false, defaultView: true, fields: ['Notes'], rowLimit: 30, paged: true, viewQuery: '' }],
      contentTypes: [{ id: listCtId, name: 'Request', parentId: parentContentTypeId(listCtId) }],
    });
    const ps = toPnpPowerShellProvisioning(doc, {});
    const js = toPnpjs2Provisioning(doc, {});
    let compiles = true;
    try {
      // eslint-disable-next-line no-new-func
      new (Object.getPrototypeOf(async () => {}).constructor)('sp', js);
    } catch { compiles = false; }
    const schemaLine = ps.split('\n').find((l) => l.includes('$schema'));
    // The formatter JSON's own double quotes are data, not PS string
    // delimiters — the assignment itself must open with a single quote
    // (CustomFormatter='{"$schema"…), never a bare double-quoted PS string.
    const schemaInSingleQuotesOnly = !!schemaLine && schemaLine.includes(`CustomFormatter='{"$schema"`)
      && !schemaLine.includes('CustomFormatter="');
    return ps.includes('Get-PnPView') && ps.includes('$l.Update()') && ps.includes('Invoke-PnPQuery')
      && !ps.includes('-ValidationFormula') && schemaInSingleQuotesOnly
      && compiles && !js.includes('// SKIP') && js.includes('Options: 12') && js.includes('newList.defaultView');
  }));

// ---- Pure: stage 2 (document libraries) ------------------------------------

await check('pure: a library settings plan drops EnableAttachments, moves EnableMinorVersions into group A alongside ForceCheckout, and keeps MajorWithMinorVersionsLimit/DraftVersionVisibility in group B', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: {
        title: 'Docs', baseTemplate: 101, enableVersioning: true, enableMinorVersions: true,
        majorWithMinorVersionsLimit: 5, draftVersionVisibility: 1, forceCheckout: true,
        enableAttachments: true, majorVersionLimit: 20,
      },
      fields: [], views: [],
    });
    const plan = buildApplyPlan(doc, { title: 'Docs' }, { targetLists: [] });
    const { groupA, groupB } = plan.steps.find((s) => s.id === 'settings').payload;
    // A generic list, for contrast: EnableAttachments stays, EnableMinorVersions
    // stays in group B (unchanged pre-stage-2 behaviour).
    const listDoc = buildSchemaDoc({
      list: {
        title: 'Reqs', baseTemplate: 100, enableVersioning: true, enableMinorVersions: true,
        enableAttachments: true,
      },
      fields: [], views: [],
    });
    const listPlan = buildApplyPlan(listDoc, { title: 'Reqs' }, { targetLists: [] });
    const listGroups = listPlan.steps.find((s) => s.id === 'settings').payload;
    return !('EnableAttachments' in groupA) && groupA.ForceCheckout === true && groupA.EnableMinorVersions === true
      && groupB.MajorWithMinorVersionsLimit === 5 && groupB.DraftVersionVisibility === 1
      && groupB.MajorVersionLimit === 20 && !('EnableMinorVersions' in groupB)
      && listGroups.groupA.EnableAttachments === true && !('EnableMinorVersions' in listGroups.groupA)
      && listGroups.groupB.EnableMinorVersions === true;
  }));

await check('pure: the base-field fix-up never forces Required:true onto a library’s Title, even when the source captured it required', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: { title: 'Docs', baseTemplate: 101 },
      fields: [{
        internalName: 'Title', displayName: 'Document title', required: true, custom: false,
        fromBaseType: true, baseTweak: { title: 'Document title', required: true, description: '' },
      }],
      views: [],
    });
    const plan = buildApplyPlan(doc, { title: 'Docs' }, { targetLists: [] });
    const titleStep = plan.steps.find((s) => s.id === 'title');
    // The same fixture on a generic list must still forward the captured
    // Required:true untouched — the clamp is library-only.
    const listDoc = buildSchemaDoc({
      list: { title: 'Reqs', baseTemplate: 100 },
      fields: [{
        internalName: 'Title', displayName: 'Request title', required: true, custom: false,
        fromBaseType: true, baseTweak: { title: 'Request title', required: true, description: '' },
      }],
      views: [],
    });
    const listPlan = buildApplyPlan(listDoc, { title: 'Reqs' }, { targetLists: [] });
    const listTitleStep = listPlan.steps.find((s) => s.id === 'title');
    return Boolean(titleStep) && titleStep.kind === 'field.base' && titleStep.payload.internalName === 'Title'
      && titleStep.payload.required === false && titleStep.payload.displayName === 'Document title'
      && Boolean(listTitleStep) && listTitleStep.payload.required === true;
  }));

await check('pure: isBuiltinParent treats exact 0x0101 as built in but a 0x0101-derived custom parent as attachable', () =>
  page.evaluate(async () => {
    const { isBuiltinParent, parentContentTypeId } = await import('/src/workbench/list-schema.js');
    const customParent = '0x010100AABBCCDDEEFF00112233445566';
    const listScoped = `${customParent}00${'C'.repeat(32)}`;
    return isBuiltinParent('0x0101') === true && isBuiltinParent('0x0120') === true
      && isBuiltinParent(parentContentTypeId(listScoped)) === false
      && parentContentTypeId(listScoped) === customParent;
  }));

await check('pure: a library view naming LinkFilename/DocIcon/FileSizeDisplay keeps them even though capture filtered the Hidden base columns out of doc.fields — a generic list gets no such free pass', () =>
  page.evaluate(async () => {
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const libDoc = buildSchemaDoc({
      list: { title: 'Docs', baseTemplate: 101 },
      fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true }],
      views: [{
        title: 'All Documents', hidden: false, fields: ['LinkFilename', 'DocIcon', 'FileSizeDisplay'],
        rowLimit: 30, paged: true, viewQuery: '',
      }],
    });
    const libPlan = buildApplyPlan(libDoc, { title: 'Docs' }, { targetLists: [] });
    const libViewStep = libPlan.steps.find((s) => s.kind === 'view.upsert');

    const listDoc = buildSchemaDoc({
      list: { title: 'Reqs', baseTemplate: 100 },
      fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true }],
      views: [{ title: 'All Items', hidden: false, fields: ['DocIcon'], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const listPlan = buildApplyPlan(listDoc, { title: 'Reqs' }, { targetLists: [] });
    const listViewStep = listPlan.steps.find((s) => s.kind === 'view.upsert');

    return libViewStep.payload.fields.includes('LinkFilename') && libViewStep.payload.fields.includes('DocIcon')
      && libViewStep.payload.fields.includes('FileSizeDisplay') && libPlan.warnings.length === 0
      && !listViewStep.payload.fields.includes('DocIcon')
      && listPlan.warnings.some((w) => w.includes('DocIcon'));
  }));

await check('pure: capture warns about a non-default per-library Forms template, and stays quiet for SharePoint’s own OOTB default', () =>
  page.evaluate(async () => {
    const { captureListSchema } = await import('/src/workbench/list-schema-capture.js');
    function makeClient(templateUrl) {
      return {
        webUrl: () => 'https://t/sites/x',
        get: async (path) => {
          if (path === 'web') return { Id: 'w1', Title: 'W', Url: 'https://t/sites/x', ServerRelativeUrl: '/sites/x', Language: 1033 };
          return {
            Id: 'L1', Title: 'Docs', BaseTemplate: 101, BaseType: 1, ItemCount: 0,
            RootFolder: { ServerRelativeUrl: '/sites/x/Docs' }, ContentTypesEnabled: false,
            DocumentTemplateUrl: templateUrl, EnableVersioning: false, EnableAttachments: true,
            ValidationFormula: '', ValidationMessage: '', OnQuickLaunch: false, ReadSecurity: null, WriteSecurity: null,
          };
        },
        getAll: async () => ({ items: [] }),
      };
    }
    const custom = await captureListSchema(makeClient('/sites/x/Docs/Forms/custom.dotx'), 'L1');
    const stock = await captureListSchema(makeClient('/sites/x/Docs/Forms/template.dotx'), 'L1');
    const none = await captureListSchema(makeClient(''), 'L1');
    return custom.doc.warnings.some((w) => w.includes('Per-library Forms template'))
      && !stock.doc.warnings.some((w) => w.includes('Per-library Forms template'))
      && !none.doc.warnings.some((w) => w.includes('Per-library Forms template'));
  }));

await check('pure: capture also warns — naming the content type — when a content type carries its own non-empty DocumentTemplate, even with the list-level template left at SharePoint’s default', () =>
  page.evaluate(async () => {
    const { captureListSchema } = await import('/src/workbench/list-schema-capture.js');
    function makeClient(ctTemplate) {
      return {
        webUrl: () => 'https://t/sites/x',
        get: async (path) => {
          if (path === 'web') return { Id: 'w1', Title: 'W', Url: 'https://t/sites/x', ServerRelativeUrl: '/sites/x', Language: 1033 };
          return {
            Id: 'L1', Title: 'Docs', BaseTemplate: 101, BaseType: 1, ItemCount: 0,
            RootFolder: { ServerRelativeUrl: '/sites/x/Docs' }, ContentTypesEnabled: true,
            DocumentTemplateUrl: '', EnableVersioning: false, EnableAttachments: true,
            ValidationFormula: '', ValidationMessage: '', OnQuickLaunch: false, ReadSecurity: null, WriteSecurity: null,
          };
        },
        getAll: async (path) => {
          if (path.includes('/contenttypes')) {
            return {
              items: [{
                StringId: '0x010100AABBCCDDEEFF00112233445566', Name: 'Report', Group: 'Custom Content Types',
                Hidden: false, ReadOnly: false, Sealed: false, DocumentTemplate: ctTemplate,
              }],
            };
          }
          return { items: [] };
        },
      };
    }
    const named = await captureListSchema(makeClient('/sites/x/Docs/report-template.dotx'), 'L1');
    const quiet = await captureListSchema(makeClient(''), 'L1');
    return named.doc.warnings.some((w) => w.includes('Per-library Forms template') && w.includes('‘Report’'))
      && !quiet.doc.warnings.some((w) => w.includes('Per-library Forms template'));
  }));

await check('pure: script emitters use DocumentLibrary for a 101 plan, and never emit EnableAttachments', () =>
  page.evaluate(async () => {
    const { toPnpPowerShellProvisioning, toPnpjs2Provisioning } = await import('/src/workbench/list-schema-script.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const doc = buildSchemaDoc({
      list: {
        title: 'Docs', baseTemplate: 101, enableVersioning: true, enableMinorVersions: true,
        forceCheckout: true, enableAttachments: true,
      },
      fields: [], views: [],
    });
    const ps = toPnpPowerShellProvisioning(doc, {});
    const js = toPnpjs2Provisioning(doc, {});
    return ps.includes('-Template DocumentLibrary') && !ps.includes('EnableAttachments')
      && ps.includes('ForceCheckout') && ps.includes('EnableMinorVersions')
      && !js.includes('EnableAttachments') && js.includes('ForceCheckout');
  }));

// ---- Mock UI: the Schema tab on /sites/schema ------------------------------

const schemaPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await schemaPage.goto(WB_URL);
await schemaPage.waitForSelector('.wb-home-cards');
await schemaPage.fill('#wb-site-input', '/sites/schema');
await schemaPage.locator('#wb-site-open').click();
await schemaPage.waitForFunction(() =>
  document.getElementById('wb-status-context').textContent.includes('/sites/schema'));
await schemaPage.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
await schemaPage.waitForSelector('.wb-table tbody tr', { hasText: 'Requests' });
await schemaPage.locator('.wb-table tbody tr', { hasText: 'Requests' }).locator('td').first().click();
await schemaPage.waitForSelector('.wb-tab');
await schemaPage.locator('.wb-tab', { hasText: 'Schema' }).click();
await schemaPage.waitForSelector('.wb-schema-chips .wb-schema-kind');

await check('schema: the tab sits after Content types, before Permissions, and writes the route', async () => {
  const tabs = await schemaPage.locator('.wb-tab').allTextContents();
  const route = await schemaPage.evaluate(() =>
    JSON.parse(sessionStorage.getItem('dcspad.workbench.route') || '{}'));
  return tabs.join(',') === 'Fields,Views,Content types,Schema,Permissions,Items,Raw' && route.tab === 'schema';
});

await check('schema: chips show the five schema facts for Requests', async () => {
  const kind = await schemaPage.locator('.wb-schema-kind').textContent();
  const fields = await schemaPage.locator('.wb-schema-fields').textContent();
  const views = await schemaPage.locator('.wb-schema-views').textContent();
  const cts = await schemaPage.locator('.wb-schema-cts').textContent();
  const versioning = await schemaPage.locator('.wb-schema-versioning').textContent();
  return kind === 'Generic list' && fields === '21 fields · 15 custom' && views === '2 views'
    && cts === 'content types on' && versioning === 'versioning on';
});

await check('schema: chips compose .wb-info-chip — sans, unfilled, not shouty', () =>
  schemaPage.evaluate(() => {
    const chip = document.querySelector('.wb-schema-kind');
    const cs = getComputedStyle(chip);
    const bg = cs.backgroundColor;
    return chip.classList.contains('wb-info-chip')
      && !cs.fontFamily.toLowerCase().includes('mono')
      && (bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent')
      && cs.textTransform !== 'uppercase';
  }));

await check('schema: the Fields section notes system columns, managed metadata, and dependent lookups', async () => {
  const text = await schemaPage.locator('.wb-schema-section', { hasText: 'Fields' }).first().textContent();
  return text.includes('system column — not copied')
    && text.includes('managed metadata — not recreated')
    && text.includes('dependent lookup of');
});

await check('schema: the Export menu offers exactly the four document actions', async () => {
  await schemaPage.locator('.wb-schema-actions .wb-menu-wrap button', { hasText: 'Export' }).click();
  const items = await schemaPage.locator('.wb-schema-actions .wb-menu-item').allTextContents();
  return items.join(',') === 'Download schema .json,Copy schema JSON,'
    + 'Copy as PnP.PowerShell (provision),Copy as PnPjs 2 (provision)';
});

await check('schema: Download schema .json parses back as a v2 document for this list', async () => {
  const [download] = await Promise.all([
    schemaPage.waitForEvent('download'),
    schemaPage.locator('.wb-schema-actions .wb-menu-item', { hasText: 'Download schema .json' }).click(),
  ]);
  const bytes = readFileSync(await download.path());
  const doc = JSON.parse(bytes.toString('utf8'));
  return download.suggestedFilename() === 'schema-requests.json'
    && doc.kind === 'dcspad-sputils-list-schema' && doc.version === 2 && doc.list.title === 'Requests';
});

await check('schema: Copy schema JSON flashes the copied state on the Export trigger', async () => {
  // Stub the clipboard so the copy is deterministic in the sandbox (no
  // clipboard-write permission grant needed, matching workbench.mjs).
  await schemaPage.evaluate(() => {
    navigator.clipboard.writeText = (t) => { window.__COPIED = t; return Promise.resolve(); };
  });
  await schemaPage.locator('.wb-schema-actions .wb-menu-wrap button', { hasText: 'Export' }).click();
  await schemaPage.locator('.wb-schema-actions .wb-menu-item', { hasText: 'Copy schema JSON' }).click();
  return (await schemaPage.locator('.wb-schema-actions .wb-menu-wrap button.copied').count()) === 1;
});

await check('schema: Copy to… is enabled for a generic list, with the working title', async () => {
  const btn = schemaPage.locator('.wb-schema-copy');
  const disabled = await btn.isDisabled();
  const title = await btn.getAttribute('title');
  return !disabled && title === 'Create a new list from this schema, on this site or another one.';
});

await check('schema: a document library shows the library chip, and Copy to… is enabled (stage 2)', async () => {
  await schemaPage.locator('.wb-back').click();
  await schemaPage.waitForSelector('.wb-table tbody tr', { hasText: 'Documents' });
  await schemaPage.locator('.wb-table tbody tr', { hasText: 'Documents' }).locator('td').first().click();
  await schemaPage.waitForSelector('.wb-tab');
  await schemaPage.locator('.wb-tab', { hasText: 'Schema' }).click();
  await schemaPage.waitForSelector('.wb-schema-chips .wb-schema-kind');
  const libChip = await schemaPage.locator('.wb-schema-libkind').count();
  const disabled = await schemaPage.locator('.wb-schema-copy').isDisabled();
  const title = await schemaPage.locator('.wb-schema-copy').getAttribute('title');
  return libChip === 1 && !disabled
    && title === 'Create a new list from this schema, on this site or another one.';
});

await check('schema: the Documents warnings section names the non-default per-library Forms template', async () => {
  const text = await schemaPage.locator('.wb-schema-section', { hasText: 'Warnings' }).textContent();
  return text.includes('Per-library Forms template');
});

await check('dialog: Copy to… on a library shows the files-not-copied line and hides the item checkboxes', async () => {
  await schemaPage.locator('.wb-schema-copy').click();
  await schemaPage.waitForSelector('.wb-schema-dialog');
  await schemaPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);
  const note = await schemaPage.locator('.wb-schema-items-note').textContent();
  const rowHidden = await schemaPage.locator('.wb-schema-items-row').isHidden();
  const fieldsetDisabled = await schemaPage.evaluate(() => document.querySelector('.wb-schema-items')?.disabled);
  return note.trim() === 'Files are not copied — a library copy is schema only.' && rowHidden && fieldsetDisabled === true;
});

await check('dialog: a library Create on /sites/target posts BaseTemplate 101, no EnableAttachments, ForceCheckout true, attaches the 0x0101-derived content type, fails zero steps, and recreates ‘All Documents’/‘By kind’ with the exact source field sequences', async () => {
  await schemaPage.fill('.wb-schema-target', '/sites/target');
  await schemaPage.locator('.wb-schema-connect').click();
  await schemaPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value === 'Documents Copy');
  await schemaPage.evaluate(() => { window.__DCSPAD_WB_WRITES__ = []; });
  await schemaPage.locator('.wb-schema-create').click();
  await schemaPage.waitForSelector('.wb-schema-report:not([hidden])');
  const failedStepsShown = await schemaPage.locator('.wb-schema-report-failed h3').count();
  const result = await schemaPage.evaluate(() => {
    const writes = window.__DCSPAD_WB_WRITES__ || [];
    const method = (w) => w.headers?.['X-HTTP-Method'] || w.headers?.['x-http-method'] || '';
    const lower = (w) => w.url.toLowerCase();
    const create = writes.find((w) => /\/web\/lists$/.test(lower(w)) && !method(w));
    const createBody = create ? JSON.parse(create.body) : {};
    const settings = writes.find((w) => method(w) === 'MERGE' && /lists\(guid'[0-9a-f-]+'\)$/.test(lower(w)));
    const settingsBody = settings ? JSON.parse(settings.body) : {};
    const ctAttach = writes.some((w) => lower(w).includes('addavailablecontenttype'));
    return {
      baseTemplate: createBody.BaseTemplate, hasEnableAttachments: 'EnableAttachments' in settingsBody,
      forceCheckout: settingsBody.ForceCheckout, ctAttach,
    };
  });
  // Read the target list back exactly the way any other capture would — the
  // ground truth for what actually landed, not a re-parse of the write log.
  const views = await schemaPage.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { mockResolver } = await import('/src/workbench/mock-data.js');
    const { captureListSchema } = await import('/src/workbench/list-schema-capture.js');
    const target = createSpRestClient({ mockResolver });
    await target.connectWeb('/sites/target');
    const { items } = await target.getAll('web/lists', { select: ['Id', 'Title'] });
    const listId = items.find((l) => l.Title === 'Documents Copy')?.Id;
    if (!listId) return null;
    const { doc } = await captureListSchema(target, listId, { includeHidden: true });
    const byTitle = (t) => doc.views.find((v) => v.title === t);
    return { allDocuments: byTitle('All Documents')?.fields, byKind: byTitle('By kind')?.fields };
  });
  // Close the report so later schemaPage checks (which start from the
  // all-lists grid) aren't blocked by this dialog's own overlay.
  await schemaPage.locator('.wb-schema-cancel').click();
  await schemaPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });
  return result.baseTemplate === 101 && !result.hasEnableAttachments
    && result.forceCheckout === true && result.ctAttach === true
    && failedStepsShown === 0
    && (views?.allDocuments || []).join(',') === 'DocIcon,LinkFilename,DocCategory,FileSizeDisplay,Modified'
    && (views?.byKind || []).join(',') === 'DocIcon,LinkFilename,Modified,Editor';
});

// ---- Stubbed live: capture request shapes ----------------------------------

const LIVE_LIST_ID = '33333333-0000-4000-8000-000000000001';
const LIVE_LOOKUP_ID = '33333333-0000-4000-8000-000000000002';
const LIVE_DENIED_ID = '33333333-0000-4000-8000-000000000003';
const LIVE_NEW_LIST_ID = '33333333-0000-4000-8000-000000000009';
const liveReads = [];   // { url, accept }
const liveDialogWrites = [];   // { url, method, body } — the dialog's own 401 test, below
const liveDialogFlags = { expireOnField: null };

const live = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await live.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await live.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const method = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  liveReads.push({ url, accept: request.headers().accept || '' });

  // ---- writes for the apply dialog's own 401 test (below) ----------------
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'LIVE-DIALOG-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (method === 'GET' && /\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
    const base = url.slice(0, url.indexOf('/_api/'));
    return route.fulfill({ json: { Id: 'live-web', Title: 'Live Web', Url: base, ServerRelativeUrl: '/' } });
  }
  if (method === 'POST' && /\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    liveDialogWrites.push({ url, method: 'POST', body: request.postData() || '' });
    return route.fulfill({ json: {
      Id: LIVE_NEW_LIST_ID, Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/Lists/LiveTarget' },
    } });
  }
  if (method === 'GET' && url.includes(`lists(guid'${LIVE_NEW_LIST_ID}')/fields`)) {
    return route.fulfill({ json: { value: [
      { Id: 'nf1', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false },
    ] } });
  }
  if (method === 'POST' && xHttpMethod === 'MERGE' && new RegExp(`lists\\(guid'${LIVE_NEW_LIST_ID}'\\)$`).test(url)) {
    liveDialogWrites.push({ url, method: 'MERGE', body: request.postData() || '' });
    return route.fulfill({ json: {} });
  }
  if (url.includes('createfieldasxml') && url.includes(`lists(guid'${LIVE_NEW_LIST_ID}')`)) {
    liveDialogWrites.push({ url, method: 'POST', body: request.postData() || '' });
    if (liveDialogFlags.expireOnField) {
      liveDialogFlags.expireOnField = null;
      return route.fulfill({ status: 401, json: { 'odata.error': { message: { value: 'The security token is expired.' } } } });
    }
    return route.fulfill({ json: { Id: 'nf2', InternalName: 'Region' } });
  }

  if (url.includes(`lists(guid'${LIVE_LIST_ID}')/fields`)) {
    return route.fulfill({ json: { value: [
      {
        Id: 'f1', Title: 'Title', InternalName: 'Title', TypeAsString: 'Text',
        FromBaseType: true, CanBeDeleted: false, SchemaXml: '<Field Name="Title" Type="Text"/>',
      },
      {
        Id: 'f2', Title: 'Region', InternalName: 'Region', TypeAsString: 'Lookup',
        FromBaseType: false, CanBeDeleted: true, LookupList: LIVE_LOOKUP_ID, LookupField: 'Title',
        SchemaXml: `<Field Name="Region" Type="Lookup" List="{${LIVE_LOOKUP_ID}}"/>`,
      },
    ] } });
  }
  if (url.includes(`lists(guid'${LIVE_LIST_ID}')/views`)) {
    return route.fulfill({ json: { value: [
      { Id: 'v1', Title: 'All Items', DefaultView: true, PersonalView: false, Hidden: false, RowLimit: 30, Paged: true, ViewQuery: '', ViewFields: { Items: ['Title'] } },
    ] } });
  }
  if (url.includes(`lists(guid'${LIVE_LIST_ID}')/contenttypes`)) {
    return route.fulfill({ json: { value: [] } });
  }
  if (url.includes(`lists(guid'${LIVE_LIST_ID}')`) && url.includes('$expand=RootFolder')) {
    return route.fulfill({ json: {
      Id: LIVE_LIST_ID, Title: 'LiveRequests', BaseTemplate: 100, BaseType: 0, ItemCount: 4,
      RootFolder: { ServerRelativeUrl: '/Lists/LiveRequests' }, ContentTypesEnabled: false,
      EnableVersioning: false, EnableAttachments: true,
    } });
  }
  if (url.includes(`lists(guid'${LIVE_LOOKUP_ID}')`)) {
    return route.fulfill({ json: { Title: 'Regions' } });
  }
  if (url.includes(`lists(guid'${LIVE_DENIED_ID}')/fields`)) {
    return route.fulfill({ status: 403, json: { 'odata.error': { message: { value: 'Access denied.' } } } });
  }
  if (url.includes(`lists(guid'${LIVE_DENIED_ID}')`) && url.includes('$expand=RootFolder')) {
    return route.fulfill({ json: {
      Id: LIVE_DENIED_ID, Title: 'LiveDenied', BaseTemplate: 100, BaseType: 0, ItemCount: 0,
      RootFolder: { ServerRelativeUrl: '/Lists/LiveDenied' }, ContentTypesEnabled: false,
    } });
  }
  if (url.includes(`lists(guid'${LIVE_DENIED_ID}')/views`) || url.includes(`lists(guid'${LIVE_DENIED_ID}')/contenttypes`)) {
    return route.fulfill({ json: { value: [] } });
  }
  if (url.includes('/_api/web/lists') && !url.includes("lists(guid'")) {
    return route.fulfill({ json: { value: [
      { Id: LIVE_LIST_ID, Title: 'LiveRequests', BaseTemplate: 100, ItemCount: 4, Hidden: false, RootFolder: { ServerRelativeUrl: '/Lists/LiveRequests' } },
      { Id: LIVE_DENIED_ID, Title: 'LiveDenied', BaseTemplate: 100, ItemCount: 0, Hidden: false, RootFolder: { ServerRelativeUrl: '/Lists/LiveDenied' } },
      // Same web as the source (this page's dialog target defaults to the
      // host web) — gives the Region lookup's field.create step a real
      // target to resolve, so the apply-dialog 401 test below reaches it.
      { Id: 'live-regions', Title: 'Regions', BaseTemplate: 100, ItemCount: 3, Hidden: false, RootFolder: { ServerRelativeUrl: '/Lists/Regions' } },
    ] } });
  }
  return route.fulfill({ json: { value: [] } });
});

await live.goto(WB_URL);
await live.waitForSelector('.wb-home-cards');
await live.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
await live.locator('.wb-table tbody tr', { hasText: 'LiveRequests' }).first().waitFor();
await live.locator('.wb-table tbody tr', { hasText: 'LiveRequests' }).locator('td').first().click();
await live.locator('.wb-tab', { hasText: 'Schema' }).click();
await live.waitForSelector('.wb-schema-chips .wb-schema-kind');

await check('live: capture reads the list entity unprojected, with RootFolder expanded', () =>
  liveReads.some((r) => r.url.includes(`lists(guid'${LIVE_LIST_ID}')`)
    && r.url.includes('$expand=RootFolder') && !r.url.includes('$select=')));

await check('live: capture reads fields with no $select, filtered to visible', () =>
  liveReads.some((r) => r.url.includes(`lists(guid'${LIVE_LIST_ID}')/fields`)
    && !r.url.includes('$select=') && r.url.includes('Hidden')));

await check('live: capture reads views with $expand=ViewFields', () =>
  liveReads.some((r) => r.url.includes(`lists(guid'${LIVE_LIST_ID}')/views`) && r.url.includes('$expand=ViewFields')));

await check('live: a custom lookup’s target list title is resolved by a dedicated read', () =>
  liveReads.some((r) => r.url.includes(`lists(guid'${LIVE_LOOKUP_ID}')`) && r.url.includes('$select=Title')));

await check('live: every capture read sends the nometadata Accept header', () =>
  liveReads.length > 0 && liveReads.every((r) => r.accept.includes('application/json;odata=nometadata')));

await check('live: a denied fields read states the Schema tab in the neutral register', async () => {
  await live.locator('.wb-back').click();
  await live.locator('.wb-table tbody tr', { hasText: 'LiveDenied' }).locator('td').first().click();
  await live.locator('.wb-tab', { hasText: 'Schema' }).click();
  await live.waitForSelector('.wb-tab-pane .wb-denied');
  const cls = await live.locator('.wb-tab-pane .wb-denied').getAttribute('class');
  return cls.includes('wb-denied') && !cls.includes('wb-error');
});

// ---- Mock: the apply executor on /sites/target (slice 2a) ------------------
// Own page: a target client + spWrite built directly (createSpRestClient /
// createSpWriteClient / mockResolver / mockWriter), not through the UI (the
// dialog is slice 2b). writerState (mock-data.js) is a module-level
// singleton, so it persists across these sequential checks within this one
// page — the dry-run check must run before the create so probeTarget still
// finds no 'Requests' list on /sites/target.

const applyPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await applyPage.goto(WB_URL);
await applyPage.waitForSelector('.wb-home-cards');

await check('apply-mock: a dry run touches nothing', () =>
  applyPage.evaluate(async () => {
    globalThis.__DCSPAD_WB_WRITES__ = [];
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { mockResolver, mockWriter } = await import('/src/workbench/mock-data.js');
    const { captureListSchema } = await import('/src/workbench/list-schema-capture.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');

    const source = createSpRestClient({ mockResolver });
    await source.connectWeb('/sites/schema');
    const { items: srcLists } = await source.getAll('web/lists', { select: ['Id', 'Title'] });
    const requestsId = srcLists.find((l) => l.Title === 'Requests').Id;
    const { doc } = await captureListSchema(source, requestsId);
    window.__LSA_DOC__ = doc;

    const target = createSpRestClient({ mockResolver });
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target, mockWriter });

    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'Requests', dryRun: true } });
    const writes = globalThis.__DCSPAD_WB_WRITES__ || [];
    return report.dryRun === true && report.steps.length > 10 && writes.length === 0;
  }));

// The Requests doc carries one taxonomy column (RequestCategory) that always
// fails (managed metadata is never recreated) and one lookup (Region) with no
// target on /sites/target. Neither may cost the list its views or its
// validation formula: views and validation depend on the list alone.
await check('apply-mock: a Create run writes list → settings → content types → fields (tier order, Options bit 8) → views → validation last; the skipped taxonomy column blocks nothing and is never retried', () =>
  applyPage.evaluate(async () => {
    globalThis.__DCSPAD_WB_WRITES__ = [];
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { mockResolver, mockWriter } = await import('/src/workbench/mock-data.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const { fieldTier } = await import('/src/workbench/list-schema.js');
    const doc = window.__LSA_DOC__;

    const target = createSpRestClient({ mockResolver });
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target, mockWriter });

    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'Requests' } });
    window.__LSA_REPORT__ = report;
    window.__LSA_TARGET__ = target;

    const writes = globalThis.__DCSPAD_WB_WRITES__ || [];
    const lower = (w) => w.url.toLowerCase();
    const method = (w) => w.headers?.['X-HTTP-Method'] || w.headers?.['x-http-method'] || '';
    const listCreateIdx = writes.findIndex((w) => /\/web\/lists$/.test(lower(w)) && !method(w));
    const settingsIdx = writes.findIndex((w, i) => i > listCreateIdx && method(w) === 'MERGE'
      && /lists\(guid'[0-9a-f-]+'\)$/.test(lower(w)));
    const ctIdx = writes.findIndex((w) => lower(w).includes('addavailablecontenttype'));
    const fieldWrites = writes.filter((w) => lower(w).includes('createfieldasxml'));
    const firstFieldIdx = writes.findIndex((w) => lower(w).includes('createfieldasxml'));
    const lastWrite = writes[writes.length - 1];
    const validationLast = method(lastWrite) === 'MERGE'
      && JSON.parse(lastWrite.body || '{}').ValidationFormula !== undefined;
    const lastFieldIdx = writes.map((w) => lower(w).includes('createfieldasxml')).lastIndexOf(true);
    const firstViewIdx = writes.findIndex((w) => /\/views/.test(lower(w)));
    const optionsHaveBit8 = fieldWrites.every((w) => (JSON.parse(w.body).parameters.Options & 8) === 8);
    // Tier order (plain → lookup → dependent → calculated) must be
    // non-decreasing across the created-field write sequence.
    const tiersInOrder = fieldWrites.every((w, i) => {
      if (i === 0) return true;
      const name = /Name="([^"]*)"/.exec(JSON.parse(w.body).parameters.SchemaXml)?.[1];
      const prevName = /Name="([^"]*)"/.exec(JSON.parse(fieldWrites[i - 1].body).parameters.SchemaXml)?.[1];
      const f = doc.fields.find((x) => x.internalName === name);
      const prevF = doc.fields.find((x) => x.internalName === prevName);
      return !f || !prevF || fieldTier(f) >= fieldTier(prevF);
    });

    const categoryStep = report.steps.find((s) => s.id === 'field:RequestCategory');
    const regionStep = report.steps.find((s) => s.id === 'field:Region');
    const viewSteps = report.steps.filter((s) => s.kind === 'view.upsert');
    const validationStep = report.steps.find((s) => s.kind === 'list.validation');

    return report.created === true
      && report.listId
      && report.fields.added === 13 && report.fields.skipped === 2 && report.fields.failed.length === 0
      && categoryStep.status === 'skipped' && categoryStep.final === true && regionStep.status === 'skipped'
      && report.contentTypes.attached === 1
      && listCreateIdx === 0
      && settingsIdx === 1
      && ctIdx > settingsIdx
      && firstFieldIdx > ctIdx
      && optionsHaveBit8 && tiersInOrder
      && viewSteps.length > 0 && viewSteps.every((s) => s.status === 'done')
      && report.views.added + report.views.updated === viewSteps.length
      && firstViewIdx > lastFieldIdx
      && validationStep.status === 'done' && validationLast;
  }));

await check('apply-mock: the created list is resolvable by the target client afterwards', () =>
  applyPage.evaluate(async () => {
    const target = window.__LSA_TARGET__;
    const report = window.__LSA_REPORT__;
    const { items } = await target.getAll('web/lists', { select: ['Id', 'Title'] });
    const found = items.find((l) => l.Title === 'Requests');
    if (!found || found.Id !== report.listId) return false;
    const { items: fields } = await target.getAll(`web/lists(guid'${found.Id}')/fields`);
    return fields.length > 20;   // base columns + the created custom fields
  }));

// A fresh POST web/lists list is born with SharePoint's own default "All
// Items" view already on it (the mock writer seeds exactly that) — the
// plan-time probe ran before the list existed and so never saw it, but the
// executor's live between-step re-read (ctx.targetViews) must still find it
// and upsert into it instead of posting a duplicate.
await check('apply-mock: a fresh create matches the list’s own already-seeded default "All Items" view by the live re-read — no second POST to …/views, and the report counts it updated', () =>
  applyPage.evaluate(async () => {
    globalThis.__DCSPAD_WB_WRITES__ = [];
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { mockResolver, mockWriter } = await import('/src/workbench/mock-data.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'DefaultViewMatch', baseTemplate: 100, contentTypesEnabled: false },
      fields: [],
      views: [{ title: 'All Items', hidden: false, defaultView: true, fields: [], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const target = createSpRestClient({ mockResolver });
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target, mockWriter });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'DefaultViewMatch' } });
    const writes = globalThis.__DCSPAD_WB_WRITES__ || [];
    const viewPost = writes.find((w) => /\/views$/.test(w.url) && !(w.headers?.['X-HTTP-Method'] || w.headers?.['x-http-method']));
    const viewStep = report.steps.find((s) => s.kind === 'view.upsert');
    return !viewPost && viewStep.status === 'done' && viewStep.result?.created === false
      && report.views.updated === 1 && report.views.added === 0;
  }));

await applyPage.close();

// ---- Stubbed live: the apply executor's write shapes and ordering ---------

const LIST_ID = 'bbbbbbbb-1111-4000-8000-000000000001';
const REGIONS_ID = 'bbbbbbbb-1111-4000-8000-000000000002';
const CT_PARENT = '0x0100442912F2B6C7409A8FF25CE5504F1FD';

const liveApply = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await liveApply.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});

const applyPosts = [];   // { url, method, ifMatch, digest, body, order }
const contextInfoUrls = [];
const applyFlags = { failField: null, throttleField: null, expireOnField: null };
let applyCallSeq = 0;
// Fields actually created via createfieldasxml on this stub's one reused
// LIST_ID, name → TypeAsString — the fields GET route below folds these in
// on top of the two genuinely pre-existing base columns (Title, LinkTitle),
// so the executor's post-create re-read (list-schema-apply.js
// runFieldCreate's already-on-the-target check) never sees a field as
// present before its own createfieldasxml has actually landed.
const stubCreatedFields = new Map();

await liveApply.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  const digest = request.headers()['x-requestdigest'] || '';
  const ifMatch = request.headers()['if-match'] || '';

  if (url.includes('/_api/contextinfo')) {
    contextInfoUrls.push(url);
    return route.fulfill({
      json: { FormDigestValue: 'WB-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }

  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      let rel = '/';
      try { rel = new URL(base).pathname || '/'; } catch { /* keep '/' */ }
      return route.fulfill({ json: { Id: 'web-id', Title: 'Target Web', Url: base, ServerRelativeUrl: rel } });
    }
    if (url.includes('/GetList(@listUrl)')) {
      return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'List not found.' } } } });
    }
    if (url.includes('/availablecontenttypes')) {
      return route.fulfill({ json: { value: [{ StringId: CT_PARENT, Name: 'Request', Group: 'Custom' }] } });
    }
    if (url.includes(`lists(guid'${LIST_ID}')/fields`)) {
      const base = [
        { Id: 'fid-title', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false },
        { Id: 'fid-linktitle', InternalName: 'LinkTitle', Title: 'Title', TypeAsString: 'Computed', Hidden: false, ReadOnlyField: true },
      ];
      const created = [...stubCreatedFields.entries()].map(([name, typeAsString]) => ({
        Id: `fid-${name.toLowerCase()}`, InternalName: name, Title: name, TypeAsString: typeAsString,
        Hidden: false, ReadOnlyField: false,
      }));
      return route.fulfill({ json: { value: [...base, ...created] } });
    }
    if (url.includes('/viewfields') && url.includes(`lists(guid'${LIST_ID}')`)) {
      return route.fulfill({ json: { Items: [] } });
    }
    if (/\/web\/lists(\?|$)/.test(url) && !url.includes("lists(guid'")) {
      return route.fulfill({ json: { value: [
        { Id: REGIONS_ID, Title: 'Regions', BaseTemplate: 100, ContentTypesEnabled: false, RootFolder: { ServerRelativeUrl: '/sites/target/Lists/Regions' } },
      ] } });
    }
    return route.fulfill({ json: { value: [] } });
  }

  const record = () => {
    const entry = { url, method: xHttpMethod || httpMethod, ifMatch, digest, body: request.postData() || '', order: ++applyCallSeq };
    applyPosts.push(entry);
    return entry;
  };

  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    record();
    return route.fulfill({ json: {
      Id: LIST_ID, Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/sites/target/Lists/LiveTarget' },
    } });
  }

  if (url.includes('createfieldasxml')) {
    const parsed = JSON.parse(request.postData() || '{}');
    const xml = parsed?.parameters?.SchemaXml || '';
    const name = /Name="([^"]*)"/.exec(xml)?.[1] || '';
    record();
    if (applyFlags.expireOnField && name === applyFlags.expireOnField) {
      applyFlags.expireOnField = null;
      return route.fulfill({ status: 401, json: { 'odata.error': { message: { value: 'The security token is expired.' } } } });
    }
    if (applyFlags.throttleField && name === applyFlags.throttleField) {
      applyFlags.throttleField = null;
      return route.fulfill({ status: 429, headers: { 'Retry-After': '0' }, json: {} });
    }
    if (applyFlags.failField && name === applyFlags.failField) {
      return route.fulfill({ status: 400, json: { 'odata.error': { message: { value: `SharePoint could not create the column '${name}'.` } } } });
    }
    const type = /Type="([^"]*)"/.exec(xml)?.[1] || 'Text';
    stubCreatedFields.set(name, type);
    return route.fulfill({ json: { Id: `fid-${name.toLowerCase()}`, InternalName: name } });
  }

  if (url.toLowerCase().includes('addavailablecontenttype')) {
    record();
    return route.fulfill({ json: {} });
  }

  if (/\/views$/.test(url) && !xHttpMethod) {
    record();
    return route.fulfill({ json: { Id: 'view-1', Title: JSON.parse(request.postData() || '{}').Title } });
  }

  if (url.includes('removeallviewfields') || url.includes('addviewfield')) {
    record();
    return route.fulfill({ json: {} });
  }

  record();
  return route.fulfill({ json: {} });
});

await liveApply.goto(WB_URL);
await liveApply.waitForSelector('.wb-home-cards');

await check('live-apply: a create run posts web/lists with the digest, settings MERGE with if-match *, tiered createfieldasxml (Options bit 8, scrubbed XML), field merges, ct attach before fields, views, and validation last', () =>
  liveApply.evaluate(async (LIST_ID_) => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc, parentContentTypeId } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');

    const parent = '0x0100442912F2B6C7409A8FF25CE5504F1FD';
    const listCtId = `${parent}00${'d'.repeat(32)}`;
    const doc = buildSchemaDoc({
      list: {
        title: 'LiveTarget', baseTemplate: 100, contentTypesEnabled: true, enableVersioning: true,
        validationFormula: '=[Budget]>0', validationMessage: 'Must be positive',
      },
      fields: [
        { internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true, required: true },
        {
          internalName: 'Budget', displayName: 'Budget', type: 'Number', custom: true,
          indexed: true, enforceUniqueValues: true,
          schemaXml: '<Field ID="{aaa}" SourceID="{bbb}" Name="Budget" Type="Number" DisplayName="Budget" Indexed="TRUE" EnforceUniqueValues="TRUE" />',
        },
        {
          internalName: 'Region', displayName: 'Region', type: 'Lookup', custom: true, lookupList: 'Regions',
          schemaXml: '<Field ID="{ccc}" Name="Region" Type="Lookup" DisplayName="Region" List="{oldguid}" ShowField="Title" />',
        },
      ],
      views: [{ title: 'All Items', hidden: false, defaultView: true, fields: ['Title', 'Budget'], rowLimit: 30, paged: true, viewQuery: '' }],
      contentTypes: [{ id: listCtId, name: 'Request', parentId: parentContentTypeId(listCtId) }],
    });

    const target = createSpRestClient({});
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target });

    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'LiveTarget' } });
    window.__LSA_LIVE_REPORT__ = report;
    window.__LSA_LIVE_TARGET__ = target;
    window.__LSA_LIVE_SPWRITE__ = spWrite;
    window.__LSA_LIVE_DOC__ = doc;
    return { created: report.created, listId: report.listId, validationApplied: report.validation.applied };
  }, LIST_ID).then((r) => r.created === true && r.listId === LIST_ID && r.validationApplied === true));

await check('live-apply: web/lists carries the digest header, and settings MERGE follows it with if-match *', () => {
  const listCreate = applyPosts.find((p) => /\/web\/lists$/.test(p.url) && p.method === 'POST');
  const settings = applyPosts.find((p) => p.order > listCreate.order && p.method === 'MERGE'
    && new RegExp(`lists\\(guid'${LIST_ID}'\\)$`).test(p.url));
  return !!listCreate && listCreate.digest === 'WB-DIGEST'
    && !!settings && settings.ifMatch === '*' && settings.order > listCreate.order;
});

await check('live-apply: createfieldasxml bodies carry Options bit 8 and scrubbed XML with no ID=/SourceID=, in tier order', () => {
  const fieldPosts = applyPosts.filter((p) => p.url.includes('createfieldasxml'));
  const bodies = fieldPosts.map((p) => JSON.parse(p.body));
  const namesInOrder = bodies.map((b) => /Name="([^"]*)"/.exec(b.parameters.SchemaXml)?.[1]);
  return fieldPosts.length === 2
    && bodies.every((b) => (b.parameters.Options & 8) === 8)
    && bodies.every((b) => !b.parameters.SchemaXml.includes('ID=') && !b.parameters.SchemaXml.includes('SourceID='))
    && namesInOrder.join(',') === 'Budget,Region';   // plain before lookup
});

await check('live-apply: Indexed then EnforceUniqueValues MERGE onto the new field, right after its create', () => {
  const budgetCreate = applyPosts.find((p) => p.url.includes('createfieldasxml')
    && JSON.parse(p.body).parameters.SchemaXml.includes('Name="Budget"'));
  const fieldMerges = applyPosts.filter((p) => p.method === 'MERGE' && /\/fields\(guid'/.test(p.url)
    && p.order > budgetCreate.order)
    .sort((a, b) => a.order - b.order);
  const keys = fieldMerges.slice(0, 2).map((p) => Object.keys(JSON.parse(p.body))[0]);
  return keys[0] === 'Indexed' && keys[1] === 'EnforceUniqueValues';
});

await check('live-apply: addavailablecontenttype posts before any createfieldasxml', () => {
  const ct = applyPosts.find((p) => p.url.toLowerCase().includes('addavailablecontenttype'));
  const firstField = applyPosts.find((p) => p.url.includes('createfieldasxml'));
  return !!ct && !!firstField && ct.order < firstField.order;
});

await check('live-apply: the view posts POST → removeallviewfields → addviewfield×2, and validation MERGE is the final post', () => {
  const viewPost = applyPosts.find((p) => /\/views$/.test(p.url) && p.method === 'POST');
  const removeAll = applyPosts.find((p) => p.url.includes('removeallviewfields'));
  const addFields = applyPosts.filter((p) => p.url.includes('addviewfield'));
  const last = applyPosts[applyPosts.length - 1];
  const lastBody = JSON.parse(last.body || '{}');
  return !!viewPost && !!removeAll && addFields.length === 2
    && viewPost.order < removeAll.order && removeAll.order < addFields[0].order
    && lastBody.ValidationFormula === '=[Budget]>0';
});

await check('live-apply: a view with viewTypeKind 2 is POSTed with ViewTypeKind, and its Scope lands in the post-columns MERGE', async () => {
  const result = await liveApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'ViewKindTarget', baseTemplate: 100, contentTypesEnabled: false },
      fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true }],
      views: [{
        title: 'Gallery View', hidden: false, fields: ['Title'], rowLimit: 30, paged: true, viewQuery: '',
        viewTypeKind: 2, scope: 1,
      }],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'ViewKindTarget' } });
    return { created: report.created, viewStatus: report.steps.find((s) => s.kind === 'view.upsert')?.status };
  });
  const viewPost = applyPosts.find((p) => /\/views$/.test(p.url) && p.method === 'POST'
    && JSON.parse(p.body || '{}').Title === 'Gallery View');
  const scopeMerge = viewPost && applyPosts.find((p) => p.order > viewPost.order && p.method === 'MERGE'
    && /\/views\(guid'/.test(p.url) && (() => {
      try { return JSON.parse(p.body || '{}').Scope === 1; } catch { return false; }
    })());
  const viewCreateBody = viewPost ? JSON.parse(viewPost.body) : {};
  return result.created === true && result.viewStatus === 'done'
    && viewCreateBody.ViewTypeKind === 2 && Boolean(scopeMerge);
});

await liveApply.evaluate(async () => {
  const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
  const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
  const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
  const doc = buildSchemaDoc({
    list: { title: 'FailFieldTarget', baseTemplate: 100, contentTypesEnabled: false },
    fields: [
      { internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true },
      {
        internalName: 'BudgetF', displayName: 'BudgetF', type: 'Number', custom: true,
        schemaXml: '<Field Name="BudgetF" Type="Number" DisplayName="BudgetF" />',
      },
      {
        internalName: 'RegionX', displayName: 'RegionX', type: 'Lookup', custom: true, lookupList: 'Regions',
        schemaXml: '<Field Name="RegionX" Type="Lookup" DisplayName="RegionX" List="{oldguid}" ShowField="Title" />',
      },
    ],
    views: [],
  });
  const target = createSpRestClient({});
  await target.connectWeb('/sites/target');
  const spWrite = createSpWriteClient({ client: target });
  window.__LSA_RETRY_DOC__ = doc;
  window.__LSA_RETRY_TARGET__ = target;
  window.__LSA_RETRY_SPWRITE__ = spWrite;
});

await check('live-apply: a failing createfieldasxml for one field does not abort the run and is named in the report', async () => {
  applyFlags.failField = 'RegionX';
  const result = await liveApply.evaluate(async () => {
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = window.__LSA_RETRY_DOC__;
    const target = window.__LSA_RETRY_TARGET__;
    const spWrite = window.__LSA_RETRY_SPWRITE__;
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'FailFieldTarget' } });
    window.__LSA_RETRY_REPORT__ = report;
    return {
      created: report.created, aborted: report.aborted,
      budgetAdded: report.fields.added, failedNames: report.fields.failed.map((f) => f.internalName),
    };
  });
  applyFlags.failField = null;
  return result.created === true && result.aborted === '' && result.budgetAdded === 1
    && result.failedNames.join(',') === 'RegionX';
});

await check('live-apply: retrying posts exactly one createfieldasxml and no second web/lists', async () => {
  const listCreatesBefore = applyPosts.filter((p) => /\/web\/lists$/.test(p.url) && p.method === 'POST').length;
  const regionCreatesBefore = applyPosts.filter((p) => p.url.includes('createfieldasxml')
    && JSON.parse(p.body).parameters.SchemaXml.includes('Name="RegionX"')).length;

  const evalResult = await liveApply.evaluate(async () => {
    const { retryPlan } = await import('/src/workbench/list-schema.js');
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const report = window.__LSA_RETRY_REPORT__;
    const target = window.__LSA_RETRY_TARGET__;
    const spWrite = window.__LSA_RETRY_SPWRITE__;
    const plan = retryPlan(report);
    const retryReport = await runPlan(plan, { client: target, spWrite });
    return {
      stepCount: plan.steps.length, stepKind: plan.steps[0]?.kind,
      added: retryReport.fields.added, failed: retryReport.fields.failed.length,
    };
  });

  const listCreatesAfter = applyPosts.filter((p) => /\/web\/lists$/.test(p.url) && p.method === 'POST').length;
  const regionCreatesAfter = applyPosts.filter((p) => p.url.includes('createfieldasxml')
    && JSON.parse(p.body).parameters.SchemaXml.includes('Name="RegionX"')).length;

  return evalResult.stepCount === 1 && evalResult.stepKind === 'field.create'
    && evalResult.added === 1 && evalResult.failed === 0
    && listCreatesAfter === listCreatesBefore
    && regionCreatesAfter === regionCreatesBefore + 1;
});

await check('live-apply: one 429 is retried (Retry-After: 0)', async () => {
  applyFlags.throttleField = 'BudgetT';
  const before = applyPosts.filter((p) => p.url.includes('createfieldasxml')
    && JSON.parse(p.body).parameters.SchemaXml.includes('Name="BudgetT"')).length;
  const added = await liveApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'ThrottleTarget', baseTemplate: 100, contentTypesEnabled: false },
      fields: [{
        internalName: 'BudgetT', displayName: 'BudgetT', type: 'Number', custom: true,
        schemaXml: '<Field Name="BudgetT" Type="Number" DisplayName="BudgetT" />',
      }],
      views: [],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'ThrottleTarget' } });
    return report.fields.added;
  });
  const after = applyPosts.filter((p) => p.url.includes('createfieldasxml')
    && JSON.parse(p.body).parameters.SchemaXml.includes('Name="BudgetT"')).length;
  // Two POSTs recorded for the one field (the 429 attempt plus sp-write's
  // own built-in retry), but only one field counted as added.
  return added === 1 && after === before + 2;
});

await check('live-apply: a cross-web run fetches /sites/target/_api/contextinfo and writes only under /sites/target/_api/', () =>
  contextInfoUrls.length > 0 && contextInfoUrls.every((u) => u.includes('/sites/target/_api/contextinfo'))
  && applyPosts.length > 0 && applyPosts.every((p) => p.url.includes('/sites/target/_api/')));

await check('live-apply: a resumed (adopt) run posts no web/lists and MERGEs only EnableAttachments/EnableFolderCreation for settings', async () => {
  const listCreatesBefore = applyPosts.filter((p) => /\/web\/lists$/.test(p.url) && p.method === 'POST').length;
  const r = await liveApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: {
        title: 'LiveTarget', baseTemplate: 100, enableVersioning: true,
        enableAttachments: true, enableFolderCreation: true,
      },
      fields: [], views: [],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target });
    const plan = buildApplyPlan(doc, { title: 'LiveTarget', existing: 'resume' }, {
      existingList: { id: 'bbbbbbbb-1111-4000-8000-000000000001', title: 'LiveTarget' },
      targetLists: [{ title: 'LiveTarget' }],
    });
    const report = await runPlan(plan, { client: target, spWrite });
    return { settingsApplied: report.settings.applied, settingsFailed: report.settings.failed };
  });
  const listCreatesAfter = applyPosts.filter((p) => /\/web\/lists$/.test(p.url) && p.method === 'POST').length;
  return r.settingsApplied.slice().sort().join(',') === 'EnableAttachments,EnableFolderCreation'
    && r.settingsFailed.length === 0
    && listCreatesAfter === listCreatesBefore;
});

await check('live-apply: a 401 mid-run aborts with report.aborted === "auth"', () => {
  applyFlags.expireOnField = 'Region2';
  return liveApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'AuthTarget', baseTemplate: 100, contentTypesEnabled: false },
      fields: [
        {
          internalName: 'Budget2', displayName: 'Budget2', type: 'Number', custom: true,
          schemaXml: '<Field Name="Budget2" Type="Number" DisplayName="Budget2" />',
        },
        {
          internalName: 'Region2', displayName: 'Region2', type: 'Number', custom: true,
          schemaXml: '<Field Name="Region2" Type="Number" DisplayName="Region2" />',
        },
      ],
      views: [],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/target');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'AuthTarget' } });
    return { aborted: report.aborted, budget2: report.steps.find((s) => s.id === 'field:Budget2')?.status };
  }).then((r) => r.aborted === 'auth' && r.budget2 === 'done');
});

await liveApply.close();

// ---- Stubbed live: capture's named-only ValidationFormula/etc read --------
// A dedicated page/list id so the default (unprojected) list payload can be
// built by hand without them, exactly like the tenant this was verified
// against — the earlier `live` page's LIVE_LIST_ID fixture predates this
// follow-up read and doesn't stub it at all.

const NAMED_ONLY_LIST_ID = '44444444-0000-4000-8000-000000000001';
const namedOnlyReads = [];

const namedOnlyPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await namedOnlyPage.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await namedOnlyPage.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  namedOnlyReads.push(url);
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'NO-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
    const base = url.slice(0, url.indexOf('/_api/'));
    return route.fulfill({ json: { Id: 'web-id', Title: 'NamedOnly Web', Url: base, ServerRelativeUrl: '/' } });
  }
  if (url.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')`) && url.includes('$expand=RootFolder')) {
    // The default, unprojected payload — this tenant leaves the NAMED_ONLY
    // set off it entirely (verified live; see list-schema-capture.js).
    return route.fulfill({ json: {
      Id: NAMED_ONLY_LIST_ID, Title: 'NamedOnlyList', BaseTemplate: 100, BaseType: 0, ItemCount: 0,
      RootFolder: { ServerRelativeUrl: '/Lists/NamedOnlyList' }, ContentTypesEnabled: false,
      EnableVersioning: false, EnableAttachments: true,
    } });
  }
  if (url.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')`) && url.includes('$select=') && url.includes('ValidationFormula')) {
    return route.fulfill({ json: {
      ValidationFormula: '=[Budget]>0', ValidationMessage: 'Must be positive',
      OnQuickLaunch: true, ReadSecurity: 2, WriteSecurity: 4,
    } });
  }
  if (url.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')/fields`)
    || url.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')/views`)
    || url.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')/contenttypes`)) {
    return route.fulfill({ json: { value: [] } });
  }
  return route.fulfill({ json: { value: [] } });
});
await namedOnlyPage.goto(WB_URL);
await namedOnlyPage.waitForSelector('.wb-home-cards');

await check('live: capture issues the named-only $select read for ValidationFormula/etc when the default list payload lacks them, and the doc carries the values', async () => {
  const doc = await namedOnlyPage.evaluate(async (listId) => {
    const { captureListSchema } = await import('/src/workbench/list-schema-capture.js');
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const client = createSpRestClient({});
    await client.connectWeb('/sites/namedonly');
    const { doc: d } = await captureListSchema(client, listId);
    return d;
  }, NAMED_ONLY_LIST_ID);
  const followUp = namedOnlyReads.find((u) => u.includes(`lists(guid'${NAMED_ONLY_LIST_ID}')`)
    && u.includes('$select=') && u.includes('ValidationFormula') && u.includes('OnQuickLaunch'));
  return doc.list.validationFormula === '=[Budget]>0' && doc.list.validationMessage === 'Must be positive'
    && doc.list.onQuickLaunch === true && doc.list.readSecurity === 2 && doc.list.writeSecurity === 4
    && Boolean(followUp);
});

await namedOnlyPage.close();

// ---- Stubbed live: the commit boundary (urlName rename never fails the list step) ----

const urlNameApply = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await urlNameApply.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await urlNameApply.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'U-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      return route.fulfill({ json: { Id: 'web-id', Title: 'UrlName Web', Url: base, ServerRelativeUrl: '/' } });
    }
    if (url.includes('/GetList(@listUrl)')) {
      return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'List not found.' } } } });
    }
    return route.fulfill({ json: { value: [] } });
  }
  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    return route.fulfill({ json: {
      Id: 'urlname-list-id', Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/Lists/UrlSafeName' },
    } });
  }
  if (xHttpMethod === 'MERGE' && /lists\(guid'urlname-list-id'\)$/.test(url)) {
    let data = {};
    try { data = JSON.parse(request.postData() || '{}'); } catch { /* keep {} */ }
    const keys = Object.keys(data);
    // Only the urlName→title rename (a lone Title key) fails on this
    // tenant; the settings MERGE (many keys at once) succeeds normally.
    if (keys.length === 1 && keys[0] === 'Title') {
      return route.fulfill({
        status: 400,
        json: { 'odata.error': { message: { value: 'A list, survey, discussion board, or document library with the specified title already exists.' } } },
      });
    }
    return route.fulfill({ json: {} });
  }
  return route.fulfill({ json: {} });
});
await urlNameApply.goto(WB_URL);
await urlNameApply.waitForSelector('.wb-home-cards');

await check('live-apply: a failing urlName→title rename MERGE leaves the list step "done", with a warning, never "failed"', () =>
  urlNameApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'My List With Spaces', baseTemplate: 100, contentTypesEnabled: false },
      fields: [], views: [],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/urlname');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({
      doc, client: target, spWrite,
      options: { title: 'My List With Spaces', urlName: 'MyListWithSpaces' },
    });
    const listStep = report.steps.find((s) => s.id === 'list');
    return listStep.status === 'done' && report.created === true && report.listId === 'urlname-list-id'
      && report.warnings.some((w) => w.includes('MyListWithSpaces') && w.includes('My List With Spaces'));
  }));

await urlNameApply.close();

// ---- Stubbed live: a library's GetList pre-check URL has no /Lists/ segment ----

const libraryCreateApply = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await libraryCreateApply.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
const libraryGetListUrls = [];
await libraryCreateApply.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'LIB-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      return route.fulfill({ json: { Id: 'web-id', Title: 'Library Web', Url: base, ServerRelativeUrl: '/sites/librarytarget' } });
    }
    if (url.includes('/GetList(@listUrl)')) {
      libraryGetListUrls.push(url);
      return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'List not found.' } } } });
    }
    return route.fulfill({ json: { value: [] } });
  }
  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    return route.fulfill({ json: {
      Id: 'library-list-id', Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/sites/librarytarget/Reports' },
    } });
  }
  return route.fulfill({ json: {} });
});
await libraryCreateApply.goto(WB_URL);
await libraryCreateApply.waitForSelector('.wb-home-cards');

await check('live-apply: a library create’s GetList pre-check URL sits directly under the web (no /Lists/ segment), unlike a generic list’s', () =>
  libraryCreateApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const target = createSpRestClient({});
    await target.connectWeb('/sites/librarytarget');
    const spWrite = createSpWriteClient({ client: target });

    const libDoc = buildSchemaDoc({ list: { title: 'Reports', baseTemplate: 101 }, fields: [], views: [] });
    await applyListSchema({ doc: libDoc, client: target, spWrite, options: { title: 'Reports' } });

    const listDoc = buildSchemaDoc({ list: { title: 'Reports', baseTemplate: 100 }, fields: [], views: [] });
    await applyListSchema({ doc: listDoc, client: target, spWrite, options: { title: 'Reports' } });
    return true;
  }).then(() => {
    // @listUrl='…' is a query-string literal (odataPathLiteral encodes it,
    // slashes included) — decode before checking the path shape.
    const [libUrl, listUrl] = libraryGetListUrls.map((u) => decodeURIComponent(u));
    return Boolean(libUrl) && Boolean(listUrl)
      && libUrl.includes('/sites/librarytarget/Reports') && !libUrl.includes('/Lists/')
      && listUrl.includes('/sites/librarytarget/Lists/Reports');
  }));

await libraryCreateApply.close();

// ---- Stubbed live: verbose-odata retry nests __metadata correctly ---------

const verboseApply = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await verboseApply.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
const verbosePosts = [];
let verboseFieldFirstAttempt = true;
await verboseApply.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'V-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      return route.fulfill({ json: { Id: 'web-id', Title: 'Verbose Web', Url: base, ServerRelativeUrl: '/' } });
    }
    if (url.includes('/GetList(@listUrl)')) {
      return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'List not found.' } } } });
    }
    return route.fulfill({ json: { value: [] } });
  }
  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    return route.fulfill({ json: {
      Id: 'verbose-list-id', Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/Lists/VerboseTarget' },
    } });
  }
  if (url.includes('createfieldasxml')) {
    verbosePosts.push({ body: request.postData() || '', contentType: request.headers()['content-type'] || '' });
    if (verboseFieldFirstAttempt) {
      verboseFieldFirstAttempt = false;
      return route.fulfill({
        status: 400,
        json: { 'odata.error': { message: { value: 'The request could not be parsed as a valid entity of type SP.XmlSchemaFieldCreationInformation.' } } },
      });
    }
    return route.fulfill({ json: { Id: 'fid-budget', InternalName: 'Budget' } });
  }
  return route.fulfill({ json: {} });
});
await verboseApply.goto(WB_URL);
await verboseApply.waitForSelector('.wb-home-cards');

await check('live-apply: a 400 naming SP.XmlSchemaFieldCreationInformation retries createfieldasxml verbose, with __metadata nested INSIDE parameters (not top-level)', async () => {
  const r = await verboseApply.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'VerboseTarget', baseTemplate: 100, contentTypesEnabled: false },
      fields: [{
        internalName: 'Budget', displayName: 'Budget', type: 'Number', custom: true,
        schemaXml: '<Field Name="Budget" Type="Number" DisplayName="Budget" />',
      }],
      views: [],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/verbose');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'VerboseTarget' } });
    return { added: report.fields.added, failed: report.fields.failed.length };
  });
  if (r.added !== 1 || r.failed !== 0 || verbosePosts.length !== 2) return false;
  const retried = JSON.parse(verbosePosts[1].body);
  return verbosePosts[1].contentType.includes('odata=verbose')
    && retried.parameters?.__metadata?.type === 'SP.XmlSchemaFieldCreationInformation'
    && retried.__metadata === undefined;
});

await verboseApply.close();

// ---- Stubbed live: reads between steps abort cleanly (never reject) -------

const PREFLIGHT_LIST_ID = 'preflight-list-id';
const preflightPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await preflightPage.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await preflightPage.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'P-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
    const base = url.slice(0, url.indexOf('/_api/'));
    return route.fulfill({ json: { Id: 'web-id', Title: 'Preflight Web', Url: base, ServerRelativeUrl: '/' } });
  }
  if (url.includes(`lists(guid'${PREFLIGHT_LIST_ID}')/fields`)) {
    return route.fulfill({ status: 401, json: { 'odata.error': { message: { value: 'The security token is expired.' } } } });
  }
  return route.fulfill({ json: { value: [] } });
});
await preflightPage.goto(WB_URL);
await preflightPage.waitForSelector('.wb-home-cards');

await check('live-apply: a 401 on the preflight fields read (priming a resumed run) returns report.aborted "auth" — runPlan never rejects', () =>
  preflightPage.evaluate(async (listId) => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const target = createSpRestClient({});
    await target.connectWeb('/sites/preflight');
    const plan = {
      title: 'Preflight', targetWebUrl: target.webUrl(), existingListId: listId,
      steps: [{
        id: 'settings', kind: 'list.settings', label: 'Apply list settings', dependsOn: [],
        payload: { groupA: {}, groupB: {} }, refs: {}, optional: false, status: 'planned', error: '', final: false, result: null,
      }],
      warnings: [],
    };
    let threw = false;
    let report = null;
    try { report = await runPlan(plan, { client: target }); } catch { threw = true; }
    return !threw && report?.aborted === 'auth' && report.steps[0].status === 'planned';
  }, PREFLIGHT_LIST_ID));

await preflightPage.close();

const PREVIEW_LIST_ID = 'preview-list-id';
let preViewFieldsReadCount = 0;
let preViewRemoveAllViewFieldsCalls = 0;
const preViewPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await preViewPage.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await preViewPage.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';
  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'PV-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      return route.fulfill({ json: { Id: 'web-id', Title: 'PreView Web', Url: base, ServerRelativeUrl: '/' } });
    }
    if (url.includes(`lists(guid'${PREVIEW_LIST_ID}')/fields`)) {
      preViewFieldsReadCount++;
      if (preViewFieldsReadCount === 1) {
        return route.fulfill({ json: { value: [
          { Id: 'f1', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false },
        ] } });
      }
      return route.fulfill({ status: 500, json: { 'odata.error': { message: { value: 'Internal Server Error.' } } } });
    }
    if (url.includes(`lists(guid'${PREVIEW_LIST_ID}')/views`)) {
      return route.fulfill({ json: { value: [] } });
    }
    return route.fulfill({ json: { value: [] } });
  }
  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    return route.fulfill({ json: {
      Id: PREVIEW_LIST_ID, Title: JSON.parse(request.postData() || '{}').Title,
      RootFolder: { ServerRelativeUrl: '/Lists/PreView' },
    } });
  }
  if (url.includes('removeallviewfields')) {
    preViewRemoveAllViewFieldsCalls++;
    return route.fulfill({ json: {} });
  }
  return route.fulfill({ json: {} });
});
await preViewPage.goto(WB_URL);
await preViewPage.waitForSelector('.wb-home-cards');

await check('live-apply: a 500 on the pre-view fields re-read posts no removeallviewfields, leaves the view step "planned", and returns report.aborted "probe"', async () => {
  const r = await preViewPage.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildApplyPlan, buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'PreView', baseTemplate: 100, contentTypesEnabled: false },
      fields: [],
      views: [{ title: 'All Items', hidden: false, fields: [], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const plan = buildApplyPlan(doc, { title: 'PreView' }, { existingList: null, existingFields: [], targetLists: [] });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/preview');
    const spWrite = createSpWriteClient({ client: target });
    const report = await runPlan(plan, { client: target, spWrite });
    return {
      aborted: report.aborted,
      listStepStatus: report.steps.find((s) => s.id === 'list').status,
      viewStepStatus: report.steps.find((s) => s.kind === 'view.upsert').status,
    };
  });
  return r.aborted === 'probe' && r.listStepStatus === 'done' && r.viewStepStatus === 'planned'
    && preViewRemoveAllViewFieldsCalls === 0;
});

await preViewPage.close();

// ---- Stubbed live: the dialog's resume-based retry ------------------------
// Mirrors exactly what list-schema-dialog.js's runRetry() now does
// (probeTarget → buildApplyPlan in resume mode → runPlan), driven directly
// rather than through the dialog UI, against a small stateful fixture kept
// in Node (this route handler) so the second run sees what the first run
// actually left behind.

let rrCreatedList = null;
let rrFields = [];
let rrViews = [];
let rrBudgetShouldFail = true;
let rrWebListsPosts = 0;
const rrAddViewFieldCalls = [];
let rrIndexedMergeApplied = false;

const retryResumePage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await retryResumePage.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await retryResumePage.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  const httpMethod = request.method();
  const xHttpMethod = request.headers()['x-http-method'] || '';

  if (url.includes('/_api/contextinfo')) {
    return route.fulfill({
      json: { FormDigestValue: 'RR-DIGEST', FormDigestTimeoutSeconds: 1800, WebFullUrl: new URL(url).origin },
    });
  }
  if (httpMethod === 'GET') {
    if (/\/_api\/web(\?|$)/.test(url) && !url.includes('/_api/web/')) {
      const base = url.slice(0, url.indexOf('/_api/'));
      return route.fulfill({ json: { Id: 'web-id', Title: 'RetryResume Web', Url: base, ServerRelativeUrl: '/' } });
    }
    if (url.includes('/GetList(@listUrl)')) {
      return route.fulfill({ status: 404, json: { 'odata.error': { message: { value: 'List not found.' } } } });
    }
    if (rrCreatedList && url.includes(`lists(guid'${rrCreatedList.Id}')/fields`)) {
      return route.fulfill({ json: { value: rrFields } });
    }
    if (rrCreatedList && url.includes(`lists(guid'${rrCreatedList.Id}')/views`) && !url.includes('viewfields')) {
      return route.fulfill({ json: { value: rrViews } });
    }
    if (rrCreatedList && url.includes('/viewfields')) {
      return route.fulfill({ json: { Items: [] } });
    }
    if (/\/web\/lists(\?|$)/.test(url) && !url.includes("lists(guid'")) {
      return route.fulfill({ json: { value: rrCreatedList ? [rrCreatedList] : [] } });
    }
    return route.fulfill({ json: { value: [] } });
  }

  if (/\/_api\/web\/lists$/.test(url) && !xHttpMethod) {
    rrWebListsPosts++;
    const data = JSON.parse(request.postData() || '{}');
    rrCreatedList = {
      Id: 'rr-list-id', Title: data.Title, BaseTemplate: data.BaseTemplate ?? 100,
      ContentTypesEnabled: !!data.ContentTypesEnabled, RootFolder: { ServerRelativeUrl: '/Lists/RetryResume' },
    };
    rrFields = [{ Id: 'f-title', InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false, ReadOnlyField: false }];
    rrViews = [];
    return route.fulfill({ json: { Id: rrCreatedList.Id, Title: rrCreatedList.Title, RootFolder: rrCreatedList.RootFolder } });
  }

  if (url.includes('createfieldasxml')) {
    const parsed = JSON.parse(request.postData() || '{}');
    const xml = parsed?.parameters?.SchemaXml || '';
    const name = /Name="([^"]*)"/.exec(xml)?.[1] || '';
    if (name === 'Budget' && rrBudgetShouldFail) {
      rrBudgetShouldFail = false;   // fails exactly once — the retry succeeds
      return route.fulfill({ status: 400, json: { 'odata.error': { message: { value: "SharePoint could not create the column 'Budget'." } } } });
    }
    const id = `fid-${name.toLowerCase()}`;
    rrFields.push({ Id: id, InternalName: name, Title: name, TypeAsString: 'Number', Hidden: false, ReadOnlyField: false });
    return route.fulfill({ json: { Id: id, InternalName: name } });
  }

  if (xHttpMethod === 'MERGE' && /\/fields\(guid'/.test(url)) {
    let data = {};
    try { data = JSON.parse(request.postData() || '{}'); } catch { /* keep {} */ }
    if ('Indexed' in data) rrIndexedMergeApplied = true;
    return route.fulfill({ json: {} });
  }

  if (/\/views$/.test(url) && !xHttpMethod) {
    const data = JSON.parse(request.postData() || '{}');
    const id = 'view-all-items';
    rrViews.push({ Id: id, Title: data.Title, DefaultView: !!data.DefaultView });
    return route.fulfill({ json: { Id: id, Title: data.Title } });
  }

  if (url.includes('addviewfield')) {
    const m = /addviewfield\('([^']*)'\)/.exec(url);
    if (m) rrAddViewFieldCalls.push(decodeURIComponent(m[1]));
    return route.fulfill({ json: {} });
  }
  if (url.includes('removeallviewfields')) {
    return route.fulfill({ json: {} });
  }

  return route.fulfill({ json: {} });
});
await retryResumePage.goto(WB_URL);
await retryResumePage.waitForSelector('.wb-home-cards');

await check('live-apply retry: resuming after a failed field POSTs no web/lists, re-creates the field, applies its Indexed merge, and rebuilds the view with the new column', async () => {
  const first = await retryResumePage.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const { buildSchemaDoc } = await import('/src/workbench/list-schema.js');
    const { applyListSchema } = await import('/src/workbench/list-schema-apply.js');
    const doc = buildSchemaDoc({
      list: { title: 'RetryResume', baseTemplate: 100, contentTypesEnabled: false },
      fields: [{
        internalName: 'Budget', displayName: 'Budget', type: 'Number', custom: true, indexed: true,
        schemaXml: '<Field Name="Budget" Type="Number" DisplayName="Budget" Indexed="TRUE" />',
      }],
      views: [{ title: 'All Items', hidden: false, fields: ['Budget'], rowLimit: 30, paged: true, viewQuery: '' }],
    });
    const target = createSpRestClient({});
    await target.connectWeb('/sites/retryresume');
    const spWrite = createSpWriteClient({ client: target });
    const report = await applyListSchema({ doc, client: target, spWrite, options: { title: 'RetryResume' } });
    window.__RR_DOC__ = doc;
    window.__RR_REPORT__ = report;
    return { created: report.created, budgetFailed: report.fields.failed.map((f) => f.internalName) };
  });

  const webListsPostsAfterFirst = rrWebListsPosts;

  const second = await retryResumePage.evaluate(async () => {
    const { probeTarget } = await import('/src/workbench/list-schema-capture.js');
    const { buildApplyPlan } = await import('/src/workbench/list-schema.js');
    const { runPlan } = await import('/src/workbench/list-schema-apply.js');
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    const doc = window.__RR_DOC__;
    const prior = window.__RR_REPORT__;
    const target = createSpRestClient({});
    await target.connectWeb('/sites/retryresume');
    const spWrite = createSpWriteClient({ client: target });
    const resumeProbe = await probeTarget(target, { title: prior.title, doc });
    const plan = buildApplyPlan(doc, { title: resumeProbe.existingList.title, existing: 'resume' }, resumeProbe);
    const report = await runPlan(plan, { client: target, spWrite });
    return {
      adopted: report.adopted,
      fieldsAdded: report.fields.added, fieldsFailed: report.fields.failed.length,
      viewsUpdated: report.views.updated, viewsAdded: report.views.added,
      firstStepKind: plan.steps[0]?.kind,
    };
  });

  return first.created === true && first.budgetFailed.join(',') === 'Budget'
    && second.firstStepKind === 'list.adopt'
    && second.fieldsAdded === 1 && second.fieldsFailed === 0
    && (second.viewsUpdated + second.viewsAdded) === 1
    && rrWebListsPosts === webListsPostsAfterFirst
    && rrIndexedMergeApplied === true
    && rrAddViewFieldCalls.includes('Budget');
});

await retryResumePage.close();

// ---- Mock UI: the apply dialog (slice 2b) ----------------------------------
// Own page: opens the Schema tab's "Copy to…" and the all-lists grid's "New
// from schema…" fresh, so it never inherits state from the read-only
// schemaPage tests above (which end on the Documents list).

const dialogPage = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await dialogPage.goto(WB_URL);
await dialogPage.waitForSelector('.wb-home-cards');
await dialogPage.fill('#wb-site-input', '/sites/schema');
await dialogPage.locator('#wb-site-open').click();
await dialogPage.waitForFunction(() =>
  document.getElementById('wb-status-context').textContent.includes('/sites/schema'));
await dialogPage.locator('.wb-rail-btn', { hasText: 'Lists' }).click();
await dialogPage.waitForSelector('.wb-table tbody tr', { hasText: 'Requests' });
await dialogPage.locator('.wb-table tbody tr', { hasText: 'Requests' }).locator('td').first().click();
await dialogPage.waitForSelector('.wb-tab');
await dialogPage.locator('.wb-tab', { hasText: 'Schema' }).click();
await dialogPage.waitForSelector('.wb-schema-copy');

// Every dialog below except the final one (Dialog A) targets /sites/target
// or an invalid URL, so the /sites/schema all-lists grid still shows exactly
// one 'Requests' row until then — the last block deliberately runs the
// same-web copy (which adds a 'Requests Copy' row) after everything else
// that re-selects 'Requests' from that grid.

// -- Dialog B: cross-web target (/sites/target) — lookups, dry run, create ordering

await dialogPage.locator('.wb-schema-copy').click();
await dialogPage.waitForSelector('.wb-schema-dialog');
// Wait out the dialog's own auto-connect (to the same web) before driving it
// further — racing a second connect() against the boot-time one is exactly
// what a real user's fast typing could also trigger, but the assertions
// below need one settled connection to reason about.
await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);

await check('dialog: connecting to /sites/target keeps the source view alive and defaults the title to ‘Requests’', async () => {
  const before = await dialogPage.locator('#wb-status-context').textContent();
  await dialogPage.fill('.wb-schema-target', '/sites/target');
  await dialogPage.locator('.wb-schema-connect').click();
  await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value === 'Requests');
  // The dialog's target client is a second, independent connection — the
  // shell's own status line (the source view's "inspecting …" note) must
  // read exactly as it did before the dialog ever touched /sites/target.
  const after = await dialogPage.locator('#wb-status-context').textContent();
  return after === before && !after.includes('/sites/target');
});

await check('dialog: lookup rows show Clients present and Region missing, with the policy radio', () =>
  dialogPage.evaluate(() => {
    const rows = [...document.querySelectorAll('.wb-schema-lookups tbody tr')];
    const clientRow = rows.find((tr) => tr.cells[0].textContent.trim() === 'Client');
    const regionRow = rows.find((tr) => tr.cells[0].textContent.trim() === 'Region');
    const clientSelect = clientRow?.querySelector('select');
    const regionSelect = regionRow?.querySelector('select');
    const policyVisible = !document.querySelector('.wb-schema-lookup-policy').hidden;
    return clientSelect?.value === 'Clients' && regionSelect?.value === '' && policyVisible;
  }));

await check('dialog: Dry run lists ordered steps and writes nothing', async () => {
  await dialogPage.evaluate(() => { window.__DCSPAD_WB_WRITES__ = []; });
  await dialogPage.locator('.wb-schema-dryrun').click();
  await dialogPage.waitForFunction(() => document.querySelectorAll('.wb-schema-steps li').length > 3);
  const steps = await dialogPage.locator('.wb-schema-steps li').allTextContents();
  const writes = await dialogPage.evaluate(() => (window.__DCSPAD_WB_WRITES__ || []).length);
  return steps.length > 3 && writes === 0;
});

await check('dialog: the text policy changes the Region plan line', async () => {
  await dialogPage.locator('.wb-schema-lookup-policy input[value="text"]').check();
  await dialogPage.locator('.wb-schema-dryrun').click();
  await dialogPage.waitForFunction(() => [...document.querySelectorAll('.wb-schema-steps li')]
    .some((li) => li.textContent.includes('Region') && li.textContent.includes('as text')));
  await dialogPage.locator('.wb-schema-lookup-policy input[value="skip"]').check();
  return true;
});

await check('dialog: editing the target input after connect disables Create until Connect is clicked again', async () => {
  const beforeDisabled = await dialogPage.locator('.wb-schema-create').isDisabled();
  // A trivial edit (still names the same site) — the point is that the
  // input changed at all, not that it now names somewhere else.
  await dialogPage.fill('.wb-schema-target', '/sites/target/');
  const afterEditDisabled = await dialogPage.locator('.wb-schema-create').isDisabled();
  const dryRunDisabledToo = await dialogPage.locator('.wb-schema-dryrun').isDisabled();
  const statusText = await dialogPage.locator('.wb-schema-target-status').textContent();
  // Restore the connection so the next check (Create) still has one.
  await dialogPage.fill('.wb-schema-target', '/sites/target');
  await dialogPage.locator('.wb-schema-connect').click();
  await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value === 'Requests');
  return !beforeDisabled && afterEditDisabled && dryRunDisabledToo
    && statusText.includes('Connect to check this site');
});

await check('dialog: Create records web/lists first and the validation MERGE last, and shows the mock-mode sentence with Cancel relabeled Close', async () => {
  await dialogPage.evaluate(() => { window.__DCSPAD_WB_WRITES__ = []; });
  await dialogPage.locator('.wb-schema-create').click();
  await dialogPage.waitForSelector('.wb-schema-report:not([hidden])');
  const headline = await dialogPage.locator('.wb-schema-report-headline').textContent();
  const cancelText = await dialogPage.locator('.wb-schema-cancel').textContent();
  const result = await dialogPage.evaluate(() => {
    const writes = window.__DCSPAD_WB_WRITES__ || [];
    const method = (w) => w.headers?.['X-HTTP-Method'] || w.headers?.['x-http-method'] || '';
    const lower = (w) => w.url.toLowerCase();
    const listIdx = writes.findIndex((w) => /\/web\/lists$/.test(lower(w)) && !method(w));
    const settingsIdx = writes.findIndex((w, i) => i > listIdx && method(w) === 'MERGE'
      && /lists\(guid'[0-9a-f-]+'\)$/.test(lower(w)));
    const last = writes[writes.length - 1];
    const validationLast = !!last && method(last) === 'MERGE' && JSON.parse(last.body || '{}').ValidationFormula !== undefined;
    return { listIdx, settingsIdx, validationLast };
  });
  return headline.includes('mock mode') && cancelText.trim() === 'Close'
    && result.listIdx === 0 && result.settingsIdx > result.listIdx && result.validationLast;
});

await check('dialog: the report .md downloads', async () => {
  const [download] = await Promise.all([
    dialogPage.waitForEvent('download'),
    dialogPage.locator('.wb-schema-report-download').click(),
  ]);
  return /^schema-report-.*\.md$/.test(download.suggestedFilename());
});

await dialogPage.locator('.wb-schema-cancel').click();
await dialogPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });

// -- Dialog C: cross-tenant URL is refused, Create stays disabled

await dialogPage.locator('.wb-schema-copy').click();
await dialogPage.waitForSelector('.wb-schema-dialog');
await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);
await check('dialog: a cross-tenant URL shows the different-tenant sentence and Create stays disabled', async () => {
  await dialogPage.fill('.wb-schema-target', 'https://not-this-tenant.example.com/sites/x');
  await dialogPage.locator('.wb-schema-connect').click();
  await dialogPage.waitForFunction(() =>
    document.querySelector('.wb-schema-target-status')?.textContent.includes('different tenant'));
  const disabled = await dialogPage.locator('.wb-schema-create').isDisabled();
  return disabled;
});
await dialogPage.locator('.wb-schema-close').click();
await dialogPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });

// -- Dialog D: an existing title on the target gates reconcile behind the consent box

await dialogPage.locator('.wb-schema-copy').click();
await dialogPage.waitForSelector('.wb-schema-dialog');
await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);
await dialogPage.fill('.wb-schema-target', '/sites/target');
await dialogPage.locator('.wb-schema-connect').click();
// Dialog B already created a 'Requests' list on /sites/target (this page's
// mock writer state persists across dialogs), so the default title here is
// 'Requests Copy', not 'Requests' — wait for the probe to settle (the
// lookup rows only render once it has) rather than pinning that value.
await dialogPage.waitForFunction(() => document.querySelectorAll('.wb-schema-lookups tbody tr').length > 0);

await check('dialog: an existing title (Archive Requests) gates reconcile behind the consent box and relabels Create', async () => {
  await dialogPage.fill('.wb-schema-title', 'Archive Requests');
  await dialogPage.waitForFunction(() =>
    document.querySelector('.wb-schema-title-status')?.textContent.includes('already exists'));
  const createDisabledDefault = await dialogPage.locator('.wb-schema-create').isDisabled();
  await dialogPage.locator('.wb-schema-existing input[value="resume"]').check();
  await dialogPage.waitForSelector('.wb-schema-gate:not([hidden])');
  const createStillDisabled = await dialogPage.locator('.wb-schema-create').isDisabled();
  await dialogPage.locator('.wb-schema-gate input[type="checkbox"]').check();
  const label = await dialogPage.locator('.wb-schema-create').textContent();
  const createEnabled = await dialogPage.locator('.wb-schema-create').isEnabled();
  return createDisabledDefault && createStillDisabled && createEnabled && label.trim() === 'Add to existing list';
});

await check('dialog: consent resets when the colliding title changes from one existing list to another (Archive Requests → Clients)', async () => {
  await dialogPage.fill('.wb-schema-title', 'Clients');
  await dialogPage.waitForFunction(() =>
    document.querySelector('.wb-schema-title-status')?.textContent.includes('Clients'));
  const newRadioChecked = await dialogPage.locator('.wb-schema-existing input[value="new"]').isChecked();
  const resumeRadioChecked = await dialogPage.locator('.wb-schema-existing input[value="resume"]').isChecked();
  const gateHidden = await dialogPage.locator('.wb-schema-gate').isHidden();
  const gateChecked = await dialogPage.locator('.wb-schema-gate input[type="checkbox"]').isChecked();
  return newRadioChecked && !resumeRadioChecked && gateHidden && !gateChecked;
});

await dialogPage.locator('.wb-schema-close').click();
await dialogPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });

// -- New from schema… on the all-lists grid

await dialogPage.locator('.wb-back').click();
await dialogPage.waitForSelector('.wb-schema-new');

await check('dialog: New from schema… rejects a non-JSON file inline', async () => {
  await dialogPage.setInputFiles('.wb-schema-file', {
    name: 'not-json.txt', mimeType: 'text/plain', buffer: Buffer.from('not json'),
  });
  await dialogPage.waitForSelector('.wb-schema-import-notice:not([hidden])');
  const text = await dialogPage.locator('.wb-schema-import-notice').textContent();
  return text.includes('valid JSON');
});

await check('dialog: New from schema… rejects an oversized file inline', async () => {
  await dialogPage.locator('.wb-schema-import-notice button', { hasText: 'Dismiss' }).click();
  await dialogPage.setInputFiles('.wb-schema-file', {
    name: 'huge.json', mimeType: 'application/json', buffer: Buffer.alloc(6 * 1024 * 1024, '1'),
  });
  await dialogPage.waitForSelector('.wb-schema-import-notice:not([hidden])');
  const text = await dialogPage.locator('.wb-schema-import-notice').textContent();
  return text.includes('MB') && text.includes('import limit');
});

await check('dialog: New from schema… imports a v1 doc titled like an existing list and opens with ‘<Title> Copy’', async () => {
  await dialogPage.locator('.wb-schema-import-notice button', { hasText: 'Dismiss' }).click();
  const v1Doc = {
    kind: 'dcspad-sputils-list-schema', version: 1,
    source: { listTitle: 'Clients' },
    list: { title: 'Clients', baseTemplate: 100 },
    fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true }],
    views: [{ title: 'All Items', fields: ['Title'] }],
    contentTypes: [], warnings: [],
  };
  await dialogPage.setInputFiles('.wb-schema-file', {
    name: 'clients-schema.json', mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(v1Doc)),
  });
  await dialogPage.waitForSelector('.wb-schema-dialog');
  await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value === 'Clients Copy');
  return true;
});
await dialogPage.locator('.wb-schema-close').click();
await dialogPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });

// -- Dialog A: same-web copy → default title, Items gate, Create, Open the new list
// Last: the created 'Requests Copy' would otherwise make every earlier
// re-selection of the 'Requests' row ambiguous.

await dialogPage.locator('.wb-table tbody tr', { hasText: 'Requests' }).locator('td').first().click();
await dialogPage.locator('.wb-tab', { hasText: 'Schema' }).click();
await dialogPage.waitForSelector('.wb-schema-copy');

await check('dialog: Copy to… on the same web defaults the title to ‘Requests Copy’', async () => {
  await dialogPage.locator('.wb-schema-copy').click();
  await dialogPage.waitForSelector('.wb-schema-dialog');
  await dialogPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value === 'Requests Copy');
  return true;
});

await check('dialog: the Items fieldset is enabled in copy mode (stage 1b-b)', async () => {
  const disabled = await dialogPage.locator('.wb-schema-items').evaluate((f) => f.disabled);
  return disabled === false;
});

await check('dialog: Create (same web) writes the list and Open the new list navigates to it', async () => {
  await dialogPage.locator('.wb-schema-create').click();
  await dialogPage.waitForSelector('.wb-schema-report:not([hidden])');
  await dialogPage.locator('.wb-schema-report-open').click();
  await dialogPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });
  await dialogPage.waitForFunction(() =>
    document.querySelector('.wb-detail-head h2')?.textContent === 'Requests Copy');
  return true;
});

await dialogPage.close();

// ---- Stubbed live: the apply dialog's own 401 mid-run ----------------------
// Reuses the existing `live` page (already SP-context-stubbed for the
// capture checks above) — its route handler was extended, above, with the
// write endpoints a live Create needs (web/lists POST, settings MERGE,
// createfieldasxml) plus a Regions list so the Region lookup resolves.

await check('live-dialog: a 401 mid-run shows EXPIRED_SESSION_NOTE', async () => {
  await live.locator('.wb-back').click();
  await live.locator('.wb-table tbody tr', { hasText: 'LiveRequests' }).locator('td').first().click();
  await live.locator('.wb-tab', { hasText: 'Schema' }).click();
  await live.waitForSelector('.wb-schema-copy');
  liveDialogFlags.expireOnField = true;
  await live.locator('.wb-schema-copy').click();
  await live.waitForSelector('.wb-schema-dialog');
  await live.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);
  await live.locator('.wb-schema-create').click();
  await live.waitForSelector('.wb-schema-report:not([hidden])');
  const headline = await live.locator('.wb-schema-report-headline').textContent();
  return headline.includes('expired') && headline.includes('reload the page');
});

// ---- Stage 1b-a: list item DATA capture + export (pure list-data.js) ------

await check('pure: a v1 data document normalizes to v2 (items/folders split, warnings kept); the wrong kind is refused', () =>
  page.evaluate(async () => {
    const { normalizeDataDoc, DATA_KIND } = await import('/src/workbench/list-data.js');
    const v1 = {
      kind: DATA_KIND, version: 1,
      source: { siteUrl: 'https://t', listTitle: 'Requests', listId: 'abc' },
      fields: { Title: { type: 'Text', custom: false, readOnly: false } },
      items: [
        { Id: 1, ID: 1, Title: 'A', _resolved: {}, _dir: '' },
        { Id: 2, ID: 2, Title: 'Sub', _folder: true, _folderPath: 'Sub', _resolved: {}, _dir: '' },
      ],
      warnings: ['a warning'],
    };
    const v2 = normalizeDataDoc(v1);
    let refused = false;
    try { normalizeDataDoc({ kind: 'not-a-data-doc' }); } catch (e) { refused = e?.code === 'bad-data'; }
    return v2.version === 2 && v2.items.length === 1 && v2.items[0].Id === 1
      && v2.folders.length === 1 && v2.folders[0]._folderPath === 'Sub'
      && v2.warnings[0] === 'a warning' && refused;
  }));

await check('pure: folderOrder sorts parents before children, stable within a depth', () =>
  page.evaluate(async () => {
    const { folderOrder } = await import('/src/workbench/list-data.js');
    const folders = [
      { _folderPath: 'A/B/C' }, { _folderPath: 'A' }, { _folderPath: 'A/B' }, { _folderPath: 'X' },
    ];
    const order = folderOrder(folders).map((f) => f._folderPath);
    return order.join(',') === 'A,X,A/B,A/B/C';
  }));

await check('pure: writableFields drops NEVER_WRITE names/types, read-only columns, and anything missing on the target', () =>
  page.evaluate(async () => {
    const { writableFields } = await import('/src/workbench/list-data.js');
    const schemaFields = {
      Title: { type: 'Text', custom: false },
      Status: { type: 'Choice', custom: true },
      ID: { type: 'Counter', custom: false },
      Modified: { type: 'DateTime', custom: false },
      Ghost: { type: 'Text', custom: true },
      Locked: { type: 'Text', custom: true },
    };
    const targetFields = [
      { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
      { InternalName: 'Status', TypeAsString: 'Choice', ReadOnlyField: false },
      { InternalName: 'ID', TypeAsString: 'Counter', ReadOnlyField: true },
      { InternalName: 'Modified', TypeAsString: 'DateTime', ReadOnlyField: true },
      { InternalName: 'Locked', TypeAsString: 'Text', ReadOnlyField: true },
    ];
    const names = writableFields(schemaFields, targetFields).map((x) => x.name);
    return names.includes('Title') && names.includes('Status')
      && !names.includes('ID') && !names.includes('Modified')
      && !names.includes('Ghost') && !names.includes('Locked');
  }));

await check('pure: toImportFormValues ports SPUtils’ per-type FieldValue conventions (Choice, MultiChoice, Number, Boolean, calibrated DateTime, User, Lookup, LookupMulti, URL)', () =>
  page.evaluate(async () => {
    const { toImportFormValues } = await import('/src/workbench/list-data.js');
    const fields = [
      { name: 'Status', typeAsString: 'Choice' },
      { name: 'Tags', typeAsString: 'MultiChoice' },
      { name: 'Budget', typeAsString: 'Number' },
      { name: 'Approved', typeAsString: 'Boolean' },
      { name: 'Due', typeAsString: 'DateTime', tf: {} },
      { name: 'Owner', typeAsString: 'User' },
      { name: 'Client', typeAsString: 'Lookup' },
      { name: 'Regions', typeAsString: 'LookupMulti' },
      { name: 'Link', typeAsString: 'URL' },
    ];
    const item = {
      Status: 'Active', Tags: ['A', 'B'], Budget: 42, Approved: true,
      Due: '2026-03-04T05:06:00Z',
      Link: { Url: 'https://example.com', Description: 'Example' },
      _resolved: {
        Owner: [{ Email: 'pat@mock.local', LoginName: 'i:0#.f|membership|pat@mock.local' }],
        Client: [{ Id: 9, value: 'Acme' }],
        Regions: [{ Id: 1, value: 'North' }, { Id: 2, value: 'South' }],
      },
    };
    const userIds = new Map([['pat@mock.local', 'i:0#.f|membership|pat@mock.local']]);
    const lookupIds = new Map([
      ['Client', new Map([['Acme', [9]]])],
      ['Regions', new Map([['North', [1]], ['South', [2]]])],
    ]);
    const { values, errors } = toImportFormValues(item, fields, {
      dateFormat: { order: 'dmy', sep: '/', timeSep: ':', offsetMinutes: 0 },
      userIds, lookupIds,
    });
    const byName = Object.fromEntries(values.map((v) => [v.FieldName, v.FieldValue]));
    return errors.length === 0
      && byName.Status === 'Active'
      && byName.Tags === ';#A;#B;#'
      && byName.Budget === '42'
      && byName.Approved === '1'
      && byName.Due === '04/03/2026 05:06'
      && byName.Owner === JSON.stringify([{ Key: 'i:0#.f|membership|pat@mock.local' }])
      && byName.Client === '9'
      && byName.Regions === '1;#2;#'
      && byName.Link === 'https://example.com, Example';
  }));

await check('pure: resolveLookupByShownValue resolves a unique match, accepts an ambiguous same-list match carrying the source id, and reports everything else', () =>
  page.evaluate(async () => {
    const { resolveLookupByShownValue } = await import('/src/workbench/list-data.js');
    const unique = resolveLookupByShownValue('Acme', [9]);
    const ambiguousSameList = resolveLookupByShownValue('Dup', [5, 7], { sourceId: 7, sameLookupList: true });
    const ambiguousOtherList = resolveLookupByShownValue('Dup', [5, 7], { sourceId: 7, sameLookupList: false });
    const notFound = resolveLookupByShownValue('Missing', [], { sourceId: 3 });
    const empty = resolveLookupByShownValue('', [1, 2], { sourceId: 4 });
    return unique.id === 9 && !unique.error
      && ambiguousSameList.id === 7 && !ambiguousSameList.error
      && ambiguousOtherList.id === null && ambiguousOtherList.error.includes('matches 2 items')
      && notFound.id === null && notFound.error.includes('was not found')
      && empty.id === null && empty.error.includes('had no shown value');
  }));

await check('pure: partitionPasses keeps self-lookups out of pass1 and only lists items that actually need pass2/pass3', () =>
  page.evaluate(async () => {
    const { partitionPasses } = await import('/src/workbench/list-data.js');
    const fields = [
      { name: 'Status', typeAsString: 'Choice', meta: { custom: true } },
      { name: 'ParentRequest', typeAsString: 'Lookup', meta: { custom: true, isSelfLookup: true } },
    ];
    const items = [
      { Id: 1, ID: 1, Status: 'A', _resolved: {} },
      { Id: 2, ID: 2, Status: 'B', _resolved: { ParentRequest: [{ Id: 1, value: 'A' }] } },
      { Id: 3, ID: 3, Status: 'C', Created: '2026-01-01T00:00:00Z', _resolved: { Author: [{ Email: 'x@y' }] } },
    ];
    const { pass1, pass2, pass3 } = partitionPasses(items, fields);
    const pass1Names = pass1.fields.map((f) => f.name);
    const pass2Ids = pass2.items.map((i) => i.Id);
    const pass3Ids = pass3.items.map((i) => i.Id);
    return pass1Names.includes('Status') && !pass1Names.includes('ParentRequest')
      && pass1.items.length === 3
      && pass2Ids.length === 1 && pass2Ids[0] === 2
      && pass3Ids.includes(3) && !pass3Ids.includes(1);
  }));

await check('pure: sp-rest getAll follows odata.nextLink past the default 5000 cap only when allowLargeCap is set', () =>
  page.evaluate(async () => {
    const { createSpRestClient } = await import('/src/workbench/sp-rest.js');
    const PAGE = 3000;
    const pageRows = (n) => Array.from({ length: PAGE }, (_, i) => ({ Id: n * PAGE + i + 1 }));
    let callN = 0;
    const fetchImpl = async () => {
      const n = callN++;
      const body = n < 2
        ? { value: pageRows(n), 'odata.nextLink': `https://t/_api/next${n + 1}` }
        : { value: pageRows(n) };
      return {
        ok: true, status: 200, headers: { get: () => null },
        json: async () => body, clone() { return this; }, text: async () => '',
      };
    };
    const client = createSpRestClient({
      getContext: () => ({ live: true, pageContext: { webAbsoluteUrl: 'https://t/sites/x' } }),
      fetchImpl,
    });
    const capped = await client.getAll("web/lists(guid'L1')/items", {});
    callN = 0;
    const large = await client.getAll("web/lists(guid'L1')/items", {}, { cap: 100000, allowLargeCap: true });
    return capped.items.length === 5000 && capped.partial === true
      && large.items.length === 9000 && large.partial === false;
  }));

await check('live: captureListData reads items with an unprojected $select=* (never a bare lookup/user internal name) and asks getAll for the opt-in 100000 cap', () =>
  page.evaluate(async () => {
    const { captureListData } = await import('/src/workbench/list-data-capture.js');
    const { SCHEMA_KIND, DATA_KIND } = await import('/src/workbench/list-schema.js');
    const schemaDoc = {
      kind: SCHEMA_KIND, version: 2,
      source: { siteUrl: 'https://t/sites/x', listTitle: 'Big', listId: 'L1', rootFolder: '/sites/x/Lists/Big', itemCount: 6000 },
      list: { title: 'Big', baseTemplate: 100 },
      fields: [
        { internalName: 'Title', type: 'Text', custom: false, readOnly: false },
        { internalName: 'Owner', type: 'User', custom: true, readOnly: false },
      ],
      views: [], contentTypes: [], warnings: [],
    };
    const calls = [];
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async (path, opts) => { calls.push({ kind: 'get', path, opts }); return { Id: 'w1' }; },
      getAll: async (path, opts, capOpts) => {
        calls.push({ kind: 'getAll', path, opts, capOpts });
        if (path.includes('/items')) return { items: [{ Id: 1, ID: 1, Title: 'One' }, { Id: 2, ID: 2, Title: 'Two' }], partial: false };
        return { items: [] };
      },
    };
    const { doc } = await captureListData(client, 'L1', { schemaDoc });
    const itemsCall = calls.find((c) => c.kind === 'getAll' && c.path.includes('/items'));
    const select = (itemsCall?.opts?.select || []).join(',');
    return doc.kind === DATA_KIND
      && select.includes('*') && !select.includes('Owner')
      && itemsCall?.capOpts?.allowLargeCap === true && itemsCall?.capOpts?.cap === 100000;
  }));

// ---- Stage 1b-a: mock UI — Items tab whole-list data export ---------------

await check('items: Download data .json exports the whole list via captureListData, not just the grid’s visible rows', async () => {
  await schemaPage.locator('.wb-back').click();
  await schemaPage.waitForSelector('.wb-table tbody tr', { hasText: 'Requests' });
  await schemaPage.locator('.wb-table tbody tr', { hasText: 'Requests' }).locator('td').first().click();
  await schemaPage.waitForSelector('.wb-tab');
  await schemaPage.locator('.wb-tab', { hasText: 'Items' }).click();
  await schemaPage.waitForSelector('.wb-items-grid .wb-table tbody tr');
  await schemaPage.locator('.wb-items-grid .wb-grid-actions .wb-menu-wrap button', { hasText: 'Export' }).click();
  const [download] = await Promise.all([
    schemaPage.waitForEvent('download'),
    schemaPage.locator('.wb-menu-item', { hasText: 'Download data .json' }).click(),
  ]);
  const bytes = readFileSync(await download.path());
  const doc = JSON.parse(bytes.toString('utf8'));
  return download.suggestedFilename() === 'data-requests.json'
    && doc.kind === 'dcspad-sputils-list-data' && doc.version === 2
    && doc.items.length === 3 && doc.folders.length === 1
    && doc.items.some((i) => i.Title === 'Server upgrade — phase 2' && i._resolved?.ParentRequest?.[0]?.value === 'Server upgrade')
    && doc.items.some((i) => i._attachments?.[0]?.name === 'quote.pdf')
    && doc.folders[0]._folderPath === 'Archive';
});

// ---- Stage 1b-b: list item DATA apply (pure list-data-apply.js) -----------

await check('pure: applyListData executor — folders before items in depth order, a per-field rejection retries once without that field, self-lookups written in pass 2 with new ids, authorship applied only with bNewDocumentUpdate:true', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const { SpFileError } = await import('/src/sp-odata.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: {
        Title: { type: 'Text', custom: false },
        Notes: { type: 'Text', custom: true },
        ParentRequest: { type: 'Lookup', custom: true, isSelfLookup: true, lookupListId: 'L1' },
      },
      items: [
        { Id: 1, Title: 'Parent', Created: '2026-01-01T00:00:00Z', Modified: '2026-01-01T00:00:00Z', _resolved: {} },
        { Id: 2, Title: 'Child', Notes: 'child notes', _resolved: { ParentRequest: [{ Id: 1, value: 'Parent' }] } },
        { Id: 3, Title: 'RejectMe', Notes: 'bad notes', _resolved: {} },
      ],
      folders: [
        { Id: 10, Title: 'Sub', _folder: true, _folderPath: 'Sub', _resolved: {} },
        { Id: 11, Title: 'SubSub', _folder: true, _folderPath: 'Sub/SubSub', _resolved: {} },
      ],
      users: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async (path) => {
        if (path.includes("lists(guid'L1')") && !path.includes('/fields')) {
          return { BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } };
        }
        return {};
      },
      getAll: async (path) => {
        if (path.includes("lists(guid'L1')/fields")) {
          return { items: [
            { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
            { InternalName: 'Notes', TypeAsString: 'Text', ReadOnlyField: false },
            { InternalName: 'ParentRequest', TypeAsString: 'Lookup', ReadOnlyField: false, LookupList: 'L1' },
          ] };
        }
        return { items: [] };
      },
    };
    const posts = [];
    const rejectedOnce = new Set();
    let idSeq = 1000;
    const spWrite = {
      isMock: () => true,
      addValidateUpdateItem: async (listId, spec) => {
        posts.push({ ...spec });
        const titleVal = spec.formValues.find((v) => v.FieldName === 'Title')?.FieldValue;
        if (titleVal === 'RejectMe' && !rejectedOnce.has('RejectMe')) {
          rejectedOnce.add('RejectMe');
          const err = new SpFileError('rejected', { code: 'metadata-write' });
          err.fieldErrors = { Notes: 'bad value' };
          throw err;
        }
        return { id: idSeq++ };
      },
      validateUpdateListItem: async (pathKind, formValues, opts) => {
        posts.push({ kind: 'update', pathKind, formValues, opts: opts || {} });
        return { updated: formValues.map((f) => f.FieldName) };
      },
      ensureUser: async () => null,
    };
    const report = await applyListData({
      dataDoc, client, spWrite, listId: 'L1', options: { preserveAuthorship: true },
    });
    const creates = posts.filter((p) => !p.kind);
    const updates = posts.filter((p) => p.kind === 'update');
    const folderCreates = creates.filter((c) => c.underlyingObjectType === 1);
    const itemCreates = creates.filter((c) => c.underlyingObjectType === 0);
    const foldersFirst = creates.indexOf(folderCreates[0]) === 0 && creates.indexOf(folderCreates[1]) === 1
      && creates.indexOf(itemCreates[0]) >= 2;
    const rejectAttempts = itemCreates.filter((c) => c.formValues.find((v) => v.FieldName === 'Title')?.FieldValue === 'RejectMe');
    const retryDroppedField = rejectAttempts.length === 2
      && rejectAttempts[0].formValues.some((v) => v.FieldName === 'Notes')
      && !rejectAttempts[1].formValues.some((v) => v.FieldName === 'Notes');
    const parentNewId = report.idMap[1];
    const selfLookupUpdate = updates.find((u) => u.formValues.some((v) => v.FieldName === 'ParentRequest'));
    const selfLookupOk = selfLookupUpdate
      && selfLookupUpdate.formValues.find((v) => v.FieldName === 'ParentRequest').FieldValue === String(parentNewId);
    const authorshipUpdate = updates.find((u) => u.formValues.some((v) => v.FieldName === 'Created'));
    const authorshipOk = authorshipUpdate && authorshipUpdate.opts.newDocumentUpdate === true;
    return foldersFirst && retryDroppedField && selfLookupOk && authorshipOk
      && report.items.added === 3 && report.items.failed.length === 0
      && report.folders.created === 2
      && report.fieldErrors.some((e) => e.sourceId === 3 && e.field === 'Notes')
      && report.authorship.applied >= 1;
  }));

await check('pure: applyListData refuses a document-library target (library-items) and aborts on a 401', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const { SpFileError } = await import('/src/sp-odata.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Docs', listId: 'SRC2' },
      fields: { Title: { type: 'Text', custom: false } },
      items: [{ Id: 1, Title: 'A', _resolved: {} }],
      folders: [], users: [], warnings: [],
    };
    const libraryClient = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({ BaseType: 1, Title: 'Docs', RootFolder: { ServerRelativeUrl: '/sites/x/Documents' } }),
      getAll: async () => ({ items: [] }),
    };
    let refused = false;
    try {
      await applyListData({ dataDoc, client: libraryClient, spWrite: { isMock: () => true }, listId: 'LIB1' });
    } catch (e) { refused = e?.code === 'library-items'; }

    const authClient = {
      webUrl: () => 'https://t/sites/x',
      get: async () => { throw new SpFileError('expired', { code: 'auth', status: 401 }); },
      getAll: async () => ({ items: [] }),
    };
    const authReport = await applyListData({ dataDoc, client: authClient, spWrite: { isMock: () => true }, listId: 'L1' });
    return refused && authReport.aborted === 'auth';
  }));

// ---- Mock UI: item copy through the apply dialog (Requests → target) ------

await check('dialog: Copy to… with Include items copies items to a NEW target list and the report counts them', async () => {
  await schemaPage.locator('.wb-back').click();
  await schemaPage.waitForSelector('.wb-table tbody tr', { hasText: 'Requests' });
  await schemaPage.locator('.wb-table tbody tr', { hasText: 'Requests' }).locator('td').first().click();
  await schemaPage.locator('.wb-tab', { hasText: 'Schema' }).click();
  await schemaPage.waitForSelector('.wb-schema-copy');
  await schemaPage.locator('.wb-schema-copy').click();
  await schemaPage.waitForSelector('.wb-schema-dialog');
  await schemaPage.fill('.wb-schema-target', '/sites/target');
  await schemaPage.locator('.wb-schema-connect').click();
  await schemaPage.waitForFunction(() => document.querySelector('.wb-schema-title')?.value?.length > 0);
  await schemaPage.fill('.wb-schema-title', 'Requests Import Test');
  await schemaPage.locator('.wb-schema-items input[type="checkbox"]').first().check();
  await schemaPage.locator('.wb-schema-create').click();
  await schemaPage.waitForSelector('.wb-schema-report:not([hidden])');
  const headline = await schemaPage.locator('.wb-schema-report-headline').textContent();
  const itemsCountText = await schemaPage.locator('.wb-schema-report-counts').textContent();
  const wroteAddItem = await schemaPage.evaluate(() =>
    (window.__DCSPAD_WB_WRITES__ || []).some((w) => String(w.url).toLowerCase().includes('addvalidateupdateitemusingpath')));
  await schemaPage.locator('.wb-schema-close').click();
  await schemaPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });
  return !headline.includes('expired') && itemsCountText.includes('Items') && wroteAddItem;
});

await check('dialog: New from schema… enables Include items when a matching data file is also chosen, and refuses a mismatched one inline', async () => {
  await schemaPage.locator('.wb-back').click();
  await schemaPage.waitForSelector('.wb-schema-new');
  const schemaDoc = {
    kind: 'dcspad-sputils-list-schema', version: 2,
    source: { listTitle: 'Clients', listId: 'clients-src-id' },
    list: { title: 'Clients', baseTemplate: 100 },
    fields: [{ internalName: 'Title', displayName: 'Title', type: 'Text', custom: false, fromBaseType: true }],
    views: [{ title: 'All Items', fields: ['Title'] }],
    contentTypes: [], warnings: [],
  };
  const mismatchedDataDoc = {
    kind: 'dcspad-sputils-list-data', version: 2,
    source: { listTitle: 'Other list', listId: 'not-clients' },
    fields: { Title: { type: 'Text', custom: false } },
    items: [], folders: [], users: [], warnings: [],
  };
  await schemaPage.setInputFiles('.wb-schema-file', [
    { name: 'clients-schema.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(schemaDoc)) },
    { name: 'other-data.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(mismatchedDataDoc)) },
  ]);
  await schemaPage.waitForSelector('.wb-schema-import-notice:not([hidden])');
  const mismatchText = await schemaPage.locator('.wb-schema-import-notice').textContent();
  const mismatchRefused = mismatchText.includes('match');

  await schemaPage.locator('.wb-schema-import-notice button', { hasText: 'Dismiss' }).click();
  const matchingDataDoc = {
    kind: 'dcspad-sputils-list-data', version: 2,
    source: { listTitle: 'Clients', listId: 'clients-src-id' },
    fields: { Title: { type: 'Text', custom: false } },
    items: [], folders: [], users: [], warnings: [],
  };
  await schemaPage.setInputFiles('.wb-schema-file', [
    { name: 'clients-schema2.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(schemaDoc)) },
    { name: 'clients-data.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(matchingDataDoc)) },
  ]);
  await schemaPage.waitForSelector('.wb-schema-dialog');
  const itemsEnabled = await schemaPage.locator('.wb-schema-items').evaluate((f) => f.disabled === false);
  await schemaPage.locator('.wb-schema-close').click();
  await schemaPage.waitForSelector('.wb-schema-dialog', { state: 'detached' });
  return mismatchRefused && itemsEnabled;
});

// ---- Stubbed live: AddValidateUpdateItemUsingPath body shape + ensureuser -

// The request BODY sp-write.js's addValidateUpdateItem/ensureUser build is
// identical whether post() then sends it live or (as here) hands it to the
// mock writer — post() only branches AFTER the body is JSON.stringify'd — so
// recording it through __DCSPAD_WB_WRITES__ pins the exact same shape a live
// POST would carry, without needing to also stand up a digest fetch.
await check('mock: AddValidateUpdateItemUsingPath body carries FolderPath.DecodedUrl/UnderlyingObjectType/string formValues, and ensureuser posts once per login', () =>
  page.evaluate(async () => {
    const { createSpWriteClient } = await import('/src/workbench/sp-write.js');
    globalThis.__DCSPAD_WB_WRITES__ = [];
    const client = { webUrl: () => 'https://t/sites/x', context: () => ({ live: false }) };
    const spWrite = createSpWriteClient({ client });
    const { id } = await spWrite.addValidateUpdateItem('L1', {
      folderPath: '/sites/x/Lists/Requests', underlyingObjectType: 0,
      formValues: [{ FieldName: 'Title', FieldValue: 'Hello' }, { FieldName: 'Budget', FieldValue: '10' }],
    });
    const writes = globalThis.__DCSPAD_WB_WRITES__;
    const createWrite = writes.find((w) => String(w.url).toLowerCase().includes('addvalidateupdateitemusingpath'));
    const createBody = JSON.parse(createWrite.body);
    const shapeOk = createBody.listItemCreateInfo.FolderPath.DecodedUrl === '/sites/x/Lists/Requests'
      && createBody.listItemCreateInfo.UnderlyingObjectType === 0
      && createBody.formValues.every((v) => typeof v.FieldValue === 'string')
      && typeof id === 'number';
    await spWrite.ensureUser('pat@t.local');
    await spWrite.ensureUser('pat@t.local');
    const ensureCount = writes.filter((w) => String(w.url).toLowerCase().includes('/ensureuser')).length;
    return shapeOk && ensureCount === 1;
  }));

// ---- Stage 1b review + live-tenant findings fixes -------------------------

await check('pure: hasBlockingFieldFailures holds items on a retryable/blocked field.create, clears on final refusals and skips', () =>
  page.evaluate(async () => {
    const { hasBlockingFieldFailures } = await import('/src/workbench/list-schema-dialog.js');
    const retryable = { steps: [
      { kind: 'field.create', status: 'failed', final: false },
      { kind: 'view.upsert', status: 'done' },
    ] };
    const blocked = { steps: [{ kind: 'field.create', status: 'blocked' }] };
    const finalRefusal = { steps: [{ kind: 'field.create', status: 'failed', final: true }] };
    const skippedPolicy = { steps: [{ kind: 'field.create', status: 'skipped' }] };
    const clean = { steps: [{ kind: 'field.create', status: 'done' }] };
    return hasBlockingFieldFailures(retryable) === true
      && hasBlockingFieldFailures(blocked) === true
      && hasBlockingFieldFailures(finalRefusal) === false
      && hasBlockingFieldFailures(skippedPolicy) === false
      && hasBlockingFieldFailures(clean) === false;
  }));

await check('pure: a v1 doc with only _resolved people (no top-level users) still ensures users and writes the User field + authorship', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 1,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: {
        Title: { type: 'Text', custom: false },
        Owner: { type: 'User', custom: true },
      },
      items: [{
        Id: 1, Title: 'Has owner', Created: '2026-01-01T00:00:00Z', Modified: '2026-01-01T00:00:00Z',
        _resolved: {
          Owner: [{ Id: 11, Email: 'pat@mock.local', LoginName: 'i:0#.f|membership|pat@mock.local', Title: 'Pat' }],
          Author: [{ Id: 11, Email: 'pat@mock.local', LoginName: 'i:0#.f|membership|pat@mock.local', Title: 'Pat' }],
        },
      }],
      // No top-level `users` key — the SPUtils v1 shape.
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({ BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } }),
      getAll: async (path) => (path.includes('/fields')
        ? { items: [
          { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
          { InternalName: 'Owner', TypeAsString: 'User', ReadOnlyField: false },
        ] }
        : { items: [] }),
    };
    const ensureCalls = [];
    let idSeq = 2000;
    const posts = [];
    const spWrite = {
      isMock: () => true,
      addValidateUpdateItem: async (listId, spec) => { posts.push({ ...spec }); return { id: idSeq++ }; },
      validateUpdateListItem: async (pathKind, formValues, opts) => {
        posts.push({ kind: 'update', formValues, opts: opts || {} });
        return {};
      },
      ensureUser: async (logon) => { ensureCalls.push(logon); return { loginName: `i:0#.f|membership|${logon.toLowerCase()}` }; },
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1', options: { preserveAuthorship: true } });
    const create = posts.find((p) => !p.kind);
    const ownerValue = create?.formValues.find((v) => v.FieldName === 'Owner')?.FieldValue;
    const authorshipUpdate = posts.find((p) => p.kind === 'update' && p.formValues.some((v) => v.FieldName === 'Author'));
    return ensureCalls.length === 1 && ensureCalls[0] === 'pat@mock.local'
      && ownerValue === JSON.stringify([{ Key: 'i:0#.f|membership|pat@mock.local' }])
      && Boolean(authorshipUpdate) && report.items.added === 1;
  }));

await check('live: captureListData embeds a small attachment as base64 and degrades an over-cap one to a url-only link with a warning', () =>
  page.evaluate(async () => {
    const { captureListData } = await import('/src/workbench/list-data-capture.js');
    const { SCHEMA_KIND } = await import('/src/workbench/list-schema.js');
    const schemaDoc = {
      kind: SCHEMA_KIND, version: 2,
      source: { siteUrl: 'https://t/sites/x', listTitle: 'Requests', listId: 'L1', rootFolder: '/sites/x/Lists/Requests', itemCount: 1 },
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'Title', type: 'Text', custom: false, readOnly: false },
        { internalName: 'Attachments', type: 'Attachments', custom: false, readOnly: false, fromBaseType: true },
      ],
      views: [], contentTypes: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({}),
      getAll: async (path) => {
        if (path.includes('/items')) {
          return {
            items: [{
              Id: 1, ID: 1, Title: 'One', FSObjType: 0,
              FileDirRef: '/sites/x/Lists/Requests', FileRef: '/sites/x/Lists/Requests/1_.000',
              AttachmentFiles: [
                { FileName: 'small.txt', ServerRelativeUrl: '/attach/small.txt' },
                { FileName: 'huge.bin', ServerRelativeUrl: '/attach/huge.bin' },
              ],
            }],
            partial: false,
          };
        }
        return { items: [] };
      },
    };
    const originalFetch = window.fetch;
    window.fetch = async (url) => {
      const s = String(url);
      if (s.includes('small.txt')) return { ok: true, status: 200, arrayBuffer: async () => new TextEncoder().encode('hello').buffer };
      if (s.includes('huge.bin')) return { ok: true, status: 200, arrayBuffer: async () => new ArrayBuffer(10 * 1024 * 1024 + 1) };
      return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
    };
    let doc;
    try {
      ({ doc } = await captureListData(client, 'L1', { schemaDoc }));
    } finally {
      window.fetch = originalFetch;
    }
    const item = doc.items[0];
    const small = item._attachments.find((a) => a.name === 'small.txt');
    const huge = item._attachments.find((a) => a.name === 'huge.bin');
    return Boolean(small) && typeof small.base64 === 'string' && !small.url
      && Boolean(huge) && !huge.base64 && huge.url === '/attach/huge.bin'
      && doc.warnings.some((w) => w.includes('huge.bin') && w.includes('cap'));
  }));

await check('pure: applyListData writes a base64 attachment via AttachmentFiles/add in import mode (no source client needed)', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: { Title: { type: 'Text', custom: false } },
      items: [{ Id: 1, Title: 'Has file', _resolved: {}, _attachments: [{ name: 'a.txt', base64: btoa('hello') }] }],
      folders: [], users: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({ BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } }),
      getAll: async (path) => (path.includes('/fields')
        ? { items: [{ InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false }] }
        : { items: [] }),
    };
    const attachCalls = [];
    let idSeq = 3000;
    const spWrite = {
      isMock: () => true,
      addValidateUpdateItem: async () => ({ id: idSeq++ }),
      validateUpdateListItem: async () => ({}),
      ensureUser: async () => null,
      addAttachment: async (listId, itemId, fileName, bytes) => {
        attachCalls.push({ listId, itemId, fileName, text: new TextDecoder().decode(bytes) });
        return { fileName };
      },
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1', options: { includeAttachments: true } });
    return attachCalls.length === 1 && attachCalls[0].fileName === 'a.txt' && attachCalls[0].text === 'hello'
      && report.attachments.added === 1;
  }));

await check('pure: a target date-format calibration failure drops DateTime fields with one warning, never guesses mdy/UTC', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: {
        Title: { type: 'Text', custom: false },
        Due: { type: 'DateTime', custom: true },
      },
      items: [{ Id: 1, Title: 'Item', Due: '2026-03-04T05:06:00Z', _resolved: {} }],
      folders: [], users: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async (path) => {
        if (path.startsWith('web/RegionalSettings') && !path.includes('utctolocaltime')) {
          throw new Error('regional settings unreadable');
        }
        return { BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } };
      },
      getAll: async (path) => (path.includes('/fields')
        ? { items: [
          { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
          { InternalName: 'Due', TypeAsString: 'DateTime', ReadOnlyField: false },
        ] }
        : { items: [] }),
    };
    let idSeq = 4000;
    const posts = [];
    const spWrite = {
      isMock: () => false,
      addValidateUpdateItem: async (listId, spec) => { posts.push({ ...spec }); return { id: idSeq++ }; },
      validateUpdateListItem: async () => ({}),
      ensureUser: async () => null,
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1' });
    const create = posts.find((p) => !p.kind);
    const hasDue = create?.formValues.some((v) => v.FieldName === 'Due');
    return !hasDue
      && report.warnings.some((w) => w.includes('Due') && w.includes('could not be learned'))
      && report.items.added === 1;
  }));

await check('pure: captureListData resolves a lookup id beyond the first page via a direct fetch, and warns when the target index is partial', () =>
  page.evaluate(async () => {
    const { captureListData } = await import('/src/workbench/list-data-capture.js');
    const { SCHEMA_KIND } = await import('/src/workbench/list-schema.js');
    const schemaDoc = {
      kind: SCHEMA_KIND, version: 2,
      source: { siteUrl: 'https://t/sites/x', listTitle: 'Requests', listId: 'L1', rootFolder: '/sites/x/Lists/Requests', itemCount: 1 },
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'Title', type: 'Text', custom: false, readOnly: false },
        { internalName: 'Client', type: 'Lookup', custom: true, readOnly: false, lookupListId: 'CL1', lookupField: 'Title' },
      ],
      views: [], contentTypes: [], warnings: [],
    };
    const directGetCalls = [];
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async (path) => {
        directGetCalls.push(path);
        if (path.includes('items(9)')) return { Id: 9, Title: 'FarAway' };
        return null;
      },
      getAll: async (path) => {
        if (path.includes("lists(guid'L1')/items")) {
          return {
            items: [{
              Id: 1, ID: 1, Title: 'One', ClientId: 9, FSObjType: 0,
              FileDirRef: '/sites/x/Lists/Requests', FileRef: '/sites/x/Lists/Requests/1_.000',
            }],
            partial: false,
          };
        }
        if (path.includes("lists(guid'CL1')/items")) {
          return { items: [{ Id: 1, Title: 'Near' }, { Id: 2, Title: 'Also near' }], partial: true };
        }
        return { items: [] };
      },
    };
    const { doc } = await captureListData(client, 'L1', { schemaDoc });
    const resolved = doc.items[0]._resolved.Client;
    return resolved?.[0]?.Id === 9 && resolved[0].value === 'FarAway'
      && directGetCalls.some((p) => p.includes('items(9)'))
      && doc.warnings.some((w) => w.includes('Client') && w.includes('indexed'));
  }));

await check('pure: captureListData turns a 403 on a referenced lookup list into a warning, not a throw', () =>
  page.evaluate(async () => {
    const { captureListData } = await import('/src/workbench/list-data-capture.js');
    const { SCHEMA_KIND } = await import('/src/workbench/list-schema.js');
    const { SpFileError } = await import('/src/sp-odata.js');
    const schemaDoc = {
      kind: SCHEMA_KIND, version: 2,
      source: { siteUrl: 'https://t/sites/x', listTitle: 'Requests', listId: 'L1', rootFolder: '/sites/x/Lists/Requests', itemCount: 1 },
      list: { title: 'Requests', baseTemplate: 100 },
      fields: [
        { internalName: 'Title', type: 'Text', custom: false, readOnly: false },
        { internalName: 'Client', type: 'Lookup', custom: true, readOnly: false, lookupListId: 'CL1', lookupField: 'Title' },
      ],
      views: [], contentTypes: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => null,
      getAll: async (path) => {
        if (path.includes("lists(guid'L1')/items")) {
          return {
            items: [{
              Id: 1, ID: 1, Title: 'One', ClientId: 5, FSObjType: 0,
              FileDirRef: '/sites/x/Lists/Requests', FileRef: '/sites/x/Lists/Requests/1_.000',
            }],
            partial: false,
          };
        }
        if (path.includes("lists(guid'CL1')/items")) throw new SpFileError('denied', { code: 'permission', status: 403 });
        return { items: [] };
      },
    };
    let threw = false;
    let doc;
    try {
      ({ doc } = await captureListData(client, 'L1', { schemaDoc }));
    } catch { threw = true; }
    return !threw
      && doc.warnings.some((w) => w.includes('Client') && w.includes('could not be read'))
      && doc.items[0]._resolved.Client?.[0]?.value === null;
  }));

await check('pure: buildApplyReport lists items field errors and warnings (dropped-field items) in the downloaded report .md', () =>
  page.evaluate(async () => {
    const { buildApplyReport, newReport } = await import('/src/workbench/list-schema.js');
    const plan = { title: 'Requests', targetWebUrl: 'https://t/sites/target', existingListId: null, steps: [], warnings: [] };
    const report = newReport(plan);
    report.created = true;
    report.itemsReport = {
      items: { added: 2, failed: [] },
      folders: { created: 0, failed: 0 },
      fieldErrors: [{ sourceId: 2, field: 'Status', message: 'SharePoint rejected the value.' }],
      warnings: ['Item 2 → 1002: created without Status (see field errors).'],
    };
    const md = buildApplyReport({ report, doc: { list: { title: 'Requests' } }, targetWebUrl: 'https://t/sites/target' });
    return md.includes('## Items') && md.includes('Added: 2')
      && md.includes('Source id 2 — Status')
      && md.includes('created without Status');
  }));

await check('pure: applyListData resolves a cross-list lookup by shown value to the TARGET id (braced LookupList GUID handled) and never writes a dependent lookup', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: {
        Title: { type: 'Text', custom: false },
        ReqClient: { type: 'Lookup', custom: true, lookupListId: 'schema-clients-id', lookupField: 'Title' },
        ReqClientCode: { type: 'Lookup', custom: true, isDependentLookup: true, lookupListId: 'schema-clients-id', lookupField: 'ClientCode' },
      },
      items: [{
        Id: 1, Title: 'First request', _resolved: {
          ReqClient: [{ Id: 5, value: 'Contoso' }],
          ReqClientCode: [{ Id: 5, value: 'CO-1' }],
        },
      }],
      folders: [], users: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({ BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } }),
      getAll: async (path) => {
        if (path.includes("lists(guid'L1')/fields")) {
          return {
            items: [
              { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
              // A tenant that hands LookupList back braced — cleanGuid() must
              // strip it before building the guid'...' OData path (finding #8).
              { InternalName: 'ReqClient', TypeAsString: 'Lookup', ReadOnlyField: false, LookupList: '{target-clients-id}', LookupField: 'Title' },
              // The dependent lookup is read-only on the target — never written.
              { InternalName: 'ReqClientCode', TypeAsString: 'Lookup', ReadOnlyField: true, LookupList: '{target-clients-id}', LookupField: 'ClientCode' },
            ],
          };
        }
        if (path.includes("lists(guid'target-clients-id')/items")) {
          return { items: [{ Id: 42, Title: 'Contoso' }], partial: false };
        }
        return { items: [] };
      },
    };
    let idSeq = 5000;
    const posts = [];
    const spWrite = {
      isMock: () => true,
      addValidateUpdateItem: async (listId, spec) => { posts.push({ ...spec }); return { id: idSeq++ }; },
      validateUpdateListItem: async () => ({}),
      ensureUser: async () => null,
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1' });
    const create = posts.find((p) => !p.kind);
    const clientValue = create?.formValues.find((v) => v.FieldName === 'ReqClient')?.FieldValue;
    const hasDependent = create?.formValues.some((v) => v.FieldName === 'ReqClientCode');
    return clientValue === '42' && !hasDependent && report.items.added === 1 && report.fieldErrors.length === 0;
  }));

await check('pure: applyListData writes pass-1 item creates sequentially, in ascending SOURCE id order (never scrambled by concurrency)', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: { Title: { type: 'Text', custom: false } },
      items: [
        { Id: 2, Title: 'Second', _resolved: {} },
        { Id: 3, Title: 'Third', _resolved: {} },
        { Id: 1, Title: 'First', _resolved: {} },
      ],
      folders: [], users: [], warnings: [],
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async () => ({ BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } }),
      getAll: async (path) => (path.includes('/fields')
        ? { items: [{ InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false }] }
        : { items: [] }),
    };
    const order = [];
    let idSeq = 6000;
    const spWrite = {
      isMock: () => true,
      addValidateUpdateItem: async (listId, spec) => {
        // Source item "Second" is queued first but deliberately takes the
        // longest — a concurrent pool would finish "Third" before it and
        // hand out ids out of source order; sequential execution must not.
        const title = spec.formValues.find((v) => v.FieldName === 'Title')?.FieldValue;
        const delay = title === 'Second' ? 15 : title === 'Third' ? 5 : 0;
        await new Promise((r) => setTimeout(r, delay));
        order.push(title);
        return { id: idSeq++ };
      },
      validateUpdateListItem: async () => ({}),
      ensureUser: async () => null,
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1' });
    return order.join(',') === 'First,Second,Third' && report.items.added === 3;
  }));

await check('pure: applyListData computes the web-local offset per VALUE, not once per item — two dates straddling a DST change get different offsets', () =>
  page.evaluate(async () => {
    const { applyListData } = await import('/src/workbench/list-data-apply.js');
    const dataDoc = {
      kind: 'dcspad-sputils-list-data', version: 2,
      source: { listTitle: 'Requests', listId: 'SRC1' },
      fields: {
        Title: { type: 'Text', custom: false },
        DueOn: { type: 'DateTime', custom: true },
      },
      items: [{
        Id: 1, Title: 'Item',
        // Created is BEFORE a DST change (offset -300 min), DueOn is AFTER
        // it (offset -240 min) — one item, two dates, two different offsets.
        Created: '2026-03-01T12:00:00Z', DueOn: '2026-03-10T12:00:00Z',
        _resolved: {},
      }],
      folders: [], users: [], warnings: [],
    };
    const offsetByDay = {
      '2026-02-28': -300, '2026-03-01': -300, '2026-03-02': -300,
      '2026-03-09': -300, '2026-03-10': -240, '2026-03-11': -240,
    };
    const client = {
      webUrl: () => 'https://t/sites/x',
      get: async (path) => {
        // Case-sensitive: "utcToLocalTime" must be checked before the
        // startsWith('web/RegionalSettings') fallback below matches too.
        const m = /utcToLocalTime\(@d\)\?@d='([^']+)'/.exec(path);
        if (m) {
          const day = m[1].slice(0, 10);
          const offsetMin = offsetByDay[day] ?? -300;
          const local = new Date(new Date(m[1]).getTime() + offsetMin * 60000).toISOString().replace('Z', '');
          return { value: local };
        }
        if (path.startsWith('web/RegionalSettings')) {
          return { DateFormat: 0, DateSeparator: '/', TimeSeparator: ':' };
        }
        return { BaseType: 0, Title: 'Target', RootFolder: { ServerRelativeUrl: '/sites/x/Lists/Target' } };
      },
      getAll: async (path) => (path.includes('/fields')
        ? { items: [
          { InternalName: 'Title', TypeAsString: 'Text', ReadOnlyField: false },
          { InternalName: 'DueOn', TypeAsString: 'DateTime', ReadOnlyField: false },
        ] }
        : { items: [] }),
    };
    let idSeq = 7000;
    const posts = [];
    const spWrite = {
      isMock: () => false,
      addValidateUpdateItem: async (listId, spec) => {
        if (spec.formValues.some((v) => v.FieldValue === 'not-a-date')) {
          const err = new Error('rejected'); err.fieldErrors = {}; throw err;
        }
        posts.push({ ...spec });
        return { id: idSeq++ };
      },
      validateUpdateListItem: async (pathKind, formValues) => { posts.push({ kind: 'update', formValues }); return {}; },
      ensureUser: async () => null,
    };
    const report = await applyListData({ dataDoc, client, spWrite, listId: 'L1', options: { preserveAuthorship: true } });
    const create = posts.find((p) => !p.kind);
    const dueValue = create?.formValues.find((v) => v.FieldName === 'DueOn')?.FieldValue;
    const authUpdate = posts.find((p) => p.kind === 'update' && p.formValues.some((v) => v.FieldName === 'Created'));
    const createdValue = authUpdate?.formValues.find((v) => v.FieldName === 'Created')?.FieldValue;
    // -300 min = -5h: 12:00 UTC -> 07:00 local. -240 min = -4h: 12:00 -> 08:00 local.
    return Boolean(dueValue) && Boolean(createdValue) && dueValue.includes('08:00') && createdValue.includes('07:00')
      && dueValue !== createdValue && report.items.added === 1;
  }));

await page.close();
await schemaPage.close();
await live.close();
await browser.close();
exitWithResult();
