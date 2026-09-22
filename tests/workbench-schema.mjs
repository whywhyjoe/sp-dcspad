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
  return kind === 'Generic list' && fields === '20 fields · 15 custom' && views === '2 views'
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

await check('schema: a document library shows the library chip and gates Copy to… behind stage 2', async () => {
  await schemaPage.locator('.wb-back').click();
  await schemaPage.waitForSelector('.wb-table tbody tr', { hasText: 'Documents' });
  await schemaPage.locator('.wb-table tbody tr', { hasText: 'Documents' }).locator('td').first().click();
  await schemaPage.waitForSelector('.wb-tab');
  await schemaPage.locator('.wb-tab', { hasText: 'Schema' }).click();
  await schemaPage.waitForSelector('.wb-schema-chips .wb-schema-kind');
  const libChip = await schemaPage.locator('.wb-schema-libkind').count();
  const disabled = await schemaPage.locator('.wb-schema-copy').isDisabled();
  const title = await schemaPage.locator('.wb-schema-copy').getAttribute('title');
  return libChip === 1 && disabled && title.includes('stage 2');
});

// ---- Stubbed live: capture request shapes ----------------------------------

const LIVE_LIST_ID = '33333333-0000-4000-8000-000000000001';
const LIVE_LOOKUP_ID = '33333333-0000-4000-8000-000000000002';
const LIVE_DENIED_ID = '33333333-0000-4000-8000-000000000003';
const liveReads = [];   // { url, accept }

const live = await browser.newPage({ viewport: { width: 1400, height: 900 } });
await live.addInitScript(() => {
  window.__DCSPAD_SP_CONTEXT__ = { webAbsoluteUrl: location.origin, userDisplayName: 'Stub User' };
});
await live.route('**/_api/**', async (route) => {
  const request = route.request();
  const url = request.url();
  liveReads.push({ url, accept: request.headers().accept || '' });

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

await page.close();
await schemaPage.close();
await live.close();
await browser.close();
exitWithResult();
