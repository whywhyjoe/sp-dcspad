// List schema — the "Copy to…" / "New from schema…" apply dialog.
//
// Mirrors upload-metadata.js's recipe (`<dialog class="app-dialog
// sp-metadata-dialog …">`, two-phase exit, retry) and browser.js's
// `.wb-consent` inline notices, so this reads like the rest of the
// Workbench instead of inventing new vocabulary. The dialog itself does no
// planning or writing: Dry run calls probeTarget()+buildApplyPlan() (reads
// only) to render the same Plan the Create button then hands to runPlan().
// Retry re-plans only the failed steps (retryPlan) and folds their results
// back into the report already on screen.
//
// The shell client (`client`) and the view that opened this dialog are
// never touched here — a second, independent client (`createClient()`) is
// what talks to the target web, so inspecting a different site to copy into
// never flips "inspecting" onto the source view's status line.

import {
  normalizeSchemaDoc, SCHEMA_KIND, defaultTargetTitle, buildApplyPlan, retryPlan,
  buildApplyReport, parentContentTypeId, isBuiltinParent, LOOKUP_TYPES,
} from './list-schema.js';
import { probeTarget } from './list-schema-capture.js';
import { runPlan } from './list-schema-apply.js';
import { createSpWriteClient } from './sp-write.js';
import { EXPIRED_SESSION_NOTE } from './denied.js';
import { getFavorites, getRecents } from './favorites.js';
import { LIST_SETTINGS, linkUrl } from './config-links.js';
import { bindNewTab } from './grid.js';
import { downloadText } from '../io.js?v=2';

export { SCHEMA_KIND };

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const cleanId = (v) => String(v || '').replace(/[{}]/g, '').toLowerCase();
const canonUrl = (u) => String(u || '').replace(/\/+$/, '').toLowerCase();
const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

const STATE_MAP = {
  planned: 'pending', running: 'running', done: 'ok',
  skipped: 'skip', failed: 'failed', blocked: 'blocked',
};

function verbFor(step) {
  switch (step.kind) {
    case 'list.create': return 'Creating the list';
    case 'list.adopt': return 'Opening the existing list';
    case 'list.settings': return 'Applying list settings';
    case 'ct.attach': return 'Attaching content type';
    case 'field.create': return 'Creating field';
    case 'field.merge': return 'Updating field';
    case 'field.base': return 'Updating the Title column';
    case 'view.upsert': return 'Creating view';
    case 'list.validation': return 'Applying the validation formula';
    default: return 'Working';
  }
}

function stepLine(step) {
  return step.error ? `${step.label} — ${step.error}` : step.label;
}

// Fold a retry report's (small, failed-steps-only) results back onto the
// displayed report: steps replaced by id, counters re-tallied from the
// merged step list so a second retry still reports honest totals.
function mergeRetryReport(base, retry) {
  const merged = { ...base };
  merged.steps = base.steps.map((s) => retry.steps.find((r) => r.id === s.id) || s);
  merged.fields = { added: 0, skipped: 0, failed: [] };
  merged.views = { added: 0, updated: 0, failed: [] };
  merged.contentTypes = { attached: 0, skipped: 0, failed: 0 };
  merged.validation = { applied: false, error: '' };
  for (const s of merged.steps) {
    if (s.kind === 'field.create') {
      if (s.status === 'done') merged.fields.added++;
      else if (s.status === 'skipped') merged.fields.skipped++;
      else if (s.status === 'failed') {
        merged.fields.failed.push({ internalName: s.payload?.field?.internalName, error: s.error });
      }
    } else if (s.kind === 'view.upsert') {
      if (s.status === 'done') { if (s.result?.created) merged.views.added++; else merged.views.updated++; }
      else if (s.status === 'failed') merged.views.failed.push({ title: s.payload?.title, error: s.error });
    } else if (s.kind === 'ct.attach') {
      if (s.status === 'done') merged.contentTypes.attached++;
      else if (s.status === 'failed') merged.contentTypes.failed++;
      else if (s.status === 'skipped') merged.contentTypes.skipped++;
    } else if (s.kind === 'list.validation') {
      merged.validation.applied = s.status === 'done';
      if (s.status === 'failed') merged.validation.error = s.error;
    }
  }
  merged.warnings = [...(base.warnings || []), ...(retry.warnings || [])];
  merged.aborted = retry.aborted || '';
  merged.listId = base.listId || retry.listId;
  merged.created = base.created || retry.created;
  merged.rootFolder = base.rootFolder || retry.rootFolder;
  return merged;
}

function buildHeadline(report, { isMock }) {
  if (report.aborted === 'auth') return EXPIRED_SESSION_NOTE;
  const failed = report.steps.filter((s) => s.status === 'failed').length;
  const retryable = report.steps.filter((s) => s.status === 'failed' && !s.final).length;
  let headline;
  if (report.adopted) {
    headline = `Updated ‘${report.title}’ — ${plural(report.fields.added, 'field')} added, `
      + `${report.fields.skipped} already present, ${plural(report.views.added + report.views.updated, 'view')} rebuilt.`;
  } else {
    const fieldsTotal = report.fields.added + report.fields.skipped + report.fields.failed.length;
    const ctPart = report.contentTypes.attached ? `, ${plural(report.contentTypes.attached, 'content type')}` : '';
    const dest = report.targetWebUrl || 'the target site';
    headline = `${report.created ? 'Created' : 'Could not create'} ‘${report.title}’ on ${dest} — `
      + `${report.fields.added} of ${fieldsTotal} fields, ${plural(report.views.added + report.views.updated, 'view')}${ctPart}.`;
  }
  if (failed) {
    headline += retryable
      ? ` ${plural(failed, 'step')} failed — see below, then Retry failed steps.`
      : ` ${plural(failed, 'step')} could not be applied — see below.`;
  }
  if (isMock) headline += ' (mock mode — the fixture web does not change).';
  return headline;
}

export function openSchemaApplyDialog({
  doc, mode = 'copy', client, createClient, navigate, inspectSite, mockWriter,
} = {}) {
  return new Promise((resolve) => {
    const d = normalizeSchemaDoc(doc);
    const targetClient = createClient();
    const isMockMode = !targetClient.context().live;
    const spWrite = createSpWriteClient({ client: targetClient, mockWriter: isMockMode ? mockWriter : undefined });

    let probe = { targetLists: [], existingList: null, existingFields: [], existingViews: [], existingContentTypeIds: [], availableContentTypes: null };
    let connected = false;
    let existingPolicy = 'new';   // 'new' | 'resume'
    let phase = 'form';           // 'form' | 'running' | 'report'
    let listExists = false;
    let lastReport = null;
    let liById = new Map();
    let abortController = null;
    let titleTouched = false;

    // ---- shell ------------------------------------------------------------
    const dialog = el('dialog', 'app-dialog sp-metadata-dialog wb-schema-dialog');
    const panel = el('div', 'app-dialog__panel');
    const head = el('div', 'app-dialog__head');
    head.append(el('h2', '', mode === 'copy' ? 'Copy list to…' : 'New list from schema'));
    const closeBtn = el('button', 'btn btn-ghost btn-xs wb-schema-close', '✕');
    closeBtn.type = 'button';
    closeBtn.setAttribute('aria-label', 'Close');
    head.append(closeBtn);
    const context = el('p', 'app-dialog__context', mode === 'copy'
      ? `Copy the schema of ‘${d.list.title}’ to a new or existing list.`
      : `Create a list from the imported schema for ‘${d.list.title}’.`);
    panel.append(head, context);

    // ---- 1. target site -----------------------------------------------
    const targetField = el('div', 'app-dialog__field wb-schema-target-row');
    targetField.append(el('label', '', 'Target site'));
    const targetRow = el('div', 'wb-schema-target-inputrow');
    const targetInput = el('input', 'wb-schema-target');
    targetInput.type = 'text';
    targetInput.autocomplete = 'off';
    targetInput.setAttribute('list', 'wb-schema-target-list');
    const datalist = el('datalist');
    datalist.id = 'wb-schema-target-list';
    targetRow.append(targetInput, datalist);
    const connectBtn = el('button', 'btn btn-xs wb-schema-connect', 'Connect');
    connectBtn.type = 'button';
    targetRow.append(connectBtn);
    targetField.append(targetRow);
    const targetStatus = el('div', 'wb-schema-target-status');
    targetStatus.hidden = true;
    targetField.append(targetStatus);
    panel.append(targetField);

    function fillDatalist() {
      datalist.textContent = '';
      const host = el('option', '', 'This site (host web)');
      host.value = '';
      datalist.append(host);
      for (const fav of getFavorites()) {
        const opt = el('option', '', fav.title || fav.url);
        opt.value = fav.url;
        datalist.append(opt);
      }
      for (const rec of getRecents()) {
        const opt = el('option', '', rec.title || rec.url);
        opt.value = rec.url;
        datalist.append(opt);
      }
    }
    fillDatalist();

    // ---- 2. title -------------------------------------------------------
    const titleField = el('div', 'app-dialog__field wb-schema-title-row');
    titleField.append(el('label', '', 'Title'));
    const titleInput = el('input', 'wb-schema-title');
    titleInput.type = 'text';
    titleInput.autocomplete = 'off';
    titleField.append(titleInput);
    const titleStatus = el('div', 'wb-schema-title-status');
    titleField.append(titleStatus);
    panel.append(titleField);

    // ---- 3. description ---------------------------------------------------
    const descField = el('div', 'app-dialog__field wb-schema-description-row');
    descField.append(el('label', '', 'Description'));
    const descInput = el('textarea', 'wb-schema-description');
    descInput.rows = 2;
    descInput.value = d.list.description || '';
    descField.append(descInput);
    panel.append(descField);

    // ---- 4. lookup targets -------------------------------------------------
    const lookupSection = el('div', 'wb-schema-lookups-section');
    lookupSection.hidden = true;
    lookupSection.append(el('h3', '', 'Lookup targets'));
    const lookupTable = el('table', 'wb-table wb-schema-lookups');
    const lookupHead = el('thead');
    const headRow = el('tr');
    headRow.append(el('th', '', 'Column'), el('th', '', 'Source list'), el('th', '', 'Target list'));
    lookupHead.append(headRow);
    const lookupBody = el('tbody');
    lookupTable.append(lookupHead, lookupBody);
    lookupSection.append(lookupTable);
    const policyRow = el('div', 'wb-schema-lookup-policy');
    policyRow.hidden = true;
    const policyLabel = el('div', '', 'When a lookup target is missing on the target:');
    const skipLabel = el('label', 'wb-schema-lookup-policy-opt');
    const skipRadio = el('input');
    skipRadio.type = 'radio';
    skipRadio.name = 'wb-schema-missing-lookup';
    skipRadio.value = 'skip';
    skipRadio.checked = true;
    skipLabel.append(skipRadio, el('span', '', 'Skip the column'));
    const textLabel = el('label', 'wb-schema-lookup-policy-opt');
    const textRadio = el('input');
    textRadio.type = 'radio';
    textRadio.name = 'wb-schema-missing-lookup';
    textRadio.value = 'text';
    textLabel.append(textRadio, el('span', '', 'Create it as a single line of text'));
    policyRow.append(policyLabel, skipLabel, textLabel);
    lookupSection.append(policyRow);
    panel.append(lookupSection);

    // ---- 5. content types ---------------------------------------------------
    const ctSection = el('div', 'wb-schema-cts-section');
    ctSection.hidden = true;
    ctSection.append(el('h3', '', 'Content types'));
    const ctList = el('ul', 'wb-schema-cts-list');
    ctSection.append(ctList);
    ctSection.append(el('p', 'wb-schema-note', 'Missing content types are not created in this stage.'));
    panel.append(ctSection);

    // ---- 6. items (stage 1b) ---------------------------------------------
    const itemsFieldset = el('fieldset', 'wb-schema-items');
    itemsFieldset.disabled = true;
    itemsFieldset.title = 'Item data arrives in stage 1b — this run copies the schema only.';
    const legend = el('legend', '', 'Items');
    itemsFieldset.append(legend);
    const itemsRow = el('div', 'wb-schema-items-row');
    for (const label of ['Include items', 'Include attachments', 'Preserve authorship']) {
      const cbLabel = el('label', 'wb-schema-items-opt');
      const cb = el('input');
      cb.type = 'checkbox';
      cbLabel.append(cb, el('span', '', label));
      itemsRow.append(cbLabel);
    }
    itemsFieldset.append(itemsRow);
    panel.append(itemsFieldset);

    // ---- 7. existing-target policy -----------------------------------------
    const existingSection = el('div', 'wb-schema-existing sp-metadata-consent');
    existingSection.hidden = true;
    const newTitleLabel = el('label', 'sp-metadata-consent__row');
    const newTitleRadio = el('input');
    newTitleRadio.type = 'radio';
    newTitleRadio.name = 'wb-schema-existing-policy';
    newTitleRadio.value = 'new';
    newTitleRadio.checked = true;
    newTitleLabel.append(newTitleRadio, el('span', 'sp-metadata-consent__label', 'Choose another title'));
    const resumeLabel = el('label', 'sp-metadata-consent__row');
    const resumeRadio = el('input');
    resumeRadio.type = 'radio';
    resumeRadio.name = 'wb-schema-existing-policy';
    resumeRadio.value = 'resume';
    resumeLabel.append(resumeRadio, el('span', 'sp-metadata-consent__label', 'Add missing fields and views to the existing list'));
    existingSection.append(newTitleLabel, resumeLabel);
    const gateRow = el('label', 'sp-metadata-consent__row wb-schema-gate');
    gateRow.hidden = true;
    const gateBox = el('input');
    gateBox.type = 'checkbox';
    const gateLabel = el('span', 'sp-metadata-consent__label');
    gateRow.append(gateBox, gateLabel);
    existingSection.append(gateRow);
    panel.append(existingSection);

    // ---- plan (dry run) -----------------------------------------------------
    const planPanel = el('div', 'wb-schema-plan');
    planPanel.hidden = true;
    const planHeading = el('p', 'wb-schema-plan-heading');
    planPanel.append(planHeading);
    const stepsList = el('ol', 'wb-schema-steps');
    planPanel.append(stepsList);
    const masterLine = el('p', 'wb-schema-master');
    masterLine.hidden = true;
    planPanel.append(masterLine);
    panel.append(planPanel);

    // ---- report ---------------------------------------------------------
    const reportPanel = el('div', 'wb-schema-report');
    reportPanel.hidden = true;
    const reportHeadline = el('p', 'wb-schema-report-headline');
    const reportCounts = el('table', 'wb-table wb-schema-report-counts');
    const reportFailed = el('div', 'wb-schema-report-failed');
    const reportWarnings = el('div', 'wb-schema-report-warnings');
    const reportLinks = el('div', 'wb-schema-report-links');
    reportPanel.append(reportHeadline, reportCounts, reportFailed, reportWarnings, reportLinks);
    panel.append(reportPanel);

    // ---- error / actions ----------------------------------------------
    const error = el('div', 'sp-files-error');
    error.setAttribute('role', 'alert');
    error.hidden = true;
    panel.append(error);

    const actions = el('div', 'app-dialog__actions sp-metadata-actions wb-schema-actions-row');
    const cancelBtn = el('button', 'btn btn-ghost wb-schema-cancel', 'Cancel');
    cancelBtn.type = 'button';
    const dryRunBtn = el('button', 'btn btn-xs wb-schema-dryrun', 'Dry run');
    dryRunBtn.type = 'button';
    const retryBtn = el('button', 'btn btn-xs wb-schema-retry', 'Retry failed steps');
    retryBtn.type = 'button';
    retryBtn.hidden = true;
    const downloadBtn = el('button', 'btn btn-xs wb-schema-report-download', 'Download report .md');
    downloadBtn.type = 'button';
    downloadBtn.hidden = true;
    const createBtn = el('button', 'btn btn-run wb-schema-create', 'Create list');
    createBtn.type = 'button';
    actions.append(cancelBtn, dryRunBtn, retryBtn, downloadBtn, createBtn);
    panel.append(actions);

    dialog.append(panel);
    document.body.append(dialog);

    const finish = (outcome) => { dialog.close(); dialog.remove(); resolve(outcome); };

    // ---- helpers ------------------------------------------------------

    function showError(message) {
      error.textContent = message;
      error.hidden = !message;
    }

    function sameAsSource() {
      return mode === 'copy' && probe.existingList
        && cleanId(probe.existingList.id) === cleanId(d.source?.listId);
    }

    function updateTitleStatus() {
      if (!probe.existingList) {
        titleStatus.textContent = connected ? 'Available' : '';
        titleStatus.classList.remove('wb-schema-title-taken');
      } else if (sameAsSource()) {
        titleStatus.textContent = 'Choose a title different from the source list — '
          + 'this is the same list on the same site.';
        titleStatus.classList.add('wb-schema-title-taken');
      } else {
        titleStatus.textContent = `A list named ‘${probe.existingList.title}’ already exists on the target.`;
        titleStatus.classList.add('wb-schema-title-taken');
      }
    }

    function updateExistingSection() {
      const show = Boolean(probe.existingList) && !sameAsSource();
      existingSection.hidden = !show;
      if (!show) { existingPolicy = 'new'; newTitleRadio.checked = true; gateRow.hidden = true; return; }
      gateLabel.textContent = `I understand this changes ‘${probe.existingList.title}’: missing columns and `
        + 'views are added, views with the same title have their columns rebuilt, and attachments or folders '
        + 'are switched on if the source needs them. Nothing is deleted.';
      gateRow.hidden = existingPolicy !== 'resume';
    }

    function customLookups() {
      return d.fields.filter((f) => f.custom && LOOKUP_TYPES.has(f.type));
    }

    function buildLookupMap() {
      const map = {};
      for (const select of lookupBody.querySelectorAll('select[data-key]')) {
        if (select.value && select.dataset.key) map[select.dataset.key] = select.value;
      }
      return map;
    }

    function renderLookupRows() {
      const lookups = customLookups();
      lookupSection.hidden = lookups.length === 0;
      lookupBody.textContent = '';
      let anyMissing = false;
      for (const f of lookups) {
        const tr = el('tr');
        tr.append(el('td', '', f.displayName || f.internalName));
        tr.append(el('td', '', f.lookupList || f.lookupListId || '—'));
        const td = el('td');
        if (f.isSelfLookup) {
          td.append(el('span', '', '(this list)'));
        } else {
          const select = el('select', 'wb-schema-lookup-target');
          const emptyOpt = el('option', '', '— missing —');
          emptyOpt.value = '';
          select.append(emptyOpt);
          for (const l of probe.targetLists || []) {
            const opt = el('option', '', l.title);
            opt.value = l.title;
            select.append(opt);
          }
          const wanted = String(f.lookupList || '').toLowerCase();
          const found = (probe.targetLists || []).find((l) => String(l.title).toLowerCase() === wanted);
          select.value = found ? found.title : '';
          if (!found) anyMissing = true;
          select.dataset.key = f.lookupList || f.lookupListId || '';
          td.append(select);
        }
        tr.append(td);
        lookupBody.append(tr);
      }
      policyRow.hidden = !anyMissing;
    }

    function renderContentTypes() {
      ctSection.hidden = !d.list.contentTypesEnabled;
      if (!d.list.contentTypesEnabled) return;
      ctList.textContent = '';
      const available = new Set((probe.availableContentTypes || []).map((c) => c.id));
      const already = new Set(probe.existingContentTypeIds || []);
      const seen = new Set();
      for (const ct of d.contentTypes) {
        const parentId = ct.parentId || parentContentTypeId(ct.id);
        if (isBuiltinParent(parentId) || seen.has(parentId)) continue;
        seen.add(parentId);
        const status = already.has(parentId) ? 'already on the target'
          : available.has(parentId) ? 'available on the target'
            : 'missing on the target; will be skipped';
        ctList.append(el('li', '', `${ct.name} — ${status}`));
      }
    }

    function canDryRun() {
      return connected && Boolean(titleInput.value.trim()) && !sameAsSource()
        && (!probe.existingList || existingPolicy === 'resume');
    }
    function canCreate() {
      return canDryRun() && (!probe.existingList || gateBox.checked);
    }

    function refreshGate() {
      dryRunBtn.disabled = !canDryRun();
      createBtn.disabled = !canCreate();
      createBtn.textContent = probe.existingList && existingPolicy === 'resume'
        ? 'Add to existing list' : 'Create list';
    }

    async function refreshProbeForTitle(title) {
      try {
        probe = await probeTarget(targetClient, { title, doc: d });
      } catch (err) {
        showError(err.message || String(err));
        return;
      }
      showError('');
      updateTitleStatus();
      updateExistingSection();
      renderLookupRows();
      renderContentTypes();
      refreshGate();
    }

    async function connect(rawUrl) {
      showError('');
      targetStatus.hidden = true;
      connectBtn.disabled = true;
      connected = false;
      try {
        await targetClient.connectWeb(rawUrl);
      } catch (err) {
        connectBtn.disabled = false;
        targetStatus.textContent = err.message || String(err);
        targetStatus.hidden = false;
        probe = { targetLists: [], existingList: null, existingFields: [], existingViews: [], existingContentTypeIds: [], availableContentTypes: null };
        updateTitleStatus();
        refreshGate();
        return false;
      }
      connectBtn.disabled = false;
      connected = true;
      let initial;
      try {
        initial = await probeTarget(targetClient, { title: '', doc: d });
      } catch (err) {
        showError(err.message || String(err));
        return false;
      }
      probe = initial;
      if (!titleTouched) titleInput.value = defaultTargetTitle(d, probe.targetLists);
      await refreshProbeForTitle(titleInput.value.trim());
      return true;
    }

    connectBtn.addEventListener('click', () => connect(targetInput.value));
    let titleDebounce = null;
    titleInput.addEventListener('input', () => {
      titleTouched = true;
      clearTimeout(titleDebounce);
      titleDebounce = setTimeout(() => {
        if (connected) refreshProbeForTitle(titleInput.value.trim());
      }, 150);
    });
    newTitleRadio.addEventListener('change', () => {
      existingPolicy = 'new'; gateBox.checked = false; updateExistingSection(); refreshGate();
    });
    resumeRadio.addEventListener('change', () => {
      existingPolicy = 'resume'; updateExistingSection(); refreshGate();
    });
    gateBox.addEventListener('change', refreshGate);

    // ---- step list / progress -------------------------------------------

    function buildStepsList(steps) {
      stepsList.textContent = '';
      liById = new Map();
      for (const step of steps) {
        const li = el('li', '', stepLine(step));
        li.dataset.state = STATE_MAP[step.status] || 'pending';
        stepsList.append(li);
        liById.set(step.id, li);
      }
    }

    function handleStep(step, i, total) {
      const li = liById.get(step.id);
      if (li) {
        li.dataset.state = STATE_MAP[step.status] || 'pending';
        li.textContent = stepLine(step);
      }
      masterLine.hidden = false;
      masterLine.textContent = `${verbFor(step)} ${i + 1} of ${total}…`;
      if (currentCtx?.listId && !listExists) {
        listExists = true;
        closeBtn.hidden = true;
        cancelBtn.hidden = true;
      }
    }

    let currentCtx = null;

    // ---- dry run ----------------------------------------------------------

    async function runDryRun() {
      if (!canDryRun()) return;
      showError('');
      dryRunBtn.disabled = true;
      try {
        probe = await probeTarget(targetClient, { title: titleInput.value.trim(), doc: d });
        updateTitleStatus();
        updateExistingSection();
        const plan = buildApplyPlan(d, buildOptions(), probe);
        planPanel.hidden = false;
        planHeading.textContent = 'Plan — nothing has been written.';
        masterLine.hidden = true;
        buildStepsList(plan.steps);
      } catch (err) {
        showError(err.message || String(err));
      } finally {
        refreshGate();
      }
    }

    function buildOptions() {
      return {
        title: titleInput.value.trim(),
        description: descInput.value,
        existing: probe.existingList && existingPolicy === 'resume' ? 'resume' : 'fail',
        missingLookup: [...policyRow.querySelectorAll('input[name="wb-schema-missing-lookup"]')]
          .find((r) => r.checked)?.value === 'text' ? 'text' : 'skip',
        lookupMap: buildLookupMap(),
        targetWebUrl: targetClient.webUrl(),
      };
    }

    // ---- create / progress / report ---------------------------------------

    function setPhase(next) {
      phase = next;
      const inForm = next === 'form';
      const inRunning = next === 'running';
      const inReport = next === 'report';
      targetField.hidden = !inForm;
      titleField.hidden = !inForm;
      descField.hidden = !inForm;
      itemsFieldset.hidden = !inForm;
      if (inForm) {
        renderLookupRows(); renderContentTypes(); updateExistingSection();
      } else {
        lookupSection.hidden = true; ctSection.hidden = true; existingSection.hidden = true;
      }
      // The plan panel is dry run's own toggle while in 'form' (left as-is
      // here), forced visible for live progress, and hidden once the report
      // takes over.
      if (inRunning) planPanel.hidden = false;
      else if (inReport) planPanel.hidden = true;
      reportPanel.hidden = !inReport;
      dryRunBtn.hidden = !inForm;
      createBtn.hidden = !inForm;
      retryBtn.hidden = true;
      downloadBtn.hidden = !inReport;
      cancelBtn.hidden = false;
      cancelBtn.textContent = inReport ? 'Close' : 'Cancel';
      closeBtn.hidden = false;
    }

    async function runCreate() {
      if (!canCreate()) return;
      showError('');
      dryRunBtn.disabled = true;
      createBtn.disabled = true;
      let plan;
      try {
        probe = await probeTarget(targetClient, { title: titleInput.value.trim(), doc: d });
        plan = buildApplyPlan(d, buildOptions(), probe);
      } catch (err) {
        showError(err.message || String(err));
        dryRunBtn.disabled = false;
        createBtn.disabled = false;
        return;
      }
      setPhase('running');
      planHeading.textContent = `Applying the plan to ‘${targetClient.webUrl()}’…`;
      buildStepsList(plan.steps);
      listExists = false;
      abortController = new AbortController();
      currentCtx = { client: targetClient, spWrite, onStep: handleStep, signal: abortController.signal };
      const report = await runPlan(plan, currentCtx);
      lastReport = report;
      renderReport(report);
      setPhase('report');
    }

    async function runRetry() {
      if (!lastReport) return;
      const plan = retryPlan(lastReport);
      setPhase('running');
      planHeading.textContent = 'Retrying the failed steps…';
      buildStepsList(plan.steps);
      listExists = Boolean(plan.existingListId);
      abortController = new AbortController();
      currentCtx = { client: targetClient, spWrite, onStep: handleStep, signal: abortController.signal };
      const retryReport = await runPlan(plan, currentCtx);
      lastReport = mergeRetryReport(lastReport, retryReport);
      renderReport(lastReport);
      setPhase('report');
    }

    function renderReportCounts(report) {
      reportCounts.textContent = '';
      const tbody = el('tbody');
      const rows = [
        ['Fields', `${report.fields.added} added · ${report.fields.skipped} skipped · ${report.fields.failed.length} failed`],
        ['Views', `${report.views.added} added · ${report.views.updated} updated · ${report.views.failed.length} failed`],
        ['Content types', `${report.contentTypes.attached} attached · ${report.contentTypes.skipped} skipped · ${report.contentTypes.failed} failed`],
        ['Validation', report.validation.applied ? 'applied' : (report.validation.error ? 'failed' : '—')],
      ];
      for (const [label, value] of rows) {
        const tr = el('tr');
        tr.append(el('td', '', label), el('td', '', value));
        tbody.append(tr);
      }
      reportCounts.append(tbody);
    }

    function renderReportFailed(report) {
      reportFailed.textContent = '';
      const failed = report.steps.filter((s) => s.status === 'failed');
      if (!failed.length) return;
      reportFailed.append(el('h3', '', 'Failed steps'));
      const table = el('table', 'wb-table');
      const tbody = el('tbody');
      for (const s of failed) {
        const tr = el('tr');
        tr.append(el('td', '', s.label), el('td', '', s.error || ''));
        tbody.append(tr);
      }
      table.append(tbody);
      reportFailed.append(table);
    }

    function renderReportWarnings(report) {
      reportWarnings.textContent = '';
      if (!report.warnings?.length) return;
      reportWarnings.append(el('h3', '', 'Warnings'));
      const list = el('ul', 'wb-grid-notice');
      for (const w of report.warnings) list.append(el('li', '', w));
      reportWarnings.append(list);
    }

    function renderReportLinks(report) {
      reportLinks.textContent = '';
      if (!report.listId) return;
      const sameWeb = canonUrl(targetClient.webUrl()) === canonUrl(client.webUrl());
      if (sameWeb) {
        const openBtn = el('button', 'btn btn-xs wb-schema-report-open', 'Open the new list');
        openBtn.type = 'button';
        openBtn.addEventListener('click', () => {
          navigate({ view: 'lists', listId: report.listId, listTitle: report.title });
          finish('created');
        });
        reportLinks.append(openBtn);
      } else {
        const settingsLink = el('a', 'btn btn-xs wb-schema-report-settings', 'Open list settings ↗');
        settingsLink.href = linkUrl(targetClient.webUrl(), LIST_SETTINGS, { guid: report.listId });
        bindNewTab(settingsLink);
        const inspectBtn = el('button', 'btn btn-xs wb-schema-report-inspect', 'Inspect the target site');
        inspectBtn.type = 'button';
        inspectBtn.addEventListener('click', async () => {
          await inspectSite(targetClient.webUrl());
          navigate({ view: 'lists', listId: report.listId, listTitle: report.title });
          finish('created');
        });
        reportLinks.append(settingsLink, inspectBtn);
      }
    }

    function renderReport(report) {
      reportHeadline.textContent = buildHeadline(report, { isMock: spWrite.isMock() });
      reportHeadline.classList.toggle('wb-schema-report-auth', report.aborted === 'auth');
      renderReportCounts(report);
      renderReportFailed(report);
      renderReportWarnings(report);
      renderReportLinks(report);
      retryBtn.hidden = !report.steps.some((s) => s.status === 'failed' && !s.final);
      downloadBtn.hidden = false;
    }

    // ---- actions ------------------------------------------------------

    dryRunBtn.addEventListener('click', runDryRun);
    createBtn.addEventListener('click', runCreate);
    retryBtn.addEventListener('click', runRetry);
    downloadBtn.addEventListener('click', () => {
      if (!lastReport) return;
      const stem = String(lastReport.title || 'list').toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '') || 'list';
      downloadText(`schema-report-${stem}.md`, buildApplyReport({ report: lastReport, doc: d, targetWebUrl: lastReport.targetWebUrl }), 'text/markdown');
    });

    cancelBtn.addEventListener('click', () => {
      if (phase === 'running') { abortController?.abort(); return; }
      if (phase === 'report') { finish(lastReport?.listId ? 'created' : 'cancelled'); return; }
      finish('cancelled');
    });
    closeBtn.addEventListener('click', () => {
      if (phase === 'running') return;
      finish(phase === 'report' && lastReport?.listId ? 'created' : 'cancelled');
    });
    dialog.addEventListener('cancel', (e) => {
      e.preventDefault();
      if (phase === 'running') { if (!listExists) abortController?.abort(); return; }
      finish(phase === 'report' && lastReport?.listId ? 'created' : 'cancelled');
    });

    // ---- boot -----------------------------------------------------------
    setPhase('form');
    dryRunBtn.disabled = true;
    createBtn.disabled = true;
    dialog.showModal();
    targetInput.value = client.webUrl();
    connect(client.webUrl());
  });
}
