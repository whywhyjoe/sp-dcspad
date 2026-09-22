// List item DATA — the Tools tab's "Import data into this list" dialog.
//
// Same recipe as list-schema-dialog.js (`<dialog class="app-dialog
// sp-metadata-dialog …">`, two-phase exit, retry-free single run, a report
// panel with a failed-steps table, warnings, and Download report .md) —
// deliberately smaller, since the target here is fixed (this list, already
// open in the drilldown): no target-site connect, no dry run, no lookup
// remap. The dialog reads this list's own fields (writableFields against a
// fresh read) to show what will be written versus dropped, then hands the
// run straight to applyListData().

import { writableFields } from './list-data.js';
import { NEVER_WRITE, NEVER_WRITE_TYPES } from './list-schema.js';
import { applyListData } from './list-data-apply.js';
import { createSpWriteClient } from './sp-write.js';
import { EXPIRED_SESSION_NOTE, showFailure } from './denied.js';
import { downloadText } from '../io.js?v=2';

// How long a run gets, after Cancel, before the dialog offers "Close
// anyway" — findings review #1: Cancel only flips the AbortSignal, which
// applyListData checks BETWEEN items (list-data-apply.js runItems/
// runFolders/runSelfLookups/runAuthorship), so a single in-flight write that
// never settles (a dropped connection, a stalled tenant) would otherwise
// leave the dialog with no way out. Mutable + exported so a test can shorten
// it deterministically instead of waiting out the real 15s.
export const DEFAULT_CLOSE_ANYWAY_MS = 15000;
let closeAnywayMs = DEFAULT_CLOSE_ANYWAY_MS;
export function setCloseAnywayMs(ms) {
  closeAnywayMs = Number.isFinite(ms) && ms >= 0 ? ms : DEFAULT_CLOSE_ANYWAY_MS;
}

const CLOSE_ANYWAY_NOTE =
  'The current request is still in flight — closing leaves it running; the Items tab will refresh.';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const guidPath = (listId, sub = '') => `web/lists(guid'${listId}')${sub}`;
const FIELD_SELECT = [
  'Id', 'InternalName', 'TypeAsString', 'ReadOnlyField', 'LookupList', 'LookupField', 'DisplayFormat',
];
const stem = (s) => String(s || 'list').toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '') || 'list';
const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

// Why a source column won't be written here — mirrors writableFields()'s own
// rule (list-data.js) so the table can never disagree with the run itself.
function dropReason(name, meta, targetFieldRow) {
  if (!targetFieldRow) return 'missing on this list';
  if (targetFieldRow.ReadOnlyField) return 'read-only on this list';
  if (NEVER_WRITE.has(name) || NEVER_WRITE_TYPES.has(targetFieldRow.TypeAsString)) return 'system column — not imported';
  if (!(meta?.custom || name === 'Title')) return 'system column — not imported';
  return 'not writable';
}

// The four authorship-pass fields (list-data-apply.js runAuthorship) never
// go through writableFields()/dropReason() — they're not in dataDoc.fields
// at all, SPUtils restores them as a dedicated post-create pass — so their
// row is driven directly by the "Preserve authorship" checkbox instead of
// the source/target field comparison every other row uses. Recomputed
// (rebuildFieldsTable, below) each time that checkbox toggles so the
// preview never lies about what a run will actually write.
const AUTHORSHIP_FIELD_NAMES = ['Author', 'Editor', 'Created', 'Modified'];

function buildFieldTable(dataDoc, targetFieldRows, writable, preserveAuthorship) {
  const table = el('table', 'wb-table wb-import-fields');
  const thead = el('thead');
  const headRow = el('tr');
  headRow.append(el('th', '', 'Column'), el('th', '', 'Status'));
  thead.append(headRow);
  const tbody = el('tbody');
  const byName = new Map(targetFieldRows.map((f) => [f.InternalName, f]));
  const writtenNames = new Set(writable.map((w) => w.name));
  for (const [name, meta] of Object.entries(dataDoc.fields || {})) {
    const tr = el('tr');
    tr.append(el('td', '', name));
    const status = writtenNames.has(name) ? 'written' : dropReason(name, meta, byName.get(name));
    const td = el('td');
    td.className = writtenNames.has(name) ? '' : 'wb-import-dropped';
    td.textContent = status;
    tr.append(td);
    tbody.append(tr);
  }
  for (const name of AUTHORSHIP_FIELD_NAMES) {
    const tr = el('tr', 'wb-import-authorship-row');
    tr.append(el('td', '', name));
    const td = el('td');
    td.className = preserveAuthorship ? '' : 'wb-import-dropped';
    td.textContent = preserveAuthorship ? 'written — preserve authorship' : 'not written (authorship off)';
    tr.append(td);
    tbody.append(tr);
  }
  table.append(thead, tbody);
  return table;
}

// Failure totals must count the items the report couldn't even keep a row
// for (items.failedTruncated, list-data-apply.js's FAILED_CAP) — not just
// the ones it lists — everywhere a total is shown, on screen and in the
// downloaded .md, with a note whenever some are missing from the itemised
// list below. Exported (single source of truth for both surfaces) so it can
// be unit-tested directly instead of only through the rendered dialog.
export function itemFailureSummary(items) {
  const total = (items.failed?.length || 0) + (items.failedTruncated || 0);
  return items.failedTruncated ? `${total} failed (${items.failedTruncated} not itemised below)` : `${plural(total, 'failure')}`;
}

function buildHeadline(report, { listTitle, isMock }) {
  if (report.aborted === 'auth') return EXPIRED_SESSION_NOTE;
  if (report.aborted === 'user') return 'Import cancelled.';
  let headline = `Imported into ‘${listTitle}’ — ${plural(report.items.added, 'item')} added, ${itemFailureSummary(report.items)}.`;
  if (isMock) headline += ' (mock mode — the fixture web does not change).';
  return headline;
}

function buildReportMarkdown(report, { listTitle }) {
  const lines = [`# Import report — ${listTitle}`, '', buildHeadline(report, { listTitle, isMock: false }), ''];
  lines.push('## Counts', '');
  lines.push(`- Items: ${report.items.added} added, ${itemFailureSummary(report.items)}`);
  lines.push(`- Folders: ${report.folders.created} created, ${report.folders.failed} failed`);
  lines.push(`- Attachments: ${report.attachments.added} added, ${report.attachments.skipped} skipped, ${report.attachments.failed} failed`);
  lines.push(`- Authorship: ${report.authorship.applied} applied, ${report.authorship.failed} failed`);
  lines.push('');
  if (report.items.failed.length) {
    lines.push('## Failed items', '');
    for (const f of report.items.failed) lines.push(`- Source id ${f.sourceId}: ${f.error}`);
    lines.push('');
  }
  if (report.fieldErrors.length) {
    lines.push('## Field errors', '');
    for (const fe of report.fieldErrors) lines.push(`- Source id ${fe.sourceId} — ${fe.field}: ${fe.message}`);
    lines.push('');
  }
  if (report.warnings.length) {
    lines.push('## Warnings', '');
    for (const w of report.warnings) lines.push(`- ${w}`);
    lines.push('');
  }
  return lines.join('\n');
}

// { dataDoc, client, listId, listTitle, mockWriter, invalidateItems } ->
// Promise<'imported'|'cancelled'|'closed-during-run'>. `invalidateItems`
// (ctx.invalidateItems from list-tools.js) is called by the dialog itself,
// at most once, the moment a report shows the list may have changed —
// findings #3: added items or created folders, from ANY outcome (a normal
// finish, a cancel that already wrote, or Close anyway) — never left to the
// caller to infer from the returned outcome string.
export function openImportDataDialog({
  dataDoc, client, listId, listTitle, mockWriter, invalidateItems,
} = {}) {
  return new Promise((resolve) => {
    const isMock = !client.context().live;
    const spWrite = createSpWriteClient({ client, mockWriter: isMock ? mockWriter : undefined });

    const dialog = el('dialog', 'app-dialog sp-metadata-dialog wb-schema-dialog wb-import-data-dialog');
    const panel = el('div', 'app-dialog__panel');
    const head = el('div', 'app-dialog__head');
    head.append(el('h2', '', 'Import data'));
    const closeBtn = el('button', 'btn btn-ghost btn-xs wb-schema-close', '✕');
    closeBtn.type = 'button';
    closeBtn.setAttribute('aria-label', 'Close');
    head.append(closeBtn);
    const context = el('p', 'app-dialog__context',
      `Add items from ‘${dataDoc.source?.listTitle || 'a list'}’${dataDoc.source?.siteUrl ? ` (${dataDoc.source.siteUrl})` : ''} to ‘${listTitle}’.`);
    panel.append(head, context);

    const countsLine = el('p', 'wb-schema-note',
      `${plural(dataDoc.items.length, 'item')}, ${plural(dataDoc.folders.length, 'folder')} in the source file.`);
    panel.append(countsLine);

    const fieldsSection = el('div', 'wb-schema-section');
    fieldsSection.append(el('h3', '', 'Columns'));
    const fieldsStatus = el('p', 'wb-schema-note', 'Reading this list’s columns…');
    fieldsSection.append(fieldsStatus);
    panel.append(fieldsSection);

    const optsFieldset = el('fieldset', 'wb-schema-items');
    optsFieldset.append(el('legend', '', 'Options'));
    const optsRow = el('div', 'wb-schema-items-row');
    const attachCb = el('input');
    attachCb.type = 'checkbox';
    const attachLabel = el('label', 'wb-schema-items-opt');
    attachLabel.append(attachCb, el('span', '', 'Include attachments'));
    const authorshipCb = el('input');
    authorshipCb.type = 'checkbox';
    const authorshipLabel = el('label', 'wb-schema-items-opt');
    authorshipLabel.append(authorshipCb, el('span', '', 'Preserve authorship'));
    optsRow.append(attachLabel, authorshipLabel);
    optsFieldset.append(optsRow);
    panel.append(optsFieldset);

    const consentSection = el('div', 'sp-metadata-consent');
    const consentLabel = el('label', 'sp-metadata-consent__row');
    const consentBox = el('input');
    consentBox.type = 'checkbox';
    const consentText = el('span', 'sp-metadata-consent__label',
      `I understand this adds ${dataDoc.items.length} item${dataDoc.items.length === 1 ? '' : 's'} to ‘${listTitle}’. `
      + 'Existing items are not changed or de-duplicated.');
    consentLabel.append(consentBox, consentText);
    consentSection.append(consentLabel);
    panel.append(consentSection);

    // ---- progress ----
    const planPanel = el('div', 'wb-schema-plan');
    planPanel.hidden = true;
    const masterLine = el('p', 'wb-schema-master');
    // Findings #1: Cancel only aborts the signal — a request already in
    // flight settles on its own timing. If it hasn't by closeAnywayMs, this
    // becomes the escape hatch so the dialog can never trap the operator.
    const closeAnywayNote = el('p', 'wb-schema-note wb-import-closeanyway-note', CLOSE_ANYWAY_NOTE);
    closeAnywayNote.hidden = true;
    const closeAnywayBtn = el('button', 'btn btn-xs wb-import-closeanyway', 'Close anyway');
    closeAnywayBtn.type = 'button';
    closeAnywayBtn.hidden = true;
    planPanel.append(masterLine, closeAnywayNote, closeAnywayBtn);
    panel.append(planPanel);

    // ---- report ----
    const reportPanel = el('div', 'wb-schema-report');
    reportPanel.hidden = true;
    const reportHeadline = el('p', 'wb-schema-report-headline');
    const reportCounts = el('table', 'wb-table wb-schema-report-counts');
    const reportFailed = el('div', 'wb-schema-report-failed');
    const reportWarnings = el('div', 'wb-schema-report-warnings');
    reportPanel.append(reportHeadline, reportCounts, reportFailed, reportWarnings);
    panel.append(reportPanel);

    const error = el('div', 'sp-files-error');
    error.setAttribute('role', 'alert');
    error.hidden = true;
    panel.append(error);

    const actions = el('div', 'app-dialog__actions sp-metadata-actions wb-schema-actions-row');
    const cancelBtn = el('button', 'btn btn-ghost wb-schema-cancel', 'Cancel');
    cancelBtn.type = 'button';
    const downloadBtn = el('button', 'btn btn-xs wb-schema-report-download', 'Download report .md');
    downloadBtn.type = 'button';
    downloadBtn.hidden = true;
    const importBtn = el('button', 'btn btn-run wb-import-run', 'Import');
    importBtn.type = 'button';
    importBtn.disabled = true;
    actions.append(cancelBtn, downloadBtn, importBtn);
    panel.append(actions);

    dialog.append(panel);
    document.body.append(dialog);

    let phase = 'form';
    let lastReport = null;
    let abortController = null;
    let cancelRequested = false;
    let closeAnywayTimer = null;
    let invalidated = false;
    let previewOk = false;
    let fieldsTableEl = null;
    let cachedTargetFieldRows = null;
    let cachedWritable = null;

    const finish = (outcome) => { dialog.close(); dialog.remove(); resolve(outcome); };

    // Findings #3: a report is grounds to invalidate the Items tab the
    // moment items were added or folders created — regardless of how the
    // run ended (finished clean, cancelled after writing, or 401'd
    // mid-write) — never gated on the outcome string the dialog resolves
    // with. Idempotent: Close anyway (below) always invalidates too, and a
    // run that later settles on its own must not invalidate a second time.
    const mutated = (report) => Boolean(report && ((report.items?.added || 0) > 0 || (report.folders?.created || 0) > 0));
    function invalidateOnce() {
      if (invalidated) return;
      invalidated = true;
      invalidateItems?.();
    }

    function hideCloseAnyway() {
      closeAnywayNote.hidden = true;
      closeAnywayBtn.hidden = true;
      if (closeAnywayTimer) { clearTimeout(closeAnywayTimer); closeAnywayTimer = null; }
    }

    // Findings #1: Cancel flips the signal (checked BETWEEN items by
    // list-data-apply.js's loops) but cannot interrupt a single request
    // already in flight. If the run hasn't settled closeAnywayMs after
    // Cancel, "Close anyway" appears so the dialog is never stuck open.
    function requestCancel() {
      if (phase !== 'running' || cancelRequested) return;
      cancelRequested = true;
      abortController?.abort();
      closeAnywayTimer = setTimeout(() => {
        closeAnywayTimer = null;
        if (phase === 'running') { closeAnywayNote.hidden = false; closeAnywayBtn.hidden = false; }
      }, closeAnywayMs);
    }

    function setPhase(next) {
      phase = next;
      const inForm = next === 'form';
      const inReport = next === 'report';
      countsLine.hidden = !inForm;
      fieldsSection.hidden = !inForm;
      optsFieldset.hidden = !inForm;
      consentSection.hidden = !inForm;
      planPanel.hidden = next !== 'running';
      reportPanel.hidden = !inReport;
      importBtn.hidden = !inForm;
      downloadBtn.hidden = !inReport;
      cancelBtn.textContent = inReport ? 'Close' : 'Cancel';
      closeBtn.hidden = next === 'running';
    }

    // Findings #2: Import needs BOTH consent and a successfully loaded
    // column preview — a preview failure (denied, expired session, a
    // network error) must never leave Import clickable against a table the
    // operator never actually saw.
    function updateImportEnabled() { importBtn.disabled = !(consentBox.checked && previewOk); }
    consentBox.addEventListener('change', updateImportEnabled);

    // Findings #5: the authorship rows (Author/Editor/Created/Modified)
    // reflect the "Preserve authorship" checkbox, not the source/target
    // field comparison every other row uses — rebuilt whenever it toggles
    // so the preview can never say something the run won't actually do.
    function rebuildFieldsTable() {
      if (!cachedTargetFieldRows) return;
      fieldsTableEl?.remove();
      fieldsTableEl = buildFieldTable(dataDoc, cachedTargetFieldRows, cachedWritable, authorshipCb.checked);
      fieldsSection.append(fieldsTableEl);
    }
    authorshipCb.addEventListener('change', rebuildFieldsTable);

    client.getAll(guidPath(listId, '/fields'), { select: FIELD_SELECT })
      .then(({ items: targetFieldRows }) => {
        fieldsStatus.remove();
        cachedTargetFieldRows = targetFieldRows;
        cachedWritable = writableFields(dataDoc.fields, targetFieldRows);
        rebuildFieldsTable();
        previewOk = true;
        updateImportEnabled();
      })
      .catch((err) => {
        previewOk = false;
        showFailure(fieldsStatus, err, 'this list’s columns');
        updateImportEnabled();
      });

    function renderReportCounts(report) {
      reportCounts.textContent = '';
      const tbody = el('tbody');
      const rows = [
        ['Items', `${report.items.added} added · ${itemFailureSummary(report.items)}`],
        ['Folders', `${report.folders.created} created · ${report.folders.failed} failed`],
      ];
      if (attachCb.checked) rows.push(['Attachments', `${report.attachments.added} added · ${report.attachments.skipped} skipped · ${report.attachments.failed} failed`]);
      if (authorshipCb.checked) rows.push(['Authorship', `${report.authorship.applied} applied · ${report.authorship.failed} failed`]);
      for (const [label, value] of rows) {
        const tr = el('tr');
        tr.append(el('td', '', label), el('td', '', value));
        tbody.append(tr);
      }
      reportCounts.append(tbody);
    }

    function renderReportFailed(report) {
      reportFailed.textContent = '';
      if (report.items.failed.length) {
        reportFailed.append(el('h3', '', 'Failed items'));
        const table = el('table', 'wb-table');
        const tbody = el('tbody');
        for (const f of report.items.failed) {
          const tr = el('tr');
          tr.append(el('td', '', `Source id ${f.sourceId}`), el('td', '', f.error || ''));
          tbody.append(tr);
        }
        table.append(tbody);
        reportFailed.append(table);
        if (report.items.failedTruncated) {
          reportFailed.append(el('p', 'wb-schema-note', `+${report.items.failedTruncated} more not shown.`));
        }
      }
      if (report.fieldErrors.length) {
        reportFailed.append(el('h3', '', 'Field errors'));
        const table = el('table', 'wb-table');
        const tbody = el('tbody');
        for (const fe of report.fieldErrors) {
          const tr = el('tr');
          tr.append(el('td', '', `Source id ${fe.sourceId} — ${fe.field}`), el('td', '', fe.message || ''));
          tbody.append(tr);
        }
        table.append(tbody);
        reportFailed.append(table);
      }
    }

    function renderReportWarnings(report) {
      reportWarnings.textContent = '';
      if (!report.warnings?.length) return;
      reportWarnings.append(el('h3', '', 'Warnings'));
      const list = el('ul', 'wb-grid-notice');
      for (const w of report.warnings) list.append(el('li', '', w));
      reportWarnings.append(list);
    }

    function renderReport(report) {
      reportHeadline.textContent = buildHeadline(report, { listTitle, isMock });
      reportHeadline.classList.toggle('wb-schema-report-auth', report.aborted === 'auth');
      renderReportCounts(report);
      renderReportFailed(report);
      renderReportWarnings(report);
      downloadBtn.hidden = false;
    }

    importBtn.addEventListener('click', async () => {
      if (importBtn.disabled) return;
      error.hidden = true;
      importBtn.disabled = true;
      cancelRequested = false;
      hideCloseAnyway();
      setPhase('running');
      masterLine.textContent = 'Importing…';
      abortController = new AbortController();
      try {
        const report = await applyListData({
          dataDoc,
          client,
          spWrite,
          listId,
          options: { includeAttachments: attachCb.checked, preserveAuthorship: authorshipCb.checked },
          onStep: (info) => { masterLine.textContent = info.label; },
          signal: abortController.signal,
        });
        // The run settled on its own — a pending close-anyway timer (Cancel
        // was clicked, then the request resolved before closeAnywayMs) is
        // moot now that a report exists to show instead.
        hideCloseAnyway();
        lastReport = report;
        setPhase('report');
        renderReport(report);
        if (mutated(report)) invalidateOnce();
      } catch (err) {
        hideCloseAnyway();
        setPhase('form');
        importBtn.disabled = false;
        error.textContent = err?.message || String(err);
        error.hidden = false;
      }
    });

    downloadBtn.addEventListener('click', () => {
      if (!lastReport) return;
      downloadText(`import-report-${stem(listTitle)}.md`, buildReportMarkdown(lastReport, { listTitle }), 'text/markdown');
    });

    closeAnywayBtn.addEventListener('click', () => {
      // The awaited applyListData() in the importBtn handler above keeps
      // running after this — its eventual report (or throw) lands on a
      // detached dialog and is discarded; invalidateOnce() already covers
      // the Items tab regardless of what that report turns out to say.
      invalidateOnce();
      finish('closed-during-run');
    });

    cancelBtn.addEventListener('click', () => {
      if (phase === 'running') { requestCancel(); return; }
      if (phase === 'report') { finish(lastReport && !lastReport.aborted ? 'imported' : 'cancelled'); return; }
      finish('cancelled');
    });
    closeBtn.addEventListener('click', () => {
      if (phase === 'running') return;
      finish(phase === 'report' && lastReport && !lastReport.aborted ? 'imported' : 'cancelled');
    });
    dialog.addEventListener('cancel', (e) => {
      e.preventDefault();
      // Esc = Cancel while running (findings #1) — previously a no-op that
      // left the dialog stuck exactly like an unreachable Cancel button.
      if (phase === 'running') { requestCancel(); return; }
      finish(phase === 'report' && lastReport && !lastReport.aborted ? 'imported' : 'cancelled');
    });

    setPhase('form');
    dialog.showModal();
  });
}
