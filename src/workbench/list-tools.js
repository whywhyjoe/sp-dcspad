// SP Workbench — list Tools tab: a small data-driven registry of card-shaped
// list-level actions (Lists → drill-down → Tools). One entry here is one
// card; a future tool is one more entry, not a new tab layout.
//
// Each card composes the Panels view's card look (.wb-linkgroup —
// styles/workbench.css) rather than inventing a new one: title, one-line
// description, action button(s), and — when unavailable — a disabled action
// with the reason stated as plain text in the neutral register (never
// coloured; design/INFO-CHIP.md's failure-register split applies here too —
// this is a fact about the list, not something gone wrong).
//
// ctx (built by views/lists.js showDetail()):
//   { client, listId, listTitle, getSchema, openCopy, createClient,
//     navigate, inspectSite, mockWriter, invalidateItems }
//   getSchema()        -> Promise<{ doc }> — the SAME cached capture the
//                         Schema tab reads (lists.js's per-list tabCache),
//                         so opening both tabs costs one read, not two.
//   openCopy(doc)      -> already defined in lists.js: opens the schema
//                         apply dialog in 'copy' mode.
//   invalidateItems()  -> drops the Items tab's built pane so it rebuilds
//                         (and reloads its rows) the next time it's opened —
//                         the one hook Import data needs into a sibling tab.

import { schemaBaseType, schemaSummary } from './list-schema.js';
import { captureListData } from './list-data-capture.js';
import { normalizeDataDoc, DATA_KIND } from './list-data.js';
import { toPnpPowerShellProvisioning, toPnpjs2Provisioning } from './list-schema-script.js';
import { createMenuButton } from './grid.js';
import { copyText } from './export.js';
import { downloadText, MAX_IMPORT_BYTES } from '../io.js?v=2';
import { showFailure } from './denied.js';
import { openImportDataDialog } from './list-data-import-dialog.js';

const el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;
  return n;
};

const fileStem = (s) =>
  String(s || 'list').toLowerCase().replace(/[^a-z0-9-_]+/g, '-').replace(/^-+|-+$/g, '') || 'list';

// Same eligibility rule the old Schema-tab "Copy to…" button used: generic
// lists (base type 0, template exactly 100) and any base-type-1 template
// (document/picture/form/wiki/publishing/asset library — schema only).
// Mirrors list-schema.js buildApplyPlan's own gate.
function copyEligible(doc) {
  const bt = schemaBaseType(doc);
  return (bt === 0 && Number(doc.list.baseTemplate) === 100) || bt === 1;
}

function cardShell(id, title, description) {
  const node = el('div', `wb-linkgroup wb-tool-card wb-tool-${id}`);
  node.append(el('h3', '', title));
  node.append(el('p', 'wb-tool-desc', description));
  const body = el('div', 'wb-tool-body');
  node.append(body);
  return { node, body };
}

// A disabled action + its reason, in the neutral register (plain text, no
// colour) — the shared shape every card's "why not" state uses.
function disabledLine(reason) {
  const line = el('p', 'wb-tool-reason', reason);
  return line;
}

// ---- 1. Copy this list… ----------------------------------------------------

const copyTool = {
  id: 'copy',
  title: 'Copy this list…',
  description: 'Create a new list or library from this one’s schema, on this site or another.',
  render(card, ctx) {
    const { body } = card;
    const status = el('p', 'wb-tool-status', 'Reading the list schema…');
    body.append(status);
    ctx.getSchema().then(({ doc }) => {
      status.remove();
      const summary = schemaSummary(doc);
      if (copyEligible(doc)) {
        const btn = el('button', 'btn btn-xs wb-tools-copy', 'Copy to…');
        btn.type = 'button';
        btn.title = 'Create a new list from this schema, on this site or another one.';
        btn.addEventListener('click', () => ctx.openCopy(doc));
        body.append(btn);
      } else {
        const btn = el('button', 'btn btn-xs wb-tools-copy', 'Copy to…');
        btn.type = 'button';
        btn.disabled = true;
        body.append(btn);
        body.append(disabledLine(
          `Only generic lists and document libraries can be copied — this is a ${summary.kind.toLowerCase()}.`,
        ));
      }
    }).catch((err) => showFailure(status, err, 'this list’s schema'));
  },
};

// ---- 2. Export schema -------------------------------------------------------

const exportSchemaTool = {
  id: 'export-schema',
  title: 'Export schema',
  description: 'Download or copy this list’s schema, or generate a provisioning script.',
  render(card, ctx) {
    const { body } = card;
    const status = el('p', 'wb-tool-status', 'Reading the list schema…');
    body.append(status);
    ctx.getSchema().then(({ doc }) => {
      status.remove();
      const stem = fileStem(ctx.listTitle);
      body.append(createMenuButton('Export ▾', 'Export this list’s schema', [
        ['Download schema .json', () => downloadText(`schema-${stem}.json`, JSON.stringify(doc, null, 2), 'application/json')],
        ['Copy schema JSON', (btn) => copyText(JSON.stringify(doc, null, 2), btn)],
        ['Copy as PnP.PowerShell (provision)', (btn) => copyText(toPnpPowerShellProvisioning(doc, { targetWebUrl: ctx.client.webUrl() }), btn)],
        ['Copy as PnPjs 2 (provision)', (btn) => copyText(toPnpjs2Provisioning(doc, {}), btn)],
      ]));
    }).catch((err) => showFailure(status, err, 'this list’s schema'));
  },
};

// ---- 3. Export data ----------------------------------------------------------

const exportDataTool = {
  id: 'export-data',
  title: 'Export data',
  description: 'Read every item in this list (not just what’s on screen elsewhere) and export it as JSON.',
  render(card, ctx) {
    const { body } = card;
    const stem = fileStem(ctx.listTitle);
    const status = el('div', 'wb-grid-status wb-tool-status');
    status.hidden = true;
    let busy = false;
    async function run(useDoc) {
      if (busy) return;
      busy = true;
      status.hidden = false;
      status.classList.remove('wb-error', 'wb-denied');
      status.removeAttribute('title');
      let n = null;
      try {
        const { doc: schemaDoc } = await ctx.getSchema();
        n = schemaDoc?.source?.itemCount;
        status.textContent = n != null ? `Reading ${n} item${n === 1 ? '' : 's'}…` : 'Reading items…';
        const { doc } = await captureListData(ctx.client, ctx.listId, { schemaDoc });
        status.hidden = true;
        useDoc(doc);
      } catch (err) {
        showFailure(status, err, 'this list’s item data');
      } finally {
        busy = false;
      }
    }
    const downloadBtn = el('button', 'btn btn-xs', 'Download data .json');
    downloadBtn.type = 'button';
    downloadBtn.addEventListener('click', () =>
      run((doc) => downloadText(`data-${stem}.json`, JSON.stringify(doc, null, 2), 'application/json')));
    const copyBtn = el('button', 'btn btn-xs', 'Copy data JSON');
    copyBtn.type = 'button';
    copyBtn.addEventListener('click', () => run((doc) => copyText(JSON.stringify(doc, null, 2), copyBtn)));
    body.append(downloadBtn, copyBtn, status);
  },
};

// ---- 4. Import data into this list… -----------------------------------------

function showInlineNotice(host, message) {
  host.textContent = '';
  host.hidden = false;
  host.classList.add('wb-consent-error');
  host.append(el('span', 'wb-consent-text', message));
  const dismiss = el('button', 'btn btn-xs', 'Dismiss');
  dismiss.type = 'button';
  dismiss.addEventListener('click', () => { host.hidden = true; });
  host.append(dismiss);
}

const importDataTool = {
  id: 'import-data',
  title: 'Import data into this list',
  description: 'Add items from a data .json exported by this Workbench, or by SPUtils.',
  render(card, ctx) {
    const { body } = card;
    const status = el('p', 'wb-tool-status', 'Reading the list schema…');
    body.append(status);
    ctx.getSchema().then(({ doc: schemaDoc }) => {
      status.remove();
      const isLibrary = schemaBaseType(schemaDoc) === 1;
      const btn = el('button', 'btn btn-xs wb-tools-import', 'Choose data .json…');
      btn.type = 'button';
      const input = el('input', 'wb-tools-import-file');
      input.type = 'file';
      input.accept = '.json';
      input.hidden = true;
      const notice = el('div', 'wb-consent wb-tool-import-notice');
      notice.hidden = true;
      if (isLibrary) {
        btn.disabled = true;
        body.append(btn, disabledLine(
          'Item import into a document library isn’t supported — a library copy is schema only.',
        ));
        return;
      }
      btn.addEventListener('click', () => input.click());
      input.addEventListener('change', async () => {
        const file = input.files?.[0];
        input.value = '';
        if (!file) return;
        notice.hidden = true;
        if (file.size > MAX_IMPORT_BYTES) {
          showInlineNotice(notice, `‘${file.name}’ is ${(file.size / 1048576).toFixed(1)} MB — above the 5 MB import limit.`);
          return;
        }
        let json;
        try { json = JSON.parse(await file.text()); }
        catch { showInlineNotice(notice, `‘${file.name}’ isn’t valid JSON.`); return; }
        let dataDoc;
        try { dataDoc = normalizeDataDoc(json); }
        catch {
          showInlineNotice(notice, `‘${file.name}’ is not a list data document (expected kind ${DATA_KIND}).`);
          return;
        }
        const outcome = await openImportDataDialog({
          dataDoc, client: ctx.client, listId: ctx.listId, listTitle: ctx.listTitle, mockWriter: ctx.mockWriter,
        });
        if (outcome === 'imported') ctx.invalidateItems?.();
      });
      body.append(btn, input, notice);
    }).catch((err) => showFailure(status, err, 'this list’s schema'));
  },
};

export const LIST_TOOLS = [copyTool, exportSchemaTool, exportDataTool, importDataTool];

export function buildToolsPane(wrap, ctx) {
  const grid = el('div', 'wb-tools');
  wrap.append(grid);
  for (const tool of LIST_TOOLS) {
    const shell = cardShell(tool.id, tool.title, tool.description);
    grid.append(shell.node);
    tool.render(shell, ctx);
  }
}
