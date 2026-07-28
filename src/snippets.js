// Snippet library: named single-pane fragments stored as one JSON
// document. Click a snippet to insert it at the cursor of its editor;
// ＋ saves the current selection (or the whole pane when nothing is
// selected). The library can be saved to / loaded from a .json file.

import { getState, loadDoc, saveDoc, newId, SNIPPETS_KEY } from './state.js';
import { downloadText, wireJsonImport } from './io.js';
import { el } from './inspect/tree-view.js';

let doc = null;
let deps = {};
const snippetNameCollator = new Intl.Collator(undefined, {
  sensitivity: 'base',
  numeric: true,
});

export function initSnippets({ getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError }) {
  deps = { getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError };
  doc = loadDoc(SNIPPETS_KEY) || { v: 1, items: [] };
  render();

  const dialog = document.getElementById('snippet-name-dialog');
  const form = document.getElementById('snippet-name-form');
  const input = document.getElementById('snippet-name-input');
  const context = document.getElementById('snippet-name-context');
  let pendingSnippet = null;

  const closeNamingDialog = () => {
    pendingSnippet = null;
    if (dialog.open) dialog.close();
  };

  document.getElementById('btn-snippet-add').addEventListener('click', () => {
    const lang = getState().layout.editorTab;
    const selection = deps.getSelection(lang);
    const code = selection || deps.getDocs()[lang];
    if (!code.trim()) return;

    pendingSnippet = { lang, code };
    input.value = '';
    context.textContent = `Save ${selection ? 'the selected' : 'all'} ${lang.toUpperCase()} code as a reusable snippet.`;
    if (!dialog.open) dialog.showModal();
    requestAnimationFrame(() => input.focus());
  });

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const name = input.value.trim();
    if (!pendingSnippet || !name) return;
    const { lang, code } = pendingSnippet;
    doc.items.push({ id: newId('snip'), name, lang, code, createdAt: Date.now() });
    persist();
    render();
    pendingSnippet = null;
    dialog.close();
  });

  document.getElementById('snippet-name-cancel').addEventListener('click', closeNamingDialog);
  document.getElementById('snippet-name-close').addEventListener('click', closeNamingDialog);
  dialog.addEventListener('cancel', () => { pendingSnippet = null; });

  document.getElementById('btn-snippets-export').addEventListener('click', () => {
    downloadText('dcspad-snippets.json', JSON.stringify(doc, null, 2));
  });
  document.getElementById('btn-snippets-import').addEventListener('click', () => {
    document.getElementById('import-snippets-file').click();
  });
  wireJsonImport('import-snippets-file', (imported) => {
    const items = imported && Array.isArray(imported.items)
      ? imported.items.filter((s) => s && typeof s.name === 'string' && typeof s.code === 'string'
          && ['html', 'css', 'js'].includes(s.lang))
      : null;
    if (!items) { alert('Not a DCSPad snippet library file.'); return; }
    if (doc.items.length && !confirm(`Replace your ${doc.items.length} snippet(s) with the ${items.length} from this file?`)) return;
    doc = { v: 1, items: items.map((s) => ({ ...s, id: s.id || newId('snip') })) };
    persist();
    render();
  });
}

function persist() {
  if (!saveDoc(SNIPPETS_KEY, doc)) {
    deps.onStorageError?.('snippet library save failed (storage full?)');
  }
}

const DEL_ICON = '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path d="m4 4 8 8M12 4l-8 8"/></svg>';

function render() {
  const host = document.getElementById('snippet-list');
  host.textContent = '';
  document.getElementById('snippet-empty').hidden = doc.items.length > 0;
  const countEl = document.getElementById('snippets-count');
  if (countEl) countEl.textContent = String(doc.items.length);

  // Display order is deliberately independent of language and save/import
  // order. Keep the stored document untouched so exports remain lossless.
  const sortedItems = [...doc.items].sort((a, b) =>
    snippetNameCollator.compare(a.name, b.name) || a.id.localeCompare(b.id));

  for (const snip of sortedItems) {
    const item = el('div', 'lib-item snippet-item');
    const lang = el('span', 'snippet-lang', snip.lang);
    lang.dataset.lang = snip.lang;
    const name = el('span', 'lib-name', snip.name);
    name.title = `Insert into the ${snip.lang.toUpperCase()} editor at the cursor\n\n${snip.code.slice(0, 400)}`;
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'lib-del';
    del.innerHTML = DEL_ICON;
    del.title = 'Delete snippet';
    del.addEventListener('click', (e) => {
      e.stopPropagation();
      if (!confirm(`Delete snippet "${snip.name}"?`)) return;
      doc.items = doc.items.filter((s) => s.id !== snip.id);
      persist();
      render();
    });
    item.addEventListener('click', () => {
      deps.selectEditorTab(snip.lang);
      deps.insertAtCursor(snip.lang, snip.code);
    });
    item.append(lang, name, del);
    host.append(item);
  }
}
