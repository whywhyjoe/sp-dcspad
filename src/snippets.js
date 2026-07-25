// Snippet library: named single-pane fragments stored as one JSON
// document. Click a snippet to insert it at the cursor of its editor;
// ＋ saves the current selection (or the whole pane when nothing is
// selected). The library can be saved to / loaded from a .json file.

import { getState, loadDoc, saveDoc, newId, SNIPPETS_KEY } from './state.js';
import { downloadText, wireJsonImport } from './io.js';
import { el } from './inspect/tree-view.js';

let doc = null;
let deps = {};

export function initSnippets({ getSelection, getDocs, insertAtCursor, selectEditorTab }) {
  deps = { getSelection, getDocs, insertAtCursor, selectEditorTab };
  doc = loadDoc(SNIPPETS_KEY) || { v: 1, items: [] };
  render();

  document.getElementById('btn-snippet-add').addEventListener('click', () => {
    const lang = getState().layout.editorTab;
    const code = deps.getSelection(lang) || deps.getDocs()[lang];
    if (!code.trim()) return;
    const name = prompt('Snippet name:');
    if (!name || !name.trim()) return;
    doc.items.push({ id: newId('snip'), name: name.trim(), lang, code, createdAt: Date.now() });
    persist();
    render();
  });

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
    document.getElementById('status-save').textContent = 'snippet save failed — storage full';
  }
}

function render() {
  const host = document.getElementById('snippet-list');
  host.textContent = '';
  document.getElementById('snippet-empty').hidden = doc.items.length > 0;

  for (const snip of doc.items) {
    const item = el('div', 'lib-item snippet-item');
    const lang = el('span', 'snippet-lang', snip.lang);
    const name = el('span', 'lib-name', snip.name);
    name.title = `Insert into the ${snip.lang.toUpperCase()} editor at the cursor\n\n${snip.code.slice(0, 400)}`;
    const del = el('span', 'lib-del', '✕');
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
