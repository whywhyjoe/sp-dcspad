// Snippet library: named single-pane fragments stored as one JSON
// document. Click a snippet to insert it at the cursor of its editor;
// ＋ saves the current selection (or the whole pane when nothing is
// selected). The library can be saved to / loaded from a .json file.

import { getState, loadDoc, saveDoc, newId, SNIPPETS_KEY } from './state.js';
import { downloadText, wireJsonImport } from './io.js';
import { el } from './inspect/tree-view.js';
import {
  SNIPPET_LIBRARY_KIND,
  validateSnippetLibrary,
} from './library-files.js';

let doc = null;
let deps = {};
const snippetNameCollator = new Intl.Collator(undefined, {
  sensitivity: 'base',
  numeric: true,
});

const defaultSnippetLibrary = () => ({
  kind: SNIPPET_LIBRARY_KIND,
  v: 1,
  items: [],
});

function starterLibraryUrl() {
  const appRoot = window.__DCSPAD_ASSET_BASE__
    || new URL('../', import.meta.url).href;
  return new URL('examples/dcspad-starter-snippets.json', appRoot).href;
}

async function loadDefaultSnippetLibrary() {
  const response = await fetch(starterLibraryUrl(), {
    credentials: 'same-origin',
    cache: 'no-cache',
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} loading the starter snippet library`);
  }
  const validation = validateSnippetLibrary(await response.json());
  if (!validation.ok) throw new Error(validation.message);
  return validation.doc;
}

export async function initSnippets({ getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError }) {
  deps = { getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError };
  const storedDoc = loadDoc(SNIPPETS_KEY);
  const storedValidation = storedDoc
    ? validateSnippetLibrary(storedDoc, { allowUnsignedEmpty: true })
    : null;
  if (storedValidation?.ok) {
    doc = storedValidation.doc;
  } else {
    if (storedDoc) {
      console.warn('DCSPad: invalid stored snippet library was reset', storedValidation.message);
    }
    try {
      doc = await loadDefaultSnippetLibrary();
    } catch (error) {
      console.warn('DCSPad: starter snippet library could not be loaded', error);
      doc = defaultSnippetLibrary();
    }
    saveDoc(SNIPPETS_KEY, doc);
  }
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
  wireJsonImport('import-snippets-file', (imported, fileName) => {
    const validation = validateSnippetLibrary(imported);
    if (!validation.ok) {
      alert(`"${fileName}" was not imported.\n\n${validation.message}`);
      return;
    }
    const items = validation.doc.items;
    if (doc.items.length && !confirm(`Replace your ${doc.items.length} snippet(s) with the ${items.length} from this file?`)) return;
    doc = {
      ...validation.doc,
      items: items.map((s) => ({ ...s, id: s.id || newId('snip') })),
    };
    persist();
    render();
  });

  document.getElementById('btn-snippets-reset').addEventListener('click', async () => {
    if (!confirm(
      `Reset Snippets to the built-in starter library?\n\nThis will replace all ${doc.items.length} currently saved snippet(s).`,
    )) return;
    const resetButton = document.getElementById('btn-snippets-reset');
    resetButton.disabled = true;
    try {
      doc = await loadDefaultSnippetLibrary();
      persist();
      render();
    } catch (error) {
      alert(`The starter snippet library could not be loaded.\n\n${error.message || error}`);
    } finally {
      resetButton.disabled = false;
    }
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
