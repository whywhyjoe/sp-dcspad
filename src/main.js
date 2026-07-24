// DCSPad bootstrap — wires modules together. Grows as milestones land.

import { getState, onSaveStatus } from './state.js';
import { initLayout } from './layout.js';
import { initEditors } from './editors.js';

const layoutApi = initLayout({
  onEditorTabChange: (name) => editorsApi.focus(name),
});

const editorsApi = initEditors({
  onChange: () => {},           // autorun hooks in later
  onRunShortcut: () => run(),
});

function run() {
  // Runner lands in the next milestone.
  console.info('DCSPad: run requested', editorsApi.getDocs());
}
document.getElementById('btn-run').addEventListener('click', run);

// Autosave tick in the status bar.
const saveEl = document.getElementById('status-save');
onSaveStatus((status) => {
  if (status === 'dirty') {
    saveEl.textContent = 'saving…';
    saveEl.classList.remove('saved');
  } else {
    saveEl.textContent = '✓ saved';
    saveEl.classList.add('saved');
  }
});

export { layoutApi, editorsApi };
