// DCSPad bootstrap — wires modules together. Grows as milestones land.

import { getState, onSaveStatus } from './state.js';
import { initLayout } from './layout.js';

const layoutApi = initLayout();

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

export { layoutApi };
