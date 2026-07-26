// Layout: splitter dragging, sidebar collapse, tab switching,
// preview/diagnostics maximize. Positions persist via state.layout.

import { getState, updateNested } from './state.js';

const px = (n) => `${n}px`;

export function initLayout({ onEditorTabChange } = {}) {
  const main = document.getElementById('main');
  const root = document.documentElement;
  const layout = getState().layout;

  // ----- restore -----
  root.style.setProperty('--sidebar-w', px(layout.sidebarW));
  root.style.setProperty('--editors-w', `${layout.editorsFr}fr`);
  root.style.setProperty('--runtime-w', `${layout.runtimeFr}fr`);
  root.style.setProperty('--preview-h', `${layout.previewFr}fr`);
  root.style.setProperty('--diag-h', px(layout.diagH));
  if (layout.sidebarCollapsed) collapseSidebar(true);
  if (layout.diagCollapsed) collapseDiag(true);
  selectEditorTab(layout.editorTab, { silent: true });
  selectDiagTab(layout.diagTab);

  // ----- splitters -----
  dragSplitter(document.getElementById('split-sidebar'), 'x', (dx, start) => {
    const w = Math.min(420, Math.max(140, start.sidebarW + dx));
    root.style.setProperty('--sidebar-w', px(w));
    updateNested('layout', { sidebarW: w });
  }, () => ({ sidebarW: parseFloat(getComputedStyle(root).getPropertyValue('--sidebar-w')) }));

  dragSplitter(document.getElementById('split-center'), 'x', (dx, start) => {
    // Reapportion the two fr columns based on pixel delta.
    const total = start.editorsPx + start.runtimePx;
    const editorsPx = Math.min(total - 260, Math.max(200, start.editorsPx + dx));
    const fr = editorsPx / (total - editorsPx);
    root.style.setProperty('--editors-w', `${fr}fr`);
    root.style.setProperty('--runtime-w', `1fr`);
    updateNested('layout', { editorsFr: fr, runtimeFr: 1 });
  }, () => ({
    editorsPx: document.getElementById('editors').getBoundingClientRect().width,
    runtimePx: document.getElementById('runtime').getBoundingClientRect().width,
  }));

  dragSplitter(document.getElementById('split-runtime'), 'y', (dy, start) => {
    const h = Math.min(start.runtimeH - 80, Math.max(100, start.diagH - dy));
    root.style.setProperty('--diag-h', px(h));
    updateNested('layout', { diagH: h });
  }, () => ({
    diagH: document.getElementById('diag-panel').getBoundingClientRect().height,
    runtimeH: document.getElementById('runtime').getBoundingClientRect().height,
  }));

  // ----- sidebar collapse -----
  document.getElementById('btn-collapse-sidebar').addEventListener('click', () => collapseSidebar(true));
  document.getElementById('btn-expand-sidebar').addEventListener('click', () => collapseSidebar(false));

  function collapseSidebar(collapsed) {
    main.classList.toggle('sidebar-collapsed', collapsed);
    document.getElementById('btn-expand-sidebar').hidden = !collapsed;
    updateNested('layout', { sidebarCollapsed: collapsed });
  }

  // ----- diagnostics collapse (tabs bar stays; clicking a tab reopens) -----
  document.getElementById('btn-collapse-diag').addEventListener('click', () => collapseDiag(true));

  function collapseDiag(collapsed) {
    main.classList.toggle('diag-collapsed', collapsed);
    updateNested('layout', { diagCollapsed: collapsed });
  }

  // ----- editor tabs -----
  document.getElementById('editor-tabs').addEventListener('click', (e) => {
    const tab = e.target.closest('.tab');
    if (tab) selectEditorTab(tab.dataset.editor);
  });

  function selectEditorTab(name, { silent } = {}) {
    for (const t of document.querySelectorAll('#editor-tabs .tab'))
      t.classList.toggle('active', t.dataset.editor === name);
    updateNested('layout', { editorTab: name });
    if (!silent) onEditorTabChange?.(name);
  }

  // ----- diagnostics tabs -----
  document.getElementById('diag-tabs').addEventListener('click', (e) => {
    const tab = e.target.closest('.tab');
    if (!tab) return;
    if (main.classList.contains('diag-collapsed')) collapseDiag(false);
    selectDiagTab(tab.dataset.diag);
  });

  // Note: the tab dots (#console-badge / #network-badge) are error
  // indicators owned by the panels — they clear with the output, not on
  // tab focus.
  function selectDiagTab(name) {
    for (const t of document.querySelectorAll('#diag-tabs .tab'))
      t.classList.toggle('active', t.dataset.diag === name);
    for (const v of document.querySelectorAll('.diag-view'))
      v.classList.toggle('active', v.id === `view-${name}`);
    document.getElementById('console-tools').hidden = name !== 'console';
    document.getElementById('network-tools').hidden = name !== 'network';
    updateNested('layout', { diagTab: name });
  }

  // ----- maximize toggles -----
  document.getElementById('btn-max-preview').addEventListener('click', () => {
    main.classList.remove('max-diag');
    main.classList.toggle('max-preview');
  });
  document.getElementById('btn-max-diag').addEventListener('click', () => {
    if (main.classList.contains('diag-collapsed')) collapseDiag(false);
    main.classList.remove('max-preview');
    main.classList.toggle('max-diag');
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') main.classList.remove('max-preview', 'max-diag');
  });

  return { selectEditorTab, selectDiagTab };
}

function dragSplitter(el, axis, onMove, getStart) {
  el.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    el.setPointerCapture(e.pointerId);
    el.classList.add('dragging');
    const origin = axis === 'x' ? e.clientX : e.clientY;
    const start = getStart();
    const move = (ev) => {
      const delta = (axis === 'x' ? ev.clientX : ev.clientY) - origin;
      onMove(delta, start);
    };
    const up = () => {
      el.classList.remove('dragging');
      el.removeEventListener('pointermove', move);
      el.removeEventListener('pointerup', up);
    };
    el.addEventListener('pointermove', move);
    el.addEventListener('pointerup', up);
  });
}
