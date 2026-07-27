// Layout: splitter dragging, pane visibility (topbar segmented toggles),
// tab switching, editor/preview/diagnostics maximize. Positions persist
// via state.layout.

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
  selectEditorTab(layout.editorTab, { silent: true });
  selectDiagTab(layout.diagTab);

  // ----- pane visibility (topbar segmented toggles) -----
  // Seed from the legacy collapse flags the first time a pre-panes
  // workspace loads, then retire them so a later re-open can't be undone
  // by a stale flag on the next boot.
  const panes = { ...layout.panes };
  if (layout.sidebarCollapsed) panes.resources = false;
  if (layout.diagCollapsed) panes.console = false;
  if (layout.sidebarCollapsed || layout.diagCollapsed) {
    updateNested('layout', { panes: { ...panes }, sidebarCollapsed: false, diagCollapsed: false });
  }

  function applyPanes() {
    main.classList.toggle('hide-resources', !panes.resources);
    main.classList.toggle('hide-preview', !panes.preview);
    main.classList.toggle('hide-console', !panes.console);
    for (const name of ['resources', 'preview', 'console']) {
      const seg = document.getElementById(`seg-${name}`);
      seg.classList.toggle('active', !!panes[name]);
      seg.setAttribute('aria-pressed', String(!!panes[name]));
    }
  }
  applyPanes();

  function setPaneVisible(name, on) {
    panes[name] = !!on;
    applyPanes();
    updateNested('layout', { panes: { ...panes } });
  }
  function togglePane(name) { setPaneVisible(name, !panes[name]); }

  document.getElementById('pane-toggles').addEventListener('click', (e) => {
    const seg = e.target.closest('.pane-seg');
    if (seg) togglePane(seg.dataset.pane);
  });

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

  // ----- sidebar vertical split (Frameworks / Snippets) -----
  const DEFAULT_SNIPPETS_H = 210;
  root.style.setProperty('--snippets-h', px(layout.snippetsPanelH || DEFAULT_SNIPPETS_H));
  const splitSide = document.getElementById('split-side');
  dragSplitter(splitSide, 'y', (dy, start) => {
    // Dragging down grows Frameworks / shrinks Snippets. Clamp so neither
    // panel can collapse (Frameworks list ≥ 120px + chrome, Snippets ≥ 150px).
    const max = start.sidebarH - 160;
    const h = Math.min(max, Math.max(150, start.snippetsH - dy));
    root.style.setProperty('--snippets-h', px(h));
    updateNested('layout', { snippetsPanelH: h });
  }, () => ({
    snippetsH: document.getElementById('panel-snippets').getBoundingClientRect().height,
    sidebarH: document.getElementById('sidebar').getBoundingClientRect().height,
  }));
  splitSide.addEventListener('dblclick', () => {
    root.style.setProperty('--snippets-h', px(DEFAULT_SNIPPETS_H));
    updateNested('layout', { snippetsPanelH: DEFAULT_SNIPPETS_H });
  });

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
    if (tab) selectDiagTab(tab.dataset.diag);
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
    main.classList.remove('max-diag', 'max-editor');
    main.classList.toggle('max-preview');
  });
  document.getElementById('btn-max-diag').addEventListener('click', () => {
    main.classList.remove('max-preview', 'max-editor');
    main.classList.toggle('max-diag');
  });
  document.getElementById('btn-max-editor').addEventListener('click', () => {
    main.classList.remove('max-preview', 'max-diag');
    main.classList.toggle('max-editor');
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { main.classList.remove('max-preview', 'max-diag', 'max-editor'); return; }
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    // Monaco registers the same chords as editor actions — this listener
    // covers everywhere else.
    if (e.code === 'Backslash') {
      e.preventDefault();
      togglePane(e.shiftKey ? 'preview' : 'resources');
    } else if (e.code === 'KeyJ' && !e.shiftKey) {
      e.preventDefault();
      togglePane('console');
    }
  });

  return { selectEditorTab, selectDiagTab, togglePane, setPaneVisible };
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
