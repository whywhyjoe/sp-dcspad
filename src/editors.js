// CodeMirror 6 editors for the HTML / CSS / JS panes.
// All CM6 imports come from the single vendored bundle (see tools/build-vendor.mjs).

import {
  EditorView, EditorState, basicSetup, keymap, indentWithTab,
  html, css, javascript, oneDark,
} from '../vendor/codemirror.js';
import { getState, updateNested, update } from './state.js';

const LANGS = { html: html, css: css, js: javascript };

export function initEditors({ onChange, onRunShortcut }) {
  const state = getState();
  const views = {};

  const runKeymap = keymap.of([{
    key: 'Mod-Enter',
    preventDefault: true,
    run: () => { onRunShortcut?.(); return true; },
  }]);

  for (const name of ['html', 'css', 'js']) {
    views[name] = new EditorView({
      parent: document.getElementById(`pane-${name}`),
      state: EditorState.create({
        doc: state[name],
        extensions: [
          basicSetup,
          oneDark,
          LANGS[name](),
          keymap.of([indentWithTab]),
          runKeymap,
          EditorView.updateListener.of((u) => {
            if (u.docChanged) {
              update({ [name]: u.state.doc.toString() });
              onChange?.(name);
            }
            if (u.selectionSet || u.docChanged) reportCursor(views[name]);
          }),
        ],
      }),
    });
  }

  const cursorEl = document.getElementById('status-cursor');
  function reportCursor(view) {
    const pos = view.state.selection.main.head;
    const line = view.state.doc.lineAt(pos);
    cursorEl.textContent = `Ln ${line.number}, Col ${pos - line.from + 1}`;
  }

  return {
    getDocs: () => ({
      html: views.html.state.doc.toString(),
      css: views.css.state.doc.toString(),
      js: views.js.state.doc.toString(),
    }),
    focus: (name) => views[name]?.focus(),
    // Replace pane contents wholesale (project load). Goes through the
    // normal dispatch path so autosave and autorun behave as if typed.
    setDocs: (docs) => {
      for (const name of ['html', 'css', 'js']) {
        if (typeof docs[name] !== 'string') continue;
        const view = views[name];
        view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: docs[name] } });
      }
    },
    // Selected text in a pane ('' when the selection is empty).
    getSelection: (name) => {
      const sel = views[name].state.selection.main;
      return views[name].state.sliceDoc(sel.from, sel.to);
    },
    insertAtCursor: (name, text) => {
      const view = views[name];
      view.dispatch(view.state.replaceSelection(text));
      view.focus();
    },
    // Jump the JS editor to a 1-based line (used by clickable stack frames).
    gotoJsLine: (lineNo) => {
      const view = views.js;
      const line = view.state.doc.line(Math.min(lineNo, view.state.doc.lines));
      view.dispatch({
        selection: { anchor: line.from },
        effects: EditorView.scrollIntoView(line.from, { y: 'center' }),
      });
      view.focus();
    },
  };
}
