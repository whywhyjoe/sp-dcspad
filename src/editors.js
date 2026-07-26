// Monaco adapter for the HTML / CSS / JS panes.
// One editor swaps among three persistent models so each document keeps its
// own undo history and view state without paying for three editor instances.

import { getState, update } from './state.js';
import { fetchPnpTypeLibraries, loadMonacoRuntime } from './monaco-runtime.js';

const NAMES = ['html', 'css', 'js'];
const LANGUAGES = { html: 'html', css: 'css', js: 'javascript' };
const MODEL_URIS = {
  html: 'file:///dcspad/index.html',
  css: 'file:///dcspad/styles.css',
  js: 'file:///dcspad/script.js',
};

export async function initEditors({ onChange, onRunShortcut }) {
  const monaco = await loadMonacoRuntime();
  const state = getState();
  const host = document.getElementById('pane-editor');
  const cursorEl = document.getElementById('status-cursor');
  const models = {};
  const viewStates = {};
  const selections = {};
  let active = NAMES.includes(state.layout.editorTab) ? state.layout.editorTab : 'js';
  let desiredPnpTypes = false;
  let pnpTypesGeneration = 0;

  monaco.editor.defineTheme('dcspad-dark', {
    base: 'vs-dark',
    inherit: true,
    rules: [
      { token: 'comment', foreground: '6A9955' },
      { token: 'keyword', foreground: 'C586C0' },
      { token: 'number', foreground: 'B5CEA8' },
      { token: 'string', foreground: 'CE9178' },
      { token: 'type', foreground: '4EC9B0' },
      { token: 'type.identifier', foreground: '4EC9B0' },
      { token: 'identifier', foreground: 'DCDCAA' },
      { token: 'tag', foreground: '569CD6' },
      { token: 'attribute.name', foreground: '9CDCFE' },
    ],
    colors: {
      // Preserve the prior One Dark editor ground while Monaco replaces its
      // rendering and language-service layers.
      'editor.background': '#282c34',
      'editor.foreground': '#d6d9e0',
      'editorGutter.background': '#282c34',
      'editorLineNumber.foreground': '#5c6270',
      'editorLineNumber.activeForeground': '#b8bdc9',
      'editor.lineHighlightBackground': '#2c313a',
      'editor.selectionBackground': '#264f78',
      'editor.inactiveSelectionBackground': '#264f7855',
      'editorCursor.foreground': '#4ec9b0',
      'editorIndentGuide.background1': '#33374255',
      'editorIndentGuide.activeBackground1': '#4b5263',
      'editorSuggestWidget.background': '#23262e',
      'editorSuggestWidget.border': '#3c4150',
      'editorSuggestWidget.selectedBackground': '#2b4058',
      'editorHoverWidget.background': '#23262e',
      'editorHoverWidget.border': '#3c4150',
      'editorWidget.background': '#23262e',
      'editorWidget.border': '#3c4150',
    },
  });

  const jsDefaults = monaco.typescript.javascriptDefaults;
  function setJsAsModule(enabled) {
    jsDefaults.setCompilerOptions({
      allowJs: true,
      checkJs: true,
      allowNonTsExtensions: true,
      target: monaco.typescript.ScriptTarget.ES2022,
      module: monaco.typescript.ModuleKind.ESNext,
      moduleResolution: monaco.typescript.ModuleResolutionKind.NodeJs,
      // TypeScript enum: Auto = 2, Force = 3. Monaco does not currently
      // re-export ModuleDetectionKind from its standalone contribution.
      moduleDetection: enabled ? 3 : 2,
      skipLibCheck: true,
    });
  }
  setJsAsModule(state.settings.jsAsModule);
  jsDefaults.setDiagnosticsOptions({
    noSyntaxValidation: false,
    noSemanticValidation: false,
    noSuggestionDiagnostics: false,
  });
  jsDefaults.setEagerModelSync(true);

  for (const name of NAMES) {
    models[name] = monaco.editor.createModel(
      state[name],
      LANGUAGES[name],
      monaco.Uri.parse(MODEL_URIS[name]),
    );
    models[name].updateOptions({ tabSize: 2, insertSpaces: true });
    selections[name] = new monaco.Selection(1, 1, 1, 1);
    models[name].onDidChangeContent(() => {
      update({ [name]: models[name].getValue() });
      onChange?.(name);
      if (name === active) reportCursor();
    });
  }

  const editor = monaco.editor.create(host, {
    model: models[active],
    theme: 'dcspad-dark',
    ariaLabel: `${active.toUpperCase()} code editor`,
    automaticLayout: true,
    fixedOverflowWidgets: true,
    fontFamily: '"Cascadia Code", "Consolas", "SF Mono", Menlo, monospace',
    fontSize: 13,
    lineHeight: 20,
    lineNumbersMinChars: 3,
    minimap: { enabled: false },
    overviewRulerLanes: 0,
    hideCursorInOverviewRuler: true,
    scrollBeyondLastLine: false,
    stickyScroll: { enabled: false },
    renderWhitespace: 'selection',
    quickSuggestions: { other: true, comments: false, strings: false },
    suggestOnTriggerCharacters: true,
    parameterHints: { enabled: true },
    folding: true,
    padding: { top: 6, bottom: 6 },
  });

  editor.addAction({
    id: 'dcspad.run',
    label: 'Run DCSPad',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter],
    run: () => onRunShortcut?.(),
  });

  editor.onDidChangeCursorPosition(() => {
    selections[active] = editor.getSelection() || selections[active];
    reportCursor();
  });
  editor.onDidChangeCursorSelection(() => {
    selections[active] = editor.getSelection() || selections[active];
  });

  const resizeObserver = new ResizeObserver(() => editor.layout());
  resizeObserver.observe(host);

  function reportCursor() {
    const pos = editor.getPosition();
    if (pos) cursorEl.textContent = `Ln ${pos.lineNumber}, Col ${pos.column}`;
  }

  function activate(name, { focus = true } = {}) {
    if (!models[name]) return;
    if (name !== active) {
      viewStates[active] = editor.saveViewState();
      selections[active] = editor.getSelection() || selections[active];
      active = name;
      editor.setModel(models[name]);
      editor.updateOptions({ ariaLabel: `${name.toUpperCase()} code editor` });
      if (viewStates[name]) editor.restoreViewState(viewStates[name]);
      else editor.setSelection(selections[name]);
    }
    editor.layout();
    reportCursor();
    if (focus) editor.focus();
  }

  async function setPnpTypesEnabled(enabled) {
    desiredPnpTypes = !!enabled;
    const generation = ++pnpTypesGeneration;
    if (!desiredPnpTypes) {
      jsDefaults.setExtraLibs([]);
      document.documentElement.dataset.pnpTypes = 'disabled';
      return;
    }
    document.documentElement.dataset.pnpTypes = 'loading';
    try {
      const libs = await fetchPnpTypeLibraries();
      if (!desiredPnpTypes || generation !== pnpTypesGeneration) return;
      jsDefaults.setExtraLibs(libs);
      document.documentElement.dataset.pnpTypes = 'ready';
    } catch (error) {
      if (generation !== pnpTypesGeneration) return;
      document.documentElement.dataset.pnpTypes = 'error';
      console.warn('DCSPad: PnPjs IntelliSense could not be loaded', error);
    }
  }

  reportCursor();

  return {
    activate,
    getDocs: () => ({
      html: models.html.getValue(),
      css: models.css.getValue(),
      js: models.js.getValue(),
    }),
    focus: (name) => activate(name),
    setDocs: (docs) => {
      for (const name of NAMES) {
        if (typeof docs[name] !== 'string') continue;
        const model = models[name];
        model.pushStackElement();
        model.pushEditOperations([], [{
          range: model.getFullModelRange(),
          text: docs[name],
          forceMoveMarkers: true,
        }], () => null);
        model.pushStackElement();
      }
    },
    getSelection: (name) => {
      const selection = name === active
        ? editor.getSelection()
        : selections[name];
      return selection ? models[name].getValueInRange(selection) : '';
    },
    insertAtCursor: (name, text) => {
      activate(name, { focus: false });
      const model = models[name];
      const selection = editor.getSelection() || selections[name];
      model.pushStackElement();
      editor.executeEdits('dcspad-snippet', [{
        range: selection,
        text,
        forceMoveMarkers: true,
      }]);
      model.pushStackElement();
      editor.focus();
    },
    gotoJsLine: (lineNo) => {
      activate('js', { focus: false });
      const line = Math.max(1, Math.min(lineNo, models.js.getLineCount()));
      editor.setPosition({ lineNumber: line, column: 1 });
      editor.revealLineInCenter(line);
      editor.focus();
    },
    setJsAsModule,
    setPnpTypesEnabled,
    dispose: () => {
      resizeObserver.disconnect();
      editor.dispose();
      for (const model of Object.values(models)) model.dispose();
    },
  };
}
