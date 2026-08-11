// Monaco adapter for the HTML / CSS / JS panes.
// One editor swaps among three persistent models so each document keeps its
// own undo history and view state without paying for three editor instances.

import { getState, update } from './state.js';
import { fetchPnpTypeLibraries, loadMonacoRuntime } from './monaco-runtime.js';
import { JSON_THEME_RULES, installJsonLanguage } from './json-language.js';
import {
  ALPINE_HTML_DATA,
  ALPINE_JS_LIBRARIES,
  ALPINE_PACK_ID,
  createAlpineHtmlCompletionProvider,
} from './intelligence/alpine.js';
import {
  BSP_PACK_ID,
  createBspCssCompletionProvider,
  createBspCssHoverProvider,
  createBspHtmlClassCompletionProvider,
  createBspHtmlClassHoverProvider,
  fetchBspIntelligence,
} from './intelligence/bsp.js';
import {
  FLUENT_ICONS_HTML_DATA,
  FLUENT_ICONS_PACK_ID,
  collectFluentIconMarkers,
  createFluentIconCompletionProvider,
  createFluentIconHoverProvider,
  fetchFluentIconIntelligence,
} from './intelligence/fluent-icons.js';

const NAMES = ['html', 'css', 'js'];
const LANGUAGES = { html: 'html', css: 'css', js: 'javascript' };
const MODEL_URIS = {
  html: 'file:///dcspad/index.html',
  css: 'file:///dcspad/styles.css',
  js: 'file:///dcspad/script.js',
};

export async function initEditors({ onChange, onRunShortcut, onTogglePane, onFontStep }) {
  const monaco = await loadMonacoRuntime();
  // JSON validation/highlighting comes from Monaco's own language service in
  // json.worker.js; configure it before any model is created so imported JSON
  // tokenizes on first paint.
  const jsonLanguage = installJsonLanguage(monaco);
  const state = getState();
  const host = document.getElementById('pane-editor');
  const cursorEl = document.getElementById('status-cursor');
  const models = {};
  const viewStates = {};
  const selections = {};
  let active = NAMES.includes(state.layout.editorTab) ? state.layout.editorTab : 'js';
  let desiredPnpTypes = false;
  let pnpTypesGeneration = 0;
  let desiredBspIntelligence = false;
  let bspIntelligenceGeneration = 0;
  let bspIntelligence = null;
  let desiredFluentIconIntelligence = false;
  let fluentIconIntelligenceGeneration = 0;
  let fluentIconIntelligence = null;
  const jsLibraryPacks = new Map();
  const htmlDataPacks = new Map();
  const enabledIntelligence = new Set();

  // Design-system syntax palette: VS Code Dark Modern hues on the --bg-editor
  // ground (#17191f). Widget surfaces mirror the app's popover tokens
  // (--bg-2 / --border-strong / --fg-row). :root CSS variables cannot reach
  // in here — keep this table and styles/app.css tokens in sync by hand.
  monaco.editor.defineTheme('dcspad-dark', {
    base: 'vs-dark',
    inherit: true,
    rules: [
      { token: 'comment', foreground: '6A9955', fontStyle: 'italic' },
      { token: 'comment.doc', foreground: '6A9955', fontStyle: 'italic' },
      { token: 'keyword', foreground: '569CD6' },
      { token: 'keyword.flow', foreground: 'C586C0' },
      { token: 'number', foreground: 'B5CEA8' },
      { token: 'string', foreground: 'CE9178' },
      { token: 'string.escape', foreground: 'D7BA7D' },
      { token: 'regexp', foreground: 'D16969' },
      { token: 'type', foreground: '4EC9B0' },
      { token: 'type.identifier', foreground: '4EC9B0' },
      { token: 'identifier', foreground: '9CDCFE' },
      { token: 'constant', foreground: '4FC1FF' },
      { token: 'tag', foreground: '569CD6' },
      { token: 'tag.css', foreground: 'D7BA7D' },
      { token: 'attribute.name', foreground: '9CDCFE' },
      { token: 'attribute.value', foreground: 'CE9178' },
      { token: 'attribute.value.number.css', foreground: 'B5CEA8' },
      { token: 'attribute.value.unit.css', foreground: 'B5CEA8' },
      { token: 'delimiter', foreground: 'D4D4D4' },
      { token: 'operator', foreground: 'D4D4D4' },
      { token: 'invalid', foreground: 'F44747' },
      ...JSON_THEME_RULES,
    ],
    colors: {
      'editor.background': '#17191f',
      'editor.foreground': '#cccccc',
      'editorGutter.background': '#14161b',
      'editorLineNumber.foreground': '#6d7484',
      'editorLineNumber.activeForeground': '#a2a9b8',
      'editor.lineHighlightBackground': '#1f232b',
      'editor.lineHighlightBorder': '#262b34',
      'editor.selectionBackground': '#264f78',
      'editor.inactiveSelectionBackground': '#264f7855',
      'editorCursor.foreground': '#aeafad',
      'editorBracketMatch.background': '#00000000',
      'editorBracketMatch.border': '#888888',
      'editorIndentGuide.background1': '#2a2e38',
      'editorIndentGuide.activeBackground1': '#3a4150',
      'editorWhitespace.foreground': '#333947',
      'editorWidget.background': '#20242c',
      'editorWidget.border': '#3a4150',
      'editorWidget.foreground': '#d4d9e2',
      'editorSuggestWidget.background': '#20242c',
      'editorSuggestWidget.border': '#3a4150',
      'editorSuggestWidget.foreground': '#d4d9e2',
      'editorSuggestWidget.selectedBackground': '#2a2f3a',
      'editorSuggestWidget.highlightForeground': '#5ee3c4',
      'editorSuggestWidget.focusHighlightForeground': '#5ee3c4',
      'editorHoverWidget.background': '#20242c',
      'editorHoverWidget.border': '#3a4150',
      'editorHoverWidget.foreground': '#d4d9e2',
      'list.hoverBackground': '#2a2f3a',
      'list.highlightForeground': '#5ee3c4',
      'input.background': '#14161b',
      'input.border': '#3a4150',
      'input.foreground': '#e6e9ef',
      'inputOption.activeBorder': '#3fd8b4',
      'editor.findMatchBackground': '#2c6a5c66',
      'editor.findMatchBorder': '#3fd8b4',
      'editor.findMatchHighlightBackground': '#1e3b35',
      'editorError.foreground': '#ff6b62',
      'editorWarning.foreground': '#e8b660',
      'editorInfo.foreground': '#67a7f7',
      // Overview-ruler marks are themed separately from the squiggles above;
      // without these they fall back to vs-dark's palette and clash.
      'editorOverviewRuler.errorForeground': '#ff6b62',
      'editorOverviewRuler.warningForeground': '#e8b660',
      'editorOverviewRuler.infoForeground': '#67a7f7',
      'editorOverviewRuler.findMatchForeground': '#3fd8b4',
      'editorOverviewRuler.bracketMatchForeground': '#888888',
      'editorOverviewRuler.background': '#14161b',
      'editorOverviewRuler.border': '#00000000',
      'editorLink.activeForeground': '#67a7f7',
      'menu.background': '#20242c',
      'menu.foreground': '#d4d9e2',
      'menu.selectionBackground': '#2a2f3a',
      'scrollbarSlider.background': '#3a415055',
      'scrollbarSlider.hoverBackground': '#3a415088',
      'scrollbarSlider.activeBackground': '#3a4150aa',
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
  const htmlDefaults = monaco.html.htmlDefaults;
  const baseHtmlOptions = htmlDefaults.options;

  function applyJsLibraries() {
    jsDefaults.setExtraLibs([...jsLibraryPacks.values()].flat());
  }

  function applyHtmlData() {
    htmlDefaults.setOptions({
      ...baseHtmlOptions,
      data: {
        useDefaultDataProvider: true,
        dataProviders: Object.fromEntries(htmlDataPacks),
      },
    });
  }

  function setAlpineIntelligenceEnabled(enabled) {
    if (enabled) {
      jsLibraryPacks.set(ALPINE_PACK_ID, ALPINE_JS_LIBRARIES);
      htmlDataPacks.set(ALPINE_PACK_ID, ALPINE_HTML_DATA);
      enabledIntelligence.add(ALPINE_PACK_ID);
      document.documentElement.dataset.alpineIntelligence = 'ready';
    } else {
      jsLibraryPacks.delete(ALPINE_PACK_ID);
      htmlDataPacks.delete(ALPINE_PACK_ID);
      enabledIntelligence.delete(ALPINE_PACK_ID);
      document.documentElement.dataset.alpineIntelligence = 'disabled';
    }
    applyJsLibraries();
    applyHtmlData();
  }

  const alpineCompletionRegistration = monaco.languages.registerCompletionItemProvider(
    'html',
    createAlpineHtmlCompletionProvider(
      monaco,
      () => enabledIntelligence.has(ALPINE_PACK_ID),
    ),
  );
  const bspRegistrations = [
    monaco.languages.registerCompletionItemProvider(
      'css',
      createBspCssCompletionProvider(monaco, () => bspIntelligence),
    ),
    monaco.languages.registerHoverProvider(
      'css',
      createBspCssHoverProvider(() => bspIntelligence),
    ),
    monaco.languages.registerCompletionItemProvider(
      'html',
      createBspHtmlClassCompletionProvider(monaco, () => bspIntelligence),
    ),
    monaco.languages.registerHoverProvider(
      'html',
      createBspHtmlClassHoverProvider(() => bspIntelligence),
    ),
  ];
  const fluentIconRegistrations = [
    monaco.languages.registerCompletionItemProvider(
      'html',
      createFluentIconCompletionProvider(monaco, () => fluentIconIntelligence),
    ),
    monaco.languages.registerHoverProvider(
      'html',
      createFluentIconHoverProvider(() => fluentIconIntelligence),
    ),
  ];

  function applyFluentIconMarkers() {
    if (!models.html) return;
    monaco.editor.setModelMarkers(
      models.html,
      FLUENT_ICONS_PACK_ID,
      collectFluentIconMarkers(
        monaco,
        models.html,
        fluentIconIntelligence,
      ),
    );
  }

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
      if (name === 'html') applyFluentIconMarkers();
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
    // Editor text size is a setting (11–18); line height locks to 1.7×.
    fontSize: state.settings.editorFontSize || 13,
    lineHeight: Math.round((state.settings.editorFontSize || 13) * 1.7),
    wordWrap: state.settings.wordWrap ? 'on' : 'off',
    lineNumbersMinChars: 3,
    minimap: { enabled: false },
    // The overview ruler is the scrollbar-strip map of markers. The option is
    // a lane count, not a boolean: 0 renders nothing at all, which hid every
    // diagnostic the language services produce. Cursor marks stay suppressed —
    // the status bar already reports the position.
    overviewRulerLanes: 3,
    overviewRulerBorder: false,
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

  // Pane toggles + font stepping must also work while Monaco has focus —
  // it swallows document-level keydown for bound chords.
  editor.addAction({
    id: 'dcspad.togglePane.resources',
    label: 'Toggle resources pane',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Backslash],
    run: () => onTogglePane?.('resources'),
  });
  editor.addAction({
    id: 'dcspad.togglePane.preview',
    label: 'Toggle preview pane',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Backslash],
    run: () => onTogglePane?.('preview'),
  });
  editor.addAction({
    id: 'dcspad.togglePane.console',
    label: 'Toggle console pane',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyJ],
    run: () => onTogglePane?.('console'),
  });
  editor.addAction({
    id: 'dcspad.fontLarger',
    label: 'Larger editor text',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Equal],
    run: () => onFontStep?.(+1),
  });
  editor.addAction({
    id: 'dcspad.fontSmaller',
    label: 'Smaller editor text',
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Minus],
    run: () => onFontStep?.(-1),
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
      jsLibraryPacks.delete('pnpjs-2.15.0');
      applyJsLibraries();
      document.documentElement.dataset.pnpTypes = 'disabled';
      return;
    }
    document.documentElement.dataset.pnpTypes = 'loading';
    try {
      const libs = await fetchPnpTypeLibraries();
      if (!desiredPnpTypes || generation !== pnpTypesGeneration) return;
      jsLibraryPacks.set('pnpjs-2.15.0', libs);
      applyJsLibraries();
      document.documentElement.dataset.pnpTypes = 'ready';
    } catch (error) {
      if (generation !== pnpTypesGeneration) return;
      document.documentElement.dataset.pnpTypes = 'error';
      console.warn('DCSPad: PnPjs IntelliSense could not be loaded', error);
    }
  }

  async function setBspIntelligenceEnabled(enabled) {
    desiredBspIntelligence = !!enabled;
    const generation = ++bspIntelligenceGeneration;
    if (!desiredBspIntelligence) {
      bspIntelligence = null;
      enabledIntelligence.delete(BSP_PACK_ID);
      document.documentElement.dataset.bspIntelligence = 'disabled';
      return;
    }
    enabledIntelligence.add(BSP_PACK_ID);
    document.documentElement.dataset.bspIntelligence = 'loading';
    try {
      const data = await fetchBspIntelligence();
      if (!desiredBspIntelligence || generation !== bspIntelligenceGeneration) return;
      bspIntelligence = data;
      document.documentElement.dataset.bspIntelligence = 'ready';
    } catch (error) {
      if (generation !== bspIntelligenceGeneration) return;
      bspIntelligence = null;
      enabledIntelligence.delete(BSP_PACK_ID);
      document.documentElement.dataset.bspIntelligence = 'error';
      console.warn('DCSPad: BMO design-system intelligence could not be loaded', error);
    }
  }

  async function setFluentIconIntelligenceEnabled(enabled) {
    desiredFluentIconIntelligence = !!enabled;
    const generation = ++fluentIconIntelligenceGeneration;
    if (!desiredFluentIconIntelligence) {
      fluentIconIntelligence = null;
      htmlDataPacks.delete(FLUENT_ICONS_PACK_ID);
      enabledIntelligence.delete(FLUENT_ICONS_PACK_ID);
      applyHtmlData();
      applyFluentIconMarkers();
      document.documentElement.dataset.fluentIconIntelligence = 'disabled';
      return;
    }

    enabledIntelligence.add(FLUENT_ICONS_PACK_ID);
    htmlDataPacks.set(FLUENT_ICONS_PACK_ID, FLUENT_ICONS_HTML_DATA);
    applyHtmlData();
    document.documentElement.dataset.fluentIconIntelligence = 'loading';
    try {
      const data = await fetchFluentIconIntelligence();
      if (!desiredFluentIconIntelligence
          || generation !== fluentIconIntelligenceGeneration) return;
      fluentIconIntelligence = data;
      applyFluentIconMarkers();
      document.documentElement.dataset.fluentIconIntelligence = 'ready';
    } catch (error) {
      if (generation !== fluentIconIntelligenceGeneration) return;
      fluentIconIntelligence = null;
      enabledIntelligence.delete(FLUENT_ICONS_PACK_ID);
      applyFluentIconMarkers();
      document.documentElement.dataset.fluentIconIntelligence = 'error';
      console.warn('DCSPad: Fluent icon intelligence could not be loaded', error);
    }
  }

  function setIntelligencePacks(packIds) {
    const requested = new Set(packIds || []);
    setAlpineIntelligenceEnabled(requested.has(ALPINE_PACK_ID));
    setPnpTypesEnabled(requested.has('pnpjs-2.15.0'));
    setBspIntelligenceEnabled(requested.has(BSP_PACK_ID));
    setFluentIconIntelligenceEnabled(requested.has(FLUENT_ICONS_PACK_ID));
  }

  reportCursor();

  return {
    activate,
    setFontSize: (px) => {
      editor.updateOptions({ fontSize: px, lineHeight: Math.round(px * 1.7) });
    },
    setWordWrap: (on) => {
      editor.updateOptions({ wordWrap: on ? 'on' : 'off' });
    },
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
    setIntelligencePacks,
    setPnpTypesEnabled,
    dispose: () => {
      resizeObserver.disconnect();
      jsonLanguage.dispose();
      alpineCompletionRegistration.dispose();
      for (const registration of bspRegistrations) registration.dispose();
      for (const registration of fluentIconRegistrations) registration.dispose();
      monaco.editor.setModelMarkers(models.html, FLUENT_ICONS_PACK_ID, []);
      editor.dispose();
      for (const model of Object.values(models)) model.dispose();
    },
  };
}
