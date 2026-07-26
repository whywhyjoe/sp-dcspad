// ../src/state.js
var STORAGE_KEY = "dcspad.v2.workspace";
var CATALOG_KEY = "dcspad.v2.catalog";
var SNIPPETS_KEY = "dcspad.v2.snippets";
var SAVE_DEBOUNCE_MS = 600;
var DEFAULTS = {
  html: '<div id="app">\n  <h2>Hello from DCSPad</h2>\n  <p>Edit HTML, CSS and JS, then press Run.</p>\n</div>\n',
  css: 'body {\n  font-family: "Segoe UI", sans-serif;\n  padding: 1rem;\n}\n',
  js: 'console.log("DCSPad ready", { when: new Date().toISOString() });\n',
  libraries: { enabled: [], pinned: ["pnpjs2"], custom: [] },
  settings: { autorun: false, jsAsModule: false, autoClearConsole: true, seenSplash: false, previewDark: true, diagFontSize: 12 },
  layout: {
    sidebarW: 230,
    sidebarCollapsed: false,
    editorsFr: 1,
    runtimeFr: 1,
    previewFr: 1,
    diagH: 260,
    diagCollapsed: false,
    editorTab: "js",
    diagTab: "console"
  }
};
var state = load();
var saveTimer = null;
var listeners = /* @__PURE__ */ new Set();
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return structuredClone(DEFAULTS);
    const parsed = JSON.parse(raw);
    return {
      ...structuredClone(DEFAULTS),
      ...parsed,
      libraries: { ...structuredClone(DEFAULTS.libraries), ...parsed.libraries || {} },
      settings: { ...structuredClone(DEFAULTS.settings), ...parsed.settings || {} },
      layout: { ...structuredClone(DEFAULTS.layout), ...parsed.layout || {} }
    };
  } catch {
    return structuredClone(DEFAULTS);
  }
}
function persist() {
  saveTimer = null;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    for (const fn of listeners) fn("saved");
  } catch (e) {
    console.warn("DCSPad: autosave failed", e);
    for (const fn of listeners) fn("error");
  }
}
function getState() {
  return state;
}
function update(patch) {
  Object.assign(state, patch);
  scheduleSave();
}
function updateNested(section, patch) {
  Object.assign(state[section], patch);
  scheduleSave();
}
function scheduleSave() {
  for (const fn of listeners) fn("dirty");
  clearTimeout(saveTimer);
  saveTimer = setTimeout(persist, SAVE_DEBOUNCE_MS);
}
function saveNow() {
  clearTimeout(saveTimer);
  persist();
}
function onSaveStatus(fn) {
  listeners.add(fn);
}
function loadDoc(key2) {
  try {
    const raw = localStorage.getItem(key2);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && Array.isArray(parsed.items) ? parsed : null;
  } catch {
    return null;
  }
}
function saveDoc(key2, doc2) {
  try {
    localStorage.setItem(key2, JSON.stringify(doc2));
    return true;
  } catch (e) {
    console.warn(`DCSPad: saving ${key2} failed`, e);
    return false;
  }
}
var idSeed = Math.random().toString(36).slice(2, 6);
var idCounter = 0;
function newId(prefix) {
  return `${prefix}_${Date.now().toString(36)}${idSeed}${(idCounter++).toString(36)}`;
}
window.addEventListener("pagehide", () => {
  if (saveTimer) saveNow();
});

// ../src/layout.js
var px = (n) => `${n}px`;
function initLayout({ onEditorTabChange } = {}) {
  const main = document.getElementById("main");
  const root = document.documentElement;
  const layout = getState().layout;
  root.style.setProperty("--sidebar-w", px(layout.sidebarW));
  root.style.setProperty("--editors-w", `${layout.editorsFr}fr`);
  root.style.setProperty("--runtime-w", `${layout.runtimeFr}fr`);
  root.style.setProperty("--preview-h", `${layout.previewFr}fr`);
  root.style.setProperty("--diag-h", px(layout.diagH));
  if (layout.sidebarCollapsed) collapseSidebar(true);
  if (layout.diagCollapsed) collapseDiag(true);
  selectEditorTab(layout.editorTab, { silent: true });
  selectDiagTab(layout.diagTab);
  dragSplitter(document.getElementById("split-sidebar"), "x", (dx, start) => {
    const w = Math.min(420, Math.max(140, start.sidebarW + dx));
    root.style.setProperty("--sidebar-w", px(w));
    updateNested("layout", { sidebarW: w });
  }, () => ({ sidebarW: parseFloat(getComputedStyle(root).getPropertyValue("--sidebar-w")) }));
  dragSplitter(document.getElementById("split-center"), "x", (dx, start) => {
    const total = start.editorsPx + start.runtimePx;
    const editorsPx = Math.min(total - 260, Math.max(200, start.editorsPx + dx));
    const fr = editorsPx / (total - editorsPx);
    root.style.setProperty("--editors-w", `${fr}fr`);
    root.style.setProperty("--runtime-w", `1fr`);
    updateNested("layout", { editorsFr: fr, runtimeFr: 1 });
  }, () => ({
    editorsPx: document.getElementById("editors").getBoundingClientRect().width,
    runtimePx: document.getElementById("runtime").getBoundingClientRect().width
  }));
  dragSplitter(document.getElementById("split-runtime"), "y", (dy, start) => {
    const h = Math.min(start.runtimeH - 80, Math.max(100, start.diagH - dy));
    root.style.setProperty("--diag-h", px(h));
    updateNested("layout", { diagH: h });
  }, () => ({
    diagH: document.getElementById("diag-panel").getBoundingClientRect().height,
    runtimeH: document.getElementById("runtime").getBoundingClientRect().height
  }));
  document.getElementById("btn-collapse-sidebar").addEventListener("click", () => collapseSidebar(true));
  document.getElementById("btn-expand-sidebar").addEventListener("click", () => collapseSidebar(false));
  function collapseSidebar(collapsed) {
    main.classList.toggle("sidebar-collapsed", collapsed);
    document.getElementById("btn-expand-sidebar").hidden = !collapsed;
    updateNested("layout", { sidebarCollapsed: collapsed });
  }
  document.getElementById("btn-collapse-diag").addEventListener("click", () => collapseDiag(true));
  function collapseDiag(collapsed) {
    main.classList.toggle("diag-collapsed", collapsed);
    updateNested("layout", { diagCollapsed: collapsed });
  }
  document.getElementById("editor-tabs").addEventListener("click", (e) => {
    const tab = e.target.closest(".tab");
    if (tab) selectEditorTab(tab.dataset.editor);
  });
  function selectEditorTab(name, { silent } = {}) {
    for (const t of document.querySelectorAll("#editor-tabs .tab"))
      t.classList.toggle("active", t.dataset.editor === name);
    updateNested("layout", { editorTab: name });
    if (!silent) onEditorTabChange?.(name);
  }
  document.getElementById("diag-tabs").addEventListener("click", (e) => {
    const tab = e.target.closest(".tab");
    if (!tab) return;
    if (main.classList.contains("diag-collapsed")) collapseDiag(false);
    selectDiagTab(tab.dataset.diag);
  });
  function selectDiagTab(name) {
    for (const t of document.querySelectorAll("#diag-tabs .tab"))
      t.classList.toggle("active", t.dataset.diag === name);
    for (const v of document.querySelectorAll(".diag-view"))
      v.classList.toggle("active", v.id === `view-${name}`);
    document.getElementById("console-tools").hidden = name !== "console";
    document.getElementById("network-tools").hidden = name !== "network";
    updateNested("layout", { diagTab: name });
  }
  document.getElementById("btn-max-preview").addEventListener("click", () => {
    main.classList.remove("max-diag");
    main.classList.toggle("max-preview");
  });
  document.getElementById("btn-max-diag").addEventListener("click", () => {
    if (main.classList.contains("diag-collapsed")) collapseDiag(false);
    main.classList.remove("max-preview");
    main.classList.toggle("max-diag");
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") main.classList.remove("max-preview", "max-diag");
  });
  return { selectEditorTab, selectDiagTab };
}
function dragSplitter(el2, axis, onMove, getStart) {
  el2.addEventListener("pointerdown", (e) => {
    e.preventDefault();
    el2.setPointerCapture(e.pointerId);
    el2.classList.add("dragging");
    const origin = axis === "x" ? e.clientX : e.clientY;
    const start = getStart();
    const move = (ev) => {
      const delta = (axis === "x" ? ev.clientX : ev.clientY) - origin;
      onMove(delta, start);
    };
    const up = () => {
      el2.classList.remove("dragging");
      el2.removeEventListener("pointermove", move);
      el2.removeEventListener("pointerup", up);
    };
    el2.addEventListener("pointermove", move);
    el2.addEventListener("pointerup", up);
  });
}

// ../src/monaco-runtime.js
var runtimePromise = null;
function runtimeBase() {
  if (window.__DCSPAD_ASSET_BASE__) {
    return new URL("vendor/monaco/", window.__DCSPAD_ASSET_BASE__);
  }
  return new URL("../vendor/monaco/", import.meta.url);
}
function assetUrl(name) {
  const url = new URL(name, runtimeBase());
  const version = window.__DCSPAD_MONACO_VERSION__;
  if (version) url.searchParams.set("v", version);
  return url.href;
}
function ensureStylesheet() {
  const existing = document.getElementById("dcspad-monaco-style");
  if (existing) {
    if (existing.dataset.loaded === "true") return Promise.resolve();
    return new Promise((resolve, reject) => {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", () => reject(new Error("Monaco stylesheet failed to load")), { once: true });
    });
  }
  return new Promise((resolve, reject) => {
    const link = document.createElement("link");
    link.id = "dcspad-monaco-style";
    link.rel = "stylesheet";
    link.href = assetUrl("monaco.css");
    link.addEventListener("load", () => {
      link.dataset.loaded = "true";
      resolve();
    }, { once: true });
    link.addEventListener("error", () => reject(new Error("Monaco stylesheet failed to load")), { once: true });
    document.head.appendChild(link);
  });
}
function configureWorkers() {
  self.MonacoEnvironment = {
    getWorker(_moduleId, label) {
      let file = "editor.worker.js";
      if (label === "css" || label === "scss" || label === "less") file = "css.worker.js";
      else if (label === "html" || label === "handlebars" || label === "razor") file = "html.worker.js";
      else if (label === "typescript" || label === "javascript") file = "ts.worker.js";
      const worker = new Worker(assetUrl(file), {
        name: `dcspad-monaco-${label || "editor"}`
      });
      worker.addEventListener("error", () => {
        document.documentElement.dataset.monacoWorkerError = label || "editor";
        const status = document.getElementById("status-run");
        if (status) {
          status.textContent = "editor worker unavailable \u2014 language tools limited";
          status.className = "status-item error";
        }
      }, { once: true });
      return worker;
    }
  };
}
function loadMonacoRuntime() {
  if (runtimePromise) return runtimePromise;
  runtimePromise = (async () => {
    configureWorkers();
    await ensureStylesheet();
    const monaco = await import(assetUrl("monaco.js"));
    document.documentElement.dataset.monacoReady = "true";
    return monaco;
  })();
  return runtimePromise;
}
async function fetchPnpTypeLibraries() {
  const response = await fetch(assetUrl("pnpjs-types.json"), {
    credentials: "same-origin",
    cache: "force-cache"
  });
  if (!response.ok) throw new Error(`HTTP ${response.status} loading PnPjs types`);
  const payload = await response.json();
  if (payload.version !== "2.15.0" || !Array.isArray(payload.libs)) {
    throw new Error("PnPjs type payload is invalid or does not match runtime 2.15.0");
  }
  return payload.libs;
}

// ../src/editors.js
var NAMES = ["html", "css", "js"];
var LANGUAGES = { html: "html", css: "css", js: "javascript" };
var MODEL_URIS = {
  html: "file:///dcspad/index.html",
  css: "file:///dcspad/styles.css",
  js: "file:///dcspad/script.js"
};
async function initEditors({ onChange, onRunShortcut }) {
  const monaco = await loadMonacoRuntime();
  const state3 = getState();
  const host = document.getElementById("pane-editor");
  const cursorEl = document.getElementById("status-cursor");
  const models = {};
  const viewStates = {};
  const selections = {};
  let active = NAMES.includes(state3.layout.editorTab) ? state3.layout.editorTab : "js";
  let desiredPnpTypes = false;
  let pnpTypesGeneration = 0;
  monaco.editor.defineTheme("dcspad-dark", {
    base: "vs-dark",
    inherit: true,
    rules: [
      { token: "comment", foreground: "6A9955" },
      { token: "keyword", foreground: "C586C0" },
      { token: "number", foreground: "B5CEA8" },
      { token: "string", foreground: "CE9178" },
      { token: "type", foreground: "4EC9B0" },
      { token: "type.identifier", foreground: "4EC9B0" },
      { token: "identifier", foreground: "DCDCAA" },
      { token: "tag", foreground: "569CD6" },
      { token: "attribute.name", foreground: "9CDCFE" }
    ],
    colors: {
      // Preserve the prior One Dark editor ground while Monaco replaces its
      // rendering and language-service layers.
      "editor.background": "#282c34",
      "editor.foreground": "#d6d9e0",
      "editorGutter.background": "#282c34",
      "editorLineNumber.foreground": "#5c6270",
      "editorLineNumber.activeForeground": "#b8bdc9",
      "editor.lineHighlightBackground": "#2c313a",
      "editor.selectionBackground": "#264f78",
      "editor.inactiveSelectionBackground": "#264f7855",
      "editorCursor.foreground": "#4ec9b0",
      "editorIndentGuide.background1": "#33374255",
      "editorIndentGuide.activeBackground1": "#4b5263",
      "editorSuggestWidget.background": "#23262e",
      "editorSuggestWidget.border": "#3c4150",
      "editorSuggestWidget.selectedBackground": "#2b4058",
      "editorHoverWidget.background": "#23262e",
      "editorHoverWidget.border": "#3c4150",
      "editorWidget.background": "#23262e",
      "editorWidget.border": "#3c4150"
    }
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
      skipLibCheck: true
    });
  }
  setJsAsModule(state3.settings.jsAsModule);
  jsDefaults.setDiagnosticsOptions({
    noSyntaxValidation: false,
    noSemanticValidation: false,
    noSuggestionDiagnostics: false
  });
  jsDefaults.setEagerModelSync(true);
  for (const name of NAMES) {
    models[name] = monaco.editor.createModel(
      state3[name],
      LANGUAGES[name],
      monaco.Uri.parse(MODEL_URIS[name])
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
    theme: "dcspad-dark",
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
    renderWhitespace: "selection",
    quickSuggestions: { other: true, comments: false, strings: false },
    suggestOnTriggerCharacters: true,
    parameterHints: { enabled: true },
    folding: true,
    padding: { top: 6, bottom: 6 }
  });
  editor.addAction({
    id: "dcspad.run",
    label: "Run DCSPad",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Enter],
    run: () => onRunShortcut?.()
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
      document.documentElement.dataset.pnpTypes = "disabled";
      return;
    }
    document.documentElement.dataset.pnpTypes = "loading";
    try {
      const libs = await fetchPnpTypeLibraries();
      if (!desiredPnpTypes || generation !== pnpTypesGeneration) return;
      jsDefaults.setExtraLibs(libs);
      document.documentElement.dataset.pnpTypes = "ready";
    } catch (error) {
      if (generation !== pnpTypesGeneration) return;
      document.documentElement.dataset.pnpTypes = "error";
      console.warn("DCSPad: PnPjs IntelliSense could not be loaded", error);
    }
  }
  reportCursor();
  return {
    activate,
    getDocs: () => ({
      html: models.html.getValue(),
      css: models.css.getValue(),
      js: models.js.getValue()
    }),
    focus: (name) => activate(name),
    setDocs: (docs) => {
      for (const name of NAMES) {
        if (typeof docs[name] !== "string") continue;
        const model = models[name];
        model.pushStackElement();
        model.pushEditOperations([], [{
          range: model.getFullModelRange(),
          text: docs[name],
          forceMoveMarkers: true
        }], () => null);
        model.pushStackElement();
      }
    },
    getSelection: (name) => {
      const selection = name === active ? editor.getSelection() : selections[name];
      return selection ? models[name].getValueInRange(selection) : "";
    },
    insertAtCursor: (name, text) => {
      activate(name, { focus: false });
      const selection = editor.getSelection() || selections[name];
      editor.executeEdits("dcspad-snippet", [{
        range: selection,
        text,
        forceMoveMarkers: true
      }]);
      editor.focus();
    },
    gotoJsLine: (lineNo) => {
      activate("js", { focus: false });
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
    }
  };
}

// ../src/runner.js
var harnessText = null;
var currentToken = null;
var currentFrame = null;
var runCounter = 0;
var userJsLine = 0;
var evalCallbacks = /* @__PURE__ */ new Map();
var evalCounter = 0;
var handlers = {};
async function initRunner(messageHandlers) {
  handlers = messageHandlers;
  const harnessUrl = window.__DCSPAD_SRC_BASE__ ? window.__DCSPAD_SRC_BASE__ + "bridge/harness.js" : new URL("./bridge/harness.js", import.meta.url);
  const res = await fetch(harnessUrl, { cache: "no-cache" });
  if (!res.ok) {
    throw new Error(`preview harness failed to load (HTTP ${res.status} for bridge/harness.js) \u2014 check the deployed folder structure`);
  }
  harnessText = await res.text();
  window.addEventListener("message", (e) => {
    const d = e.data;
    if (!d || d.dcspad !== currentToken) return;
    if (d.kind === "eval-result") {
      const cb = evalCallbacks.get(d.id);
      if (cb) {
        evalCallbacks.delete(d.id);
        cb(d);
      }
      return;
    }
    handlers[d.kind]?.(d);
  });
}
var escScript = (s) => s.replace(/<\/script/gi, "<\\/script");
var escStyle = (s) => s.replace(/<\/style/gi, "<\\/style");
function hostNonce() {
  for (const s of document.scripts) if (s.nonce) return s.nonce;
  return "";
}
function assemble({ docs, libraries, spContext: spContext2, settings, token }) {
  const nonce = hostNonce();
  const nonceAttr = nonce ? ` nonce="${nonce}"` : "";
  const cssLinks = libraries.filter((l) => l.css).map((l) => (Array.isArray(l.css) ? l.css : [l.css]).map((u) => `<link rel="stylesheet" href="${u}">`).join("\n")).join("\n");
  const jsTags = libraries.filter((l) => l.js).map((l) => (Array.isArray(l.js) ? l.js : [l.js]).map((u) => `<script src="${u}"${nonceAttr}><\/script>`).join("\n")).join("\n");
  const chromeStyle = settings.previewDark ? `<style data-dcspad-chrome>
:root { color-scheme: dark; }
html { background: #1d2026; color: #d6d9e0; }
</style>
` : "";
  const contextScript = spContext2 ? `<script${nonceAttr}>window._spPageContextInfo = ${JSON.stringify(spContext2.pageContext)};<\/script>
` + (spContext2.baseHref ? `<base href="${spContext2.baseHref}">
` : "") : "";
  const head = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script${nonceAttr}>${escScript(harnessText.replaceAll("__DCSPAD_TOKEN__", token))}<\/script>
${contextScript}${chromeStyle}${cssLinks}
<style>
${escStyle(docs.css)}
</style>
</head>
<body>
${docs.html}
${jsTags}
`;
  const scriptOpen = settings.jsAsModule ? `<script type="module"${nonceAttr}>` : `<script${nonceAttr}>`;
  userJsLine = head.split("\n").length + 1;
  return `${head}${scriptOpen}
${escScript(docs.js)}
<\/script>
</body>
</html>`;
}
function run(opts) {
  runCounter += 1;
  currentToken = `run-${runCounter}-${Math.random().toString(36).slice(2)}`;
  const doc2 = assemble({ ...opts, token: currentToken });
  const host = document.getElementById("preview-host");
  document.getElementById("preview-empty")?.remove();
  if (currentFrame) currentFrame.remove();
  for (const cb of evalCallbacks.values()) {
    cb({ ok: false, cancelled: true, value: { t: "str", v: "(cancelled \u2014 a new run replaced the frame before this settled)" } });
  }
  evalCallbacks.clear();
  const frame = document.createElement("iframe");
  frame.setAttribute("sandbox", "allow-scripts allow-same-origin allow-forms allow-modals allow-popups");
  frame.srcdoc = doc2;
  host.appendChild(frame);
  currentFrame = frame;
  return { runNumber: runCounter, token: currentToken };
}
function evalInFrame(code) {
  return new Promise((resolve) => {
    if (!currentFrame) {
      resolve({ ok: false, value: { t: "str", v: "Nothing is running \u2014 press Run first." }, noRun: true });
      return;
    }
    const id = ++evalCounter;
    evalCallbacks.set(id, resolve);
    currentFrame.contentWindow.postMessage({ dcspad: currentToken, kind: "eval", code, id }, "*");
  });
}
function mapSrcdocLineToUserJs(line) {
  const mapped = line - userJsLine + 1;
  return mapped >= 1 ? mapped : null;
}
function hasRun() {
  return !!currentFrame;
}

// ../src/inspect/tree-view.js
var el = (tag, cls, text) => {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== void 0) n.textContent = text;
  return n;
};
function renderValue(node, opts = {}) {
  if (!node) return el("span", "t-undef", "undefined");
  switch (node.t) {
    case "str": {
      const s = el("span", "t-str", opts.bare ? node.v : JSON.stringify(node.v));
      if (node.trunc) s.append(el("span", "t-truncated", ` \u2026(${node.trunc} chars)`));
      return s;
    }
    case "num":
      return el("span", "t-num", String(node.v));
    case "bool":
      return el("span", "t-bool", String(node.v));
    case "null":
      return el("span", "t-null", "null");
    case "undef":
      return el("span", "t-undef", "undefined");
    case "sym":
      return el("span", "t-str", node.v);
    case "fn":
      return el("span", "t-fn", `\u0192 ${node.v}()`);
    case "date":
      return el("span", "t-node", node.v);
    case "regex":
      return el("span", "t-str", node.v);
    case "node":
      return el("span", "t-node", node.v);
    case "circ":
      return el("span", "t-circular", "[circular]");
    case "maxdepth":
      return el("span", "t-preview", node.v);
    case "err":
      return renderError(node);
    case "arr":
      return renderExpandable(node, `Array(${node.n})`, node.items.map((item, i) => [String(i), item]), opts);
    case "obj":
      return renderExpandable(node, node.cls === "Object" ? "" : node.cls, node.keys, opts);
    default:
      return el("span", "t-preview", JSON.stringify(node));
  }
}
function renderError(node) {
  const wrap = el("span");
  let head = `${node.name}: ${node.msg}`;
  if (node.status !== void 0) head += ` (HTTP ${node.status}${node.statusText ? " " + node.statusText : ""})`;
  wrap.append(el("span", "t-err", head));
  if (node.stack) {
    const stack = el("div", "stack-frame");
    stack.textContent = node.stack.split("\n").slice(1, 6).join("\n");
    wrap.append(stack);
  }
  return wrap;
}
function previewOf(node) {
  switch (node.t) {
    case "str": {
      const v = node.v.length > 24 ? node.v.slice(0, 24) + "\u2026" : node.v;
      return JSON.stringify(v);
    }
    case "num":
    case "bool":
      return String(node.v);
    case "null":
      return "null";
    case "undef":
      return "undefined";
    case "fn":
      return "\u0192";
    case "arr":
      return `Array(${node.n})`;
    case "obj":
      return node.cls === "Object" ? "{\u2026}" : `${node.cls}`;
    case "err":
      return node.name;
    case "node":
      return node.v;
    case "date":
      return node.v;
    case "maxdepth":
      return node.v;
    case "circ":
      return "[circular]";
    default:
      return "\u2026";
  }
}
function renderExpandable(node, label, entries, opts = {}) {
  const wrap = el("span", "tree-node");
  const row = el("span", "tree-row expandable");
  row.append(el("span", "twist", "\u25B6"));
  if (label) row.append(el("span", "", label + " "));
  const parts = entries.slice(0, 5).map(([k, v]) => (node.t === "arr" ? "" : `${k}: `) + previewOf(v));
  const openBrace = node.t === "arr" ? "[" : "{";
  const closeBrace = node.t === "arr" ? "]" : "}";
  const more = entries.length > 5 || node.trunc ? ", \u2026" : "";
  row.append(el("span", "t-preview", `${openBrace}${parts.join(", ")}${more}${closeBrace}`));
  wrap.append(row);
  const children = el("div", "tree-children");
  wrap.append(children);
  let built = false;
  row.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = row.classList.toggle("open");
    if (open && !built) {
      built = true;
      for (const [key2, val] of entries) {
        const line = el("div");
        const keySpan = el("span", "tree-key" + (opts.dimKeys?.has?.(key2) ? " dim-key" : ""), key2);
        line.append(keySpan, el("span", "", ": "), renderValue(val, opts));
        children.append(line);
      }
      if (node.trunc) children.append(el("div", "t-truncated", "\u2026 truncated"));
    }
  });
  return wrap;
}
function renderTable(dataNode, columns) {
  if (!dataNode || dataNode.t !== "arr" && dataNode.t !== "obj") {
    return renderValue(dataNode);
  }
  const rows = dataNode.t === "arr" ? dataNode.items.map((item, i) => [String(i), item]) : dataNode.keys;
  let cols = columns ? [...columns] : [];
  if (!cols.length) {
    const seen = /* @__PURE__ */ new Set();
    for (const [, v] of rows) {
      if (v.t === "obj") for (const [k] of v.keys) seen.add(k);
      else if (v.t === "arr") v.items.forEach((_, i) => seen.add(String(i)));
      else seen.add("Value");
    }
    cols = [...seen].slice(0, 20);
  }
  const wrap = el("div", "console-table-wrap");
  const table = el("table", "console-table");
  const thead = el("thead");
  const hr = el("tr");
  hr.append(el("th", "", "(index)"));
  cols.forEach((c) => hr.append(el("th", "", c)));
  thead.append(hr);
  table.append(thead);
  const tbody = el("tbody");
  for (const [key2, v] of rows) {
    const tr = el("tr");
    tr.append(el("td", "", key2));
    for (const c of cols) {
      const td = el("td");
      let cell;
      if (v.t === "obj") cell = v.keys.find(([k]) => k === c)?.[1];
      else if (v.t === "arr") cell = v.items[Number(c)];
      else if (c === "Value") cell = v;
      td.textContent = cell ? previewOf(cell) : "";
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  wrap.append(table);
  return wrap;
}

// ../src/inspect/sp-shapes.js
var NOISE_KEYS = /* @__PURE__ */ new Set(["__metadata", "__deferred", "odata.metadata", "odata.type", "odata.id", "odata.etag", "odata.editLink", "@odata.context", "@odata.type", "@odata.id", "@odata.etag", "@odata.editLink", "FirstUniqueAncestorSecurableObject", "RoleAssignments"]);
var key = (node, k) => node?.t === "obj" ? node.keys.find(([n]) => n === k)?.[1] : void 0;
var keyNames = (node) => node?.t === "obj" ? node.keys.map(([n]) => n) : [];
var str = (node) => node && (node.t === "str" || node.t === "num") ? String(node.v) : void 0;
function enhance(node) {
  if (!node || node.t !== "obj" && node.t !== "arr") return null;
  const d = key(node, "d");
  if (d && node.keys.length === 1) {
    return envelope("OData verbose", d, node);
  }
  const value = key(node, "value");
  if (value?.t === "arr" && keyNames(node).every((k) => k === "value" || k.startsWith("odata.") || k.startsWith("@odata."))) {
    return envelope("OData", value, node);
  }
  const results = key(node, "results");
  if (results?.t === "arr") {
    return collection(results, node);
  }
  if (node.t === "arr" && node.items.length && node.items.every(looksLikeSpObject)) {
    return collection(node, null);
  }
  if (looksLikeSpObject(node)) {
    return entity(node);
  }
  return null;
}
function looksLikeSpObject(node) {
  if (node?.t !== "obj") return false;
  if (key(node, "__metadata")) return true;
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  return has("InternalName", "TypeAsString") || has("BaseTemplate", "EntityTypeName") || has("ServerRelativeUrl", "WebTemplate") || has("LoginName", "PrincipalType") || names.includes("odata.type") || names.includes("@odata.type");
}
function spType(node) {
  const meta = key(node, "__metadata");
  return str(key(meta, "type")) || str(key(node, "odata.type")) || str(key(node, "@odata.type")) || detectShape(node);
}
function detectShape(node) {
  const names = keyNames(node);
  const has = (...ks) => ks.every((k) => names.includes(k));
  if (has("InternalName", "TypeAsString")) return "SP.Field";
  if (has("BaseTemplate", "EntityTypeName")) return "SP.List";
  if (has("ServerRelativeUrl", "WebTemplate")) return "SP.Web";
  if (has("LoginName", "PrincipalType")) {
    return str(key(node, "OwnerTitle")) !== void 0 ? "SP.Group" : "SP.User";
  }
  return null;
}
function envelope(label, inner, outer) {
  const wrap = el("div", "tree-node");
  const head = el("div");
  head.append(badge(label));
  wrap.append(head);
  const enhanced = enhance(inner);
  wrap.append(enhanced ?? renderValue(inner, { dimKeys: NOISE_KEYS }));
  const metaKeys = outer.keys.filter(([k]) => k !== "d" && k !== "value");
  if (metaKeys.length) {
    const fold = el("div", "sp-meta-fold");
    fold.append(renderValue({ t: "obj", cls: "envelope metadata", keys: metaKeys }, { dimKeys: NOISE_KEYS }));
    wrap.append(fold);
  }
  return wrap;
}
function collection(arrNode, parentNode) {
  const wrap = el("div", "tree-node");
  const head = el("div");
  const type = arrNode.items.length ? spType(arrNode.items[0]) : null;
  head.append(badge(`${arrNode.n} item${arrNode.n === 1 ? "" : "s"}`));
  if (type) head.append(el("span", "sp-entity-head", shortType(type)));
  const toggle = el("span", "table-toggle", "\u229E table view");
  head.append(toggle);
  wrap.append(head);
  const treeEl = el("div");
  if (arrNode.items.length && arrNode.items.every((i) => i.t === "obj")) {
    const list = el("div");
    arrNode.items.forEach((item, i) => {
      const row = el("div");
      row.append(el("span", "tree-key dim-key", `${i}: `));
      row.append(enhance(item) ?? renderValue(item, { dimKeys: NOISE_KEYS }));
      list.append(row);
    });
    if (arrNode.trunc) list.append(el("div", "t-truncated", `\u2026 showing first ${arrNode.items.length} of ${arrNode.n}`));
    treeEl.append(list);
  } else {
    treeEl.append(renderValue(arrNode, { dimKeys: NOISE_KEYS }));
  }
  const tableEl = el("div");
  tableEl.hidden = true;
  let tableBuilt = false;
  toggle.addEventListener("click", () => {
    const showTable = tableEl.hidden;
    if (showTable && !tableBuilt) {
      tableBuilt = true;
      tableEl.append(renderTable(filterNoise(arrNode)));
    }
    tableEl.hidden = !showTable;
    treeEl.hidden = showTable;
    toggle.textContent = showTable ? "\u2261 tree view" : "\u229E table view";
  });
  wrap.append(treeEl, tableEl);
  const next = str(key(parentNode, "__next")) || str(key(parentNode, "odata.nextLink")) || str(key(parentNode, "@odata.nextLink"));
  if (next) {
    const warn = el("div", "sp-next-link");
    warn.append(el("span", "", "\u26A0 partial result set \u2014 next page: "));
    warn.append(copySpan(next, next.length > 80 ? next.slice(0, 80) + "\u2026" : next));
    wrap.append(warn);
  }
  return wrap;
}
function filterNoise(arrNode) {
  return {
    ...arrNode,
    items: arrNode.items.map((item) => item.t === "obj" ? { ...item, keys: item.keys.filter(([k]) => !NOISE_KEYS.has(k)) } : item)
  };
}
var ENTITY_FIELDS = {
  "SP.List": [
    ["Title", false],
    ["Id", true],
    ["EntityTypeName", true],
    ["BaseTemplate", false],
    ["ItemCount", false]
  ],
  "SP.Field": [
    ["Title", false],
    ["InternalName", true],
    ["TypeAsString", false],
    ["Required", false]
  ],
  "SP.Web": [
    ["Title", false],
    ["ServerRelativeUrl", true],
    ["WebTemplate", false]
  ],
  "SP.User": [
    ["Title", false],
    ["LoginName", true],
    ["Email", true]
  ],
  "SP.Group": [
    ["Title", false],
    ["Id", true],
    ["OwnerTitle", false]
  ],
  "SP.ListItem": [
    ["Title", false],
    ["Id", false]
  ]
};
function shortType(type) {
  if (!type) return "";
  if (type.startsWith("SP.Data.") && type.endsWith("Item")) return "SP.ListItem \xB7 " + type.slice(8);
  return type;
}
function entityKind(type) {
  if (!type) return null;
  if (ENTITY_FIELDS[type]) return type;
  for (const known of Object.keys(ENTITY_FIELDS)) {
    if (type.startsWith(known)) return known;
  }
  if (type.startsWith("SP.Data.")) return "SP.ListItem";
  return null;
}
function entity(node) {
  const type = spType(node);
  const kind = entityKind(type);
  const wrap = el("div", "tree-node");
  const head = el("div", "sp-entity-head");
  head.append(badge(shortType(type) || "SP"));
  if (kind) {
    for (const [field, copyable] of ENTITY_FIELDS[kind]) {
      const v = key(node, field);
      if (v === void 0) continue;
      const fieldEl = el("span", "sp-field");
      fieldEl.append(el("span", "dim-key tree-key", `${field}: `));
      const text = v.t === "str" || v.t === "num" || v.t === "bool" ? String(v.v) : previewOf(v);
      fieldEl.append(copyable ? copySpan(text, text) : el("span", "", text));
      head.append(fieldEl);
    }
  }
  wrap.append(head);
  wrap.append(renderValue(node, { dimKeys: NOISE_KEYS }));
  return wrap;
}
function badge(text) {
  return el("span", "sp-badge", text);
}
function copySpan(copyText, displayText) {
  const s = el("span", "sp-copy", displayText);
  s.title = "Click to copy";
  s.addEventListener("click", async (e) => {
    e.stopPropagation();
    try {
      await navigator.clipboard.writeText(copyText);
      s.classList.add("copied");
      setTimeout(() => s.classList.remove("copied"), 800);
    } catch {
    }
  });
  return s;
}

// ../src/console-panel.js
var out;
var groupStack;
var replHistory;
var replIndex;
var deps = {};
var FILTER_DEBOUNCE_MS = 150;
function initConsolePanel({ evalInFrame: evalInFrame2, mapSrcdocLine, gotoJsLine, isConsoleVisible }) {
  deps = { evalInFrame: evalInFrame2, mapSrcdocLine, gotoJsLine, isConsoleVisible };
  out = document.getElementById("console-out");
  groupStack = [];
  replHistory = [];
  replIndex = -1;
  document.getElementById("btn-clear-console").addEventListener("click", clear);
  for (const btn of document.querySelectorAll(".lvl-filter")) {
    btn.addEventListener("click", () => {
      btn.classList.toggle("active");
      applyFilters();
    });
  }
  let filterTimer = null;
  document.getElementById("console-filter-text").addEventListener("input", () => {
    clearTimeout(filterTimer);
    filterTimer = setTimeout(applyFilters, FILTER_DEBOUNCE_MS);
  });
  refreshFilterState();
  const input = document.getElementById("console-input");
  input.addEventListener("keydown", async (e) => {
    if (e.key === "Enter" && input.value.trim()) {
      const code = input.value;
      input.value = "";
      replHistory.push(code);
      replIndex = replHistory.length;
      addEntry("log", [el("span", "", code)], { cls: "repl-echo" });
      const res = await deps.evalInFrame(code);
      const body = renderNodeSmart(res.value);
      if (res.awaited) {
        const tag = el("span", "sp-badge", "awaited");
        addEntry(res.ok ? "log" : "error", [tag, body], { cls: "repl-result" });
      } else {
        addEntry(res.ok ? "log" : "error", [body], { cls: "repl-result" });
      }
    } else if (e.key === "ArrowUp") {
      if (replIndex > 0) {
        replIndex--;
        input.value = replHistory[replIndex];
        e.preventDefault();
      }
    } else if (e.key === "ArrowDown") {
      if (replIndex < replHistory.length - 1) {
        replIndex++;
        input.value = replHistory[replIndex];
      } else {
        replIndex = replHistory.length;
        input.value = "";
      }
      e.preventDefault();
    }
  });
  return { handlers: makeHandlers(), clear, runDivider };
}
function renderNodeSmart(node) {
  return enhance(node) ?? renderValue(node);
}
function makeHandlers() {
  return {
    console: (d) => {
      const parts = d.args.map((a) => a.t === "str" ? el("span", "", a.v) : renderNodeSmart(a));
      addEntry(d.level, parts);
    },
    table: (d) => addEntry("log", [renderTable(d.data, d.columns)]),
    group: (d) => {
      const header = el("div", "console-entry group-header" + (d.collapsed ? " collapsed" : ""));
      header.append(el("span", "twist", "\u25B6"));
      header.append(el("span", "entry-body", d.label));
      const container = el("div", "console-group");
      header.addEventListener("click", () => header.classList.toggle("collapsed"));
      currentContainer().append(header, container);
      groupStack.push(container);
      scrollIfPinned();
    },
    groupEnd: () => {
      groupStack.pop();
    },
    clear: () => clear(),
    error: (d) => {
      const parts = [];
      let msg = d.message || "Error";
      if (d.rejection) msg = "Uncaught (in promise): " + msg;
      parts.push(el("span", "", msg));
      if (d.rejection && d.reason && d.reason.t !== "str" && d.reason.t !== "err") {
        parts.push(renderNodeSmart(d.reason));
      }
      const stackEl = renderStack(d);
      if (stackEl) parts.push(stackEl);
      addEntry("error", parts);
    }
  };
}
function renderStack(d) {
  const lines = [];
  if (d.stack) lines.push(...d.stack.split("\n").slice(1, 8));
  else if (d.source && d.line) lines.push(`    at ${d.source}:${d.line}:${d.col ?? 0}`);
  if (!lines.length) return null;
  const wrap = el("div", "stack-frame");
  for (const line of lines) {
    const div = el("div");
    const m = line.match(/(about:srcdoc):(\d+):(\d+)/);
    if (m) {
      const userLine = deps.mapSrcdocLine(Number(m[2]));
      if (userLine) {
        const pre = line.slice(0, m.index);
        const link = el("a", "", `js:${userLine}:${m[3]}`);
        link.addEventListener("click", () => deps.gotoJsLine(userLine));
        div.append(el("span", "", pre), link);
        wrap.append(div);
        continue;
      }
    }
    div.textContent = line;
    wrap.append(div);
  }
  return wrap;
}
function currentContainer() {
  return groupStack.length ? groupStack[groupStack.length - 1] : out;
}
function addEntry(level, parts, { cls } = {}) {
  const entry = el("div", `console-entry lvl-${level} new-entry${cls ? " " + cls : ""}`);
  entry.dataset.lvl = level === "info" || level === "debug" ? "log" : level;
  const ts = /* @__PURE__ */ new Date();
  entry.append(el(
    "span",
    "entry-ts",
    `${String(ts.getHours()).padStart(2, "0")}:${String(ts.getMinutes()).padStart(2, "0")}:${String(ts.getSeconds()).padStart(2, "0")}`
  ));
  const body = el("span", "entry-body");
  for (const p of parts) body.append(p, " ");
  entry.append(body);
  applyFilterTo(entry);
  currentContainer().append(entry);
  scrollIfPinned();
  if (level === "error") document.getElementById("console-badge").hidden = false;
}
function runDivider(runNumber) {
  groupStack = [];
  const div = el("div", "run-divider new-entry");
  const ts = (/* @__PURE__ */ new Date()).toLocaleTimeString();
  div.append(el("span", "rd-mark", "\u259E \u25B6"), el("span", "", `run #${runNumber} \xB7 ${ts}`));
  out.append(div);
  scrollIfPinned(true);
}
function clear() {
  out.textContent = "";
  groupStack = [];
  document.getElementById("console-badge").hidden = true;
}
var filterState = { lvls: /* @__PURE__ */ new Set(), text: "" };
function refreshFilterState() {
  filterState = {
    lvls: new Set([...document.querySelectorAll(".lvl-filter.active")].map((b) => b.dataset.lvl)),
    text: document.getElementById("console-filter-text").value.trim().toLowerCase()
  };
}
function applyFilters() {
  refreshFilterState();
  for (const entry of out.querySelectorAll(".console-entry")) applyFilterTo(entry);
}
function applyFilterTo(entry) {
  if (!entry.dataset.lvl) return;
  const lvl = entry.dataset.lvl === "error" ? "error" : entry.dataset.lvl === "warn" ? "warn" : "log";
  entry.classList.toggle("hidden-lvl", !filterState.lvls.has(lvl));
  const body = entry.querySelector(".entry-body");
  const bodyText = body ? body.textContent.toLowerCase() : "";
  entry.classList.toggle("hidden-txt", !!filterState.text && !bodyText.includes(filterState.text));
}
function scrollIfPinned(force) {
  const nearBottom = out.scrollHeight - out.scrollTop - out.clientHeight < 60;
  if (nearBottom || force) out.scrollTop = out.scrollHeight;
}

// ../src/network-panel.js
var requests = /* @__PURE__ */ new Map();
var selectedId = null;
var deps2 = {};
function initNetworkPanel({ isNetworkVisible }) {
  deps2 = { isNetworkVisible };
  document.getElementById("btn-clear-network").addEventListener("click", clear2);
  document.getElementById("chk-api-only").addEventListener("change", applyApiFilter);
  return { handlers: { "net-start": onStart, "net-end": onEnd }, clear: clear2 };
}
var isApiUrl = (url) => /\/_api\/|\/_vti_bin\//i.test(url);
var fmtSize = (bytes) => bytes < 1024 ? `${bytes} B` : bytes < 1024 * 1024 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
function onStart(d) {
  const tbody = document.getElementById("network-rows");
  const row = el("tr", "network-row net-pending");
  row.dataset.id = d.id;
  if (isApiUrl(d.url)) row.classList.add("is-api");
  const tdMethod = el("td", "net-method", d.method);
  const tdUrl = el("td", "net-url", d.url);
  tdUrl.title = d.url;
  const tdStatus = el("td", "net-status net-status-pending", "\u2026");
  const tdTime = el("td", "", "");
  const tdSize = el("td", "", "");
  row.append(tdMethod, tdUrl, tdStatus, tdTime, tdSize);
  row.addEventListener("click", () => select(d.id));
  requests.set(d.id, { row, data: { ...d } });
  applyApiFilterTo(row);
  tbody.append(row);
  const wrap = document.getElementById("network-table-wrap");
  wrap.scrollTop = wrap.scrollHeight;
}
function onEnd(d) {
  const entry = requests.get(d.id);
  if (!entry) return;
  Object.assign(entry.data, d);
  const [, , tdStatus, tdTime, tdSize] = entry.row.children;
  tdStatus.textContent = d.failed ? "\u2715 failed" : String(d.status);
  tdStatus.className = "net-status " + (d.ok ? "net-status-ok" : "net-status-err");
  tdTime.textContent = `${d.ms} ms`;
  tdSize.textContent = d.size != null ? fmtSize(d.size) : "\u2014";
  if (selectedId === d.id) renderDetail(entry.data);
  if (!d.ok) document.getElementById("network-badge").hidden = false;
}
function select(id) {
  selectedId = id;
  for (const { row } of requests.values()) row.classList.toggle("selected", row.dataset.id === id);
  renderDetail(requests.get(id).data);
}
function renderDetail(data) {
  const detail = document.getElementById("network-detail");
  detail.hidden = false;
  detail.textContent = "";
  const close = el("span", "nd-close", "\u2715");
  close.addEventListener("click", () => {
    detail.hidden = true;
    selectedId = null;
    for (const { row } of requests.values()) row.classList.remove("selected");
  });
  detail.append(close);
  detail.append(el("h4", "", "Request"));
  const kv = (label, value) => {
    const div = el("div", "nd-kv");
    div.append(el("b", "", label + ": "), el("span", "", value ?? "\u2014"));
    return div;
  };
  detail.append(kv("Method", data.method), kv("URL", data.url), kv("Via", data.api === "xhr" ? "XMLHttpRequest" : "fetch"));
  detail.append(el("h4", "", "Response"));
  if (data.cancelled) {
    detail.append(el("div", "net-status-err", "cancelled \u2014 the frame was replaced by a new run before the response arrived"));
    return;
  }
  if (data.status === void 0) {
    detail.append(el("div", "net-status-pending", "pending\u2026"));
    return;
  }
  detail.append(
    kv("Status", `${data.status} ${data.statusText || ""}`),
    kv("Duration", `${data.ms} ms`),
    kv("Size", data.size != null ? fmtSize(data.size) : "\u2014"),
    kv("Content-Type", data.contentType || "\u2014")
  );
  detail.append(el("h4", "", "Body"));
  if (!data.preview) {
    detail.append(el("div", "t-preview", data.failed ? `request failed: ${data.statusText}` : "(no text preview)"));
    return;
  }
  if ((data.contentType || "").includes("json") || /^[\[{]/.test(data.preview.trim())) {
    try {
      const parsed = JSON.parse(data.preview);
      const node = toNode(parsed, 0);
      detail.append(enhance(node) ?? renderValue(node));
      return;
    } catch {
    }
  }
  const pre = el("pre", "", data.preview.slice(0, 5e3));
  pre.style.whiteSpace = "pre-wrap";
  detail.append(pre);
}
function toNode(v, depth) {
  if (v === null) return { t: "null" };
  switch (typeof v) {
    case "string":
      return { t: "str", v };
    case "number":
      return { t: "num", v };
    case "boolean":
      return { t: "bool", v };
    case "undefined":
      return { t: "undef" };
  }
  if (depth >= 6) return { t: "maxdepth", v: Array.isArray(v) ? `Array(${v.length})` : "{\u2026}" };
  if (Array.isArray(v)) {
    return { t: "arr", n: v.length, items: v.slice(0, 100).map((x) => toNode(x, depth + 1)), trunc: v.length > 100 };
  }
  const keys = Object.keys(v);
  return {
    t: "obj",
    cls: "Object",
    keys: keys.slice(0, 100).map((k) => [k, toNode(v[k], depth + 1)]),
    trunc: keys.length > 100
  };
}
function clear2() {
  requests.clear();
  selectedId = null;
  document.getElementById("network-rows").textContent = "";
  document.getElementById("network-detail").hidden = true;
  document.getElementById("network-badge").hidden = true;
}
function applyApiFilter() {
  for (const { row } of requests.values()) applyApiFilterTo(row);
}
function applyApiFilterTo(row) {
  const apiOnly = document.getElementById("chk-api-only").checked;
  row.classList.toggle("hidden-api", apiOnly && !row.classList.contains("is-api"));
}
function markRun() {
  for (const { row, data } of requests.values()) {
    if (data.status !== void 0 || data.cancelled) continue;
    data.cancelled = true;
    const tdStatus = row.children[2];
    tdStatus.textContent = "\u2715 cancelled";
    tdStatus.className = "net-status net-status-err";
    row.classList.remove("net-pending");
    if (selectedId === data.id) renderDetail(data);
  }
  const tbody = document.getElementById("network-rows");
  if (tbody.children.length) {
    const sep = el("tr", "net-run-sep");
    const td = el("td", "", "\u2014 new run \u2014");
    td.colSpan = 5;
    td.style.color = "var(--fg-faint)";
    td.style.textAlign = "center";
    sep.append(td);
    tbody.append(sep);
  }
}

// ../src/libraries.js
var PRESETS = [
  {
    id: "dcs-standard",
    name: "DCS Standard Include",
    needsConfig: true,
    hint: "Set your org include URL once; stored with your workspace."
  },
  {
    id: "pnpjs2",
    name: "PnPjs v2 (classic)",
    js: "https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/2.15.0/pnpjs.es5.umd.bundle.min.js",
    hint: "Exposes global pnp \u2014 use const { sp } = pnp;"
  },
  { id: "alpine", name: "Alpine.js", js: "https://cdn.jsdelivr.net/npm/alpinejs@3/dist/cdn.min.js" },
  { id: "chartjs", name: "Chart.js", js: "https://cdn.jsdelivr.net/npm/chart.js@4/dist/chart.umd.min.js" },
  { id: "lodash", name: "Lodash", js: "https://cdn.jsdelivr.net/npm/lodash@4/lodash.min.js" },
  { id: "exceljs", name: "ExcelJS", js: "https://cdn.jsdelivr.net/npm/exceljs@4/dist/exceljs.min.js" },
  { id: "dayjs", name: "Day.js", js: "https://cdn.jsdelivr.net/npm/dayjs@1/dayjs.min.js" },
  { id: "fusejs", name: "Fuse.js", js: "https://cdn.jsdelivr.net/npm/fuse.js@7/dist/fuse.min.js" },
  { id: "marked", name: "Marked", js: "https://cdn.jsdelivr.net/npm/marked@12/marked.min.js" },
  { id: "sortable", name: "Sortable.js", js: "https://cdn.jsdelivr.net/npm/sortablejs@1/Sortable.min.js" },
  { id: "fabric", name: "Fluent/Fabric Icons (CSS)", css: "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/11.0.0/css/fabric.min.css" }
];
var catalog = null;
var onChangeCb = null;
var onStorageErrorCb = null;
var isCssUrl = (url) => /\.css(\?|$)/i.test(url);
var entryFromUrl = (url, name) => ({
  id: newId("lib"),
  name: name || url.split("/").pop() || url,
  js: isCssUrl(url) ? void 0 : url,
  css: isCssUrl(url) ? url : void 0
});
function initLibraries({ onChange, onStorageError }) {
  onChangeCb = onChange;
  onStorageErrorCb = onStorageError;
  catalog = loadDoc(CATALOG_KEY);
  if (!catalog) {
    catalog = { v: 1, items: structuredClone(PRESETS) };
    const libs = getState().libraries;
    if (libs.custom?.length) {
      const migratedIds = [];
      for (const url of libs.custom) {
        const entry = entryFromUrl(url);
        catalog.items.push(entry);
        migratedIds.push(entry.id);
      }
      updateNested("libraries", { custom: [], enabled: [...libs.enabled, ...migratedIds] });
    }
    persistCatalog();
  }
  render();
  document.getElementById("lib-custom-form").addEventListener("submit", (e) => {
    e.preventDefault();
    const urlInput = document.getElementById("lib-custom-url");
    const nameInput = document.getElementById("lib-custom-name");
    const url = urlInput.value.trim();
    if (!url) return;
    const entry = entryFromUrl(url, nameInput.value.trim());
    catalog.items.push(entry);
    persistCatalog();
    const enabled = new Set(getState().libraries.enabled);
    enabled.add(entry.id);
    updateNested("libraries", { enabled: [...enabled] });
    urlInput.value = "";
    nameInput.value = "";
    render();
    onChangeCb?.();
  });
  return { getEnabledLibraries };
}
function persistCatalog() {
  if (!saveDoc(CATALOG_KEY, catalog)) {
    onStorageErrorCb?.("framework catalog save failed (storage full?)");
  }
}
function render() {
  const libs = getState().libraries;
  const pinnedHost = document.getElementById("lib-pinned");
  const listHost = document.getElementById("lib-list");
  pinnedHost.textContent = "";
  listHost.textContent = "";
  for (const entry of catalog.items) {
    const pinned = libs.pinned.includes(entry.id);
    (pinned ? pinnedHost : listHost).append(catalogItem(entry, libs, pinned));
  }
}
function catalogItem(entry, libs, pinned) {
  const item = el("label", "lib-item");
  const chk = document.createElement("input");
  chk.type = "checkbox";
  chk.checked = libs.enabled.includes(entry.id);
  const name = el("span", "lib-name", entry.name);
  if (entry.hint) name.title = entry.hint;
  else if (entry.js || entry.css) name.title = entry.js || entry.css;
  if (entry.needsConfig && !libs.dcsUrl) {
    item.classList.add("needs-config");
    name.title = entry.hint || "Needs a URL";
  }
  chk.addEventListener("change", () => {
    if (entry.needsConfig && !getState().libraries.dcsUrl && chk.checked) {
      const url = prompt("URL for the DCS Standard Include (your org\u2019s script/CSS bundle):");
      if (!url) {
        chk.checked = false;
        return;
      }
      updateNested("libraries", { dcsUrl: url.trim() });
      item.classList.remove("needs-config");
    }
    const enabled = new Set(getState().libraries.enabled);
    chk.checked ? enabled.add(entry.id) : enabled.delete(entry.id);
    updateNested("libraries", { enabled: [...enabled] });
    onChangeCb?.();
  });
  const tools = el("span", "lib-tools");
  const tool = (cls, text, title, fn) => {
    const s = el("span", cls, text);
    s.title = title;
    s.addEventListener("click", (e) => {
      e.preventDefault();
      fn();
    });
    return s;
  };
  const liveIdx = () => catalog.items.indexOf(entry);
  tools.append(
    tool("lib-move", "\u2191", "Move up (injection order)", () => moveEntry(liveIdx(), -1)),
    tool("lib-move", "\u2193", "Move down (injection order)", () => moveEntry(liveIdx(), 1)),
    tool("lib-pin" + (pinned ? " pinned" : ""), pinned ? "\u2605" : "\u2606", pinned ? "Unpin" : "Pin to top", () => {
      const pins = new Set(getState().libraries.pinned);
      pinned ? pins.delete(entry.id) : pins.add(entry.id);
      updateNested("libraries", { pinned: [...pins] });
      render();
    }),
    tool("lib-del", "\u2715", "Remove from catalog", () => {
      const idx = liveIdx();
      if (idx === -1) return;
      if (!confirm(`Remove "${entry.name}" from the framework catalog?`)) return;
      catalog.items.splice(idx, 1);
      persistCatalog();
      const cur = getState().libraries;
      updateNested("libraries", {
        enabled: cur.enabled.filter((id) => id !== entry.id),
        pinned: cur.pinned.filter((id) => id !== entry.id)
      });
      render();
      onChangeCb?.();
    })
  );
  item.append(chk, name, tools);
  return item;
}
function moveEntry(idx, delta) {
  const to = idx + delta;
  if (idx < 0 || to < 0 || to >= catalog.items.length) return;
  const [entry] = catalog.items.splice(idx, 1);
  catalog.items.splice(to, 0, entry);
  persistCatalog();
  render();
  onChangeCb?.();
}
function getEnabledLibraries() {
  const libs = getState().libraries;
  const result = [];
  for (const entry of catalog.items) {
    if (!libs.enabled.includes(entry.id)) continue;
    if (entry.needsConfig) {
      if (libs.dcsUrl) {
        result.push({ name: entry.name, js: isCssUrl(libs.dcsUrl) ? void 0 : libs.dcsUrl, css: isCssUrl(libs.dcsUrl) ? libs.dcsUrl : void 0 });
      }
      continue;
    }
    result.push({ name: entry.name, js: entry.js, css: entry.css });
  }
  return result;
}
function isPnpjs215Runtime(entry) {
  const url = String(entry?.js || "").toLowerCase();
  return url.includes("@pnp/pnpjs@2.15.0/") || url.includes("/pnp-pnpjs/2.15.0/") || url.includes("/pnpjs/2.15.0/");
}
function hasEnabledPnpjs215Runtime() {
  const enabled = new Set(getState().libraries.enabled);
  return catalog.items.some((entry) => enabled.has(entry.id) && isPnpjs215Runtime(entry));
}
function getCatalogDoc() {
  return catalog;
}
function replaceCatalog(doc2) {
  if (!doc2 || !Array.isArray(doc2.items)) return false;
  const items = doc2.items.filter((it) => it && typeof it.id === "string" && typeof it.name === "string");
  catalog = { v: 1, items };
  persistCatalog();
  const known = new Set(items.map((it) => it.id));
  const cur = getState().libraries;
  updateNested("libraries", {
    enabled: cur.enabled.filter((id) => known.has(id)),
    pinned: cur.pinned.filter((id) => known.has(id))
  });
  render();
  onChangeCb?.();
  return true;
}
function unknownLibraryIds(ids) {
  const known = new Set(catalog.items.map((it) => it.id));
  return ids.filter((id) => !known.has(id));
}
function refreshLibraryUI() {
  render();
}

// ../src/io.js
var MAX_IMPORT_BYTES = 5 * 1024 * 1024;
function downloadText(filename, text, type = "application/json") {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1e3);
}
function wireJsonImport(inputId, onDoc) {
  const input = document.getElementById(inputId);
  input.addEventListener("change", async () => {
    const file = input.files?.[0];
    input.value = "";
    if (!file) return;
    if (file.size > MAX_IMPORT_BYTES) {
      alert(`"${file.name}" is ${(file.size / 1048576).toFixed(1)} MB \u2014 too large to be a DCSPad file.`);
      return;
    }
    let doc2;
    try {
      doc2 = JSON.parse(await file.text());
    } catch {
      alert(`"${file.name}" isn't valid JSON.`);
      return;
    }
    onDoc(doc2, file.name);
  });
  return input;
}

// ../src/snippets.js
var doc = null;
var deps3 = {};
function initSnippets({ getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError }) {
  deps3 = { getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError };
  doc = loadDoc(SNIPPETS_KEY) || { v: 1, items: [] };
  render2();
  document.getElementById("btn-snippet-add").addEventListener("click", () => {
    const lang = getState().layout.editorTab;
    const code = deps3.getSelection(lang) || deps3.getDocs()[lang];
    if (!code.trim()) return;
    const name = prompt("Snippet name:");
    if (!name || !name.trim()) return;
    doc.items.push({ id: newId("snip"), name: name.trim(), lang, code, createdAt: Date.now() });
    persist2();
    render2();
  });
  document.getElementById("btn-snippets-export").addEventListener("click", () => {
    downloadText("dcspad-snippets.json", JSON.stringify(doc, null, 2));
  });
  document.getElementById("btn-snippets-import").addEventListener("click", () => {
    document.getElementById("import-snippets-file").click();
  });
  wireJsonImport("import-snippets-file", (imported) => {
    const items = imported && Array.isArray(imported.items) ? imported.items.filter((s) => s && typeof s.name === "string" && typeof s.code === "string" && ["html", "css", "js"].includes(s.lang)) : null;
    if (!items) {
      alert("Not a DCSPad snippet library file.");
      return;
    }
    if (doc.items.length && !confirm(`Replace your ${doc.items.length} snippet(s) with the ${items.length} from this file?`)) return;
    doc = { v: 1, items: items.map((s) => ({ ...s, id: s.id || newId("snip") })) };
    persist2();
    render2();
  });
}
function persist2() {
  if (!saveDoc(SNIPPETS_KEY, doc)) {
    deps3.onStorageError?.("snippet library save failed (storage full?)");
  }
}
function render2() {
  const host = document.getElementById("snippet-list");
  host.textContent = "";
  document.getElementById("snippet-empty").hidden = doc.items.length > 0;
  for (const snip of doc.items) {
    const item = el("div", "lib-item snippet-item");
    const lang = el("span", "snippet-lang", snip.lang);
    const name = el("span", "lib-name", snip.name);
    name.title = `Insert into the ${snip.lang.toUpperCase()} editor at the cursor

${snip.code.slice(0, 400)}`;
    const del = el("span", "lib-del", "\u2715");
    del.title = "Delete snippet";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!confirm(`Delete snippet "${snip.name}"?`)) return;
      doc.items = doc.items.filter((s) => s.id !== snip.id);
      persist2();
      render2();
    });
    item.addEventListener("click", () => {
      deps3.selectEditorTab(snip.lang);
      deps3.insertAtCursor(snip.lang, snip.code);
    });
    item.append(lang, name, del);
    host.append(item);
  }
}

// ../src/bridge/sp-context.js
var cached = null;
function getSpContext({ refresh = false } = {}) {
  if (cached && !refresh) return cached;
  const real = window._spPageContextInfo;
  if (real && real.webAbsoluteUrl) {
    let pageContext;
    try {
      pageContext = JSON.parse(JSON.stringify(real));
    } catch {
      pageContext = {};
      for (const k of ["webAbsoluteUrl", "webServerRelativeUrl", "siteAbsoluteUrl", "siteServerRelativeUrl", "webTitle", "userId", "userLoginName", "userDisplayName", "currentLanguage", "currentCultureName", "layoutsUrl", "webUIVersion", "siteClientTag", "formDigestValue", "formDigestTimeoutSeconds"]) {
        if (real[k] !== void 0) pageContext[k] = real[k];
      }
    }
    const digestEl = document.getElementById("__REQUESTDIGEST");
    if (digestEl?.value) pageContext.formDigestValue = digestEl.value;
    cached = {
      live: true,
      pageContext,
      baseHref: real.webAbsoluteUrl.replace(/\/$/, "") + "/",
      label: real.webAbsoluteUrl,
      user: real.userDisplayName || real.userLoginName || ""
    };
    return cached;
  }
  cached = {
    live: false,
    pageContext: {
      isDcsPadMock: true,
      webAbsoluteUrl: location.origin,
      webServerRelativeUrl: "/",
      siteAbsoluteUrl: location.origin,
      siteServerRelativeUrl: "/",
      webTitle: "DCSPad Mock Web",
      userId: 1,
      userLoginName: "i:0#.f|membership|dev@mock.local",
      userDisplayName: "Mock Developer",
      currentLanguage: 1033,
      currentCultureName: "en-US",
      layoutsUrl: "_layouts/15",
      formDigestValue: "MOCK-DIGEST-0x0000",
      formDigestTimeoutSeconds: 1800
    },
    baseHref: null,
    // keep relative URLs pointed at the local server
    label: "mock (not in SharePoint)",
    user: "Mock Developer"
  };
  return cached;
}
function applyContextIndicators() {
  const ctx = getSpContext();
  const chip = document.getElementById("sp-chip");
  const chipText = document.getElementById("sp-chip-text");
  const statusCtx = document.getElementById("status-context");
  chip.classList.toggle("sp-chip-live", ctx.live);
  chip.classList.toggle("sp-chip-mock", !ctx.live);
  chipText.textContent = ctx.live ? "SP: Live" : "SP: Mock";
  chip.title = ctx.live ? `Connected to ${ctx.label} as ${ctx.user} \u2014 _spPageContextInfo is injected into every run` : "Not hosted in SharePoint \u2014 a mock _spPageContextInfo (correct shape) is injected; _api calls will fail here";
  statusCtx.textContent = ctx.live ? `SP: ${ctx.label} \xB7 ${ctx.user}` : "SP: mock context (deploy to SharePoint for live APIs)";
  return ctx;
}

// ../src/splash.js
var LOGO = String.raw`
 ██████╗  ██████╗ ███████╗ ██████╗  █████╗  ██████╗
 ██╔══██╗██╔════╝ ██╔════╝ ██╔══██╗██╔══██╗ ██╔══██╗
 ██║  ██║██║      ███████╗ ██████╔╝███████║ ██║  ██║
 ██║  ██║██║      ╚════██║ ██╔═══╝ ██╔══██║ ██║  ██║
 ██████╔╝╚██████╗ ███████║ ██║     ██║  ██║ ██████╔╝
 ╚═════╝  ╚═════╝ ╚══════╝ ╚═╝     ╚═╝  ╚═╝ ╚═════╝`.slice(1);
function showSplash() {
  const splash = document.getElementById("splash");
  const logoEl = document.getElementById("splash-logo");
  if (!splash) return;
  const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
  logoEl.textContent = LOGO;
  splash.hidden = false;
  let done = false;
  const finish = () => {
    if (done) return;
    done = true;
    splash.classList.add("fading");
    setTimeout(() => splash.remove(), reduced ? 0 : 400);
    if (!getState().settings.seenSplash) updateNested("settings", { seenSplash: true });
  };
  splash.addEventListener("click", finish);
  setTimeout(finish, reduced ? 150 : 700);
}

// ../src/main.js
var state2 = getState();
var editorsApi = null;
var layoutApi = initLayout({
  onEditorTabChange: (name) => editorsApi?.activate(name)
});
var isDiagVisible = (name) => document.querySelector(`#diag-tabs .tab[data-diag="${name}"]`).classList.contains("active");
editorsApi = await initEditors({
  onChange: () => scheduleAutorun(),
  onRunShortcut: () => run2()
});
var consoleApi = initConsolePanel({
  evalInFrame,
  mapSrcdocLine: mapSrcdocLineToUserJs,
  gotoJsLine: (line) => {
    layoutApi.selectEditorTab("js");
    editorsApi.gotoJsLine(line);
  },
  isConsoleVisible: () => isDiagVisible("console")
});
var networkApi = initNetworkPanel({
  isNetworkVisible: () => isDiagVisible("network")
});
initLibraries({
  onChange: () => {
    scheduleAutorun();
    editorsApi.setPnpTypesEnabled(hasEnabledPnpjs215Runtime());
  },
  onStorageError: (msg) => reportStorageError(msg)
});
editorsApi.setPnpTypesEnabled(hasEnabledPnpjs215Runtime());
initSnippets({
  getSelection: (name) => editorsApi.getSelection(name),
  getDocs: () => editorsApi.getDocs(),
  insertAtCursor: (name, text) => editorsApi.insertAtCursor(name, text),
  selectEditorTab: (name) => layoutApi.selectEditorTab(name),
  onStorageError: (msg) => reportStorageError(msg)
});
var spContext = applyContextIndicators();
showSplash();
var statusRun = document.getElementById("status-run");
var SPINNER = ["\u280B", "\u2819", "\u2839", "\u2838", "\u283C", "\u2834", "\u2826", "\u2827", "\u2807", "\u280F"];
var spinnerTimer = null;
var runnerReady = initRunner({
  ...consoleApi.handlers,
  ...networkApi.handlers,
  loaded: (d) => {
    stopSpinner();
    statusRun.textContent = `ran in ${d.ms} ms`;
    statusRun.className = "status-item ok";
  }
});
function startSpinner() {
  let i = 0;
  statusRun.className = "status-item running";
  clearInterval(spinnerTimer);
  spinnerTimer = setInterval(() => {
    statusRun.textContent = `${SPINNER[i++ % SPINNER.length]} running`;
  }, 80);
}
function stopSpinner() {
  clearInterval(spinnerTimer);
  spinnerTimer = null;
}
async function run2() {
  try {
    await runnerReady;
  } catch (e) {
    statusRun.textContent = e.message;
    statusRun.className = "status-item error";
    return;
  }
  const settings = getState().settings;
  if (settings.autoClearConsole) consoleApi.clear();
  markRun();
  startSpinner();
  document.getElementById("btn-run").classList.remove("running");
  void document.getElementById("btn-run").offsetWidth;
  document.getElementById("btn-run").classList.add("running");
  const panel = document.getElementById("preview-panel");
  panel.classList.remove("sweeping");
  void panel.offsetWidth;
  panel.classList.add("sweeping");
  const { runNumber } = run({
    docs: editorsApi.getDocs(),
    libraries: getEnabledLibraries(),
    // Re-capture per run: on classic pages the host rewrites the
    // #__REQUESTDIGEST form field, and a bootstrap-time digest expires.
    spContext: getSpContext({ refresh: true }),
    settings
  });
  if (!settings.autoClearConsole) consoleApi.runDivider(runNumber);
  setTimeout(() => {
    if (spinnerTimer) {
      stopSpinner();
      statusRun.textContent = "still loading\u2026";
      statusRun.className = "status-item";
    }
  }, 15e3);
}
document.getElementById("btn-run").addEventListener("click", run2);
document.getElementById("btn-rerun").addEventListener("click", run2);
var btnPreviewTheme = document.getElementById("btn-preview-theme");
function applyPreviewTheme() {
  const dark = getState().settings.previewDark;
  btnPreviewTheme.textContent = dark ? "\u2600" : "\u{1F319}";
  btnPreviewTheme.title = dark ? "Switch preview to light \u2014 pad-only canvas color; your CSS still wins, and SharePoint pages are typically light" : "Switch preview to dark \u2014 pad-only canvas color; your CSS still wins";
  document.getElementById("preview-host").classList.toggle("dark", dark);
}
applyPreviewTheme();
btnPreviewTheme.addEventListener("click", () => {
  updateNested("settings", { previewDark: !getState().settings.previewDark });
  applyPreviewTheme();
  if (hasRun()) run2();
});
var AUTORUN_DEBOUNCE_MS = 800;
var autorunTimer = null;
var chkAutorun = document.getElementById("chk-autorun");
chkAutorun.checked = state2.settings.autorun;
document.getElementById("live-dot").classList.toggle("on", state2.settings.autorun);
chkAutorun.addEventListener("change", () => {
  updateNested("settings", { autorun: chkAutorun.checked });
  document.getElementById("live-dot").classList.toggle("on", chkAutorun.checked);
  if (chkAutorun.checked) scheduleAutorun();
});
function scheduleAutorun() {
  if (!getState().settings.autorun) return;
  clearTimeout(autorunTimer);
  autorunTimer = setTimeout(run2, AUTORUN_DEBOUNCE_MS);
}
var menus = [
  { btn: document.getElementById("btn-settings"), menu: document.getElementById("settings-menu") },
  { btn: document.getElementById("btn-file"), menu: document.getElementById("file-menu") }
];
for (const { btn, menu } of menus) {
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    for (const m of menus) m.menu.hidden = true;
    menu.hidden = !open;
  });
}
document.addEventListener("click", (e) => {
  for (const { menu } of menus) {
    if (!menu.hidden && !menu.contains(e.target)) menu.hidden = true;
  }
});
var closeFileMenu = () => {
  document.getElementById("file-menu").hidden = true;
};
function padWarn(msg) {
  consoleApi.handlers.console({ level: "warn", args: [{ t: "str", v: `DCSPad: ${msg}` }] });
}
document.getElementById("mi-save-project").addEventListener("click", () => {
  closeFileMenu();
  const s = getState();
  const file = {
    app: "dcspad",
    kind: "project",
    v: 1,
    savedAt: (/* @__PURE__ */ new Date()).toISOString(),
    docs: editorsApi.getDocs(),
    libraries: { enabled: s.libraries.enabled, dcsUrl: s.libraries.dcsUrl },
    jsAsModule: s.settings.jsAsModule
  };
  downloadText("dcspad-project.json", JSON.stringify(file, null, 2));
});
document.getElementById("mi-load-project").addEventListener("click", () => {
  closeFileMenu();
  document.getElementById("import-project-file").click();
});
wireJsonImport("import-project-file", (doc2) => {
  if (!doc2 || doc2.kind !== "project" || typeof doc2.docs !== "object" || doc2.docs === null) {
    alert("Not a DCSPad project file.");
    return;
  }
  const str2 = (v) => typeof v === "string" ? v : "";
  editorsApi.setDocs({ html: str2(doc2.docs.html), css: str2(doc2.docs.css), js: str2(doc2.docs.js) });
  const libs = doc2.libraries || {};
  const enabled = Array.isArray(libs.enabled) ? libs.enabled.filter((id) => typeof id === "string") : [];
  updateNested("libraries", {
    enabled,
    ...typeof libs.dcsUrl === "string" ? { dcsUrl: libs.dcsUrl } : {}
  });
  if (typeof doc2.jsAsModule === "boolean") {
    updateNested("settings", { jsAsModule: doc2.jsAsModule });
    document.getElementById("chk-module").checked = doc2.jsAsModule;
    editorsApi.setJsAsModule(doc2.jsAsModule);
  }
  refreshLibraryUI();
  editorsApi.setPnpTypesEnabled(hasEnabledPnpjs215Runtime());
  const missing = unknownLibraryIds(enabled);
  if (missing.length) {
    padWarn(`this project references framework(s) not in your catalog: ${missing.join(", ")} \u2014 re-add them under Frameworks, or the run will fail where they're used`);
  }
  statusRun.textContent = "project loaded \u2014 press Run";
  statusRun.className = "status-item";
});
var PANE_EXPORTS = [
  ["mi-export-html", "html", "dcspad.html", "text/html"],
  ["mi-export-css", "css", "dcspad.css", "text/css"],
  ["mi-export-js", "js", "dcspad.js", "text/javascript"]
];
for (const [id, pane, filename, type] of PANE_EXPORTS) {
  document.getElementById(id).addEventListener("click", () => {
    closeFileMenu();
    downloadText(filename, editorsApi.getDocs()[pane], type);
  });
}
document.getElementById("btn-catalog-export").addEventListener("click", () => {
  downloadText("dcspad-catalog.json", JSON.stringify(getCatalogDoc(), null, 2));
});
document.getElementById("btn-catalog-import").addEventListener("click", () => {
  document.getElementById("import-catalog-file").click();
});
wireJsonImport("import-catalog-file", (doc2) => {
  if (!doc2 || !Array.isArray(doc2.items)) {
    alert("Not a DCSPad catalog file.");
    return;
  }
  const cur = getCatalogDoc().items.length;
  if (!confirm(`Replace your framework catalog (${cur} entries) with this file (${doc2.items.length} entries)?`)) return;
  replaceCatalog(doc2);
});
var chkModule = document.getElementById("chk-module");
chkModule.checked = state2.settings.jsAsModule;
chkModule.addEventListener("change", () => {
  updateNested("settings", { jsAsModule: chkModule.checked });
  editorsApi.setJsAsModule(chkModule.checked);
});
var chkAutoclear = document.getElementById("chk-autoclear");
chkAutoclear.checked = state2.settings.autoClearConsole;
chkAutoclear.addEventListener("change", () => updateNested("settings", { autoClearConsole: chkAutoclear.checked }));
var DIAG_FS_MIN = 10;
var DIAG_FS_MAX = 18;
function applyDiagFontSize(px2) {
  document.documentElement.style.setProperty("--diag-fs", `${px2}px`);
  document.getElementById("diag-font-val").textContent = String(px2);
}
applyDiagFontSize(state2.settings.diagFontSize);
function stepDiagFontSize(delta) {
  const cur = getState().settings.diagFontSize;
  const next = Math.min(DIAG_FS_MAX, Math.max(DIAG_FS_MIN, cur + delta));
  if (next === cur) return;
  updateNested("settings", { diagFontSize: next });
  applyDiagFontSize(next);
}
document.getElementById("btn-diag-font-dec").addEventListener("click", () => stepDiagFontSize(-1));
document.getElementById("btn-diag-font-inc").addEventListener("click", () => stepDiagFontSize(1));
var saveEl = document.getElementById("status-save");
onSaveStatus((status) => {
  saveEl.classList.remove("saved", "error");
  if (status === "dirty") {
    saveEl.textContent = "saving\u2026";
  } else if (status === "error") {
    saveEl.textContent = "save failed \u2014 use File \u25B8 Save project";
    saveEl.classList.add("error");
  } else {
    saveEl.textContent = "\u2713 saved";
    saveEl.classList.add("saved");
  }
});
function reportStorageError(msg) {
  padWarn(`${msg} \u2014 use the \u2913 export buttons or File \u25B8 Save project to keep your work`);
  saveEl.textContent = "save failed";
  saveEl.classList.add("error");
  saveEl.classList.remove("saved");
}
