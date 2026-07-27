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
  settings: {
    autorun: false,
    jsAsModule: false,
    autoClearConsole: true,
    seenSplash: false,
    previewDark: true,
    diagFontSize: 12,
    editorFontSize: 13,
    wordWrap: false
  },
  layout: {
    sidebarW: 230,
    sidebarCollapsed: false,
    editorsFr: 1,
    runtimeFr: 1,
    previewFr: 1,
    diagH: 260,
    diagCollapsed: false,
    editorTab: "js",
    diagTab: "console",
    // Pane visibility (the topbar segmented toggles). sidebarCollapsed /
    // diagCollapsed above are legacy flags kept for shape stability: layout.js
    // reads them once to seed `panes` for pre-existing workspaces, then only
    // writes `panes`.
    panes: { resources: true, preview: true, console: true },
    snippetsPanelH: 210
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
  selectEditorTab(layout.editorTab, { silent: true });
  selectDiagTab(layout.diagTab);
  const panes = { ...layout.panes };
  if (layout.sidebarCollapsed) panes.resources = false;
  if (layout.diagCollapsed) panes.console = false;
  if (layout.sidebarCollapsed || layout.diagCollapsed) {
    updateNested("layout", { panes: { ...panes }, sidebarCollapsed: false, diagCollapsed: false });
  }
  function applyPanes() {
    main.classList.toggle("hide-resources", !panes.resources);
    main.classList.toggle("hide-preview", !panes.preview);
    main.classList.toggle("hide-console", !panes.console);
    for (const name of ["resources", "preview", "console"]) {
      const seg = document.getElementById(`seg-${name}`);
      seg.classList.toggle("active", !!panes[name]);
      seg.setAttribute("aria-pressed", String(!!panes[name]));
    }
  }
  applyPanes();
  function setPaneVisible(name, on) {
    panes[name] = !!on;
    applyPanes();
    updateNested("layout", { panes: { ...panes } });
  }
  function togglePane(name) {
    setPaneVisible(name, !panes[name]);
  }
  document.getElementById("pane-toggles").addEventListener("click", (e) => {
    const seg = e.target.closest(".pane-seg");
    if (seg) togglePane(seg.dataset.pane);
  });
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
  const DEFAULT_SNIPPETS_H = 210;
  root.style.setProperty("--snippets-h", px(layout.snippetsPanelH || DEFAULT_SNIPPETS_H));
  const splitSide = document.getElementById("split-side");
  dragSplitter(splitSide, "y", (dy, start) => {
    const max = start.sidebarH - 160;
    const h = Math.min(max, Math.max(150, start.snippetsH - dy));
    root.style.setProperty("--snippets-h", px(h));
    updateNested("layout", { snippetsPanelH: h });
  }, () => ({
    snippetsH: document.getElementById("panel-snippets").getBoundingClientRect().height,
    sidebarH: document.getElementById("sidebar").getBoundingClientRect().height
  }));
  splitSide.addEventListener("dblclick", () => {
    root.style.setProperty("--snippets-h", px(DEFAULT_SNIPPETS_H));
    updateNested("layout", { snippetsPanelH: DEFAULT_SNIPPETS_H });
  });
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
    if (tab) selectDiagTab(tab.dataset.diag);
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
    main.classList.remove("max-diag", "max-editor");
    main.classList.toggle("max-preview");
  });
  document.getElementById("btn-max-diag").addEventListener("click", () => {
    main.classList.remove("max-preview", "max-editor");
    main.classList.toggle("max-diag");
  });
  document.getElementById("btn-max-editor").addEventListener("click", () => {
    main.classList.remove("max-preview", "max-diag");
    main.classList.toggle("max-editor");
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      main.classList.remove("max-preview", "max-diag", "max-editor");
      return;
    }
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    if (e.code === "Backslash") {
      e.preventDefault();
      togglePane(e.shiftKey ? "preview" : "resources");
    } else if (e.code === "KeyJ" && !e.shiftKey) {
      e.preventDefault();
      togglePane("console");
    }
  });
  return { selectEditorTab, selectDiagTab, togglePane, setPaneVisible };
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
function setEditorStatus(text, state3 = "", title = "") {
  const status = document.getElementById("status-editor");
  if (!status) return;
  status.textContent = text;
  status.className = `status-item status-editor${state3 ? ` ${state3}` : ""}`;
  status.title = title;
}
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
        setEditorStatus(
          "Monaco \u26A0",
          "warning",
          `${label || "editor"} worker unavailable \u2014 language tools limited`
        );
      }, { once: true });
      return worker;
    }
  };
}
function loadMonacoRuntime() {
  if (runtimePromise) return runtimePromise;
  runtimePromise = (async () => {
    setEditorStatus("Monaco \u2026", "", "Loading Monaco editor");
    configureWorkers();
    await ensureStylesheet();
    const monaco = await import(assetUrl("monaco.js"));
    document.documentElement.dataset.monacoReady = "true";
    setEditorStatus("Monaco \u2713", "ok", "Monaco editor ready");
    return monaco;
  })().catch((error) => {
    setEditorStatus("Monaco \u2715", "warning", error.message || "Monaco failed to load");
    throw error;
  });
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

// ../src/intelligence/alpine.js
var ALPINE_PACK_ID = "alpine-3";
var ALPINE_JS_LIBRARIES = [{
  filePath: "file:///node_modules/@types/dcspad-alpine/index.d.ts",
  content: `/**
 * Alpine.js v3 public browser API.
 * @see https://alpinejs.dev/globals/
 */
interface AlpineMagicProperties {
  /** The current DOM element. */
  readonly $el: HTMLElement;
  /** Elements marked with x-ref in the current component. */
  readonly $refs: Record<string, HTMLElement>;
  /** Global stores registered with Alpine.store(). */
  readonly $store: Record<string, any>;
  /** The nearest Alpine component root. */
  readonly $root: HTMLElement;
  /** The current merged Alpine data scope. */
  readonly $data: Record<string, any>;
  /** Watch a dot-notated component property for changes. */
  $watch<T = any>(property: string, callback: (value: T, oldValue: T) => void): () => void;
  /** Dispatch a bubbling CustomEvent from the current element. */
  $dispatch(name: string, detail?: any): boolean;
  /** Run work after Alpine has flushed reactive DOM updates. */
  $nextTick(callback?: () => void): Promise<void>;
  /** Generate a component-scoped element id. */
  $id(name: string, key?: string | number): string;
}

type AlpinePlugin = (alpine: AlpineStatic) => void;
type AlpineDataProvider<T extends object = Record<string, any>> =
  (...parameters: any[]) => T & ThisType<T & AlpineMagicProperties>;
type AlpineBindProvider =
  () => Record<string, any> & ThisType<Record<string, any> & AlpineMagicProperties>;

interface AlpineStatic {
  /** The loaded Alpine runtime version. */
  readonly version: string;
  /** Start Alpine and initialize the current document. */
  start(): void;

  /**
   * Register a reusable x-data provider.
   * @see https://alpinejs.dev/globals/alpine-data
   */
  data<T extends object>(name: string, provider: AlpineDataProvider<T>): void;

  /**
   * Read a global store.
   * @see https://alpinejs.dev/globals/alpine-store
   */
  store<T = any>(name: string): T;
  /** Register or replace a global store. */
  store<T extends object>(name: string, value: T & ThisType<T & AlpineMagicProperties>): T;

  /**
   * Register a reusable x-bind object.
   * @see https://alpinejs.dev/globals/alpine-bind
   */
  bind(name: string, provider: AlpineBindProvider): void;

  /** Install one or more Alpine plugins before Alpine.start(). */
  plugin(plugin: AlpinePlugin | AlpinePlugin[]): void;
  /** Register a custom magic property. */
  magic(name: string, callback: (element: HTMLElement, utilities: Record<string, any>) => any): void;
  /** Register a custom x-* directive. */
  directive(name: string, callback: (...args: any[]) => void): { before(other: string): void };
  /** Change or return Alpine's directive prefix. */
  prefix(prefix: string): void;
  prefixed(subject?: string): string;

  /** Create a deeply reactive proxy. */
  reactive<T extends object>(target: T): T;
  /** Return the original object behind a reactive proxy. */
  raw<T>(value: T): T;
  /** Run and track a reactive effect. */
  effect(callback: () => void): any;
  /** Stop a reactive effect. */
  release(effect: any): void;
  /** Watch a reactive getter. */
  watch<T>(getter: () => T, callback: (value: T, oldValue: T) => void): () => void;
  /** Schedule work after Alpine's next DOM update. */
  nextTick(callback?: () => void): Promise<void>;

  /** Return the merged Alpine data scope for an element. */
  $data<T extends object = Record<string, any>>(element: Element): T;
  /** Initialize Alpine behavior under an element added dynamically. */
  initTree(element: Element): void;
  /** Tear down Alpine behavior under an element. */
  destroyTree(element: Element): void;
}

declare const Alpine: AlpineStatic;

interface Window {
  Alpine: AlpineStatic;
}
`
}];
var directive = (name, description, {
  insertText = `${name}="\${1:expression}"`,
  url = `https://alpinejs.dev/directives/${name.slice(2).split(":")[0]}`
} = {}) => ({ name, description, insertText, url });
var ALPINE_DIRECTIVES = [
  directive("x-data", "Declares a new Alpine component and its reactive state.", {
    insertText: `x-data="{ \${1:open}: \${2:false} }"`
  }),
  directive("x-init", "Runs an expression while Alpine initializes the element."),
  directive("x-show", "Toggles element visibility from a truthy expression."),
  directive("x-bind", "Binds an object of attributes, or a reusable Alpine.bind() provider."),
  directive("x-bind:class", "Reactively binds the element class attribute."),
  directive("x-bind:aria-expanded", "Reactively binds aria-expanded.", {
    insertText: `x-bind:aria-expanded="\${1:open}"`,
    url: "https://alpinejs.dev/directives/bind"
  }),
  directive("x-bind:aria-selected", "Reactively binds aria-selected.", {
    insertText: `x-bind:aria-selected="\${1:selected}"`,
    url: "https://alpinejs.dev/directives/bind"
  }),
  directive("x-bind:aria-pressed", "Reactively binds aria-pressed.", {
    insertText: `x-bind:aria-pressed="\${1:pressed}"`,
    url: "https://alpinejs.dev/directives/bind"
  }),
  directive("x-on", "Attaches an event listener. Add an event name, such as x-on:click."),
  directive("x-on:click", "Runs an expression when the element is clicked."),
  directive("x-on:submit.prevent", "Prevents form submission and runs an expression.", {
    url: "https://alpinejs.dev/directives/on"
  }),
  directive("x-on:click.outside", "Runs when a click occurs outside the element.", {
    url: "https://alpinejs.dev/directives/on"
  }),
  directive("x-on:keydown.escape.window", "Runs on Escape keydown from window.", {
    url: "https://alpinejs.dev/directives/on"
  }),
  directive("x-text", "Sets textContent from an expression."),
  directive("x-html", "Sets innerHTML from an expression. Only use trusted content."),
  directive("x-model", "Creates two-way binding between form state and component data."),
  directive("x-modelable", "Exposes an internal property to an outer x-model binding."),
  directive("x-for", "Repeats a template for each item in an iterable.", {
    insertText: `x-for="(\${1:item}, \${2:index}) in \${3:items}"`
  }),
  directive("x-transition", "Adds Alpine transition classes around x-show changes.", {
    insertText: "x-transition"
  }),
  directive("x-effect", "Re-runs an expression whenever its reactive dependencies change."),
  directive("x-ignore", "Prevents Alpine from initializing this element subtree.", {
    insertText: "x-ignore"
  }),
  directive("x-ref", "Names an element for access through $refs.", {
    insertText: `x-ref="\${1:name}"`
  }),
  directive("x-cloak", "Keeps an element hidden until Alpine initializes it.", {
    insertText: "x-cloak"
  }),
  directive("x-teleport", "Moves a template to another DOM location.", {
    insertText: `x-teleport="\${1:body}"`
  }),
  directive("x-if", "Conditionally adds or removes a template from the DOM."),
  directive("x-id", "Declares names used by the component-scoped $id() helper.", {
    insertText: `x-id="['\${1:control}']"`
  })
];
var shorthand = (name, insertText, description, url) => ({
  name,
  insertText,
  description,
  url
});
var ALPINE_SHORTHANDS = [
  shorthand("@click", `@click="\${1:expression}"`, "Shorthand for x-on:click.", "https://alpinejs.dev/directives/on"),
  shorthand("@click.outside", `@click.outside="\${1:open = false}"`, "Runs when clicking outside the element.", "https://alpinejs.dev/directives/on"),
  shorthand("@submit.prevent", `@submit.prevent="\${1:submit()}"`, "Prevents submission and runs an expression.", "https://alpinejs.dev/directives/on"),
  shorthand("@keydown.escape.window", `@keydown.escape.window="\${1:open = false}"`, "Runs on Escape from window.", "https://alpinejs.dev/directives/on"),
  shorthand(":class", `:class="{ '\${1:is-active}': \${2:active} }"`, "Shorthand for x-bind:class.", "https://alpinejs.dev/directives/bind"),
  shorthand(":disabled", `:disabled="\${1:disabled}"`, "Shorthand for x-bind:disabled.", "https://alpinejs.dev/directives/bind"),
  shorthand(":aria-expanded", `:aria-expanded="\${1:open}"`, "Shorthand for x-bind:aria-expanded.", "https://alpinejs.dev/directives/bind"),
  shorthand(":aria-selected", `:aria-selected="\${1:selected}"`, "Shorthand for x-bind:aria-selected.", "https://alpinejs.dev/directives/bind"),
  shorthand(":aria-pressed", `:aria-pressed="\${1:pressed}"`, "Shorthand for x-bind:aria-pressed.", "https://alpinejs.dev/directives/bind")
];
var magic = (name, description, insertText = name) => ({
  name,
  description,
  insertText,
  url: name === "$event" ? "https://alpinejs.dev/directives/on" : `https://alpinejs.dev/magics/${name.slice(1).toLowerCase()}`
});
var ALPINE_MAGICS = [
  magic("$el", "The current DOM element."),
  magic("$refs", "Elements marked with x-ref in the current component.", `$refs.\${1:name}`),
  magic("$store", "Global stores registered with Alpine.store().", `$store.\${1:name}`),
  magic("$watch", "Watches a component property for changes.", `$watch('\${1:property}', (\${2:value}, \${3:oldValue}) => { \${0} })`),
  magic("$dispatch", "Dispatches a bubbling CustomEvent from the current element.", `$dispatch('\${1:event}', { \${0} })`),
  magic("$nextTick", "Runs work after Alpine flushes reactive DOM updates.", `$nextTick(() => { \${0} })`),
  magic("$root", "The nearest Alpine component root element."),
  magic("$data", "The current merged Alpine data scope."),
  magic("$id", "Generates a component-scoped element id.", `$id('\${1:name}')`),
  magic("$event", "The native event available inside x-on expressions.")
];
var htmlAttribute = (item) => ({
  name: item.name,
  description: item.description,
  references: [{ name: "Alpine.js documentation", url: item.url }]
});
var ALPINE_HTML_DATA = {
  version: 1.1,
  globalAttributes: [
    ...ALPINE_DIRECTIVES.map(htmlAttribute),
    ...ALPINE_SHORTHANDS.map(htmlAttribute)
  ]
};
function completionRange(monaco, model, position, prefix) {
  const end = model.getOffsetAt(position);
  const start = Math.max(0, end - prefix.length);
  return new monaco.Range(
    model.getPositionAt(start).lineNumber,
    model.getPositionAt(start).column,
    position.lineNumber,
    position.column
  );
}
function alpineAttributeValue(fragment) {
  const match = fragment.match(
    /(?:^|\s)(x-[\w:.-]+|@[\w:.-]+|:[\w:.-]+)\s*=\s*(["'])([\s\S]*)$/
  );
  if (!match || match[3].includes(match[2])) return null;
  return match[3];
}
function createAlpineHtmlCompletionProvider(monaco, isEnabled) {
  return {
    triggerCharacters: ["x", "-", "@", ":", "$", "."],
    provideCompletionItems(model, position) {
      if (!isEnabled()) return { suggestions: [] };
      const offset = model.getOffsetAt(position);
      const before = model.getValue().slice(Math.max(0, offset - 6e3), offset);
      const lastOpen = before.lastIndexOf("<");
      const lastClose = before.lastIndexOf(">");
      if (lastOpen <= lastClose) return { suggestions: [] };
      const fragment = before.slice(lastOpen + 1);
      const value = alpineAttributeValue(fragment);
      if (value !== null) {
        const magicPrefix = value.match(/\$[A-Za-z]*$/)?.[0];
        if (!magicPrefix) return { suggestions: [] };
        const range2 = completionRange(monaco, model, position, magicPrefix);
        return {
          suggestions: ALPINE_MAGICS.map((item) => ({
            label: item.name,
            kind: monaco.languages.CompletionItemKind.Variable,
            detail: "Alpine magic property",
            documentation: {
              value: `${item.description}

[Alpine.js documentation](${item.url})`
            },
            insertText: item.insertText,
            insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
            range: range2
          }))
        };
      }
      const prefix = fragment.match(/(?:^|\s)(x-[^\s"'=<>]*|@[^\s"'=<>]*|:[^\s"'=<>]*)$/)?.[1];
      if (!prefix) return { suggestions: [] };
      const source = prefix.startsWith("x-") ? ALPINE_DIRECTIVES : ALPINE_SHORTHANDS;
      const range = completionRange(monaco, model, position, prefix);
      return {
        suggestions: source.map((item) => ({
          label: item.name,
          kind: monaco.languages.CompletionItemKind.Property,
          detail: prefix.startsWith("x-") ? "Alpine directive" : "Alpine shorthand",
          documentation: {
            value: `${item.description}

[Alpine.js documentation](${item.url})`
          },
          insertText: item.insertText,
          filterText: item.name,
          insertTextRules: monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet,
          range
        }))
      };
    }
  };
}

// ../src/intelligence/bsp.js
var BSP_PACK_ID = "bsp-design";
var dataPromise = null;
function artifactUrl() {
  const root = window.__DCSPAD_ASSET_BASE__ || new URL("../../", import.meta.url).href;
  const url = new URL("vendor/intelligence/bsp-design.json", root);
  const version = window.__DCSPAD_INTELLIGENCE_VERSION__;
  if (version) url.searchParams.set("v", version);
  return url.href;
}
function prepareData(raw) {
  if (raw?.schemaVersion !== 1 || raw?.pack !== BSP_PACK_ID || !Array.isArray(raw.tokens) || !Array.isArray(raw.classes)) {
    throw new Error("unsupported or malformed bsp-design.json");
  }
  return {
    ...raw,
    tokenByName: new Map(raw.tokens.map((token) => [token.name, token])),
    classByName: new Map(raw.classes.map((item) => [item.name, item]))
  };
}
function fetchBspIntelligence() {
  if (!dataPromise) {
    dataPromise = fetch(artifactUrl(), {
      credentials: "same-origin",
      cache: "no-cache"
    }).then((response) => {
      if (!response.ok) throw new Error(`HTTP ${response.status} loading BSP intelligence`);
      return response.json();
    }).then(prepareData).catch((error) => {
      dataPromise = null;
      throw error;
    });
  }
  return dataPromise;
}
function markdownToken(token) {
  const scope = token.scope === "editorial" ? "\n\nScoped to **`.editorial`**." : "";
  return [
    `**\`${token.name}\`** \xB7 ${token.category}`,
    token.description,
    `Value: \`${token.value}\`${scope}`,
    `Source: \`${token.source.file}:${token.source.line}\``
  ].filter(Boolean).join("\n\n");
}
function markdownClass(item) {
  const labels = {
    base: "base/component",
    element: "BEM element",
    modifier: "BEM modifier",
    state: "state",
    utility: "utility"
  };
  const requirements = item.base ? `

Compose with **\`.${item.base}\`**.` : "";
  const scope = item.scopes?.includes("editorial") ? "\n\nAvailable in **Editorial mode**." : "";
  return [
    `**\`.${item.name}\`** \xB7 ${labels[item.kind] || item.kind}`,
    `${item.description}${requirements}${scope}`,
    `Source: \`${item.source.file}:${item.source.line}\``
  ].filter(Boolean).join("\n\n");
}
function cssTokenAt(model, position) {
  const line = model.getLineContent(position.lineNumber);
  const offset = position.column - 1;
  const pattern = /--[\w-]+/g;
  let match;
  while (match = pattern.exec(line)) {
    if (offset >= match.index && offset <= match.index + match[0].length) {
      return {
        name: match[0],
        range: {
          startLineNumber: position.lineNumber,
          startColumn: match.index + 1,
          endLineNumber: position.lineNumber,
          endColumn: match.index + match[0].length + 1
        }
      };
    }
  }
  return null;
}
function cssCompletionContext(model, position) {
  const before = model.getLineContent(position.lineNumber).slice(0, position.column - 1);
  const variable = before.match(/var\(\s*(--[\w-]*)?$/);
  const declaration = before.match(/^\s*(--[\w-]*)$/);
  const prefix = variable ? variable[1] || "" : declaration?.[1];
  if (prefix === void 0) return null;
  return {
    prefix,
    range: {
      startLineNumber: position.lineNumber,
      startColumn: position.column - prefix.length,
      endLineNumber: position.lineNumber,
      endColumn: position.column
    }
  };
}
function createBspCssCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: ["-", "("],
    provideCompletionItems(model, position) {
      const data = getData();
      const context = data && cssCompletionContext(model, position);
      if (!context) return { suggestions: [] };
      return {
        suggestions: data.tokens.map((token) => ({
          label: token.name,
          kind: monaco.languages.CompletionItemKind.Variable,
          detail: `${token.category} \xB7 ${token.value}`,
          documentation: { value: markdownToken(token) },
          insertText: token.name,
          filterText: token.name,
          sortText: token.name,
          range: context.range
        }))
      };
    }
  };
}
function createBspCssHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && cssTokenAt(model, position);
      const token = target && data.tokenByName.get(target.name);
      if (!token) return null;
      return {
        range: target.range,
        contents: [{ value: markdownToken(token) }]
      };
    }
  };
}
function classAttributeBeforeCursor(model, position) {
  const offset = model.getOffsetAt(position);
  const start = Math.max(0, offset - 6e3);
  const before = model.getValue().slice(start, offset);
  const match = before.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const value = match[2];
  const prefix = value.match(/[^\s]*$/)?.[0] || "";
  return {
    value,
    prefix,
    startOffset: offset - prefix.length
  };
}
function classTokenAt(model, position) {
  const value = model.getValue();
  const offset = model.getOffsetAt(position);
  const before = value.slice(Math.max(0, offset - 6e3), offset);
  const open = before.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!open) return null;
  const quote = open[1];
  const valueStart = offset - open[2].length;
  const close = value.indexOf(quote, offset);
  if (close < 0) return null;
  let start = offset;
  let end = offset;
  while (start > valueStart && !/\s/.test(value[start - 1])) start--;
  while (end < close && !/\s/.test(value[end])) end++;
  const name = value.slice(start, end);
  if (!name) return null;
  const startPosition = model.getPositionAt(start);
  const endPosition = model.getPositionAt(end);
  return {
    name,
    range: {
      startLineNumber: startPosition.lineNumber,
      startColumn: startPosition.column,
      endLineNumber: endPosition.lineNumber,
      endColumn: endPosition.column
    }
  };
}
function createBspHtmlClassCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: [" ", "-", "_"],
    provideCompletionItems(model, position) {
      const data = getData();
      const context = data && classAttributeBeforeCursor(model, position);
      if (!context) return { suggestions: [] };
      const present = new Set(context.value.trim().split(/\s+/).filter(Boolean));
      present.delete(context.prefix);
      const startPosition = model.getPositionAt(context.startOffset);
      const range = new monaco.Range(
        startPosition.lineNumber,
        startPosition.column,
        position.lineNumber,
        position.column
      );
      return {
        suggestions: data.classes.filter((item) => !present.has(item.name)).map((item) => ({
          label: item.name,
          kind: monaco.languages.CompletionItemKind.Class,
          detail: `BMO design system \xB7 ${item.kind}`,
          documentation: { value: markdownClass(item) },
          insertText: item.name,
          filterText: item.name,
          sortText: `${item.kind === "base" ? "0" : "1"}-${item.name}`,
          range
        }))
      };
    }
  };
}
function createBspHtmlClassHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && classTokenAt(model, position);
      const item = target && data.classByName.get(target.name);
      if (!item) return null;
      return {
        range: target.range,
        contents: [{ value: markdownClass(item) }]
      };
    }
  };
}

// ../src/intelligence/fluent-icons.js
var FLUENT_ICONS_PACK_ID = "fluent-icons";
var FLUENT_ICONS_HTML_DATA = {
  version: 1.1,
  tags: [{
    name: "fluent-icon",
    description: [
      "Font-backed Fluent System Icon. The name uses",
      "`{icon-name}-{source-size}-{regular|filled|light}`.",
      "Decorative by default; add `label` when the icon itself conveys meaning."
    ].join(" "),
    attributes: [
      {
        name: "name",
        description: "A real Fluent icon token, for example `home-24-regular`."
      },
      {
        name: "label",
        description: "Accessible label for a meaningful icon. Omit for decorative icons."
      }
    ]
  }]
};
var dataPromise2 = null;
function artifactUrl2() {
  const root = window.__DCSPAD_ASSET_BASE__ || new URL("../../", import.meta.url).href;
  const url = new URL("vendor/intelligence/fluent-icons.json", root);
  const version = window.__DCSPAD_INTELLIGENCE_VERSION__;
  if (version) url.searchParams.set("v", version);
  return url.href;
}
function variantPreference(variant) {
  return (variant.size === 24 ? 0 : variant.size === 20 ? 1 : 2) + (variant.style === "regular" ? 0 : variant.style === "filled" ? 0.1 : variant.style === "light" ? 0.2 : 0.3);
}
function prepareData2(raw) {
  if (raw?.schemaVersion !== 1 || raw?.pack !== FLUENT_ICONS_PACK_ID || !Array.isArray(raw.icons)) {
    throw new Error("unsupported or malformed fluent-icons.json");
  }
  const variants = [];
  const variantByToken = /* @__PURE__ */ new Map();
  const variantById = /* @__PURE__ */ new Map();
  const variantByClass = /* @__PURE__ */ new Map();
  const defaultVariants = [];
  const defaultFontVariants = [];
  const fontVariants = [];
  for (const icon of raw.icons) {
    if (!icon?.slug || !icon?.idBase || !Array.isArray(icon.variants)) continue;
    const iconVariants = [];
    const svgOnly = new Set(icon.svgOnly || []);
    for (const suffix of icon.variants) {
      const match = String(suffix).match(
        /^(\d+)-(regular|filled|light|color)(?:-(ltr|rtl))?$/
      );
      if (!match) continue;
      const size = Number(match[1]);
      const style = match[2];
      const direction = match[3] || "";
      const token = `${icon.slug}-${size}-${style}`;
      const directionalToken = direction ? `${token}-${direction}` : token;
      const id = `ic_fluent_${icon.idBase}_${size}_${style}${direction ? `_${direction}` : ""}`;
      const className = `icon-${id}`;
      const variant = {
        name: icon.name,
        slug: icon.slug,
        description: icon.description || "",
        metaphors: icon.metaphors || [],
        size,
        style,
        direction,
        token: directionalToken,
        id,
        className,
        filename: `${id}.svg`,
        fontAvailable: !svgOnly.has(suffix)
      };
      variant.searchText = [
        icon.name,
        icon.slug,
        icon.description,
        ...icon.metaphors || [],
        directionalToken
      ].filter(Boolean).join(" ").toLowerCase();
      variants.push(variant);
      iconVariants.push(variant);
      variantByToken.set(directionalToken, variant);
      variantById.set(id, variant);
      if (variant.fontAvailable) {
        fontVariants.push(variant);
        variantByClass.set(className, variant);
      }
    }
    iconVariants.sort((a, b) => variantPreference(a) - variantPreference(b) || a.token.localeCompare(b.token));
    if (iconVariants[0]) defaultVariants.push(iconVariants[0]);
    const defaultFont = iconVariants.find((variant) => variant.fontAvailable);
    if (defaultFont) defaultFontVariants.push(defaultFont);
  }
  defaultVariants.sort((a, b) => a.name.localeCompare(b.name));
  defaultFontVariants.sort((a, b) => a.name.localeCompare(b.name));
  return {
    ...raw,
    variants,
    fontVariants,
    defaultVariants,
    defaultFontVariants,
    variantByToken,
    variantById,
    variantByClass
  };
}
function fetchFluentIconIntelligence() {
  if (!dataPromise2) {
    dataPromise2 = fetch(artifactUrl2(), {
      credentials: "same-origin",
      cache: "no-cache"
    }).then((response) => {
      if (!response.ok) throw new Error(`HTTP ${response.status} loading Fluent icon intelligence`);
      return response.json();
    }).then(prepareData2).catch((error) => {
      dataPromise2 = null;
      throw error;
    });
  }
  return dataPromise2;
}
function markdownVariant(variant) {
  const summary = variant.description || (variant.metaphors.length ? `Related concepts: ${variant.metaphors.slice(0, 8).join(", ")}.` : "Microsoft Fluent System Icon.");
  const fontMarkup = variant.fontAvailable ? [
    "**Fluent icon font**",
    `\`<i class="${variant.className}" aria-hidden="true"></i>\``,
    "Keep the `<i>` decorative and label its parent control."
  ] : [
    "**SVG only**",
    "This color/direction-specific asset is not available in the configured icon fonts."
  ];
  return [
    `**${variant.name}** \xB7 ${variant.size}px ${variant.style}`,
    summary,
    `Token: \`${variant.token}\``,
    `Font class: \`${variant.className}\``,
    `SVG file: \`${variant.filename}\``,
    "**Custom element**",
    `\`<fluent-icon name="${variant.token}"></fluent-icon>\``,
    "**Fluent SVG sprite**",
    `\`<svg class="icon"><use href="#${variant.id}"></use></svg>\``,
    "The `<use>` form requires a Fluent sprite containing that symbol id.",
    ...fontMarkup
  ].join("\n\n");
}
function completionRange2(monaco, model, position, startOffset) {
  const start = model.getPositionAt(startOffset);
  return new monaco.Range(
    start.lineNumber,
    start.column,
    position.lineNumber,
    position.column
  );
}
function currentTagFragment(model, position) {
  const offset = model.getOffsetAt(position);
  const start = Math.max(0, offset - 12e3);
  const before = model.getValue().slice(start, offset);
  const lastOpen = before.lastIndexOf("<");
  const lastClose = before.lastIndexOf(">");
  if (lastOpen <= lastClose) return null;
  return {
    offset,
    fragment: before.slice(lastOpen)
  };
}
function nameAttributeContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag || !/^<fluent-icon\b/i.test(tag.fragment)) return null;
  const match = tag.fragment.match(/\bname\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  return {
    kind: "name",
    prefix: match[2],
    startOffset: tag.offset - match[2].length
  };
}
function useHrefContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag || !/^<use\b/i.test(tag.fragment)) return null;
  const match = tag.fragment.match(/\bhref\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const hash = match[2].lastIndexOf("#");
  if (hash < 0) return null;
  const prefix = match[2].slice(hash + 1);
  return {
    kind: "use",
    prefix,
    startOffset: tag.offset - prefix.length
  };
}
function fontClassContext(model, position) {
  const tag = currentTagFragment(model, position);
  if (!tag) return null;
  const match = tag.fragment.match(/\bclass\s*=\s*(["'])([^"']*)$/i);
  if (!match) return null;
  const prefix = match[2].match(/[^\s]*$/)?.[0] || "";
  if (!prefix.startsWith("icon-")) return null;
  return {
    kind: "class",
    prefix,
    startOffset: tag.offset - prefix.length
  };
}
function normalizedQuery(context) {
  if (context.kind === "class") {
    return context.prefix.replace(/^icon-ic_fluent_/, "").replaceAll("_", "-").toLowerCase();
  }
  if (context.kind === "use") {
    return context.prefix.replace(/^ic_fluent_/, "").replaceAll("_", "-").toLowerCase();
  }
  return context.prefix.toLowerCase().replaceAll("_", "-");
}
function matchingVariants(data, context) {
  const query = normalizedQuery(context);
  const supportsSvgOnly = context.kind === "use";
  const source = query ? supportsSvgOnly ? data.variants : data.fontVariants : supportsSvgOnly ? data.defaultVariants : data.defaultFontVariants;
  const ranked = [];
  for (const variant of source) {
    const token = variant.token.toLowerCase();
    let score = 9;
    if (!query) score = 3;
    else if (token.startsWith(query)) score = 0;
    else if (token.includes(query)) score = 1;
    else if (variant.searchText.includes(query.replaceAll("-", " "))) score = 2;
    if (score < 9) ranked.push({ variant, score });
  }
  ranked.sort((a, b) => a.score - b.score || variantPreference(a.variant) - variantPreference(b.variant) || a.variant.token.localeCompare(b.variant.token));
  return ranked.slice(0, 600);
}
function createFluentIconCompletionProvider(monaco, getData) {
  return {
    triggerCharacters: ['"', "'", "-", "_", "#"],
    provideCompletionItems(model, position) {
      const data = getData();
      if (!data) return { suggestions: [] };
      const context = nameAttributeContext(model, position) || useHrefContext(model, position) || fontClassContext(model, position);
      if (!context) return { suggestions: [] };
      const range = completionRange2(monaco, model, position, context.startOffset);
      return {
        suggestions: matchingVariants(data, context).map(({ variant, score }) => {
          const insertText = context.kind === "class" ? variant.className : context.kind === "use" ? variant.id : variant.token;
          return {
            label: insertText,
            kind: context.kind === "class" ? monaco.languages.CompletionItemKind.Class : monaco.languages.CompletionItemKind.EnumMember,
            detail: `${variant.name} \xB7 ${variant.size}px ${variant.style}`,
            documentation: { value: markdownVariant(variant) },
            insertText,
            filterText: `${insertText} ${variant.searchText}`,
            sortText: `${score}-${String(variantPreference(variant)).padStart(3, "0")}-${variant.token}`,
            range
          };
        })
      };
    }
  };
}
function wordAt(model, position) {
  const value = model.getValue();
  const offset = model.getOffsetAt(position);
  let start = offset;
  let end = offset;
  while (start > 0 && /[\w-]/.test(value[start - 1])) start--;
  while (end < value.length && /[\w-]/.test(value[end])) end++;
  if (start === end) return null;
  const startPosition = model.getPositionAt(start);
  const endPosition = model.getPositionAt(end);
  return {
    value: value.slice(start, end),
    range: {
      startLineNumber: startPosition.lineNumber,
      startColumn: startPosition.column,
      endLineNumber: endPosition.lineNumber,
      endColumn: endPosition.column
    }
  };
}
function createFluentIconHoverProvider(getData) {
  return {
    provideHover(model, position) {
      const data = getData();
      const target = data && wordAt(model, position);
      if (!target) return null;
      const variant = data.variantByToken.get(target.value) || data.variantById.get(target.value) || data.variantByClass.get(target.value);
      if (!variant) return null;
      return {
        range: target.range,
        contents: [{ value: markdownVariant(variant) }]
      };
    }
  };
}
function markerForValue(monaco, model, startOffset, value, message) {
  const start = model.getPositionAt(startOffset);
  const end = model.getPositionAt(startOffset + value.length);
  return {
    severity: monaco.MarkerSeverity.Warning,
    message,
    startLineNumber: start.lineNumber,
    startColumn: start.column,
    endLineNumber: end.lineNumber,
    endColumn: end.column
  };
}
function attributeValues(source, tagName, attributeName) {
  const results = [];
  const pattern = new RegExp(
    `<${tagName}\\b[^>]*\\b${attributeName}\\s*=\\s*(["'])([^"']+)\\1`,
    "gi"
  );
  let match;
  while (match = pattern.exec(source)) {
    const value = match[2];
    const relative = match[0].lastIndexOf(value);
    results.push({ value, start: match.index + relative });
  }
  return results;
}
function collectFluentIconMarkers(monaco, model, data) {
  if (!data) return [];
  const source = model.getValue();
  const markers = [];
  for (const item of attributeValues(source, "fluent-icon", "name")) {
    const variant = data.variantByToken.get(item.value);
    if (!variant) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start,
        item.value,
        `Unknown Fluent icon token "${item.value}". Use completion to select a real size/style variant.`
      ));
    } else if (!variant.fontAvailable) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start,
        item.value,
        `"${item.value}" is SVG-only and cannot render through the configured font-backed <fluent-icon>. Use a Fluent sprite <use> reference.`
      ));
    }
  }
  for (const item of attributeValues(source, "use", "href")) {
    const hash = item.value.lastIndexOf("#");
    const id = hash >= 0 ? item.value.slice(hash + 1) : "";
    if (id.startsWith("ic_fluent_") && !data.variantById.has(id)) {
      markers.push(markerForValue(
        monaco,
        model,
        item.start + hash + 1,
        id,
        `Unknown Fluent sprite symbol "${id}".`
      ));
    }
  }
  const classPattern = /\bclass\s*=\s*(["'])([^"']+)\1/gi;
  let classMatch;
  while (classMatch = classPattern.exec(source)) {
    const value = classMatch[2];
    const valueStart = classMatch.index + classMatch[0].lastIndexOf(value);
    const tokens = value.matchAll(/\S+/g);
    for (const tokenMatch of tokens) {
      const className = tokenMatch[0];
      if (className.startsWith("icon-ic_fluent_") && !data.variantByClass.has(className)) {
        markers.push(markerForValue(
          monaco,
          model,
          valueStart + tokenMatch.index,
          className,
          `Unknown Fluent icon font class "${className}".`
        ));
      }
    }
  }
  return markers;
}

// ../src/editors.js
var NAMES = ["html", "css", "js"];
var LANGUAGES = { html: "html", css: "css", js: "javascript" };
var MODEL_URIS = {
  html: "file:///dcspad/index.html",
  css: "file:///dcspad/styles.css",
  js: "file:///dcspad/script.js"
};
async function initEditors({ onChange, onRunShortcut, onTogglePane, onFontStep }) {
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
  let desiredBspIntelligence = false;
  let bspIntelligenceGeneration = 0;
  let bspIntelligence = null;
  let desiredFluentIconIntelligence = false;
  let fluentIconIntelligenceGeneration = 0;
  let fluentIconIntelligence = null;
  const jsLibraryPacks = /* @__PURE__ */ new Map();
  const htmlDataPacks = /* @__PURE__ */ new Map();
  const enabledIntelligence = /* @__PURE__ */ new Set();
  monaco.editor.defineTheme("dcspad-dark", {
    base: "vs-dark",
    inherit: true,
    rules: [
      { token: "comment", foreground: "6A9955", fontStyle: "italic" },
      { token: "comment.doc", foreground: "6A9955", fontStyle: "italic" },
      { token: "keyword", foreground: "569CD6" },
      { token: "keyword.flow", foreground: "C586C0" },
      { token: "number", foreground: "B5CEA8" },
      { token: "string", foreground: "CE9178" },
      { token: "string.escape", foreground: "D7BA7D" },
      { token: "regexp", foreground: "D16969" },
      { token: "type", foreground: "4EC9B0" },
      { token: "type.identifier", foreground: "4EC9B0" },
      { token: "identifier", foreground: "9CDCFE" },
      { token: "constant", foreground: "4FC1FF" },
      { token: "tag", foreground: "569CD6" },
      { token: "tag.css", foreground: "D7BA7D" },
      { token: "attribute.name", foreground: "9CDCFE" },
      { token: "attribute.value", foreground: "CE9178" },
      { token: "attribute.value.number.css", foreground: "B5CEA8" },
      { token: "attribute.value.unit.css", foreground: "B5CEA8" },
      { token: "delimiter", foreground: "D4D4D4" },
      { token: "operator", foreground: "D4D4D4" },
      { token: "invalid", foreground: "F44747" }
    ],
    colors: {
      "editor.background": "#17191f",
      "editor.foreground": "#cccccc",
      "editorGutter.background": "#14161b",
      "editorLineNumber.foreground": "#6d7484",
      "editorLineNumber.activeForeground": "#a2a9b8",
      "editor.lineHighlightBackground": "#1f232b",
      "editor.lineHighlightBorder": "#262b34",
      "editor.selectionBackground": "#264f78",
      "editor.inactiveSelectionBackground": "#264f7855",
      "editorCursor.foreground": "#aeafad",
      "editorBracketMatch.background": "#00000000",
      "editorBracketMatch.border": "#888888",
      "editorIndentGuide.background1": "#2a2e38",
      "editorIndentGuide.activeBackground1": "#3a4150",
      "editorWhitespace.foreground": "#333947",
      "editorWidget.background": "#20242c",
      "editorWidget.border": "#3a4150",
      "editorWidget.foreground": "#d4d9e2",
      "editorSuggestWidget.background": "#20242c",
      "editorSuggestWidget.border": "#3a4150",
      "editorSuggestWidget.foreground": "#d4d9e2",
      "editorSuggestWidget.selectedBackground": "#2a2f3a",
      "editorSuggestWidget.highlightForeground": "#5ee3c4",
      "editorSuggestWidget.focusHighlightForeground": "#5ee3c4",
      "editorHoverWidget.background": "#20242c",
      "editorHoverWidget.border": "#3a4150",
      "editorHoverWidget.foreground": "#d4d9e2",
      "list.hoverBackground": "#2a2f3a",
      "list.highlightForeground": "#5ee3c4",
      "input.background": "#14161b",
      "input.border": "#3a4150",
      "input.foreground": "#e6e9ef",
      "inputOption.activeBorder": "#3fd8b4",
      "editor.findMatchBackground": "#2c6a5c66",
      "editor.findMatchBorder": "#3fd8b4",
      "editor.findMatchHighlightBackground": "#1e3b35",
      "editorError.foreground": "#ff6b62",
      "editorWarning.foreground": "#e8b660",
      "editorInfo.foreground": "#67a7f7",
      "editorLink.activeForeground": "#67a7f7",
      "menu.background": "#20242c",
      "menu.foreground": "#d4d9e2",
      "menu.selectionBackground": "#2a2f3a",
      "scrollbarSlider.background": "#3a415055",
      "scrollbarSlider.hoverBackground": "#3a415088",
      "scrollbarSlider.activeBackground": "#3a4150aa"
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
        dataProviders: Object.fromEntries(htmlDataPacks)
      }
    });
  }
  function setAlpineIntelligenceEnabled(enabled) {
    if (enabled) {
      jsLibraryPacks.set(ALPINE_PACK_ID, ALPINE_JS_LIBRARIES);
      htmlDataPacks.set(ALPINE_PACK_ID, ALPINE_HTML_DATA);
      enabledIntelligence.add(ALPINE_PACK_ID);
      document.documentElement.dataset.alpineIntelligence = "ready";
    } else {
      jsLibraryPacks.delete(ALPINE_PACK_ID);
      htmlDataPacks.delete(ALPINE_PACK_ID);
      enabledIntelligence.delete(ALPINE_PACK_ID);
      document.documentElement.dataset.alpineIntelligence = "disabled";
    }
    applyJsLibraries();
    applyHtmlData();
  }
  const alpineCompletionRegistration = monaco.languages.registerCompletionItemProvider(
    "html",
    createAlpineHtmlCompletionProvider(
      monaco,
      () => enabledIntelligence.has(ALPINE_PACK_ID)
    )
  );
  const bspRegistrations = [
    monaco.languages.registerCompletionItemProvider(
      "css",
      createBspCssCompletionProvider(monaco, () => bspIntelligence)
    ),
    monaco.languages.registerHoverProvider(
      "css",
      createBspCssHoverProvider(() => bspIntelligence)
    ),
    monaco.languages.registerCompletionItemProvider(
      "html",
      createBspHtmlClassCompletionProvider(monaco, () => bspIntelligence)
    ),
    monaco.languages.registerHoverProvider(
      "html",
      createBspHtmlClassHoverProvider(() => bspIntelligence)
    )
  ];
  const fluentIconRegistrations = [
    monaco.languages.registerCompletionItemProvider(
      "html",
      createFluentIconCompletionProvider(monaco, () => fluentIconIntelligence)
    ),
    monaco.languages.registerHoverProvider(
      "html",
      createFluentIconHoverProvider(() => fluentIconIntelligence)
    )
  ];
  function applyFluentIconMarkers() {
    if (!models.html) return;
    monaco.editor.setModelMarkers(
      models.html,
      FLUENT_ICONS_PACK_ID,
      collectFluentIconMarkers(
        monaco,
        models.html,
        fluentIconIntelligence
      )
    );
  }
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
      if (name === "html") applyFluentIconMarkers();
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
    // Editor text size is a setting (11–18); line height locks to 1.7×.
    fontSize: state3.settings.editorFontSize || 13,
    lineHeight: Math.round((state3.settings.editorFontSize || 13) * 1.7),
    wordWrap: state3.settings.wordWrap ? "on" : "off",
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
  editor.addAction({
    id: "dcspad.togglePane.resources",
    label: "Toggle resources pane",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Backslash],
    run: () => onTogglePane?.("resources")
  });
  editor.addAction({
    id: "dcspad.togglePane.preview",
    label: "Toggle preview pane",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.Backslash],
    run: () => onTogglePane?.("preview")
  });
  editor.addAction({
    id: "dcspad.togglePane.console",
    label: "Toggle console pane",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyJ],
    run: () => onTogglePane?.("console")
  });
  editor.addAction({
    id: "dcspad.fontLarger",
    label: "Larger editor text",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Equal],
    run: () => onFontStep?.(1)
  });
  editor.addAction({
    id: "dcspad.fontSmaller",
    label: "Smaller editor text",
    keybindings: [monaco.KeyMod.CtrlCmd | monaco.KeyCode.Minus],
    run: () => onFontStep?.(-1)
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
      jsLibraryPacks.delete("pnpjs-2.15.0");
      applyJsLibraries();
      document.documentElement.dataset.pnpTypes = "disabled";
      return;
    }
    document.documentElement.dataset.pnpTypes = "loading";
    try {
      const libs = await fetchPnpTypeLibraries();
      if (!desiredPnpTypes || generation !== pnpTypesGeneration) return;
      jsLibraryPacks.set("pnpjs-2.15.0", libs);
      applyJsLibraries();
      document.documentElement.dataset.pnpTypes = "ready";
    } catch (error) {
      if (generation !== pnpTypesGeneration) return;
      document.documentElement.dataset.pnpTypes = "error";
      console.warn("DCSPad: PnPjs IntelliSense could not be loaded", error);
    }
  }
  async function setBspIntelligenceEnabled(enabled) {
    desiredBspIntelligence = !!enabled;
    const generation = ++bspIntelligenceGeneration;
    if (!desiredBspIntelligence) {
      bspIntelligence = null;
      enabledIntelligence.delete(BSP_PACK_ID);
      document.documentElement.dataset.bspIntelligence = "disabled";
      return;
    }
    enabledIntelligence.add(BSP_PACK_ID);
    document.documentElement.dataset.bspIntelligence = "loading";
    try {
      const data = await fetchBspIntelligence();
      if (!desiredBspIntelligence || generation !== bspIntelligenceGeneration) return;
      bspIntelligence = data;
      document.documentElement.dataset.bspIntelligence = "ready";
    } catch (error) {
      if (generation !== bspIntelligenceGeneration) return;
      bspIntelligence = null;
      enabledIntelligence.delete(BSP_PACK_ID);
      document.documentElement.dataset.bspIntelligence = "error";
      console.warn("DCSPad: BMO design-system intelligence could not be loaded", error);
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
      document.documentElement.dataset.fluentIconIntelligence = "disabled";
      return;
    }
    enabledIntelligence.add(FLUENT_ICONS_PACK_ID);
    htmlDataPacks.set(FLUENT_ICONS_PACK_ID, FLUENT_ICONS_HTML_DATA);
    applyHtmlData();
    document.documentElement.dataset.fluentIconIntelligence = "loading";
    try {
      const data = await fetchFluentIconIntelligence();
      if (!desiredFluentIconIntelligence || generation !== fluentIconIntelligenceGeneration) return;
      fluentIconIntelligence = data;
      applyFluentIconMarkers();
      document.documentElement.dataset.fluentIconIntelligence = "ready";
    } catch (error) {
      if (generation !== fluentIconIntelligenceGeneration) return;
      fluentIconIntelligence = null;
      enabledIntelligence.delete(FLUENT_ICONS_PACK_ID);
      applyFluentIconMarkers();
      document.documentElement.dataset.fluentIconIntelligence = "error";
      console.warn("DCSPad: Fluent icon intelligence could not be loaded", error);
    }
  }
  function setIntelligencePacks(packIds) {
    const requested = new Set(packIds || []);
    setAlpineIntelligenceEnabled(requested.has(ALPINE_PACK_ID));
    setPnpTypesEnabled(requested.has("pnpjs-2.15.0"));
    setBspIntelligenceEnabled(requested.has(BSP_PACK_ID));
    setFluentIconIntelligenceEnabled(requested.has(FLUENT_ICONS_PACK_ID));
  }
  reportCursor();
  return {
    activate,
    setFontSize: (px2) => {
      editor.updateOptions({ fontSize: px2, lineHeight: Math.round(px2 * 1.7) });
    },
    setWordWrap: (on) => {
      editor.updateOptions({ wordWrap: on ? "on" : "off" });
    },
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
      const model = models[name];
      const selection = editor.getSelection() || selections[name];
      model.pushStackElement();
      editor.executeEdits("dcspad-snippet", [{
        range: selection,
        text,
        forceMoveMarkers: true
      }]);
      model.pushStackElement();
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
    setIntelligencePacks,
    setPnpTypesEnabled,
    dispose: () => {
      resizeObserver.disconnect();
      alpineCompletionRegistration.dispose();
      for (const registration of bspRegistrations) registration.dispose();
      for (const registration of fluentIconRegistrations) registration.dispose();
      monaco.editor.setModelMarkers(models.html, FLUENT_ICONS_PACK_ID, []);
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
var escAttr = (s) => String(s).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
function externalScript(url, nonceAttr) {
  return `<script src="${escAttr(url)}"${nonceAttr}><\/script>`;
}
function libraryScript(entry, nonceAttr) {
  const primary = externalScript(entry.js, nonceAttr);
  if (!entry.fallbackJs || !entry.probeGlobal) return primary;
  const path = entry.probeGlobal.split(".").filter(Boolean);
  const fallbackTag = externalScript(entry.fallbackJs, nonceAttr);
  const message = `DCSPad: ${entry.name || entry.probeGlobal} did not expose ${entry.probeGlobal}; loading configured fallback`;
  const probe = `(function(){var value=window;var path=${JSON.stringify(path)};for(var i=0;i<path.length&&value!=null;i+=1)value=value[path[i]];if(value==null){console.warn(${JSON.stringify(message)});document.write(${JSON.stringify(fallbackTag)});}})();`;
  return `${primary}
<script${nonceAttr}>${escScript(probe)}<\/script>`;
}
function hostNonce() {
  for (const s of document.scripts) if (s.nonce) return s.nonce;
  return "";
}
function assemble({ docs, libraries, spContext, settings, token }) {
  const nonce = hostNonce();
  const nonceAttr = nonce ? ` nonce="${nonce}"` : "";
  const cssLinks = libraries.filter((l) => l.css).map((l) => (Array.isArray(l.css) ? l.css : [l.css]).map((u) => `<link rel="stylesheet" href="${escAttr(u)}">`).join("\n")).join("\n");
  const libraryStyles = libraries.filter((l) => l.cssText).map((l) => `<style data-dcspad-library="${escAttr(l.name || "configured")}">
${escStyle(l.cssText)}
</style>`).join("\n");
  const jsTags = libraries.filter((l) => l.js).map((l) => {
    if (Array.isArray(l.js)) {
      return l.js.map((u) => externalScript(u, nonceAttr)).join("\n");
    }
    return libraryScript(l, nonceAttr);
  }).join("\n");
  const chromeStyle = settings.previewDark ? `<style data-dcspad-chrome>
:root { color-scheme: dark; }
html { background: #1a1d23; color: #e6e9ef; }
</style>
` : "";
  const contextScript = spContext ? `<script${nonceAttr}>window._spPageContextInfo = ${JSON.stringify(spContext.pageContext)};<\/script>
` + (spContext.baseHref ? `<base href="${escAttr(spContext.baseHref)}">
` : "") : "";
  const head = `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<script${nonceAttr}>${escScript(harnessText.replaceAll("__DCSPAD_TOKEN__", token))}<\/script>
${contextScript}${chromeStyle}${cssLinks}
${libraryStyles}
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
var errorCount = 0;
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
  async function submitRepl() {
    if (!input.value.trim()) return;
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
  }
  document.getElementById("btn-repl-eval").addEventListener("click", submitRepl);
  input.addEventListener("keydown", async (e) => {
    if (e.key === "Enter" && input.value.trim()) {
      await submitRepl();
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
  updateConsoleEmpty();
  if (level === "error") {
    errorCount++;
    const badge2 = document.getElementById("console-badge");
    badge2.textContent = errorCount > 99 ? "99+" : String(errorCount);
    badge2.hidden = false;
  }
}
function runDivider(runNumber) {
  groupStack = [];
  const div = el("div", "run-divider new-entry");
  const ts = (/* @__PURE__ */ new Date()).toLocaleTimeString();
  div.append(el("span", "rd-mark", "\u259E \u25B6"), el("span", "", `run #${runNumber} \xB7 ${ts}`));
  out.append(div);
  updateConsoleEmpty();
  scrollIfPinned(true);
}
function clear() {
  out.textContent = "";
  groupStack = [];
  errorCount = 0;
  document.getElementById("console-badge").hidden = true;
  updateConsoleEmpty();
}
function updateConsoleEmpty() {
  const emptyEl = document.getElementById("console-empty");
  if (emptyEl) emptyEl.hidden = out.children.length > 0;
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
var errorCount2 = 0;
var deps2 = {};
function bumpErrorCount() {
  errorCount2++;
  const badge2 = document.getElementById("network-badge");
  badge2.textContent = errorCount2 > 99 ? "99+" : String(errorCount2);
  badge2.hidden = false;
}
function resetErrorCount() {
  errorCount2 = 0;
  document.getElementById("network-badge").hidden = true;
}
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
  if (!d.ok) bumpErrorCount();
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
  resetErrorCount();
}
function applyApiFilter() {
  for (const { row } of requests.values()) applyApiFilterTo(row);
}
function applyApiFilterTo(row) {
  const apiOnly = document.getElementById("chk-api-only").checked;
  row.classList.toggle("hidden-api", apiOnly && !row.classList.contains("is-api"));
}
function markRun() {
  if (getState().settings.autoClearConsole) resetErrorCount();
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

// ../src/config.js
var EMPTY_CONFIG = Object.freeze({
  version: 1,
  frameworks: Object.freeze({
    prefer: "local",
    fallbackToCdn: true,
    items: Object.freeze({})
  }),
  assets: Object.freeze({})
});
var activeConfig = EMPTY_CONFIG;
var isRecord = (value) => value !== null && typeof value === "object" && !Array.isArray(value);
var cleanString = (value) => typeof value === "string" ? value.trim() : "";
var sourcePreference = (value, fallback = "local") => value === "cdn" || value === "hosted" || value === "local" ? value : fallback;
function resolveUrl(value, configUrl2, { folder = false } = {}) {
  const source = cleanString(value);
  if (!source) return "";
  const resolved = new URL(source, configUrl2).href;
  return folder && !resolved.endsWith("/") ? `${resolved}/` : resolved;
}
function normalizeFrameworks(value, configUrl2, warnings) {
  const source = isRecord(value) ? value : {};
  const prefer = sourcePreference(source.prefer);
  const fallbackToCdn = source.fallbackToCdn !== false;
  const items = {};
  for (const [id, raw] of Object.entries(isRecord(source.items) ? source.items : {})) {
    if (!isRecord(raw)) {
      warnings.push(`framework config "${id}" was ignored because it is not an object`);
      continue;
    }
    const probeGlobal = cleanString(raw.probeGlobal).replace(/^window\./, "");
    if (probeGlobal && !/^[$A-Z_a-z][$\w]*(?:\.[$A-Z_a-z][$\w]*)*$/.test(probeGlobal)) {
      warnings.push(`framework config "${id}" has an invalid probeGlobal path`);
    }
    items[id] = {
      localUrl: resolveUrl(raw.localUrl, configUrl2),
      cdnUrl: resolveUrl(raw.cdnUrl, configUrl2),
      prefer: sourcePreference(raw.prefer, prefer),
      fallbackToCdn: typeof raw.fallbackToCdn === "boolean" ? raw.fallbackToCdn : fallbackToCdn,
      probeGlobal: /^[$A-Z_a-z][$\w]*(?:\.[$A-Z_a-z][$\w]*)*$/.test(probeGlobal) ? probeGlobal : "",
      intelligence: Array.isArray(raw.intelligence) ? [...new Set(raw.intelligence.map(cleanString).filter(Boolean))] : []
    };
  }
  return { prefer, fallbackToCdn, items };
}
function normalizeAssetGroup(raw, configUrl2, defaultPreference) {
  if (!isRecord(raw)) return null;
  const files = {};
  for (const [name, path] of Object.entries(isRecord(raw.files) ? raw.files : {})) {
    const clean = cleanString(path);
    if (clean) files[name] = clean;
  }
  const rawRuntime = isRecord(raw.runtime) ? raw.runtime : {};
  return {
    prefer: sourcePreference(raw.prefer, defaultPreference),
    localBaseUrl: resolveUrl(raw.localBaseUrl, configUrl2, { folder: true }),
    hostedBaseUrl: resolveUrl(raw.hostedBaseUrl, configUrl2, { folder: true }),
    intelligence: Array.isArray(raw.intelligence) ? [...new Set(raw.intelligence.map(cleanString).filter(Boolean))] : [],
    files,
    runtime: {
      enabled: rawRuntime.enabled === true,
      cssFiles: Array.isArray(rawRuntime.cssFiles) ? [...new Set(rawRuntime.cssFiles.map(cleanString).filter(Boolean))] : [],
      fluentIconElement: rawRuntime.fluentIconElement === true
    }
  };
}
function normalizeConfig(raw, configUrl2) {
  const warnings = [];
  if (!isRecord(raw)) {
    return { config: EMPTY_CONFIG, warnings: ["configuration root must be an object"] };
  }
  if (raw.version !== 1) {
    warnings.push(`configuration version ${JSON.stringify(raw.version)} is not supported; expected 1`);
  }
  const assets = {};
  for (const [name, value] of Object.entries(isRecord(raw.assets) ? raw.assets : {})) {
    const group = normalizeAssetGroup(value, configUrl2, "local");
    if (group) assets[name] = group;
  }
  return {
    config: {
      version: 1,
      frameworks: normalizeFrameworks(raw.frameworks, configUrl2, warnings),
      assets
    },
    warnings
  };
}
function configUrl() {
  return window.__DCSPAD_CONFIG_URL__ || new URL("../dcspad.config.json", import.meta.url).href;
}
async function loadAppConfig() {
  const url = configUrl();
  try {
    const response = await fetch(url, {
      credentials: "same-origin",
      cache: "no-cache"
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const normalized = normalizeConfig(await response.json(), url);
    activeConfig = normalized.config;
    return { ...normalized, url };
  } catch (error) {
    activeConfig = EMPTY_CONFIG;
    return {
      config: activeConfig,
      url,
      warnings: [`dcspad.config.json could not be loaded (${error.message || error}); built-in framework URLs remain active`]
    };
  }
}
function applyFrameworkConfig(entry, config = activeConfig) {
  const override = config?.frameworks?.items?.[entry?.id];
  if (!override) return { ...entry };
  const local = override.localUrl;
  const cdn = override.cdnUrl;
  const preferred = override.prefer === "cdn" ? cdn : local;
  const alternate = override.prefer === "cdn" ? local : cdn;
  const primary = preferred || alternate;
  const fallback = preferred && alternate && override.fallbackToCdn ? alternate : "";
  const effective = {
    ...entry,
    intelligence: [
      .../* @__PURE__ */ new Set([
        ...Array.isArray(entry.intelligence) ? entry.intelligence : [],
        ...override.intelligence
      ])
    ],
    configuredSources: { local, cdn },
    probeGlobal: override.probeGlobal
  };
  if (!primary) return effective;
  const primaryIsCss = /\.css(?:[?#]|$)/i.test(primary);
  if (primaryIsCss) {
    effective.css = primary;
    delete effective.js;
    return effective;
  }
  effective.js = primary;
  delete effective.css;
  if (fallback && !/\.css(?:[?#]|$)/i.test(fallback) && override.probeGlobal) {
    effective.fallbackJs = fallback;
  }
  return effective;
}
function selectedAssetBase(group) {
  if (!group) return "";
  const preferred = group.prefer === "hosted" ? group.hostedBaseUrl : group.localBaseUrl;
  return preferred || group.hostedBaseUrl || group.localBaseUrl || "";
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
    js: "https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/2.15.0/pnp.js",
    intelligence: ["pnpjs-2.15.0"],
    hint: "Exposes global pnp \u2014 use const { sp } = pnp;"
  },
  {
    id: "alpine",
    name: "Alpine.js",
    js: "https://cdn.jsdelivr.net/npm/alpinejs@3/dist/cdn.min.js",
    intelligence: ["alpine-3"]
  },
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
var appConfig = null;
var onChangeCb = null;
var onStorageErrorCb = null;
var filterText = "";
var isCssUrl = (url) => /\.css(\?|$)/i.test(url);
var entryFromUrl = (url, name) => ({
  id: newId("lib"),
  name: name || url.split("/").pop() || url,
  js: isCssUrl(url) ? void 0 : url,
  css: isCssUrl(url) ? url : void 0
});
function initLibraries({ config, onChange, onStorageError }) {
  appConfig = config;
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
  const form = document.getElementById("lib-custom-form");
  const toggleBtn = document.getElementById("btn-add-framework");
  const urlInput = document.getElementById("lib-custom-url");
  const nameInput = document.getElementById("lib-custom-name");
  const errorEl = document.getElementById("lib-custom-error");
  function setAddFormOpen(open) {
    form.hidden = !open;
    toggleBtn.hidden = open;
    if (open) {
      nameInput.focus();
    } else {
      clearAddError();
      urlInput.value = "";
      nameInput.value = "";
    }
  }
  function showAddError(msg) {
    errorEl.textContent = msg;
    errorEl.hidden = false;
    urlInput.classList.add("invalid");
  }
  function clearAddError() {
    errorEl.hidden = true;
    urlInput.classList.remove("invalid");
  }
  const filterRow = document.getElementById("frameworks-filter-row");
  const filterInput = document.getElementById("frameworks-filter");
  document.getElementById("btn-frameworks-search").addEventListener("click", () => {
    filterRow.hidden = !filterRow.hidden;
    if (!filterRow.hidden) filterInput.focus();
    else {
      filterInput.value = "";
      filterText = "";
      render();
    }
  });
  filterInput.addEventListener("input", () => {
    filterText = filterInput.value.trim().toLowerCase();
    render();
  });
  filterInput.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      e.stopPropagation();
      filterRow.hidden = true;
      filterInput.value = "";
      filterText = "";
      render();
    }
  });
  toggleBtn.addEventListener("click", () => setAddFormOpen(true));
  document.getElementById("lib-add-cancel").addEventListener("click", () => setAddFormOpen(false));
  form.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      e.stopPropagation();
      setAddFormOpen(false);
    }
  });
  urlInput.addEventListener("input", clearAddError);
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const url = urlInput.value.trim();
    if (!url) {
      showAddError("Enter a script or stylesheet URL.");
      return;
    }
    try {
      new URL(url, location.href);
    } catch {
      showAddError("That is not a valid URL.");
      return;
    }
    if (!/\.(js|css)(\?|#|$)/i.test(url)) {
      showAddError("The URL should point at a .js or .css file.");
      return;
    }
    const entry = entryFromUrl(url, nameInput.value.trim());
    catalog.items.push(entry);
    persistCatalog();
    const enabled = new Set(getState().libraries.enabled);
    enabled.add(entry.id);
    updateNested("libraries", { enabled: [...enabled] });
    setAddFormOpen(false);
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
  let shown = 0;
  for (const entry of catalog.items) {
    if (filterText && !entry.name.toLowerCase().includes(filterText)) continue;
    shown++;
    const pinned = libs.pinned.includes(entry.id);
    (pinned ? pinnedHost : listHost).append(catalogItem(entry, libs, pinned));
  }
  const noMatch = document.getElementById("frameworks-no-match");
  if (noMatch) noMatch.hidden = !(filterText && shown === 0);
  const countEl = document.getElementById("frameworks-count");
  if (countEl) {
    const known = new Set(catalog.items.map((it) => it.id));
    const enabledCount = libs.enabled.filter((id) => known.has(id)).length;
    countEl.textContent = `${enabledCount}/${catalog.items.length}`;
  }
}
var TOOL_ICONS = {
  up: '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 12.5v-9M4.5 7 8 3.5 11.5 7"/></svg>',
  down: '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 3.5v9M4.5 9 8 12.5 11.5 9"/></svg>',
  pin: '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round" aria-hidden="true"><path d="M4.5 2.5h7v11L8 10.6l-3.5 2.9z"/></svg>',
  pinFilled: '<svg width="12" height="12" viewBox="0 0 16 16" aria-hidden="true"><path d="M4.5 2.5h7v11L8 10.6l-3.5 2.9z" fill="currentColor"/></svg>',
  del: '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path d="m4 4 8 8M12 4l-8 8"/></svg>'
};
function catalogItem(entry, libs, pinned) {
  const effective = applyFrameworkConfig(entry, appConfig);
  const item = el("label", "lib-item");
  const chk = document.createElement("input");
  chk.type = "checkbox";
  chk.checked = libs.enabled.includes(entry.id);
  const name = el("span", "lib-name", effective.name);
  if (effective.hint) name.title = effective.hint;
  else if (effective.js || effective.css) {
    name.title = effective.js || effective.css;
    if (effective.fallbackJs) name.title += `
Fallback: ${effective.fallbackJs}`;
  }
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
  const tool = (cls, icon, title, fn) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = cls;
    b.innerHTML = icon;
    b.title = title;
    b.addEventListener("click", (e) => {
      e.preventDefault();
      fn();
    });
    return b;
  };
  const liveIdx = () => catalog.items.indexOf(entry);
  tools.append(
    tool("lib-move", TOOL_ICONS.up, "Move up (injection order)", () => moveEntry(liveIdx(), -1)),
    tool("lib-move", TOOL_ICONS.down, "Move down (injection order)", () => moveEntry(liveIdx(), 1)),
    tool("lib-pin" + (pinned ? " pinned" : ""), pinned ? TOOL_ICONS.pinFilled : TOOL_ICONS.pin, pinned ? "Unpin" : "Pin to top", () => {
      const pins = new Set(getState().libraries.pinned);
      pinned ? pins.delete(entry.id) : pins.add(entry.id);
      updateNested("libraries", { pinned: [...pins] });
      render();
    }),
    tool("lib-del", TOOL_ICONS.del, "Remove from catalog", () => {
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
    const effective = applyFrameworkConfig(entry, appConfig);
    result.push({
      name: effective.name,
      js: effective.js,
      css: effective.css,
      fallbackJs: effective.fallbackJs,
      probeGlobal: effective.probeGlobal
    });
  }
  result.push(...getConfiguredAssetLibraries());
  return result;
}
var FLUENT_FONT_RUNTIME_CSS = `
/* The vendored per-style font files share a generated broad selector.
   Restore the intended family per class suffix so the three styles can
   coexist in one preview document. */
i[class*="icon-ic_fluent_"][class*="_regular"]::before {
  font-family: "FluentSystemIcons-Regular" !important;
}
i[class*="icon-ic_fluent_"][class*="_filled"]::before {
  font-family: "FluentSystemIcons-Filled" !important;
}
i[class*="icon-ic_fluent_"][class*="_light"]::before {
  font-family: "FluentSystemIcons-Light" !important;
}
fluent-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  line-height: 1;
  vertical-align: -0.125em;
}
`;
function appSourceUrl(path) {
  const base = window.__DCSPAD_SRC_BASE__ || new URL("./", import.meta.url).href;
  const url = new URL(path, base);
  const version = window.__DCSPAD_INTELLIGENCE_VERSION__;
  if (version) url.searchParams.set("v", version);
  return url.href;
}
function getConfiguredAssetLibraries() {
  const result = [];
  for (const [id, group] of Object.entries(appConfig?.assets || {})) {
    if (!group.runtime?.enabled) continue;
    const base = selectedAssetBase(group);
    if (!base) continue;
    const css = group.runtime.cssFiles.map((key2) => group.files[key2] || key2).filter(Boolean).map((path) => new URL(path, base).href);
    const entry = {
      name: `${id} configured assets`,
      ...css.length ? { css } : {}
    };
    if (id === "fluentIcons" && group.runtime.fluentIconElement) {
      entry.js = appSourceUrl("bridge/fluent-icon-font.js");
      entry.cssText = FLUENT_FONT_RUNTIME_CSS;
    }
    if (entry.css || entry.js || entry.cssText) result.push(entry);
  }
  return result;
}
function isPnpjs215Runtime(entry) {
  if (entry?.intelligence?.includes("pnpjs-2.15.0")) return true;
  const urls = [
    entry?.js,
    entry?.fallbackJs,
    entry?.configuredSources?.local,
    entry?.configuredSources?.cdn
  ].map((url) => String(url || "").toLowerCase());
  return urls.some((url) => url.includes("@pnp/pnpjs@2.15.0/") || url.includes("/pnp-pnpjs/2.15.0/") || url.includes("/pnpjs/2.15.0/"));
}
function isAlpine3Runtime(entry) {
  if (entry?.intelligence?.includes("alpine-3")) return true;
  if (entry?.id === "alpine") return true;
  const urls = [
    entry?.js,
    entry?.fallbackJs,
    entry?.configuredSources?.local,
    entry?.configuredSources?.cdn
  ].map((url) => String(url || "").toLowerCase());
  return urls.some((url) => url.includes("/alpinejs@3") || url.includes("/alpinejs/3."));
}
function getEnabledIntelligence() {
  const enabled = new Set(getState().libraries.enabled);
  const packs = /* @__PURE__ */ new Set();
  for (const group of Object.values(appConfig?.assets || {})) {
    for (const pack of group.intelligence || []) packs.add(pack);
  }
  for (const entry of catalog.items) {
    if (!enabled.has(entry.id)) continue;
    const effective = applyFrameworkConfig(entry, appConfig);
    for (const pack of effective.intelligence || []) packs.add(pack);
    if (isPnpjs215Runtime(effective)) packs.add("pnpjs-2.15.0");
    if (isAlpine3Runtime(effective)) packs.add("alpine-3");
  }
  return [...packs];
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

// ../src/snippets.js?v=2
var doc = null;
var deps3 = {};
function initSnippets({ getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError }) {
  deps3 = { getSelection, getDocs, insertAtCursor, selectEditorTab, onStorageError };
  doc = loadDoc(SNIPPETS_KEY) || { v: 1, items: [] };
  render2();
  const dialog = document.getElementById("snippet-name-dialog");
  const form = document.getElementById("snippet-name-form");
  const input = document.getElementById("snippet-name-input");
  const context = document.getElementById("snippet-name-context");
  let pendingSnippet = null;
  const closeNamingDialog = () => {
    pendingSnippet = null;
    if (dialog.open) dialog.close();
  };
  document.getElementById("btn-snippet-add").addEventListener("click", () => {
    const lang = getState().layout.editorTab;
    const selection = deps3.getSelection(lang);
    const code = selection || deps3.getDocs()[lang];
    if (!code.trim()) return;
    pendingSnippet = { lang, code };
    input.value = "";
    context.textContent = `Save ${selection ? "the selected" : "all"} ${lang.toUpperCase()} code as a reusable snippet.`;
    if (!dialog.open) dialog.showModal();
    requestAnimationFrame(() => input.focus());
  });
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = input.value.trim();
    if (!pendingSnippet || !name) return;
    const { lang, code } = pendingSnippet;
    doc.items.push({ id: newId("snip"), name, lang, code, createdAt: Date.now() });
    persist2();
    render2();
    pendingSnippet = null;
    dialog.close();
  });
  document.getElementById("snippet-name-cancel").addEventListener("click", closeNamingDialog);
  document.getElementById("snippet-name-close").addEventListener("click", closeNamingDialog);
  dialog.addEventListener("cancel", () => {
    pendingSnippet = null;
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
var DEL_ICON = '<svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><path d="m4 4 8 8M12 4l-8 8"/></svg>';
function render2() {
  const host = document.getElementById("snippet-list");
  host.textContent = "";
  document.getElementById("snippet-empty").hidden = doc.items.length > 0;
  const countEl = document.getElementById("snippets-count");
  if (countEl) countEl.textContent = String(doc.items.length);
  for (const snip of doc.items) {
    const item = el("div", "lib-item snippet-item");
    const lang = el("span", "snippet-lang", snip.lang);
    lang.dataset.lang = snip.lang;
    const name = el("span", "lib-name", snip.name);
    name.title = `Insert into the ${snip.lang.toUpperCase()} editor at the cursor

${snip.code.slice(0, 400)}`;
    const del = document.createElement("button");
    del.type = "button";
    del.className = "lib-del";
    del.innerHTML = DEL_ICON;
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
function standaloneController(splash) {
  const statusEl = document.getElementById("splash-status");
  const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const seen = getState().settings.seenSplash;
  const minimumMs = reduced ? 0 : seen ? 150 : 700;
  const startedAt = performance.now();
  let settled = false;
  let finishTimer = null;
  let removing = false;
  const remove = () => {
    if (removing) return;
    removing = true;
    requestAnimationFrame(() => requestAnimationFrame(() => {
      document.documentElement.classList.add("dcspad-crossfading");
      splash.classList.add("fading");
      setTimeout(() => {
        splash.remove();
        document.documentElement.classList.remove("dcspad-crossfading");
      }, reduced ? 0 : 800);
    }));
    if (!seen) updateNested("settings", { seenSplash: true });
  };
  const controller = {
    status(message) {
      if (!settled && statusEl) statusEl.textContent = message;
    },
    finish() {
      if (settled) return;
      settled = true;
      if (statusEl) statusEl.textContent = "Editor ready";
      const remaining = Math.max(0, minimumMs - (performance.now() - startedAt));
      finishTimer = setTimeout(remove, remaining);
    },
    fail(message) {
      if (settled) return;
      settled = true;
      clearTimeout(finishTimer);
      splash.classList.add("failed");
      if (statusEl) statusEl.textContent = message;
    },
    skip() {
      if (settled) return;
      settled = true;
      clearTimeout(finishTimer);
      remove();
    }
  };
  splash.addEventListener("click", () => controller.skip());
  return controller;
}
function showSplash() {
  if (window.__DCSPAD_BOOT_SPLASH__) {
    const hosted = window.__DCSPAD_BOOT_SPLASH__;
    const reduced = matchMedia("(prefers-reduced-motion: reduce)").matches;
    const seen = getState().settings.seenSplash;
    hosted.minimum?.(reduced ? 0 : seen ? 150 : 700);
    hosted.status("Restoring workspace\u2026");
    if (!seen) updateNested("settings", { seenSplash: true });
    return hosted;
  }
  const splash = document.getElementById("splash");
  const logoEl = document.getElementById("splash-logo");
  if (!splash) {
    return { status() {
    }, finish() {
    }, fail() {
    }, skip() {
    } };
  }
  logoEl.innerHTML = LOGO.replace(/([═-╬]+)/g, '<span class="dim">$1</span>');
  if (!splash.querySelector(".splash-version")) {
    const ver = document.createElement("div");
    ver.className = "splash-version";
    ver.innerHTML = "developer workbench \xB7 <b>sharepoint</b>";
    logoEl.after(ver);
  }
  splash.hidden = false;
  splash.getBoundingClientRect();
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) {
    splash.classList.add("visible");
  } else {
    requestAnimationFrame(() => requestAnimationFrame(() => {
      splash.classList.add("visible");
    }));
  }
  return standaloneController(splash);
}

// ../src/main.js
var splashApi = showSplash();
splashApi.status("Restoring workspace\u2026");
var configReady = loadAppConfig();
var state2 = getState();
applyContextIndicators();
var editorsApi = null;
var layoutApi = initLayout({
  onEditorTabChange: (name) => {
    editorsApi?.activate(name);
    updateStatusLang(name);
  }
});
function updateStatusLang(name) {
  const badge2 = document.getElementById("status-lang");
  badge2.textContent = name;
  badge2.dataset.lang = name;
}
updateStatusLang(state2.layout.editorTab);
function markUnsaved(name) {
  const dot = document.getElementById(`unsaved-${name}`);
  if (dot) dot.hidden = false;
}
function clearUnsaved() {
  for (const name of ["html", "css", "js"]) {
    const dot = document.getElementById(`unsaved-${name}`);
    if (dot) dot.hidden = true;
  }
}
var isDiagVisible = (name) => document.querySelector(`#diag-tabs .tab[data-diag="${name}"]`).classList.contains("active");
splashApi.status("Starting Monaco editor\u2026");
try {
  editorsApi = await initEditors({
    onChange: (name) => {
      markUnsaved(name);
      scheduleAutorun();
    },
    onRunShortcut: () => run2(),
    onTogglePane: (name) => layoutApi.togglePane?.(name),
    onFontStep: (delta) => stepEditorFontSize(delta)
  });
} catch (error) {
  splashApi.fail(`Monaco failed to start \u2014 ${error.message || error}`);
  throw error;
}
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
var configResult = await configReady;
initLibraries({
  config: configResult.config,
  onChange: () => {
    scheduleAutorun();
    editorsApi.setIntelligencePacks(getEnabledIntelligence());
  },
  onStorageError: (msg) => reportStorageError(msg)
});
editorsApi.setIntelligencePacks(getEnabledIntelligence());
initSnippets({
  getSelection: (name) => editorsApi.getSelection(name),
  getDocs: () => editorsApi.getDocs(),
  insertAtCursor: (name, text) => editorsApi.insertAtCursor(name, text),
  selectEditorTab: (name) => layoutApi.selectEditorTab(name),
  onStorageError: (msg) => reportStorageError(msg)
});
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
    settleRunFeedback();
  }
});
var longRunTimer = null;
function settleRunFeedback() {
  clearTimeout(longRunTimer);
  longRunTimer = null;
  const panel = document.getElementById("preview-panel");
  panel.classList.remove("running-long");
  const chip = document.getElementById("preview-run-chip");
  document.getElementById("preview-run-time").textContent = (/* @__PURE__ */ new Date()).toTimeString().slice(0, 8);
  chip.hidden = false;
  chip.classList.remove("pop");
  void chip.offsetWidth;
  chip.classList.add("pop");
}
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
  panel.classList.remove("sweeping", "running-long");
  void panel.offsetWidth;
  panel.classList.add("sweeping");
  clearUnsaved();
  clearTimeout(longRunTimer);
  longRunTimer = setTimeout(() => panel.classList.add("running-long"), 1200);
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
      clearTimeout(longRunTimer);
      document.getElementById("preview-panel").classList.remove("running-long");
    }
  }, 15e3);
}
document.getElementById("btn-run").addEventListener("click", run2);
document.getElementById("btn-rerun").addEventListener("click", run2);
var btnPreviewTheme = document.getElementById("btn-preview-theme");
function applyPreviewTheme() {
  const dark = getState().settings.previewDark;
  btnPreviewTheme.dataset.mode = dark ? "dark" : "light";
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
for (const warning of configResult.warnings) padWarn(warning);
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
  editorsApi.setIntelligencePacks(getEnabledIntelligence());
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
  refreshStepperDisabled("btn-diag-font-dec", "btn-diag-font-inc", px2, DIAG_FS_MIN, DIAG_FS_MAX);
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
var EDITOR_FS_MIN = 11;
var EDITOR_FS_MAX = 18;
function refreshStepperDisabled(decId, incId, val, min, max) {
  document.getElementById(decId).disabled = val <= min;
  document.getElementById(incId).disabled = val >= max;
}
function applyEditorFontSize(px2) {
  editorsApi.setFontSize(px2);
  refreshStepperDisabled("btn-editor-font-dec", "btn-editor-font-inc", px2, EDITOR_FS_MIN, EDITOR_FS_MAX);
}
function stepEditorFontSize(delta) {
  const cur = getState().settings.editorFontSize;
  const next = Math.min(EDITOR_FS_MAX, Math.max(EDITOR_FS_MIN, cur + delta));
  if (next === cur) return;
  updateNested("settings", { editorFontSize: next });
  applyEditorFontSize(next);
}
refreshStepperDisabled(
  "btn-editor-font-dec",
  "btn-editor-font-inc",
  state2.settings.editorFontSize,
  EDITOR_FS_MIN,
  EDITOR_FS_MAX
);
document.getElementById("btn-editor-font-dec").addEventListener("click", () => stepEditorFontSize(-1));
document.getElementById("btn-editor-font-inc").addEventListener("click", () => stepEditorFontSize(1));
document.getElementById("diag-panel").addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  if (e.key === "=" || e.key === "+") {
    e.preventDefault();
    stepDiagFontSize(1);
  } else if (e.key === "-") {
    e.preventDefault();
    stepDiagFontSize(-1);
  }
});
var btnWordWrap = document.getElementById("btn-word-wrap");
function reflectWordWrap() {
  const on = getState().settings.wordWrap;
  btnWordWrap.classList.toggle("active", on);
  btnWordWrap.setAttribute("aria-pressed", String(on));
}
reflectWordWrap();
btnWordWrap.addEventListener("click", () => {
  const on = !getState().settings.wordWrap;
  updateNested("settings", { wordWrap: on });
  editorsApi.setWordWrap(on);
  reflectWordWrap();
});
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
splashApi.status("Editor ready");
splashApi.finish();
