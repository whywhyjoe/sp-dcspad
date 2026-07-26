# DCSPad — Claude Code guide

> History note: this repo was split out of `whywhyjoe/todo`, where the app lived
> in a `devpad/` subfolder. Commit history is preserved; paths are now at the
> repo root. Older entries in `REVIEW-LOG.md` refer to `devpad/…` paths — read
> those as repo-root-relative.

SharePoint-native, JSFiddle-style developer workbench. Pure client-side: HTML/CSS/JS editors (CodeMirror 6), live preview iframe, console + network panels with an SP-aware object inspector, library manager, REPL. No backend, no framework, no build step — deploy by uploading this folder to a SharePoint library.

**The mission, one sentence:** code written in the pad runs **unmodified** on a real SharePoint page. Every design decision below serves that.

## Architecture invariants — do not break these

1. **Fresh same-origin `srcdoc` iframe per run.** Never blob URLs (opaque origin broke PnPjs in the previous incarnation of this project), never reuse a frame (leaked state caused "sometimes it initializes, sometimes it doesn't").
2. **Deterministic assembly** in `src/runner.js` `assemble()`. Fixed order: harness script → SP context + `<base>` → preview-chrome style → library CSS → user CSS → user HTML → library JS as ordered blocking `<script src>` → **user JS last**. The browser's parser handles ordering exactly like a real page. Never inject scripts into a live preview document.
3. **`postMessage`-only across the frame boundary.** The harness (`src/bridge/harness.js`) pre-serializes everything and posts it with a per-run token; the app ignores messages whose token isn't current. Never reach into the iframe from app code (tests may read computed styles — that's it).
4. **Pad chrome must lose to user code.** Anything the pad adds to the preview (e.g. the dark-mode canvas style) is injected *before* library/user CSS so the user's styling always wins. The pad must never misrepresent how code will render on a real page.
5. **No framework, no user build.** Vanilla ES modules, plain DOM. CodeMirror is vendored as a single file (`vendor/codemirror.mjs`) because CM6 breaks subtly with duplicate `@codemirror/state` instances; regenerate only via `tools/build-vendor.mjs` (instructions at top of that file).
6. **`state.js` is the only module that touches localStorage.** It owns three documents: the workspace blob `{html, css, js, libraries, settings, layout}` (live, autosaved), the framework catalog, and the snippet library. The future SharePoint storage layer (`DevPadData/{user}.json`) replaces persistence wholesale by swapping this one seam — don't scatter storage elsewhere. Files on disk (project/catalog/snippet .json, pane exports) go through `src/io.js`, which moves bytes but stores nothing.

## File map

```
index.html                app shell (rename to devpad.aspx if a tenant blocks .html rendering)
styles/app.css            all styling; layout via CSS grid + JS-set vars (--sidebar-w etc.)
vendor/codemirror.mjs     vendored CM6 bundle — regenerate via tools/, never hand-edit
tools/build-vendor.mjs    esbuild one-liner for the vendor bundle
src/main.js               bootstrap; wires every module; run() lives here
src/layout.js             splitters, tabs, collapse/maximize; persists via state.layout
src/editors.js            CM6 editors; Mod-Enter run; gotoJsLine() for stack links
src/state.js              defaults + deep-merge load + debounced autosave; loadDoc/saveDoc
                          for the catalog + snippet documents (sole localStorage toucher)
src/io.js                 file download + JSON file-picker helpers (no storage)
src/runner.js             assemble() + iframe lifecycle + run tokens + evalInFrame()
src/libraries.js          framework catalog: single stored JSON (seeded once from PRESETS,
                          then authoritative), add/remove/reorder, getEnabledLibraries()
src/snippets.js           snippet library: save from selection, insert-at-cursor, file I/O
src/console-panel.js      console rendering, filters, groups, REPL input, stack-frame links
src/network-panel.js      request rows, _api filter, detail pane (JSON via inspector)
src/splash.js             ASCII boot splash (short shimmer after first visit)
src/inspect/tree-view.js  generic expandable trees + table renderer (serialized-node format)
src/inspect/sp-shapes.js  SP/OData/PnPjs shape detection + smart views (standalone by design —
                          the future Site Inspector reuses it)
src/bridge/harness.js     iframe-side instrumentation; plain classic script, no imports;
                          __DCSPAD_TOKEN__ placeholder replaced per run
src/bridge/sp-context.js  real _spPageContextInfo capture (live) or labeled mock
tests/                    Playwright verification suites — see tests/README.md
REVIEW-LOG.md             external-review triage record + accepted low-priority backlog
```

## Dev workflow

```bash
python3 -m http.server 8642     # from the repo root; app at http://localhost:8642/index.html
```

Outside SharePoint the SP chip shows **Mock** and `_api` calls 404 — expected. Live PnPjs/REST behavior can only be validated in a tenant (deployment + validation checklist in README.md). Run the test suites (below) after any change to runner/harness/console/inspector — they exist because this project's failure modes are timing- and boundary-shaped, not type-shaped.

## Tests

`tests/README.md` has the two-server setup (app on 8642, fixtures on 8643) and how Chromium is resolved. Suites: `smoke.mjs` (49 checks: capture, isolation, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave), `darkmode.mjs` (8), `splash.mjs` (3). All should pass; a `custom library` failure usually means the 8643 fixture server isn't running.

## Gotchas already paid for

- `harness.js` token substitution must be `replaceAll` — the placeholder also appears in a comment, and `replace` once shipped a broken token check.
- `app.css` has a global `[hidden] { display: none !important; }` guard: any element with a `display` rule plus the `hidden` attribute silently ignores `hidden` without it (an invisible splash overlay once ate every click).
- Don't put interactive controls inside a `<label>` containing a disabled input — the browser treats the whole label as disabled (custom-library ✕ button bug).
- User code is embedded raw into the assembled document: keep the `</script>` / `</style>` escaping (`escScript`/`escStyle` in runner.js) intact.
- An `about:srcdoc` document resolves relative and **fragment** URLs against the *parent's* base URL, not itself. A plain `<a href="#foo">` therefore counts as a cross-document navigation and loads the pad (or, with `<base href>`, the SP web) *inside its own preview pane*. `harness.js` intercepts fragment clicks and re-creates the same-document jump; it deliberately listens in bubble phase and honours `defaultPrevented` so user `<a href="#">` handlers still win.
- CDN egress: the Claude sandbox proxy blocks public CDNs, so library-loading tests use the local fixture; real CDN presets can only be exercised in a browser with normal egress.

## Roadmap (seams already reserved — don't build without being asked)

- **Site Inspector** sidebar section: enumerate lists/libraries (GUIDs, internal names), fields (display/internal/type), groups + members, content types; renders through `src/inspect/`; "copy as PnPjs call".
- **SharePoint JSON storage** for snippets/projects/templates/shared team resources, replacing localStorage in `state.js`.
- **Console remote handles** — lazy live-object expansion instead of eager capped serialization.
