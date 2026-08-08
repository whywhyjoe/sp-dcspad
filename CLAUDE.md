# DCSPad — Claude Code guide

> **Start here if you're new to this repo:** `HANDOFF.md` — current state plus
> the open problem (hosting inside a modern-page custom-script web part, which
> the standalone-page assumptions below do not yet account for).
>
> History note: this repo was split out of `whywhyjoe/todo`, where the app lived
> in a `devpad/` subfolder. Commit history is preserved; paths are now at the
> repo root. Older entries in `REVIEW-LOG.md` refer to `devpad/…` paths — read
> those as repo-root-relative.
>
> UI/design-system changes must also read
> `design/POST-MONACO-UI-INTEGRATION.md`. It compares the exact pre-Monaco
> baseline (`3414854`) with the current shell and is authoritative for DOM
> hooks, Monaco theming/widget integration, settings and persistence,
> splash/dialog states, and preview/runtime configuration boundaries.

SharePoint-native, JSFiddle-style developer workbench. Pure client-side: HTML/CSS/JS editors (Monaco), live preview iframe, console + network panels with an SP-aware object inspector, library manager, REPL. No backend, no framework, no user build step — deploy by uploading this folder to a SharePoint library.

**The mission, one sentence:** code written in the pad runs **unmodified** on a real SharePoint page. Every design decision below serves that.

## Architecture invariants — do not break these

1. **Fresh same-origin `srcdoc` iframe per run.** Never blob URLs (opaque origin broke PnPjs in the previous incarnation of this project), never reuse a frame (leaked state caused "sometimes it initializes, sometimes it doesn't").
2. **Deterministic assembly** in `src/runner.js` `assemble()`. Fixed order: harness script → SP context + `<base>` → preview-chrome style → library CSS → user CSS → user HTML → library JS as ordered blocking `<script src>` → **user JS last**. The browser's parser handles ordering exactly like a real page. Never inject scripts into a live preview document.
3. **`postMessage`-only across the frame boundary.** The harness (`src/bridge/harness.js`) pre-serializes everything and posts it with a per-run token; the app ignores messages whose token isn't current. Never reach into the iframe from app code (tests may read computed styles — that's it).
4. **Pad chrome must lose to user code.** Anything the pad adds to the preview (e.g. the dark-mode canvas style) is injected *before* library/user CSS so the user's styling always wins. The pad must never misrepresent how code will render on a real page.
5. **No framework, no user build — but the hosted artifact is bundled.**
   Vanilla ES modules, plain DOM; standalone `index.html` and the test suites
   load `src/` unbundled. The web-part hosting, however, loads
   `dcspad.app.js` — a single-file esbuild bundle of the same source
   (`tools/build-app.mjs`, run by `deploy/Sync-Live.ps1`) — because
   SharePoint's caching makes a multi-file module graph un-bustable (see
   Gotchas). **Rebuild the app bundle after touching `src/`; rebuild the
   Monaco vendor set after changing Monaco, PnPjs types, or
   `tools/build-monaco.mjs`; regenerate `vendor/intelligence/` after changing
   configured BMO design-system CSS or the Fluent icon catalog.** Monaco lives under `vendor/monaco/` as one
   ESM entry, CSS/font assets, same-origin classic `.js` workers, and the
   exact PnPjs 2.15.0 declaration graph. There are no CDN, blob-worker, or
   `.mjs` dependencies. `boot.js` versions the complete set through
   `vendor/monaco/version.json`.
6. **`state.js` is the only module that touches localStorage.** It owns three documents: the workspace blob `{html, css, js, libraries, settings, layout}` (live, autosaved), the framework catalog, and the snippet library. The future SharePoint storage layer (`DevPadData/{user}.json`) replaces persistence wholesale by swapping this one seam — don't scatter storage elsewhere. Files on disk (project/catalog/snippet .json, pane exports) go through `src/io.js`, which moves bytes but stores nothing.

## File map

```
index.html                app shell (single source of truth; standalone + fetched by boot.js)
dcspad.webpart.html       2-line web-part entry: anchor + absolute <script src=boot.js>
boot.js                   web-part bootstrap: fetches index.html no-store, injects shell,
                          imports the versioned bundle; adds .dcspad-hosted to <html>
dcspad.app.js             generated single-file ESM bundle of src/ (tools/build-app.mjs);
                          what the web part actually runs — rebuild on every deploy
styles/app.css            all styling; layout via CSS grid + JS-set vars (--sidebar-w etc.)
vendor/monaco/            generated Monaco runtime, workers, CSS/font + PnPjs type graph
vendor/intelligence/      generated BMO design-token/class + Fluent icon data
tools/build-monaco.mjs    reproducible Monaco/PnPjs vendor build; never hand-edit its output
tools/build-design-intelligence.mjs deterministic BMO CSS/class + Fluent icon generator
tools/build-starter-snippets.mjs generates and catalog-validates examples/dcspad-starter-snippets.json
examples/                 import-ready starter snippet library + usage/prerequisite notes
src/main.js               bootstrap; wires every module; run() lives here
src/layout.js             splitters, tabs, collapse/maximize; persists via state.layout
src/editors.js            Monaco adapter; 3 models, Mod-Enter, PnPjs types, stack jumps
src/monaco-runtime.js     hosted/standalone asset URLs + same-origin worker wiring
src/json-language.js      JSON diagnostics options + dcspad-dark token colours +
                          .json-URI model adoption (highlighting/validation itself
                          is Monaco's own service in vendor/monaco/json.worker.js)
src/config.js             loads/normalizes editable dcspad.config.json runtime URLs
src/docs.js               same-tenant HTML/Markdown/code/text Browser + configured bookmarks
src/sp-chrome.js          hosted SP chip toggle for SuiteNavWrapper + app reflow
src/intelligence/alpine.js Alpine v3 JS declarations + HTML data/completion provider
src/intelligence/bsp.js   BMO CSS-token + HTML-class completion/hover providers
src/intelligence/fluent-icons.js Fluent element/sprite/font completion, hover, diagnostics
src/state.js              defaults + deep-merge load + debounced autosave; loadDoc/saveDoc
                          for the catalog + snippet documents (sole localStorage toucher)
src/io.js                 file download + JSON file-picker helpers (no storage)
src/runner.js             assemble() + iframe lifecycle + run tokens + evalInFrame()
src/libraries.js          framework catalog: single stored JSON (seeded once from PRESETS,
                          then authoritative), add/remove/reorder, getEnabledLibraries()
src/snippets.js           snippet library: save from selection, insert-at-cursor, file I/O
src/console-panel.js      console rendering, filters, groups, REPL input, stack-frame links
src/network-panel.js      request rows, _api filter, detail pane (JSON via inspector)
src/splash.js             readiness-gated splash controller; boot.js starts it earliest when hosted
src/inspect/tree-view.js  generic expandable trees + table renderer (serialized-node format)
src/inspect/sp-shapes.js  SP/OData/PnPjs shape detection + smart views (standalone by design —
                          the future Site Inspector reuses it)
src/bridge/harness.js     iframe-side instrumentation; plain classic script, no imports;
                          __DCSPAD_TOKEN__ placeholder replaced per run
src/bridge/fluent-icon-font.js preview-only font-backed <fluent-icon> adapter
src/bridge/sp-context.js  real _spPageContextInfo capture (live) or labeled mock
src/sp-odata.js           shared OData plumbing (ACCEPT_JSON, SpFileError, requireOk…)
                          used by sp-files.js and the workbench REST client
src/sp-files.js           SharePoint document-library text transfer (digest cache,
                          same-tenant site switching)
src/inspect/to-node.js    JSON → serialized-node adapter for the inspector renderers
workbench.html            SP Workbench shell (second entry point; standalone too)
boot-workbench.js         workbench web-part bootstrap (trimmed boot.js sibling;
                          own ?v= bump rule in workbench.webpart.html)
workbench.webpart.html    2-line web-part entry for the workbench hosting page
dcspad.workbench.js       generated workbench bundle (tools/build-workbench.mjs);
                          rebuild on every deploy — Sync-Live.ps1 does it
styles/workbench.css      workbench layout (loads after app.css, .wb-scoped)
src/workbench/            SP Workbench, a site inspector (read + Tier 2 edit tools):
                          main.js bootstrap + same-tenant site switcher + favorites ·
                          shell.js nav/routing with optional rail groups
                          (sessionStorage only — invariant 6 untouched) ·
                          sp-rest.js GET client on sp-odata helpers (nometadata,
                          paging cap 5000, 429/503 retry, mock resolver,
                          connectWeb target-web switching with origin guard) ·
                          sp-write.js POST client layered on the sp-files.js digest
                          cache (ValidateUpdateListItem, binary AddUsingPath, mock
                          writer recording to __DCSPAD_WB_WRITES__) ·
                          field-editor.js per-TypeAsString metadata editors +
                          FieldValue string conventions (pure, unit-tested) ·
                          favorites.js durable site favorites/recents through
                          state.js loadDoc/saveDoc (key dcspad.v2.wbsites) ·
                          canvas.js tolerant CanvasContent1 parser (JSON + legacy
                          HTML formats, WEBPART_NAMES id map, section tree,
                          sanitizer — never throws, degrades to raw) ·
                          page-export.js content (.md) + raw (.json) page export
                          builders (pure; location line, merged content,
                          standardized metadata block) ·
                          config-links.js CURATED _layouts jump data (Joe-approved
                          set — don't grow without asking) ·
                          mock-data.js offline fixtures ·
                          grid.js sort/filter/export/copy-as + render cells ·
                          export.js CSV/JSON/Markdown · scriptgen.js descriptor →
                          PnPjs 2 / REST fetch / PnP.PowerShell · perm-kinds.js
                          64-bit SPBasePermissions decode · views/: site-home.js
                          (Site landing: web/user cards + subweb Inspect) ·
                          security.js (Permissions: groups, Members roster with
                          add/remove, roledefs, assignments, inheritance) ·
                          lists.js · pages.js (Extract/Metadata/Structure/Web
                          parts/Raw) · browser.js (Files) · query.js · links.js
                          (Panels) · site.js (Advanced sheets)
tests/                    Playwright verification suites — see tests/README.md
REVIEW-LOG.md             external-review triage record + accepted low-priority backlog
```

## Dev workflow

```bash
python3 -m http.server 8642     # from the repo root; app at http://localhost:8642/index.html
cd tools && npm run build:intelligence  # after design CSS or Fluent catalog changes
```

Outside SharePoint the SP chip shows **Mock** and `_api` calls 404 — expected. Live PnPjs/REST behavior can only be validated in a tenant (deployment + validation checklist in README.md). Run the test suites (below) after any change to runner/harness/console/inspector — they exist because this project's failure modes are timing- and boundary-shaped, not type-shaped.

## Tests

`tests/README.md` has the two-server setup (app on 8642, fixtures on 8643) and how Chromium is resolved. Suites: `smoke.mjs` (59 checks: capture, Fluent preview runtime, isolation, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave), `monaco.mjs` (44: typed models, editor integration, mirrored PnPjs 2.15/Alpine 3.15.2 runtime policy, pnp/pnp2/Alpine/BMO/Fluent completion and hover, generated-data coverage, migration-safe runtime detection, false-diagnostic coverage, composable declaration lifecycle, isolated snippet undo/redo, the JSON language service — tokenization, trailing-comma validation, .json-URI adoption, schema-request policy — the scrollbar overview ruler, persistent worker failure behavior), `config.mjs` (21: runtime URLs, same-tenant Browser formats/scripts/history/refresh/tenant rejection, Copilot launcher, framework and asset intelligence/runtime, ordered fallback), `hosted.mjs` (14: early/delayed splash, SharePoint chrome reflow, exact boot/bundle/config path, versioned hosted assets/intelligence/Fluent bridge and same-origin worker), `darkmode.mjs` (8), `splash.mjs` (3), `ux.mjs` (18: pane toggles incl. Ctrl+J through Monaco, editor text-size stepper, error count pills, REPL Eval button, add-framework footer validation, sidebar split persistence), `files.mjs` (23: local import confirmation, SharePoint context fallbacks, same-tenant site switching, ResourcePath and Browser-mode folder browsing, cross-site reads/digests, uploads, and overwrite consent), `workbench-hosted.mjs` (10: hosted Workbench boot, viewport pinning, edit-mode suspension, and boot guard), `workbench.mjs` (69: SP Workbench Site landing + role chip, mock grids/drilldown, Permissions incl. the flattened Members roster with add/remove writes, Advanced sheets, export formats, script generator, same-tenant site switching, the stubbed live path incl. Accept header + paging + re-targeted requests, plus Tier 2: curated Panels, favorites via the state.js seam, query builder, CanvasContent1 page inspector incl. folders/Extract tabs/URL copy, pages metadata editors, page content/raw exports, field-editor conventions), and `workbench-edit.mjs` (15: Files browser + write paths — listing/sort/breadcrumbs, download/copy-URL actions, per-type metadata editors, binary uploads with digest, overwrite consent incl. the 409 race, metadata-failure retry that never re-uploads, oversized rejection). All 293 should pass (per-suite counts in tests/README.md are authoritative); a `custom library` failure usually means the 8643 fixture server isn't running. The serving root must also expose `bsp-design-system/` and `bsp-fluent-icon-lib/` (sibling repos — junctions inside the repo root work).

## Gotchas already paid for

- SharePoint hosting makes a multi-file module graph un-bustable — three
  stacked facts, each proven empirically on this tenant: SPO serves library
  files with `cache-control: public, max-age=86400`; Chrome caches
  module-script requests separately from fetch() (so fetch-side
  revalidation never refreshes what `import` uses); and the modern page
  freezes import-map registration (late maps are silently ignored). Hence
  hosted mode loads ONE bundled ESM file behind a Last-Modified-versioned
  URL (boot.js `VERSIONED` list: app.css + dcspad.app.js + config + the Monaco
  and design-intelligence manifests; index.html is
  no-store; the harness is runtime-fetched with no-cache). Consequences:
  **rebuild `dcspad.app.js` on every deploy** (`deploy/Sync-Live.ps1` does
  it), and **a boot.js change needs the `?v=` bump in dcspad.webpart.html**
  (boot sits below the versioning layer; the web part busts its own HTML
  fetch with a `?pnp=` timestamp).
- Modern SharePoint pages ship a nonce-based `script-src` CSP (no
  `unsafe-inline`) **that `about:srcdoc` documents inherit**. Every script tag
  the runner assembles must carry the host page's nonce (`hostNonce()` in
  runner.js) or the preview frame silently runs nothing — no errors, no
  messages, markup renders fine. Standalone pages have no nonce and the
  attribute is omitted.

- **`app.css` styles the whole host page, not just the app.** Its unscoped
  `html, body { height: 100% }` + `body { overflow: hidden }` (top of the
  file) apply to the SharePoint page the moment a boot script injects the
  stylesheet — so the host page stops scrolling. Anything hosted must
  therefore pin itself over the viewport (`html.dcspad-hosted .app`, and
  `html.dcspad-hosted .wb` in workbench.css) and hide itself on suspension.
  Leaving a hosted root in normal document flow renders a **blank page** in
  view mode that mysteriously appears in edit mode, where `.dcspad-suspended`
  reverts those rules. Cost one live deploy; `tests/workbench-hosted.mjs`
  now guards it.
- **This machine is Windows ARM64 and has two Node installs.** Builds must run
  under the system Node (`C:\Program Files\nodejs\node.exe`, v24 arm64), not
  the nvm4w x64 one — esbuild ships a native per-arch binary, so an x64
  `node_modules` fails with "You installed esbuild for another platform" the
  moment `Sync-Live.ps1` reaches `build-app.mjs`. `tools/node_modules` was
  once committed to the repo (14 files, including the x64 `esbuild.exe`),
  which made that mismatch travel with a clone; it is now untracked and
  `.gitignore`d. If the build ever throws that error, reinstall with the
  system Node: `cd tools && "C:\Program Files\nodejs\node.exe" "C:\Program Files\nodejs\node_modules\npm\bin\npm-cli.js" install`.
- `harness.js` token substitution must be `replaceAll` — the placeholder also appears in a comment, and `replace` once shipped a broken token check.
- `app.css` has a global `[hidden] { display: none !important; }` guard: any element with a `display` rule plus the `hidden` attribute silently ignores `hidden` without it (an invisible splash overlay once ate every click).
- Don't put interactive controls inside a `<label>` containing a disabled input — the browser treats the whole label as disabled (custom-library ✕ button bug).
- User code is embedded raw into the assembled document: keep the `</script>` / `</style>` escaping (`escScript`/`escStyle` in runner.js) intact.
- An `about:srcdoc` document resolves relative and **fragment** URLs against the *parent's* base URL, not itself. A plain `<a href="#foo">` therefore counts as a cross-document navigation and loads the pad (or, with `<base href>`, the SP web) *inside its own preview pane*. `harness.js` intercepts fragment clicks and re-creates the same-document jump; it deliberately listens in bubble phase and honours `defaultPrevented` so user `<a href="#">` handlers still win.
- CDN egress: the Claude sandbox proxy blocks public CDNs, so library-loading tests use the local fixture; real CDN presets can only be exercised in a browser with normal egress.

## Roadmap (seams already reserved — don't build without being asked)

- **Site Inspector** — v1 SHIPPED as the **SP Workbench** second entry point
  (`src/workbench/`, own hosting page via `workbench.webpart.html`): lists +
  fields/views/content types, groups/role definitions/assignments +
  inheritance scan, site overview, exports, and "Copy as PnPjs 2 / REST /
  PnP.PowerShell". **Tier 2 SHIPPED**: config-links view, durable site
  favorites (state.js seam), quick query builder (descriptors round-trip to
  scriptgen), modern-page inspector (CanvasContent1 parse, structure tree,
  text extraction), pages metadata editor, and the Files browser (binary
  upload + full-field-type metadata editing over the sp-files.js digest
  cache via `sp-write.js`). Remaining seams: mounting the workbench views
  into the pad sidebar (index.html "future: Site Inspector" comment),
  User/Lookup/Taxonomy field editing (formats documented in field-editor.js),
  chunked >50 MB uploads.
- **SharePoint JSON storage** for snippets/projects/templates/shared team resources, replacing localStorage in `state.js`.
- **Console remote handles** — lazy live-object expansion instead of eager capped serialization.
