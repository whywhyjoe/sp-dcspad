# DCSPad — SharePoint Developer Workbench

A JSFiddle-style workbench for SharePoint development that runs entirely in the browser: Monaco HTML/CSS/JS editors, live preview, built-in console + network monitor, an SP-aware object inspector, and a library manager. No backend, no user build step, no SPFx — deploy by uploading this folder to a SharePoint library.

This is a from-scratch rebuild focused on the **execution environment**. The core guarantee: **code written in the pad runs unmodified on a real SharePoint page.**

## Why the old version broke, and how this one doesn't

The previous build ran user code in a `Blob` + `createObjectURL()` iframe. Blob URLs get an *opaque origin*: `window.location` becomes `blob:…`, cookies don't flow, and PnPjs/SP REST can't resolve `_api` URLs → `"Failed to parse URL from ''"`. Script injection into a live document also made load order a race → "sometimes it initializes, sometimes it doesn't."

This version:

1. **Same-origin `srcdoc` iframe.** The preview document inherits the host SharePoint page's origin — cookies, auth, and same-origin REST just work.
2. **Deterministic full-document assembly.** Every run builds one complete HTML document — harness → SP context + `<base>` → library CSS → user CSS → user HTML → ordered blocking library scripts → user JS last — and hands it to the browser's normal parser. Exactly how a real page loads; no injection races.
3. **Fresh iframe per run.** The old iframe is destroyed; no leaked globals, no half-initialized state.
4. **`postMessage`-only instrumentation.** The in-iframe harness never gets reached into from outside; all console/network/error/REPL traffic crosses the boundary as messages, tagged with a per-run token so stale frames are ignored.
5. **SP context bridge.** Context is resolved through a stable DCSPad host adapter, `_spPageContextInfo`, or a guarded Modern-page `legacyPageContext` fallback. It is re-captured at every run and `<base href>` points at the web, so PnPjs v2 can auto-resolve its base URL. SharePoint file writes obtain or refresh their digest with `/_api/contextinfo` and retry once after a digest-related 403.

## Saving, loading, exporting

Workspace state lives in your browser's localStorage. Project files and local
code exports are ordinary downloads; HTML/CSS/JS panes can also be transferred
to SharePoint. There is no DCSPad backend.

- **Project name and project file** — a new workspace starts as
  **Project (untitled)**. Click the title to set or edit it; the name is
  autosaved and becomes a safe lowercase filename slug. Saving project JSON
  requires a non-empty project name and downloads
  `<project-name>.dcspad.json`; choosing Save while untitled opens the name
  editor and resumes the download after a valid name is entered. Loading a
  `.dcspad.json` file restores its name, panes, enabled frameworks, DCS URL,
  and module setting. Project JSON remains local-only.
- **Local code files** — **Import HTML, CSS, or JS…** accepts one `.html`,
  `.htm`, `.css`, or `.js` file, infers its pane from the extension, and asks
  before replacing that pane. Export one pane or choose **Export all
  non-empty** to trigger a separate download for every non-empty pane. A named
  project supplies the filename slug; an untitled project falls back to
  `dcspad.html`, `dcspad.css`, and `dcspad.js`.
- **SharePoint code files** — when the chip reads **SP: Live**, enter any
  SharePoint site URL on the same tenant origin, then browse that web's
  libraries/folders to import or upload HTML, CSS, and JavaScript. The selected
  site and last folder persist. Imports require replacement confirmation and
  uploads require explicit overwrite confirmation. These operations use the
  signed-in user's SharePoint REST permissions—no Graph app, PnP.PowerShell,
  or separate login. Project, framework-catalog, and snippet JSON never use
  this picker.
- **Framework catalog** — the checkbox list in the sidebar is a single stored JSON document, seeded once with the built-in presets and then yours: add entries by URL (with an optional name), remove any entry, and drag rows to reorder them. Order is injection order, so put a plugin below the library it extends. The ⤓/⤒ buttons save/load the whole catalog as a file. If a loaded project references a framework you've since removed, it still loads — you get a console warning naming it, and the run fails with the usual `X is not defined` until you re-add it.
- **Snippets** — save the current editor selection (or whole pane) as a named snippet with ＋; click a snippet to insert it at the cursor of its editor. ⤓/⤒ save/load the snippet library as a file. A maintained, import-ready 37-entry starter pack lives at [`examples/dcspad-starter-snippets.json`](examples/dcspad-starter-snippets.json), with usage notes in [`examples/README.md`](examples/README.md).

## Runtime configuration

Edit [`dcspad.config.json`](dcspad.config.json) to change environment-specific
URLs without editing or rebuilding the application.

- `frameworks.items.<catalog-id>.localUrl` is the preferred self-hosted or
  organization-hosted script.
- `cdnUrl` is its backup. Set `frameworks.prefer` to `"cdn"` to reverse the
  order, or `fallbackToCdn` to `false` to disable automatic fallback.
- `probeGlobal` is the global the primary script must expose. When it is absent,
  DCSPad inserts the backup as a parser-blocking script at the same catalog
  position, so plugins and user JavaScript still run in the expected order.
- `intelligence` explicitly associates editor metadata with the runtime. The
  custom PnPjs rollup can therefore keep `["pnpjs-2.15.0"]` even when its URL
  does not contain a recognizable package/version path.
- `assets.designSystem` and `assets.fluentIcons` hold local-review and eventual
  hosted base folders. Relative local paths resolve from `dcspad.config.json`;
  hosted locations may be absolute SharePoint URLs.
- `assets.designSystem.intelligence: ["bsp-design"]` enables the generated BMO
  design-system pack independently of framework checkboxes. Remove that ID to
  disable it. The browser loads the compact versioned artifact under
  `vendor/intelligence/`, never the source CSS repository.
- `assets.fluentIcons.intelligence: ["fluent-icons"]` enables completions,
  hover, and diagnostics for real Fluent tokens, symbol IDs, and font classes.
  Its `runtime` block controls which configured font CSS files and the
  preview-only `<fluent-icon>` adapter are injected on Run. Set
  `runtime.enabled` to `false` if the consuming page supplies those assets.
- `docs` supplies the Browser menu bookmarks. Each entry has `id`, `title`,
  `url`, and an optional `type` (`html`, `markdown`, `text`, or `auto`).
  Relative URLs resolve from this configuration file. The Browser also accepts
  pasted `.html`, `.htm`, `.md`, `.markdown`, and `.txt` URLs, but enforces the
  exact origin hosting DCSPad. Same-tenant HTML runs its scripts; Markdown and
  text are rendered without script execution.
- `copilot` controls the external Microsoft 365 Copilot shortcut. It opens one
  reusable named tab because the service does not permit iframe embedding.

Blank URLs are ignored. With the supplied file, PnPjs and Alpine continue using
their existing CDN URLs until local copies are filled in. Hosted mode versions
the JSON independently, and `deploy/Sync-Live.ps1` copies it automatically.

## Local development

```bash
python3 -m http.server 8642      # from the repo root
# open http://localhost:8642/index.html
```

Outside SharePoint you get a clearly-flagged **mock** `_spPageContextInfo` (correct shape, `SP: Mock` chip). `_api` calls will 404 locally — that's expected; deploy to SharePoint for live APIs.

## Deploying to SharePoint

Sync the target SharePoint document library with OneDrive, then run the build-and-copy
deployment. See [`deploy/README.md`](deploy/README.md) for the full runbook and
first-deployment checks.

```powershell
.\deploy\Sync-Live.ps1
```

For another synced site or tenant, supply its local target folder:

```powershell
.\deploy\Sync-Live.ps1 -LivePath 'C:\path\to\the-synced-library\dcspad-live'
```

The script rebuilds the hosted bundle and design-system intelligence before
copying the complete runtime. Confirm the top-right chip reads **SP: Live**
after OneDrive finishes syncing.

### Explicit host-context adapter

DCSPad does not require a host to recreate the complete
`_spPageContextInfo`. A custom script editor or SPFx wrapper can provide only
the current web URL before `boot.js` loads:

```html
<script>
window.__DCSPAD_SP_CONTEXT__ = {
  webAbsoluteUrl: "https://contoso.sharepoint.com/sites/dev"
};
</script>
<script src="/sites/dev/SiteAssets/dcspad/boot.js"></script>
```

An SPFx host can populate it dynamically:

```js
window.__DCSPAD_SP_CONTEXT__ = {
  webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
  userDisplayName: this.context.pageContext.user.displayName
};
```

`webAbsoluteUrl` is the only required field. DCSPad validates that it is
same-origin, then uses `POST {webAbsoluteUrl}/_api/contextinfo` to obtain the
canonical web/site URLs and a request digest. Optional `pageContext` and user
display fields may also be supplied. A same-origin iframe can expose this
object on `parent`; cross-origin embedding requires an explicit same-origin
adapter and is intentionally not trusted automatically.

When the adapter is absent, DCSPad falls back to `_spPageContextInfo`, then to
the Modern page's internal `legacyPageContext`. The Modern lookup is guarded
because its `spModuleLoader` path is undocumented and may change. If the chip
still says Mock on SharePoint, add the explicit adapter above rather than
depending on that private implementation.

The adapter identifies the site hosting DCSPad; it does not limit file
transfer to that site. The SharePoint-files dialog has its own persisted
**SharePoint site** field. Opening another same-tenant site validates
`{siteUrl}/_api/contextinfo`, establishes that site's server-relative boundary,
and obtains a digest specifically for that web. A different hostname is
rejected because normal browser same-origin rules prevent reusing the current
SharePoint session there.

### Tenant validation checklist

1. Chip shows **SP: Live**; status bar shows your web URL and user name.
2. Enable the **PnPjs v2 (classic)** library, then Run:
   ```js
   const { sp } = pnp;
   sp.web.get().then(w => console.log(w));
   ```
   — unmodified, no setup call. The web object should render in the console with the SP-aware inspector (badge, Title/ServerRelativeUrl header).
3. The request appears in the **Network** tab (try the `_api only` filter); click the row — the JSON body renders through the inspector.
4. Try the REPL line at the bottom of the console while the run is alive:
   `sp.web.lists.get()` — the promise is awaited and the list collection renders with a table-view toggle and copyable GUIDs/EntityTypeNames.
5. Paste the same code into a real page (script editor / custom script) and confirm identical behavior. That's the point.
6. Choose **File ▸ Import from SharePoint…**, browse the current web, select
   an HTML/CSS/JS file, and confirm that DCSPad asks before replacing the
   matching editor. Repeat with another site URL on the same tenant.
7. In a disposable document-library folder, choose **File ▸ Export to
   SharePoint…** and verify a new upload succeeds. Repeat with the same name
   and confirm that the second action requires explicit overwrite consent.

## Using the pad

| Thing | How |
|---|---|
| Run | `Run` button or `Ctrl/Cmd+Enter` anywhere in an editor |
| Auto-run | Toggle in the toolbar; re-runs ~800 ms after you stop typing |
| Language tools | Monaco find/replace, suggestions, hover, signatures, diagnostics and navigation; enabling PnPjs v2 adds matching 2.15.0 fluent API types; enabling Alpine adds v3 JavaScript API and HTML intelligence; the configured design packs add documented BMO CSS/classes plus exact Fluent icon tokens, symbol IDs, and font classes |
| Editor status | The status bar shows `Monaco ✓`; `Monaco ⚠` remains visible if a worker is blocked, even after running code |
| Top-level `await` | Settings ⚙ → "Run JS as module" (strict mode; `var` won't become window globals) |
| REPL | Input line under the console — evaluates *inside the current run's iframe*; `↑`/`↓` history; promises are awaited |
| Stack traces | Frames pointing into your JS are clickable → jumps the editor to that line |
| Libraries | Left sidebar; checkbox = include on next run, ★ = pin to top; custom URLs (`.js`/`.css`) at the bottom |
| Browser | Globe button for configured bookmarks; the left Browser tab accepts same-tenant HTML, Markdown, and text URLs and can be maximized |
| SharePoint toolbar | In hosted mode, click the `SP: Live` chip to hide/show the suite bar and reclaim/restore its space |
| Network | `_api only` filter; click a row for the response body rendered through the SP inspector |
| Maximize | ⛶ on the preview or console panel; `Esc` restores |
| Preview dark mode | ☀/🌙 on the preview header (default dark). Pad-only canvas color injected *before* your CSS, so anything you style wins — flip to light to see how it renders on a typical SharePoint page |

Work-in-progress (editors, libraries, settings, layout) autosaves to `localStorage` and restores on load.

### Testing framework intelligence

Framework intelligence follows the enabled runtime checkbox.

- **PnPjs:** enable **PnPjs v2**, type `pnp.sp.w` in JavaScript, then press
  `Ctrl+Space`; `web` should be offered.
- **Alpine JavaScript:** enable **Alpine.js**, type `Alpine.d`, then press
  `Ctrl+Space`; `data` should be offered. `Alpine.s` and `Alpine.p` offer
  `store` and `plugin`.
- **Alpine HTML directives:** type `<div x-d`, then press `Ctrl+Space`;
  `x-data` should be offered with documentation and an editable state snippet.
- **Alpine HTML magics:** type `<button x-data @click="$d`, then press
  `Ctrl+Space`; `$dispatch` should be offered. `$refs`, `$store`, `$watch`,
  `$nextTick`, `$root`, `$data`, `$id`, `$el`, and `$event` are also included.
- **BMO CSS tokens:** type `color: var(--fg-p`, then press `Ctrl+Space`;
  `--fg-primary` should be offered with its resolved design-system value,
  category, source line, and hover documentation.
- **BMO HTML classes:** type `<button class="btn btn--p`, then press
  `Ctrl+Space`; `btn--primary` should be offered as a BEM modifier with
  `.btn` composition guidance and source documentation.
- **Fluent custom element:** type `<fluent-icon name="home-24-r`, then press
  `Ctrl+Space`; `home-24-regular` should be offered. Run it to verify the
  configured icon font renders in the preview.
- **Fluent sprite:** inside a project that supplies a symbol sprite, type
  `<svg><use href="#ic_fluent_home_24_r`, then press `Ctrl+Space`; the exact
  `ic_fluent_home_24_regular` symbol ID should be offered.
- **Fluent font class:** type
  `<i class="icon-ic_fluent_home_24_r`, then press `Ctrl+Space`; the generated
  `icon-ic_fluent_home_24_regular` class should be offered.

The first Alpine pack understands the stable core API and whether the cursor is
inside an Alpine attribute. It does not yet infer arbitrary properties declared
inside the surrounding `x-data` object. Plugin-only features such as `$persist`,
`x-collapse`, and `x-trap` require their own enabled runtime/intelligence packs.

## The SP-aware inspector

Console output and network response bodies are rendered by an inspector that understands SharePoint shapes:

- **OData envelopes** (`d`, `d.results`, `value`) are unwrapped — payload first, envelope metadata folded away, item-count badge.
- **Collections** get a tree ⇄ table toggle; plumbing fields (`__metadata`, `odata.*`) are dimmed/dropped.
- **Entities** (List, Field, Web, User, Group, list items) get compact headers with the fields you actually need — Ids, `EntityTypeName`, `InternalName`, `TypeAsString` — **click to copy**.
- **Paging** — `__next` / `odata.nextLink` is surfaced loudly so truncated result sets don't slip past.
- **PnPjs** plain results and `HttpRequestError` (with HTTP status) are recognized.

## Repo layout

```
index.html            app shell
styles/app.css        theme + layout
vendor/monaco/        generated Monaco runtime, workers, CSS/font + PnPjs 2.15 types
vendor/intelligence/  generated compact BMO token/class + Fluent icon data
tools/build-monaco.mjs reproducible Monaco/PnPjs vendor build (esbuild)
tools/build-design-intelligence.mjs deterministic BMO CSS/class + Fluent icon generator
tools/build-starter-snippets.mjs validated starter snippet-library generator
examples/              import-ready PnPjs, Alpine, BSP, Fluent, and general starters
src/
  main.js             bootstrap/wiring
  layout.js           splitters, tabs, collapse/maximize (persisted)
  editors.js          Monaco adapter: three models, language tools, editor API
  monaco-runtime.js   standalone/hosted asset and same-origin worker loading
  intelligence/bsp.js BMO CSS-token and HTML-class completion/hover providers
  intelligence/fluent-icons.js Fluent element/sprite/font completion, hover, diagnostics
  state.js            workspace state + debounced autosave
  runner.js           document assembly + iframe lifecycle
  libraries.js        preset catalog + custom URLs
  console-panel.js    console UI + REPL
  network-panel.js    network UI
  splash.js           readiness-gated boot splash controller
  inspect/tree-view.js   generic expandable trees + tables
  inspect/sp-shapes.js   SP/OData/PnPjs smart views
  bridge/harness.js   iframe-side instrumentation (injected per run)
  bridge/fluent-icon-font.js preview-only font-backed <fluent-icon> adapter
  bridge/sp-context.js   real/mock _spPageContextInfo capture
```

For UI/design-system work based on the pre-Monaco shell, read
[`design/POST-MONACO-UI-INTEGRATION.md`](design/POST-MONACO-UI-INTEGRATION.md).
It is the authoritative migration and UI contract: exact baseline, changed DOM
hooks, complete settings/persistence reference, Monaco theme/widget rules,
splash/dialog states, runtime configuration, and code ownership.

Monaco is fully vendored: the ESM runtime, CSS/font, classic same-origin `.js`
workers, and exact PnPjs 2.15.0 declaration graph. There is no runtime CDN,
blob worker, or `.mjs` dependency. From `tools/`, run `npm run build:monaco`
after changing the pinned editor/type versions or the vendor builder.

Design-system intelligence is generated rather than scraped in the browser.
From `tools/`, run `npm run build:intelligence` after changing the configured
local design-system CSS or Fluent icon catalog. `deploy/Sync-Live.ps1` runs
that generator and the app bundle build automatically.

## Tests

Playwright browser suites live in `tests/` — see `tests/README.md` for the two-server setup. Run them after touching the runner, harness, console/network panels, or inspector.

## Roadmap (deliberately not built yet)

- **Site Inspector** — sidebar panel for the discovery every SP build starts with: lists/libraries with GUIDs + internal names, fields with display/internal names + types, security groups & members, content types — rendered through `src/inspect/` with "copy as PnPjs call" so discovery flows into code.
- **Snippets / projects / templates / shared team resources** on SharePoint-backed JSON storage (`DevPadData/{user}.json`), replacing localStorage wholesale — `state.js` already keeps the workspace as a single serializable blob for exactly this.
- **Console remote handles** — lazy live-object expansion instead of eager capped serialization.
