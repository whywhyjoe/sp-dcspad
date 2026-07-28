# Handoff — DCSPad

Read `CLAUDE.md` first for architecture invariants; this file covers **state**
and **what's next**. Last updated: 2026-07-27, Monaco plus generated BMO and
Fluent design-system intelligence.

For the UI redesign handoff, also read
`design/POST-MONACO-UI-INTEGRATION.md`. It documents the exact pre-Monaco
baseline, all post-baseline changes, every setting and persistence boundary,
and the functional DOM/CSS hooks the workbench design system must preserve.

## Where things stand

**The pad runs on the real SharePoint page.** Editors, run pipeline, console,
inspector, REPL and network capture all work inside the web part; a live
`_api/web` call from the REPL returns real site data. The standalone
`index.html` still works and remains what the test suites drive.

### Hosting model (decided and working)

- Host page: `https://nervedotnet.sharepoint.com/sites/NewNerve/SitePages/DCSpad.aspx`
  (confirmed live on 2026-07-26; the former `DCSpad(1).aspx` duplicate now
  returns 404).
- Web part: PnP **Modern Script Editor** (Mikael Svenson), in
  "Use script from an external URL" mode. Despite the field name "Script URL",
  it fetches an **HTML file** and injects it with script re-creation.
- The URL points at `dcspad.webpart.html` (repo root): a 2-line anchor +
  absolute `<script src=…/boot.js>`. `boot.js` fetches `index.html`, injects
  the app shell, loads `styles/app.css`, and imports the versioned hosted app
  bundle. `index.html` stays the single source of truth.
- **Deployment = file copy.** The doc library folder
  `…/SiteAssets/Code/dcspad-live/` is OneDrive-synced to
  `C:\Users\other\NERVE\NewNerve - Code\dcspad-live`; copying files there goes
  live in seconds. **Deploy with `deploy/Sync-Live.ps1`** — it rebuilds
  `dcspad.app.js` (the
  single-file bundle the web part actually runs) and copies everything.
  ⚠ Never hand-copy `src/` changes without rebuilding the bundle: hosted
  mode won't see them. Why a bundle: SPO's `max-age=86400`, Chrome's
  separate module-request cache, and the host page ignoring late import
  maps make a multi-file graph un-bustable (full story: cache gotcha in
  CLAUDE.md). Bump `?v=` in `dcspad.webpart.html` when boot.js itself
  changes. SPO can also serve a just-uploaded file stale for ~15–30s —
  verify with a cache-busted fetch before debugging "my change didn't work".
- Visual seating (deliberate, Joe's call): the SharePoint suite bar is
  **visible by default** (desaturated while the pad runs); clicking the
  `SP: Live` chip slides it away and expands the app to the top edge. Hosted
  mode otherwise pins `.app` at
  `inset: 53px 5px 5px`, borderless over a darker surround so it reads as
  part of the page (see "Web-part hosting" in `app.css`, activated by
  `boot.js` adding `.dcspad-hosted` to `<html>`). The earlier `env=WebView`
  chrome-less plan is abandoned — the suite bar is wanted context.
- **Page-edit safety**: boot.js refuses to boot when the URL carries
  `Mode=Edit` (or under `/_layouts/`), and also watches SPA navigation
  (pushState/replaceState/popstate) to *suspend* an already-booted pad —
  `.dcspad-suspended` hides the mount and reverts the html/body overrides so
  the edit canvas looks and scrolls normally. Leaving edit mode restores it.

### Web-part-hosting constraints that were found and fixed

1. **SharePoint serves `.mjs` as `application/octet-stream`** — browsers
   refuse ES modules with that MIME. This is permanent: Monaco's runtime and
   all same-origin workers ship as `.js`.
2. **Modern pages ship a nonce-based `script-src` CSP that `about:srcdoc`
   inherits.** Preview frames rendered markup but silently executed nothing.
   `runner.js` now stamps the host page's nonce on every assembled script tag
   (`hostNonce()`); standalone pages have no nonce → no-op. If preview
   execution ever "silently dies" on SP again, look here first.

3. **Monaco language tooling needs workers.** The editor uses same-origin
   classic worker files, never `blob:` URLs. `deploy/webpart-spike.html`
   includes a worker/CSP check, and the live app surfaces a blocked worker in
   the status bar without losing the editor.

### Spike results (deploy/webpart-spike.html, kept for re-verification)

1. classic injected script executes ✅
2. `type="module"` script executes ✅
3. relative import from an absolute module URL resolves ✅
4. `_spPageContextInfo` present ❌ — absent on the modern page by default
5. same-origin classic worker allowed ✅ — proven by the live Monaco
   TypeScript worker; no `worker-src` error

### Monaco migration (2026-07-26)

- One Monaco editor swaps three typed models (`index.html`, `styles.css`,
  `script.js`) behind the previous editor adapter, so the runner, persistence,
  snippets, console stack jumps, autosave and autorun remain unchanged.
- `vendor/monaco/` contains Monaco 0.55.1, classic `.js` workers, CSS/font,
  and 254 declaration files for the exact PnPjs 2.15.0 dependency graph.
- `boot.js` paints the hosted splash immediately after its mount exists,
  before fetching the shell or probing assets. Its live status survives
  delayed Monaco loading and fades only after the editor and app wiring are
  usable. Its hidden state is committed for two paint frames before the
  entrance class is added, preventing SharePoint from batching away the
  fade-in. The fully painted app and a permanent dark hosted underlay sit
  beneath the curtain, so its fade-out cannot expose white SharePoint
  wrappers. Standalone mode follows the same readiness-gated lifecycle.
- Enabling the PnPjs v2 runtime library loads the declarations and global
  `pnp` bridge; disabling it unloads them. Fluent completions, hover,
  signatures and JS diagnostics use Monaco's TypeScript worker.
- Runtime detection follows the enabled 2.15.0 script URL rather than a fixed
  catalog id, so the live custom `Pnpjs JSD` jsDelivr entry receives the same
  matching types.
- `tools/build-monaco.mjs` is the only way to regenerate vendor assets.
  `vendor/monaco/version.json` versions the set as a unit in hosted mode.
- `tests/monaco.mjs` covers the editor contract, completions, diagnostics,
  declaration lifecycle, isolated snippet undo/redo, asset policy and
  persistent worker failure behavior. `tests/hosted.mjs` deliberately delays
  both `index.html` and Monaco to verify the earliest splash stages and that
  the app is fully painted over the dark underlay before the curtain fades.

### Runtime URL configuration (2026-07-26)

- `dcspad.config.json` is a separately versioned, editable runtime document.
  Framework entries can specify local and CDN URLs, source preference, a
  global probe, and intelligence-pack IDs without rewriting the persisted
  framework catalog.
- If a preferred JavaScript runtime does not expose its configured global,
  the backup is inserted parser-blocking at the same catalog position. This
  preserves dependency order and avoids CSP-blocked inline event handlers.
- The PnPjs 2.15.0 type pack can now follow a custom rollup through explicit
  metadata even when its URL is opaque.
- The imported BMO SharePoint design-system and Fluent-icon repositories are
  source inputs and are not copied by `Sync-Live.ps1`. Their local and eventual
  hosted base folders live under `assets` in the config. The generated compact
  runtime data is copied under `vendor/intelligence/`.
- `tests/config.mjs` covers URL resolution, config-enabled asset intelligence,
  explicit PnP intelligence, and ordered primary/fallback loading.

### Same-tenant Browser and SharePoint chrome toggle (2026-07-27)

- The header globe opens configured Browser bookmarks; the left extras column
  switches between Resources and Browser without disturbing the editor/runtime
  columns. Browser can temporarily maximize and `Esc` restores the layout.
- The address bar accepts `.html`, `.htm`, `.md`, `.markdown`, and `.txt`.
  Every configured, pasted, and followed URL must match `location.origin`,
  which is the current SharePoint tenant origin in hosted mode.
- SharePoint HTML is fetched as text and rendered through `srcdoc` with a base
  URL so relative assets, links, and scripts work despite SharePoint download
  MIME behavior. Tenant HTML is trusted to run scripts; Markdown and plain text
  use built-in scriptless renderers.
- Microsoft 365 Copilot remains an external named-tab shortcut because its
  page blocks iframe embedding.
- `src/sp-chrome.js` owns the hosted-only `SP: Live` button behavior and
  `src/docs.js` owns Browser rendering/navigation. `tests/hosted.mjs` covers
  suite-bar reflow; `tests/config.mjs` covers Browser formats, scripts,
  maximize, tenant rejection, and the Copilot launcher.

### Alpine intelligence (2026-07-26)

- Enabling the Alpine runtime now activates a composable `alpine-3` pack.
  JavaScript completion/hover/signatures cover the public browser API including
  `Alpine.data/store/bind/plugin`, reactivity hooks, and tree lifecycle.
- HTML completion and hover cover all core `x-*` directives, common `@event`
  and `:attribute` shorthands, editable attribute snippets, and the core magic
  properties (`$dispatch`, `$refs`, `$store`, `$watch`, `$nextTick`, `$root`,
  `$data`, `$id`, `$el`, `$event`).
- PnPjs and Alpine declarations live in separate registry entries. Disabling
  one no longer clears the other through `setExtraLibs([])`.
- The HTML provider is context-aware enough to offer magic properties only
  inside Alpine attribute expressions. Inferring arbitrary members from the
  nearest `x-data` object remains the later deep-expression phase.
- `tests/monaco.mjs` now verifies Alpine JS, HTML directives, magics, pack
  coexistence, unloading, legacy-catalog detection, and false diagnostics. The
  complete browser suite is 113 checks.

### BMO design-system intelligence (2026-07-26)

- `tools/build-design-intelligence.mjs` reads the configured local
  `colors_and_type.css`, `components.css`, and `editorial.css` files and emits
  deterministic `vendor/intelligence/bsp-design.json` plus `manifest.json`.
- The current artifact contains 164 CSS custom properties and 375 canonical
  classes. Records include values, categories, nearby source documentation,
  source line numbers, BEM kind/base relationships, and Editorial-mode scope.
- `assets.designSystem.intelligence: ["bsp-design"]` enables the pack
  independently of framework checkboxes. CSS completion/hover works inside
  `var(--...)`; HTML completion/hover works inside `class=""`.
- Hosted mode versions the intelligence manifest as a unit. Because `boot.js`
  changed for that version seam, `dcspad.webpart.html` is now at `boot.js?v=11`.
- `Sync-Live.ps1` regenerates intelligence before rebuilding/copying the app.
  The design-system source folder remains a development input, not a runtime
  dependency.

### Fluent icon intelligence and preview runtime (2026-07-27)

- `tools/build-design-intelligence.mjs` also projects the complete Fluent
  catalog into `vendor/intelligence/fluent-icons.json`: 2,655 normalized icon
  groups and all 18,681 real SVG variants, with font availability preserved.
- Monaco completes and explains `<fluent-icon name="">`, Fluent
  `<use href="...#symbol">` IDs, and generated `icon-ic_fluent_*` font classes.
  Unknown and SVG-only/font mismatches receive focused warnings.
- The configured runtime loads the vendored Regular, Filled, and Light font
  CSS plus `src/bridge/fluent-icon-font.js`, a preview-only adapter for
  `<fluent-icon>`. Generated suffix-specific CSS keeps all three font families
  working together.
- The imported Fluent package contains individual SVG files, not a combined
  symbol sprite. `<use>` intelligence is retained for consuming projects that
  provide a Fluent sprite. No new BMO sprite integration was added because
  that UI icon path is being deprecated.

## Open items

- **CSS bleed, both directions**: `app.css` still styles `html`/`body`
  (darkens the host page behind the pad — currently invisible and arguably
  nice; the gap around the seated app shows it). SP styles also bleed into the
  pad; nothing visibly broken, but scope properly if oddities appear.
- Low-priority backlog lives in `REVIEW-LOG.md`.

## Completed feature work (2026-07-27)

- The Claude Design visual pass is implemented. `design/DESIGN-BRIEF.md` and
  the integration notes remain as historical design inputs.
- Projects start unnamed, expose an inline top-bar name editor, and save as
  required-name `.dcspad.json` files. Local HTML/CSS/JS imports use one
  extension-aware picker with replacement confirmation; exports support one
  pane or all non-empty panes.
- SharePoint HTML/CSS/JS import and export are implemented through
  `src/sp-files.js`, including explicit/global/Modern context resolution,
  selectable same-tenant sites, per-web digest handling, ResourcePath
  browsing/reads/uploads, replacement confirmation, and overwrite consent.
  See `plans/file-sp-import-export.md`, which is now an implementation record.
- Framework rows use drag-and-drop ordering without up/down controls; snippets
  are always displayed alphabetically regardless of file type.

## Feature backlog

- **Framework/snippet row action icons only on hover** (currently always
  visible at 60% opacity).
- **Site Inspector / SP diagnostics tools** — already in the roadmap below;
  explicitly wanted.

## Roadmap (seams reserved)

- **Site Inspector** sidebar section (renders through `src/inspect/`).
- **SharePoint JSON storage** replacing localStorage via the `state.js` seam.
- **Console remote handles** — lazy live-object expansion.
