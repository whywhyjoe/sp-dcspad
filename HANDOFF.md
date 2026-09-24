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

- Host page: `https://nervedotnet.sharepoint.com/sites/NewNerve/SitePages/tools/DCSpad.aspx`
  (this page location is unchanged by the runtime-library move).
- Web part: PnP **Modern Script Editor** (Mikael Svenson), in
  "Use script from an external URL" mode. Despite the field name "Script URL",
  it fetches an **HTML file** and injects it with script re-creation.
- The URL points at `dcspad.webpart.html` (repo root): a 2-line anchor +
  absolute `<script src=…/boot.js>`. `boot.js` fetches `index.html`, injects
  the app shell, loads `styles/app.css`, and imports the versioned hosted app
  bundle. `index.html` stays the single source of truth.
- **Deployment = file copy.** The application folder
  `…/FCUPortal/Dev/tools/dcspad/` is OneDrive-synced under the local Dev mirror at
  `C:\dev\fcuportal-dev\tools\dcspad`; copying files there goes
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
  Shared BSP Design, Fluent Icons, and other organization-hosted resources
  remain in `/Code` or `/Code/tools`. The DCSPad and SP Workbench ASPX pages
  remain in `/sites/NewNerve/SitePages/tools/`; only the runtime folder moved.
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
- The address bar accepts `.html`, `.htm`, `.md`, `.markdown`, `.css`, `.js`,
  `.json`, `.csv`, and `.txt`.
  Every configured, pasted, and followed URL must match `location.origin`,
  which is the current SharePoint tenant origin in hosted mode.
- The address bar persists the 10 most recent unique URLs, provides a
  cache-bypassing refresh action, and opens the existing SharePoint file
  dialog in Browser mode. That mode retains the editable site URL, same-tenant
  enforcement, remembered folder, parent/refresh navigation, and folder
  traversal used by import/export, while filtering to HTML/Markdown/code/text.
- SharePoint HTML is fetched as text and rendered through `srcdoc` with a base
  URL so relative assets, links, and scripts work despite SharePoint download
  MIME behavior. Tenant HTML is trusted to run scripts; Markdown, code, and
  plain text use built-in scriptless renderers.
- Microsoft 365 Copilot remains an external named-tab shortcut because its
  page blocks iframe embedding.
- `src/sp-chrome.js` owns the hosted-only `SP: Live` button behavior and
  `src/docs.js` owns Browser rendering/navigation. `tests/hosted.mjs` covers
  suite-bar reflow; `tests/config.mjs` covers Browser formats, scripts,
  history, refresh, maximize, tenant rejection, and the Copilot launcher.
  `tests/files.mjs` covers Browser-mode subsite switching, folder traversal,
  resource filtering, and selection.

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

- **SP Workbench links do not reliably open in a new tab/window**:
  [GitHub issue #10](https://github.com/whywhyjoe/sp-dcspad/issues/10).
  In the SharePoint-hosted Workbench, links intended for a separate tab or
  window only rarely open there despite `target="_blank"` and the explicit
  `bindNewTab()` handler. Reproduce and validate the fix in a live SharePoint
  host; the standalone/mock popup test is not sufficient tenant evidence.
  **Reproduced live 2026-09-23 (dev, Build #215–#220) — root cause narrowed:**
  every `bindNewTab` anchor (Panels, grid ↗, the Pages drilldown's Open
  page, library links) navigates the *Workbench tab itself*; no new tab
  opens. SharePoint registers a capture-phase `click` listener on `window`
  (before the bundle loads) that routes any `<a href>` click in the same
  tab — even when the anchor's own listener calls `preventDefault()` +
  `stopPropagation()`, and even for a synthetic `element.click()`; it acts
  on `click`, not on pointer/mouse down/up. `window.open` itself is fine:
  an injected `<button>` calling `window.open(url, '_blank', 'noopener')`
  on the same page opens a new tab and leaves the Workbench in place. So
  the likely fix is a non-anchor control (button, or an anchor without
  `href` carrying the URL in a data attribute) — at the cost of native
  middle/ctrl-click, which is Joe's call.
- **Prod (bmo) reconcile follow-ups — unverified** (from the 2026-09-17
  work-prod reconcile, whose state file is retired; all of its code is on
  `main` via PR #19). Prod has since been redeployed (Joe saw a
  `201-dirty` build stamp on 2026-09-23 — `-dirty` means the work clone had
  uncommitted tracked changes at build time, most likely
  `vendor/intelligence/` regenerated from that machine's design-system
  repos; check `git status` there). Still to confirm:
  - README.md's validation step 8 (overwrite in a **Require Check Out**
    library, from the pad and from Workbench Files) on the work tenant.
  - Then delete the home machine's local-only branch
    `recovered/work-checkout-prod` (the literal old prod source, kept only
    as a reference until step 8 passes).
  - Open question for Joe: the check-out consent box now starts
    **unticked** (Overwrite disabled until ticked); the old prod build
    pre-ticked it. One-line change in `applySpCheckoutState()`
    (`src/main.js`) if pre-ticked is preferred.
- **Dev-tenant test leftovers**: the `zz-schema-*` lists on `/sites/NewNerve`
  and `/sites/NewNerve/sputils-test` (List schema live checks) are to be
  deleted by hand. `SitePages/zz-markdown-export-test.aspx` is a keeper.
- **CSS bleed, both directions**: `app.css` still styles `html`/`body`
  (darkens the host page behind the pad — currently invisible and arguably
  nice; the gap around the seated app shows it). SP styles also bleed into the
  pad; nothing visibly broken, but scope properly if oddities appear.

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
  `sharePointFiles.additionalTypes` in `dcspad.config.json` maps extra
  extensions onto an editor pane for SharePoint transfer only (the supplied
  config maps `.json` to the JS pane; built-ins can't be overridden and local
  disk import stays HTML/CSS/JS).
  Libraries that set **ForceCheckout** are probed as part of the destination
  inspection: overwriting one is gated behind a consent checkbox at the foot
  of the metadata dialog, and consenting makes the pad `CheckOut()` the file,
  upload, write metadata, then `CheckIn()` — the check-in reads the file's
  state first and only posts when SharePoint still reports it checked out, so
  it is correct whether or not the tenant's overwrite ends the check-out.
  A file held by another user is stated and refused (in any library) rather
  than offered the box; one you already hold is overwritten without a second
  check-out and checked in; a new file born checked out is checked in. Every
  stage is resumable: an upload failure after the pad's own check-out offers
  **Discard check-out** (`UndoCheckOut`, safe because nothing was uploaded), a
  failed check-in offers **Retry check-in** / **Leave checked out**, and a
  write SharePoint refuses with a check-out error the probe missed
  (`checkout-required` in sp-odata.js, classified on the unlocalized OData
  error code as well as the message) reveals the same consent for the retry —
  unless the refusal names another holder, which blocks instead.
  Two SharePoint facts this rests on: `ValidateUpdateListItem` with
  `bNewDocumentUpdate` checks a checked-out file in by itself (so the comment
  rides that write as `checkInComment`, and the explicit `CheckIn()` is usually
  a no-op); and page context carries the login as a bare UPN while
  `CheckedOutByUser.LoginName` is the full claim, so "checked out to me"
  accepts either form (`isCheckedOutByCurrentUser`).
  This lifecycle first shipped to the work prod tenant from an uncommitted
  tree (bundle stamp `47edb5d4-dirty`); it was merged here with the gate.
  See `plans/file-sp-import-export.md`, which is now an implementation record.
- Framework rows use drag-and-drop ordering without up/down controls; snippets
  are always displayed alphabetically regardless of file type.

## Feature backlog

- **Framework/snippet row action icons only on hover** (currently always
  visible at 60% opacity).
- **Site Inspector / SP diagnostics tools** — v1 SHIPPED as the SP Workbench
  (below).

## SP Workbench Tier 2 (2026-07-30)

Six features shipped on the workbench entry point (plan record:
`~/.claude/plans/i-want-to-continue-sleepy-charm.md`; everything lives in
`src/workbench/` — zero changes to pad-loaded files, which were under
concurrent debugging on main):

- **Config links** view — curated `_layouts` jumps grouped by category
  (`config-links.js` data + `views/links.js`), plus per-row ⚙ and drilldown
  "List settings" links in the Lists view. Uncertain classic URLs carry
  hints; live click-through pass still pending.
- **Site favorites + recents** — topbar star + "Sites ▾" picker; persists
  through `state.js` `loadDoc`/`saveDoc` (key `dcspad.v2.wbsites`,
  invariant 6 honored; workbench imports state.js safely — verified its
  import side effects are read-only).
- **Query builder** (`views/query.js`) — list/endpoint target, $select
  checkboxes (User/Lookup auto-expand to `/Title`), typed $filter rows,
  orderby/top/expand, raw-edit mode with descriptor round-trip parsing
  (Copy-as menu omitted when a raw string can't round-trip), per-web
  sessionStorage persistence.
- **Page inspector** (`canvas.js` + `views/pages.js`) — SitePages master
  grid (PromotedState badges), drilldown tabs: Structure (section/column
  tree, widths /12, emphasis, vertical, collapsible), Web parts (names via
  WEBPART_NAMES, unknown ids degrade to raw GUIDs), Text (sanitized render
  + raw HTML), Metadata, Raw. Parser handles BOTH storage formats (JSON
  array + legacy HTML `data-sp-controldata`) and never throws — malformed
  entries land in errors[] + an "unplaced" bucket. CanvasContent1 must be
  explicitly $select-ed (done per-item).
- **Pages metadata editor** + **Files browser** share `field-editor.js`
  (per-TypeAsString editors; FieldValue conventions unit-tested: MultiChoice
  `;#A;#B;#`, Boolean `1/0`, DateTime ISO, URL `url, desc`;
  User/Lookup/Taxonomy read-only v1) and `sp-write.js` (digest via the
  exported sp-files.js cache, 403→force-refresh retry, ValidateUpdateListItem
  with per-field error mapping, binary AddUsingPath ≤50 MB). Editors refuse
  `CanvasContent1`/`LayoutWebpartsContent` (NO_EDIT_INTERNAL).
- **Files browser** (`views/browser.js`) — every file type, paged listing,
  breadcrumbs + library picker, download via `download.aspx`, upload with
  overwrite consent (pre-flight and 409-race), post-upload metadata panel
  with keep-without-metadata / retry-that-never-reuploads. A library with
  **ForceCheckout** puts a consent checkbox on that overwrite bar,
  `CheckOut()`s the file before the upload and `CheckIn()`s it after (also for
  a new file born checked out) — the same contract as the pad's
  export, sharing `isCheckedOut` / `isCheckedOutByCurrentUser` from
  sp-files.js so the two cannot drift.

**Tier 2 refinement pass (same day, Joe's feedback on the first cut):**

- **Left nav reorganized**: Site · Permissions | Lists · Pages · Files |
  Query | Panels · Advanced. Renames: Security→Permissions, Config
  links→Panels, old Site sheets→Advanced. **New Site landing view**
  (`views/site-home.js`): web + current-user cards, role chip, subwebs with
  one-click Inspect (wired through a late-bound `inspectSite` dep).
- **Status bar** now shows the current user plus a lit "Site admin" chip
  (subtle "Site user" otherwise); refreshed on site switch.
- **Panels curated** to Joe's crossed-out screenshot — 19 links in 5 groups
  (General, Permissions & people, Recycle bins, Galleries, Search);
  label-first UI, paths in tooltips. The set is deliberate; don't grow it.
- **Permissions** gained a Members tab: flattened user-by-group roster
  (group name/id, user, email, login, admin — exportable), add-user (email
  → claims login) via `POST sitegroups(id)/users`, two-step remove via
  `removebyid`, plus panel jump links in the head.
- **Pages**: tab order Extract·Metadata·Structure·Web parts·Raw; header
  shows the server-relative URL (click copies the FULL absolute URL; id
  moved to Metadata); **Export content** (single human-readable .md —
  title/description/created/location up top, merged control content,
  standardized metadata block at bottom) and **Export raw** (.json) sit on
  the header row (`page-export.js`, pure + unit-tested). Master grid shows
  a sortable Folder column (subfolder pages) and an "Open Site Pages
  library" link.
- **Files**: per-row copy-direct-URL button beside download.

Mock mode covers all of it (`mock-data.js` fixtures incl. canvas JSON +
legacy HTML pages, subfolder pages, and a `/Shared Documents` tree; mock
writes recorded on `__DCSPAD_WB_WRITES__`). Tests: `workbench.mjs` grew
39→69, new `workbench-edit.mjs` (15), `workbench-hosted.mjs` still 10 —
all pass.
Known pre-existing failures on this branch (NOT Tier 2): smoke's "legacy
framework catalog adopts explicit preset order" and files' "explicit host
adapter enables SharePoint file actions" — both pad-side, consistent with
the in-flight "metadata save - still buggy" work on main.

**Live-tenant checklist still owed** (after deploy): config-link
click-through on a web + subweb, DateTime/Number ValidateUpdateListItem
formats on the tenant locale, legacy-format page parse, >200-file paging,
binary up/download + copy-URL, `bNewDocumentUpdate` version behavior,
SitePages Templates-folder noise, uncertain WEBPART_NAMES entries,
group add/remove (the nometadata JSON body for `sitegroups(id)/users`
needs one live confirmation), and the role chip on a non-admin account.

## SP Workbench: List schema (stage 1a) (2026-09-21)

The Workbench can now capture, export, import, create and copy a **list's
schema** (settings, fields, views, content types — no items yet). Plan
record: `~/.claude/plans/plan-to-add-a-idempotent-rossum.md`. Commits
3a7d6e3 (capture + Schema tab + scripts), 2d96bc3 (bundle rebuild),
32bce23 (apply executor), aa9418c (Copy to… / New from schema… dialog).

**What shipped:** a **Schema** tab in the Lists drill-down (after Content
types) showing settings/fields/views/content types plus an Export menu
(download/copy JSON, copy as PnP.PowerShell or PnPjs 2 provisioning script);
a **New from schema…** button on the Lists grid toolbar that imports a
schema document and opens the **Copy to… / New from schema** dialog — its
own target-web connection, title/lookup/content-type mapping, a Dry run
(reads only) preview of the ordered step plan, Create with per-step
progress, Retry failed steps, and a downloadable markdown report. New
modules: `list-schema.js` (pure core), `list-schema-capture.js`,
`list-schema-apply.js`, `list-schema-script.js`, `list-schema-dialog.js`
(all under `src/workbench/`). New suite `tests/workbench-schema.mjs` (78
checks); `workbench.mjs` stays at 133 (only its tab-order pin gained
Schema). Total across all suites: 513.

**Doc interchange contract:** `kind: "dcspad-sputils-list-schema"`. Every
v1 key is kept verbatim with its v1 type; new keys are additive; the
Workbench writes `version: 2` and reads both `version: 1` and `version: 2`.
SPUtils's own reader checks only `kind`, so a SPUtils export imports in the
Workbench and a Workbench export imports in SPUtils unchanged — this is
what makes the two tools interoperate without either one depending on the
other's code.

**SPUtils gaps this port closes** (recorded as SPUtils follow-ups; SPUtils
itself was not touched):
- The Workbench always sends `createfieldasxml` with an `Options` bitmask
  including `AddFieldInternalNameHint` (8). SPUtils calls the string
  overload of `createFieldAsXml` with no `Options`, so SharePoint may derive
  the internal name from `DisplayName` instead of the requested one (e.g.
  `Project_x0020_Status`). Needs a live check on SPUtils; back-port if
  confirmed.
- `scrubSchemaXml` strips `List=` only from Lookup fields — User fields keep
  `List="UserInfo"`, which SPUtils's blanket strip removes. Calculated
  fields' `<FieldRefs><FieldRef ID=…>` source GUIDs are stripped too
  (SharePoint resolves them by `Name`); SPUtils leaves the source GUIDs in.

**Second-client pattern:** the dialog opens its own REST connection to the
target web (`createClient()` + `connectWeb()` + its own `sp-write` client)
so a Copy to… run never disturbs the shell's client or the source list
view — the one new shell dependency is `createClient` itself
(`src/workbench/main.js`), everything else (`shell.js`, the source view)
is unchanged.

**Hooks left for later stages:** the dialog's Items fieldset (`Include
items`, `Include attachments`, `Preserve authorship`) is present but
`disabled`, titled "Item data arrives in stage 1b — this run copies the
schema only."; `Copy to…` is disabled on document libraries (`baseTemplate
!== 100`) with a stage-2 title. Stage 1b (item data, same engine) and stage
2 (document libraries, schema only) are reserved seams — see Roadmap below.

**Live-tenant checklist — run 2026-09-22 on the dev tenant, build #171.**
Source `zz-schema-requests` on `/sites/NewNerve` (12 custom columns: Choice,
FullHtml Note, Number, DateOnly, Boolean, URL, User, lookup, dependent
lookup, self-lookup, Calculated, Indexed+unique Text; validation formula;
versioning 50; an extra view). Copies: `zz-schema-requests`, `… Copy`,
`… Copy 2` on the subweb `/sites/NewNerve/sputils-test` and `… Copy` on
`/sites/NewNerve` — leftover test lists, safe to delete (with the
`zz-schema-clients` lookup targets on both webs). The first two subweb
copies came from pre-fix builds (one carries a duplicate "All Items").
- ✓ nometadata `POST web/lists`; `createfieldasxml {parameters:{SchemaXml,
  Options:8}}` keeps every internal name (verified via PnP).
- ✓ Lookup rebound to the target's list, self-lookup to the new list,
  dependent lookup carries its new primary `FieldRef`; User keeps
  `List="UserInfo"`; FullHtml Note, Calculated, and the Indexed+unique
  MERGEs land.
- ✓ `views?$expand=ViewFields`; `RowLimit`/`Paged` MERGE; the new list's
  own "All Items" is updated, not duplicated; validation applied last.
- ✓ Cross-web copy into a subweb; same-web copy (offers Open the new list).
- Found live and fixed: SPO omits `ValidationFormula`, `ValidationMessage`,
  `OnQuickLaunch`, `ReadSecurity`, `WriteSecurity` from the default SP.List
  payload (capture names them); SP.List has no `Ordered` REST property.
- Still open: `Options 8|4` on a content-types-enabled list and
  `addAvailableContentType` (no CT-enabled fixture yet); `ViewTypeKind` in a
  views POST (only HTML views tested); the 429 path under real throttling;
  `EnableRequestSignOff`/`ListExperienceOptions`/`DisableGridEditing`
  availability; whether SPUtils' string-overload `createFieldAsXml`
  regenerates internal names; a SPUtils v1 doc imported in the Workbench and
  the reverse; a reconcile (Add to existing list) run.

### Invariants to keep (from the build thread's state file)

- `buildApplyPlan`'s `dependsOn`: views and the validation-formula step
  depend on the `list` step only, never on every field — one column that
  can't be recreated (managed metadata, say) must not block every view or
  the validation formula.
- Plan-time refusals the probe makes (type clash, taken title, missing
  content type) carry `final: true` and are never re-run by "Retry failed
  steps"; only steps that failed during execution are.
- `sp-write` `post` merges caller headers OVER its base set (digest, Accept,
  content-type): a step needing `X-HTTP-Method: MERGE` or `IF-MATCH: *` must
  pass them itself — the base set never supplies them.
- The mock writer must return the ids the executor binds to (`web/lists` →
  `{Id, Title, RootFolder}`, `createfieldasxml` → `{Id, InternalName}`, a
  view add → `{Id}`): the post-create probe resolves through the same
  registry, so a write it doesn't register breaks the next read in the run.

## SP Workbench: List schema stages 1b + 2 (2026-09-22)

**Stage 1b — items.** `list-data.js` (pure: the SPUtils data document,
kind `dcspad-sputils-list-data`, v1 keys verbatim, v2 written, v1 read;
per-type form-value strings; lookup re-resolution by shown value; folders
by depth; the three-pass partition), `list-data-capture.js` (whole-list
read with `$select=*`, folders, referenced users, attachments embedded as
v1 `{name, base64}` up to 10 MB/file and 50 MB total, url-only beyond),
`list-data-apply.js` (library targets refused; live date-format
calibration — a web whose format cannot be learned gets its date fields
dropped with a warning, never guessed; per-value web-local offset;
ensureuser; folders, then items written sequentially in source-id order so
a fresh list keeps the source ids, one retry without a rejected field,
self-lookups in pass 2, authorship in pass 3). The Items tab exports the
whole list as `data-<list>.json`; the dialog's Items fieldset is live
(copy mode captures on demand; New from schema… takes a schema and a
matching data document). Items import only once the schema has no failed
column, exactly once. Reviewed by Codex (xo turns 7–8).

**Stage 2 — document libraries (schema only).** Any base-type-1 source
(Picture/Wiki libraries included, with a warning) is recreated as a
standard library (BaseTemplate 101, no `/Lists/` segment); settings drop
`EnableAttachments` and carry `ForceCheckout` and minor versioning;
Title is never forced required; `_ExtendedDescription`/Title tweaks apply
as base-field steps; 0x0101-derived content types attach; library view
columns survive; Forms templates (list or content type) are warned as a
manual step. Files are not copied. Reviewed by Codex (xo turns 9–10).

**Live on the dev tenant (builds #176, #179):** items copy with users,
cross-list lookups (and their dependent lookup), self-lookups, URL
descriptions, DateOnly local dates across a DST change, and authorship;
source ids kept. A library copies with versioning 20/5, required check-out
and its custom column. Fixed from live runs: a braced lookup-list GUID
broke the target lookup index; a column SharePoint provisions with a new
list (`_ExtendedDescription`) was duplicated as `_ExtendedDescription0` —
the executor now checks the post-create field set. Test lists
`zz-schema-*` on `/sites/NewNerve` and `/sites/NewNerve/sputils-test` are
leftovers to delete by hand.

**Live checks completed 2026-09-22 (build #185):**
- Folders and attachments: a cross-web copy with items recreated the `Archive`
  folder, placed its item inside it, and carried the attachment (`zz-note.txt`)
  onto the right item — 4 items, 1 folder, 1 attachment, 0 failures.
- Reconcile (Add to existing list): the consent sentence gates it, the button
  relabels, and the run added only the new column — 1 added, 12 already
  present, 2 views rebuilt, validation re-applied.
- Content types: a content-type-enabled list copied with its site content type
  attached (`addAvailableContentType` before the fields) and its column created
  with the internal name intact; content-type membership matches the source.
- Item import into an existing list (Tools tab): export from one list, import
  into another — ids in source order, lookups, self-lookup, user and dates
  correct, 0 failures.
- SPUtils ⇄ Workbench document round-trip, both directions: `SPUtils.getListSchema`
  → New from schema… → created (13/13 columns, internal names intact), and a
  Workbench v2 export → `SPUtils.createListFromSchema` → created (13 columns, no
  failures). Note for AUTOMATION only: a harness that injects the pnp2 bundle
  late (evaluating its text, or a createElement('script')) must hide
  `define.amd` across that load or `window.pnp2` is never set. A literal
  `<script src>` in Script Editor markup — the normal way SPUtils is used on a
  page — is early enough and needs nothing (verified on SPutils.aspx).

**Still open:** the 429 path (cannot be forced on demand — left unproven).

**New SPUtils finding from the round-trip (for the SPUtils owner):** a dependent
lookup created by `createListFromSchema` lands as `Client_x0020_code` — it goes
through `addDependentLookupField(displayName, …)`, which derives the internal
name from the display name. The `Options` fix does not cover this path (it is
not a `createFieldAsXml` call). The Workbench keeps `ReqClientCode` by creating
dependent lookups from scrubbed XML with `Options: 8`.

## SP Workbench: list Tools tab (2026-09-22)

New **Tools** tab in the list drilldown (after Items, before Raw), built
from a small data-driven registry (`list-tools.js`: `LIST_TOOLS`, one entry
per card, composing the Panels view's `.wb-linkgroup` card look) so a future
tool is one entry, not a new tab layout. Four cards: **Copy this list…**
(the same BaseType-aware gate the old Schema-head "Copy to…" carried, now
here), **Export schema** (the same four document actions the old Schema-head
"Export ▾" carried), **Export data** (the whole-list data-document export
that moved off the Items tab's own "Export ▾" menu), and **Import data into
this list** (new — reads a data `.json` this Workbench or SPUtils exported,
refuses a non-JSON/oversized/wrong-kind file inline, shows a
`writableFields()`-driven written-vs-dropped column table against a fresh
read of this list's own fields, a consent-gated run through
`applyListData()`, and a report panel; disabled with a stated reason on a
document library, since item import into a library isn't supported). The
Schema tab is now **read-only** — its head lost the Export/Copy controls and
gained a one-line "Export and copy this list from the Tools tab" hint;
chips, sections and warnings are unchanged. New module
`list-data-import-dialog.js` is the Import dialog itself: same recipe as
`list-schema-dialog.js` (two-phase exit, a report panel) but smaller — the
target is fixed to the open list, so there is no target-site connect or dry
run. `tests/workbench-schema.mjs` grew from 116 to 129 checks (moved/renamed
the Schema-head Export/Copy checks and the Items-tab whole-list-export check
onto the Tools tab, added the Import dialog's refusal/consent/mock-write/
401 checks); the suite total is 564. Not built or deployed — `dcspad.app.js`
and `dcspad.workbench.js` are unchanged by this work.


**Live on the dev tenant (build #185):** the Tools tab's Export data wrote a
v2 data document (3 items, 36 fields) for `zz-schema-requests`, and Import data
into this list added those items to the separate, schema-only
`zz-schema-requests Copy` on `/sites/NewNerve` — ids in source order, client
lookup, self-lookup rebound to the new id, user column and both DateOnly values
correct, 0 failures. The column preview listed each written column and named
the read-only ones.

## SP Workbench: markdown page export (2026-09-22)

Four pieces, shipped as one rebased set:

- **Shallow-clone build guard.** `git rev-list --count HEAD` counts only what
  a clone actually has, so a shallow checkout (cloud sessions, most CI)
  stamped a build number far too low and — silently — going *backwards*
  between releases (the bundle at `7b6ec4e` carried Build #94 against a true
  count of 152). `tools/build-app.mjs` and `tools/build-workbench.mjs` now
  refuse to stamp a build number from a shallow clone; `DCSPAD_BUILD_NUMBER`
  still wins, so an explicit override (CI, a reproducible rebuild) is
  unaffected.
- **`src/html-markdown.js`** — the one HTML→Markdown converter, replacing
  item-export.js's hand-rolled version. Backed by vendored Turndown
  (`vendor/turndown/turndown.js`, `LICENSE`, `version.json`, produced by
  `tools/build-vendor-turndown.mjs`, `npm run build:vendor` in `tools/`).
  Vendored rather than bare-imported for the same reason as `vendor/monaco/`:
  `src/` loads both bundled (esbuild) and unbundled (standalone `index.html`,
  every test suite), and only a real relative import resolves in all three —
  a bare `'turndown'` specifier would 404 outside the bundle. The build step
  is a copy with provenance (a sha256 in `version.json`) plus a bare-import
  guard and a default-export check, so a future Turndown release that adds
  its own imports fails the build instead of silently breaking the
  unbundled paths. Two named profiles so the two call sites cannot drift:
  `listField` (used by `src/workbench/item-export.js` — headings flatten to
  bold, tables join with `' | '`, because that document already owns `##`
  for its per-item structure) and `pageContent` (used by
  `src/workbench/page-export.js` — real ATX headings demoted to sit under
  the web part's own `## <label>` heading, real markdown tables). An unknown
  profile throws rather than converting under the wrong conventions.
  Security posture is unchanged and slightly stronger: `sanitizeHtml` still
  runs first, and Turndown is additionally configured to keep no raw HTML
  and to drop `script`/`style`/`iframe`/`object`/`embed`/`form` outright, so
  no tag reaches a permissive markdown renderer even if sanitization were
  ever bypassed upstream. A part markdown genuinely cannot carry (a bare
  video embed) falls back to its sanitized HTML rather than exporting an
  empty section. It bundles into `dcspad.workbench.js` and rides the
  existing `?v=` busting — no new `VERSIONED` entry in `boot.js` — and the
  pad bundle tree-shakes it out entirely.
- **Content format choice.** `buildContentExport(…, format)` in
  `page-export.js` takes `'markdown'` (default; sanitized then converted) or
  `'html'` (the sanitized source markup, exactly what this export emitted
  before conversion existed — the escape hatch for a page the conversion got
  wrong); an unrecognized format throws. Only the content blocks differ —
  the metadata frame is byte-identical between the two — so `.md` files in
  either format diff cleanly against each other. Both are `.md`, so
  `contentFileName()` keeps the stems apart: `<stem>-content.md` and
  `<stem>-content-html.md`, and the bulk zips likewise
  `sp-pages-content.zip` / `sp-pages-content-html.zip`. `bundleEntryName()`
  reuses `contentFileName()`, so a bundled page stays byte-identical to the
  same page exported alone.
- **Row-level export.** Every Pages grid row now carries always-visible
  MD / HTML buttons (`src/workbench/views/pages.js`; the drilldown's pair
  is labelled Export MD / Export HTML) that
  export a single page without opening it first — deliberately not
  hover-revealed like the Files browser's row actions, because these name a
  choice between two formats and a control you must hover to discover is one
  most people never find. They fetch the full item the same way the
  drilldown and the zip do; a page the account cannot read goes through
  `denied.js` in the neutral register like every other refused read. The
  detail pane's single "Export content" button became the same MD/HTML
  pair, and the grid's Export ▾ menu offers both zips ("Download content
  .zip (Markdown)" / "(original HTML)").
- **Action columns.** `src/workbench/grid.js` columns may now be marked
  `action` — a column that exists only to host per-row controls, like the
  new Pages row buttons. `src/workbench/export.js` drops `action` columns
  from CSV/JSON/markdown exports; before this an action column contributed
  an empty CSV column and a JSON key whose value was whatever the render
  happened to key off.

**Items vs. Pages bulk shape, written down on purpose.** A separate commit
(`09b4236`) records why the Items tab's selection-scoped export and the
Pages bulk zip differ in shape: Items exports selected rows as one merged
`.md` (a list's rows are one dataset); Pages exports a zip of one `.md` per
page (a page is its own document). Noted in `CLAUDE.md`'s `item-export.js`
file-map entry because it very nearly got "fixed" into a false
inconsistency.

**The rebase.** This work started on `origin/claude/workbench-markdown-extraction-tpuu10`
(2026-08-26) and was rebased onto `main` today, 2026-09-22, over the List
Schema stages and the list Tools tab. Seven original commits collapsed to
four source commits here, bundles regenerated from the rebased tree. The
code merged clean; the only conflicts were the test-count paragraphs in
`CLAUDE.md` and `tests/README.md`, resolved in favour of `main` and
re-pinned in this docs pass.

**Tests.** `workbench.mjs` grew from 133 (main) to **150** — 141 from the rebased branch (eight new
checks: both html-markdown profiles' structural conventions, no markup
surviving conversion, page text parts exporting as markdown, the HTML
fallback for a part markdown can't carry, the html format's structural
conventions and shared metadata frame, the naming split between the two
`.md` stems, the row buttons exporting both formats without opening the
page, the zip honouring the format, and action columns staying out of
CSV/JSON/markdown) plus nine from the two Codex review rounds below. Six existing assertions were re-pinned to the new
(correct) output, three per rebased commit: from the converter commit —
`<br>` now emits a markdown hard break; a hostile `##` inside a list item's
field value arrives escaped (`> \## forged item`) because Turndown
neutralizes it at the source, on top of the existing blockquote; and the
classic-page export carries converted markdown instead of embedded HTML;
from the format/row-button commit — the detail pane's button pair, and two
zip-menu locators that now match both entries. Every other suite is
unchanged. **All 590 checks pass across the suites** after the rebase onto
main (per-suite counts: `tests/README.md`).

**Codex rounds (xo turns 14–15).** The review found no criticals and six
real gaps, all closed in `0500a85`: a text part holding only a `<style>`
block came back through the markdown fallback because `sanitizeHtml`
never removed `style` (it now drops `style`/`noscript`, the converter's
`DROPPED` list is exported and a check holds the sanitizer to covering
it, and a style-only part is judged empty after sanitizing); a bare
`<pre>` lost its fence (a `barePre` rule in both profiles, fence longer
than any backtick run inside); link/image targets with spaces were
invalid CommonMark destinations (control characters stripped before the
scheme check, whitespace/`()<>` percent-encoded, backslash escaped rather
than `%5C`); a row export could finish against a library switched away
from (the zip path's identity guard now runs after every await and in
the catch); an `action` column's header was a phantom "Sort by" control
(inert now; the filter skips it); and the vendor guard missed
side-effect/dynamic imports. The re-review accepted those and raised
three more, closed in `8258110`: a stale bulk export still wrote to the
shared status line (it now returns without touching it, progress ticks
included), the trailing blank-line squeeze ate blank lines inside fenced
code (`squeezeBlankLines()` walks fences per CommonMark and leaves their
content alone), and the guard now strips comments — walking over
strings, template and regex literals — before matching (a bare specifier
inside a string still trips it: an accepted, loud false failure). Its
"stale hosted bundle" high was the rebuild that followed (`aaaf26b`,
Build #197). One Codex ask was declined: a JS module lexer for the guard
would add a dependency for a copy step; the comment stripper covers the
forms named without one.

### Live-tenant checklist

Run 2026-09-22 on the dev tenant, Build #197, headless Playwright against
`SitePages/zz-markdown-export-test.aspx` (a page created for this with
`Add-PnPPageTextPart`: h2/h3, ordered and unordered lists, a table with a
`<thead>`, a link, bold/italic, a `<br>`). Evidence files stayed in the
session scratchpad; the facts:

- [x] Extract content exports markdown with real tables and no raw HTML:
      `## Text`, h2 → `###`, h3 → `####`, `1.`/`2.`/`3.`, `-` bullets, a
      five-line GFM table, `[link](https://example.com/handbook)`,
      `**markdown**`, `*workstation*`, the `<br>` as a two-space hard
      break; `<[a-z][^>]*>` finds nothing in the content section. The row
      button and the drilldown button produce byte-identical files
      (1621 bytes).
- [x] The HTML format emits the sanitized HTML (h2/p/strong/a/h3/ol/li/em/
      ul/table/thead/tr/th/tbody/td/br only; no script or style), and the
      two files' metadata frames — header up to the first `---`, footer
      from the `## Metadata` block — are byte-identical (408 + 726 bytes).
- [x] Row MD / HTML buttons download without opening the page: the
      sessionStorage route stays `{"view":"pages"}`, no detail pane mounts;
      clicking the row itself does open the drilldown on Extract.
      A denied page cannot be produced live — the account is a site
      collection admin — so the neutral 403 register rests on the suite's
      stubbed checks, not on a live read.
- [x] `sp-pages-content.zip` (store-only) and `sp-pages-content-html.zip`
      each carry the test page's entry byte-equal to the solo export of
      the same format (Buffer.equals), plus the second selected page.
- Console: no errors during the Pages view, exports or zips; the two
  message-less `pageerror`s fire during SharePoint's own page load and
  fire on a plain page too.
- Observed, not a defect of the export: the page's `Description` is
  SharePoint's auto-generated summary (words run together, truncated) and
  the export quotes it verbatim.

**Still open:** the other reserved seams (chunked file transfer for a
library copy, etc. — see Roadmap below) are unchanged by this work. The
`zz-markdown-export-test.aspx` page on the dev web is a keeper for future
export checks (recorded in memory), not a leftover to delete.

## SP Workbench: page status + EEEU audit (2026-09-23)

Two features, merged to `main` 2026-09-23 (rollback point: branch
`rollback/main-before-page-status-eeeu`, Build #197). Status
(what is done, what is next) lives in `state/2026-09-23-page-status-eeeu.md`;
this section is the design and the live-tenant checklist.

**Page status is a rebuild.** It was built once on the work (bmo) machine and
lost; only Joe's original prompt and that session's plan survived. Decisions
below are Joe's answers of 2026-09-23.

### Page status (Pages view)

Scope: modern Site Pages (119) and the classic publishing Pages library (850)
— `supportsStatus()` in `views/pages.js`. A library merely titled "Pages"
keeps the plain drilldown.

- **Published = the page has EVER had a live major version**, even with a
  newer draft on top (Joe). Read from `File/MajorVersion` (fallback: the
  `OData__UIVersionString` label) — never from `_ModerationStatus` alone,
  which reads 0 ("Approved") on every item of a library without content
  approval, drafts included. With approval/scheduling, a 1.0 that is
  Rejected (1), Pending (2) or Scheduled (4) has never gone live. Two
  states only: Published / Unpublished. Rules and tests:
  `src/workbench/page-status.js` (`derivePageStatus`, `lastPublishedFrom`).
- **Scan Page Status** (grid toolbar): one library-wide query, same
  `$orderby=FileLeafRef`/`$top=5000` as the grid so the two capped sets line
  up. Adds Published, Checked out (holder's name) and Inheritance ("Broken"
  only) columns via `grid.setColumns()`; they ride into CSV/JSON exports.
  Its own four-rung query ladder (`pageStatusShapes`): a refused status
  field costs the columns/chips, never the page; a degraded rung shows the
  quiet reduced chip with SharePoint's sentence. Rows the read did not cover
  stay blank (never "Unpublished") with a line saying so.
- **Drilldown:** status chips after the kind chip in the loud register
  (`.wb-role-chip` + `.wb-status-on/-off`, `design/INFO-CHIP.md`) —
  Published (accent) / Unpublished (neutral), Inheritance broken and Checked
  out (accent) only when true. The Export MD / HTML / raw and Open page
  buttons moved to the tab row, right-aligned, outside `role=tablist`. A
  failed status read is stated in the chip row through `denied.js`.
- **Permissions tab** (between Metadata and Structure): Inherited / Broken
  inheritance chip + `items(id)/roleassignments` — the item endpoint answers
  with the inherited set when inheritance is intact, so it is right even
  under a library that breaks inheritance itself.
- **Metadata tab:** only Joe's rows, in his order (`METADATA_SPEC`); fields
  the library lacks are skipped, unlisted fields hidden. Matched by internal
  name, then display name: Item Type is `FolderType` (a column whose use
  changed), Content Category is `bmocContentCategory` (managed metadata,
  multi), Contact is a person field, Org a single choice. Editable: Title,
  Item Type, Pillar, Org. Read-only: Name (a rename breaks every link — was
  already read-only), managed metadata and person fields (no editor yet —
  Roadmap), system fields. `WikiField`/`PublishingPageContent` are now never
  editable anywhere (page bodies, like `CanvasContent1`). Publish Date is
  **last** published: the newest live `x.0` in `items(id)/versions`; on a
  moderated library, a version whose moderation the read could not return
  makes the date unknown rather than assumed approved. Scheduled start date
  (`_PublishStartDate`) was dropped (Joe).
- **Content `.md` export** gains Publish status, Last published,
  Permissions, and Checked out to lines; unknown values are left out.

### EEEU audit (Permissions → EEEU audit)

Ported loosely from Joe's standalone `DCS.SecurityGroups` content audit (bmo
FCUPortal); its group-membership viewer was out of scope — the Groups and
Members tabs already cover it. Engine: `src/workbench/eeeu-audit.js` (no DOM,
client injected); UI: `views/security.js`.

- **What counts:** EEEU and Everyone by claim
  (`c:0-.f|rolemanager|spo-grid-all-users/…`, `c:0(.s|true`), the editable
  name list (defaults include BMO's "SharePoint EEEU Visitors" convention),
  every site group whose **members** include EEEU/Everyone (found at scan
  start, once per site collection), and "People in your organization"
  sharing-link groups (`SharingLinks.*.OrganizationView|Edit.*`).
- **Scope:** four checkboxes (Site, Pages libraries, Lists, Document
  libraries), "Include items and folders", and the site's direct subsites as
  checkboxes — a ticked one is scanned with its whole tree. Hidden/system
  lists skipped (`isInternalList`).
- **Noise cut:** Limited Access-only grants dropped; only securables with
  their own permissions are read — an inheriting list or subsite is covered
  by its parent's row (the run summary says so).
- **Robustness:** every REST client now shares one three-request queue
  (`sp-rest.js`), 429/503 retried once; Cancel is cooperative and keeps
  partial results; switching site cancels a running audit (view
  `destroy()`); anything unreadable — including a list over the client's
  100,000-item cap — is a Problems row, never a silently clean result.

### Review round 1 (ChatGPT, 2026-09-23)

Nine findings (five major, four minor), all verified and fixed in
`4cfb92e`: the shared request queue with slot hand-off (the old
release-then-wake could briefly allow four), the 100,000-item cap reported,
`#` in audit links, cancel checks after every await, Scheduled (4) not live,
version moderation read from every payload shape, the scan aligned with the
grid's order/cap and its ladder rung kept, and status-read failures surfaced
instead of swallowed. Its REST claims came from Microsoft docs, not a
tenant — the checklist below settles them.

### Live-tenant checklist

Run 2026-09-23 on the dev tenant with headless Playwright (a clean
browser context carrying only the sp-env profile's auth cookies, so no
workspace was touched). First against `main` at Build #215 (ab5a0f9 —
served bundles confirmed byte-identical to the mirror, stamp #215): 19/23
Pages checks passed and six defects surfaced. They were fixed on
`claude/page-status-live-fixes` (PR below) and every item re-run green
on the branch's **Build #220** (e282515): Pages 29/29, EEEU 10/10.
Evidence stayed in the session scratchpad; the facts:

- [x] **Scan Page Status** on Site Pages: the three columns fill, no
      reduced chip. Draft-over-published pages (CPtest 1.2, DCSpad 1.3,
      SPutils 1.1, pnp2-object-test 3.1, SPWorkbench 4.2) read Published;
      the never-published 0.2 pages (`sp-pilot-listview`,
      `_harness-sp-pilot-listview`) read Unpublished; a checked-out page
      (`zz-markdown-export-test.aspx`, checked out for the run, then
      UndoCheckOut) shows "Joe Zapert (NERVE.digital)"; Inheritance reads
      Broken on SPWorkbench.aspx only — the one page REST reports unique.
- [x] **Drilldown chips** match the grid (CPtest Published; sp-pilot
      Unpublished (off); the checked-out page Published + Checked out;
      SPWorkbench Published + Inheritance broken). The tab row holds
      Extract·Metadata·Permissions·Structure·Web parts·Raw with Export MD /
      Export HTML / Export raw / Open page ↗ right-aligned outside the
      tablist; the exports download. **Open page ↗ opens in the same tab,
      not a new one** — the pre-existing issue #10 (root cause under Open
      items above), not introduced by this work.
- [x] **Permissions tab**: SPWorkbench.aspx lists its own 4 assignments
      (REST `items(7)/roleassignments` = 4; the library has 5); CPtest
      (inheriting) lists exactly the library's 5.
- [x] **Metadata tab** (after the fixes): rows in spec order; Publish Date
      is the published version under the draft — CPtest `2026-07-28` (1.0,
      not the 1.2 draft), SPWorkbench `2026-08-17` (4.0, not 4.2); Modified
      By / Created By read names. The bmo columns are absent on dev, and
      the "Item Type" row is now absent too (see fix 3).
- [x] **Versions payload shape**, on a content-approval pages library built
      for this (`zz-mod-pages`, 119, approval on; `zz-mod-page.aspx` at
      approved 1.0 with 1.1 submitted/Pending on top): versions return
      moderation only as `OData__x005f_ModerationStatus` (fix 1). After
      the fix the scan reads Published and Publish Date is the approved
      1.0 — the deployed parser returns exactly `2026-09-23T19:54:10` (1.0),
      not 1.1's `19:55:18`.
- [x] **Content `.md` export**: the checked-out page carries all four lines
      (Publish status: Published / Last published: 2026-09-23 / Permissions:
      Inherited / Checked out to: Joe Zapert (NERVE.digital)); CPtest three
      (not checked out); SPWorkbench says Broken inheritance. "Last
      published" was missing everywhere before fix 1.
- [x] **EEEU audit**, with temporary fixtures Joe approved (all removed
      after the run — lists recycled, group deleted, and the Limited Access
      entry SharePoint had added for EEEU at web level removed; the web is
      back to no EEEU assignment): the direct EEEU Read on library
      `zz-eeeu-test` found, Why "Everyone except external users"; a group
      named **SharePoint EEEU Visitors** containing EEEU shows as "Group
      containing Everyone except external users" (membership, not the name
      list); an OrganizationView link on `zz-org-link.txt` found as
      "Org-wide sharing link" (items on); ticking `sputils-test` finds
      `zz-eeeu-sub` there, unticked it is not scanned; the full item scan
      (Documents ≈ 34.6k items) completes in ~4.5 min with 0 problems
      (fixes 4, 6); Cancel mid-run keeps partial results and settles 22 s
      after starting (fix 5). **Not produced:** a locked group for the
      Problems grid — the account is a site-collection admin, so every
      group's membership is readable (stubbed path only).
- [x] Console: no errors from the Workbench across Pages and Permissions.
      The only console entry is the one deliberate 400 of the EEEU Title
      retry (fix 6); the two message-less `pageerror`s are SharePoint's own
      page load.

**Defects found live and fixed** (branch `claude/page-status-live-fixes`;
mock fixtures now mirror the tenant, and each fix has a check that fails
without it):

1. `items(id)/versions?$select=…,OData__ModerationStatus` answers 200 with
   the value **omitted** on every library; versions carry it only as
   `OData__x005f_ModerationStatus`. Publish Date / "Last published" were
   blank on every library. `versionShapes` selects the escaped key
   (`VERSION_MODERATION_FIELD`), and `moderationApplies()` lets a missing
   value withhold the date only where the list's own `EnableModeration`
   is on — the `_ModerationStatus` *field* exists (hidden) on every pages
   library, so the field probe alone never meant "approval on".
2. Metadata Modified By / Created By read "7": the inline
   `$expand=FieldValuesAsText` answers person fields with lookup ids; the
   separate `/FieldValuesAsText` endpoint returns names. Always two reads
   now (pre-existing code, surfaced by the new spec rows).
3. "Item Type" matched SharePoint's hidden built-in `FSObjType` (titled
   "Item Type") when `FolderType` is absent, showing "0". Display-name
   aliases skip hidden fields.
4. The EEEU item scan stopped at 5,000 items: `getAll`'s `cap` defaulted to
   `PAGE_CAP` even with `allowLargeCap`, so the documented 100,000 ceiling
   never applied. No `cap` now means the ceiling.
5. With 4 fixed, Cancel waited out a large library's full read (~144 s).
   `getAll` takes an optional `shouldStop`, checked between pages.
6. The EEEU item read named `Title`, which a 119 wiki library lacks — the
   whole read 400'd into a Problems row. It retries once without `Title`.

## SP Workbench: page copy (2026-09-23)

Branch `claude/page-copy` (not merged; no PR opened yet). The plan, the live spike's
answers (§11) and a dated progress log (§12) are in `design/PAGE-COPY.md` — read §1
(decisions), §11 and §12 first; this section is the orientation only.

**What it does.** Pages detail → **Copy…** (modern Site Pages, 119): duplicate the page
in its own site, or copy it to another site or site collection on the tenant. The dialog
reads its own snapshot (the sitepages DTO — never the item field), connects a fresh
client per destination, runs a bound preflight (names, folder, carry set, per-part
verdicts, assets), and executes a frozen plan with a run journal (`page-copy-run.js`).

**Decision 2 (engine), settled by the spike:** the sitepages API, created by *path A* —
`POST sitepages/pages`, a staging Title (`{stem}~copy-{id}`) that names the file, the
content save, then `MoveFileByPath` to the final name/folder, Description and the carry
set through `ValidateUpdateListItem`, `SetCommentsDisabled`, publish only when ticked
(else a minor check-in), and a read-back verify. `CopyFileByPath` is used only for a
legacy-HTML canvas within its own site. Why: a whole-file copy of a promoted page lands
still promoted (in the news query within a second), and `addTemplateFile` (path B) makes
a file without the Site Page content type.

**Live-proven behaviours (don't "fix" these away):** savepage blanks `Description` (write
it via VULI); a custom-thumbnail page's DTO `BannerImageUrl` *is* the thumbnail and
`BannerThumbnailUrl` is a tokened CDN URL (never copy it); only the DTO's
`CommentsDisabled` reflects `SetCommentsDisabled`; `Folders/AddUsingPath` refuses an
existing folder (probe first); a modern page's list-item `CanvasContent1` can be null
(the Copy button accepts 'empty' items); `AddImageFromExternalUrl` cannot fetch
same-tenant files (assets always go over the bytes path).

**Live-verified (dev, Build #237):** same-web duplicate on NewNerve; cross-site copies of
`zz-pagecopy-xsite.aspx` to `/sites/TestSiteCollection` (published, rendered, assets
fetched from the destination) and to `/sites/NewNerve/sputils-test`. No read-only account
exists on the tenant, so reader visibility was established structurally (published major,
inherited permissions, Visitors = Read) — see §12.

**Open items.**
- Analyzers still to write (they need editor-authored shapes): Image (SharePoint file),
  Image gallery, Hero, Call to action, Countdown, File viewer, Events, Highlighted content,
  Quick chart, Sites. `zz-pagecopy-shapes.aspx` on the dev web holds a default instance of
  each, waiting for its properties to be set in the page editor (§11 Q7); until then those
  parts copy verbatim with their references reported "unverified".
- User/Lookup/Taxonomy columns are not carried by the API engine (documented limitation).
- Cross-web DateTime carry uses the source's display text; a destination in a different
  regional format may reject it (reported per field, page kept).
- Every `zz-pagecopy-*` artefact on the three dev webs must be recycled when testing ends.

Suite: `tests/workbench-pages-copy.mjs` (106 checks; sections in `tests/pages-copy/`,
own mock webs in `src/workbench/mock-pagecopy.js`).

## Roadmap (seams reserved)

- **Site Inspector** — v1 + Tier 2 shipped as the **SP Workbench** (above).
  Remaining seams: mount its views into the pad sidebar,
  User/Lookup/Taxonomy metadata editing, chunked >50 MB uploads.
- **List Schema** — stage 1a shipped (above). Remaining: stage 1b (item
  data on the same engine), stage 2 (document libraries, schema only).
- **SharePoint JSON storage** replacing localStorage via the `state.js` seam.
- **Console remote handles** — lazy live-object expansion.
