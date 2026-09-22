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
## Roadmap (seams reserved)

- **Site Inspector** — v1 + Tier 2 shipped as the **SP Workbench** (above).
  Remaining seams: mount its views into the pad sidebar,
  User/Lookup/Taxonomy metadata editing, chunked >50 MB uploads.
- **List Schema** — stage 1a shipped (above). Remaining: stage 1b (item
  data on the same engine), stage 2 (document libraries, schema only).
- **SharePoint JSON storage** replacing localStorage via the `state.js` seam.
- **Console remote handles** — lazy live-object expansion.
