# DCSPad tests

Playwright-driven browser suites against a locally served copy of the app. They exist because this project's failure modes are timing-, worker-, and frame-boundary-shaped — run them after changes to the editor/runtime, runner, harness, console/network panels, or inspector.

## Setup (once)

```bash
npm init -y && npm i playwright-core     # anywhere; or a global install
```

Chromium is resolved in this order: `CHROMIUM_PATH` env var → `/opt/pw-browsers/chromium` (Claude Code sandbox) → your locally installed Chrome (`channel: 'chrome'`). `HTTPS_PROXY` is honored automatically when set.

## Run

Two static servers, then the suites:

```bash
# terminal 1 — the app (from the repo root)
python3 -m http.server 8642

# terminal 2 — test fixtures (serves tests/fixtures/ for library-injection checks)
cd tests && python3 -m http.server 8643

# terminal 3
cd tests
node smoke.mjs      # 76 checks: capture, isolation, Fluent preview runtime, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave
node monaco.mjs     # 44 checks: models, diagnostics, scrollbar overview ruler, mirrored PnPjs 2.15 + Alpine 3.15.2 runtime/version policy, pnp/pnp2 + Alpine + BMO + Fluent completion/hover, generated catalogs, isolated snippet undo/redo, JSON language service (tokenization, trailing-comma validation, .json-URI adoption, schema-request policy), assets and worker degradation
node config.mjs     # 33 checks: config/Browser/Copilot behavior, additive SharePoint file types, same-tenant formats/history/refresh/rejection, framework/asset intelligence/runtime, and fallback
node hosted.mjs     # 15 checks: early/delayed splash, SharePoint chrome reflow, exact boot/bundle/config path, hosted flag/base, versioned runtime/types/intelligence/Fluent bridge and same-origin worker
node darkmode.mjs   #  8 checks: preview theme toggle + user-CSS-wins layering
node splash.mjs     #  3 checks: boot splash lifecycle
node ux.mjs         # 18 checks: pane toggles (visibility/persistence/Ctrl+J via Monaco), editor text-size stepper, error count pills, REPL Eval button, add-framework footer validation, sidebar split persistence
node files.mjs      # 39 checks: local import confirmation/isolation, configurable SharePoint JSON import/export, SP context/site switching, ResourcePath/Browser browsing, required untitled filenames, metadata-reviewed overwrite, extension-optional upload, parent-library metadata resolution, guarded metadata prefill/write, metadata-failure fallback, and overwrite-highlight/name sync
node workbench-hosted.mjs # 10 checks: the workbench's boot-workbench.js + bundle path on a fixture whose mount sits below the fold (as on a real page) — pinned in-viewport, painted on top, edit-mode suspension, and the edit-mode boot guard
node workbench.mjs  # 118 checks: SP Workbench — Site landing view (web/user cards, subweb Inspect, status-bar role chip), mock grids incl. boot build-stamp logging + logo tooltip + status-bar build, list drilldown (incl. the Items tab: view/all-columns picker, max-items cap, ID-desc ordering, typed OData $filter/$orderby (quote-aware parsing) with inline rejection of other keys, and the markdown list-content export builder), clickable URL cells (forced new-tab opens + copy glyph), row selection with selection-scoped exports, column resizers, contained horizontal scrolling, the internal-lists expander (flag/IsCatalog/template/curated-title classification), the Pages-library resolver (visible outranks hidden, then 119/850/titled; every candidate kept) and the multi-library picker on a web carrying both a 119 and an 850 pages library (default pick, grid rebuilt for the other shape, drilldown and selection following the switch), Permissions view (groups, flattened Members roster with add/remove writes, roledefs, inheritance scan), Advanced sheets, export formats, script generator, same-tenant site switching, the stubbed live path (Accept header, paging, re-targeted requests, $select=* items projection, error-path controls mounting), and Tier 2: curated Panels, site favorites (state.js seam), query builder (compose/quote/raw round-trip/persistence), CanvasContent1 page inspector (JSON + legacy HTML formats, tolerant degrade, folders, Extract-first tabs, URL copy), pages metadata editors, page content/raw export builders, field-editor FieldValue conventions, the info-chip register guard (classification recedes; status stays the loud register, and both chip instances compose one class), and the classic-page path (library kind 119/850/titled named in the header, a field-probed query plan (BaseTemplate is never trusted as schema proof; no PromotedState in the composed query), per-item page-kind detection, Structure hidden with the reason on the kind chip, and Content Editor / Script Editor web-part content merged into Extract and the content export)
node workbench-edit.mjs # 19 checks: SP Workbench write paths — Files browser listing/sort/breadcrumbs/library picker (with picker, crumbs, filter and actions sharing one toolbar row), download + copy-direct-URL actions, per-type metadata editors, mock-writer uploads through the pad-style metadata dialog (availability/grey-out, overwrite prefill, changed-only writes, metadata retry/keep), new-folder creation with name validation, overwrite consent (pre-flight and 409 race with carried values that also write), cancel-before-upload, entity-property prefill (OData__ExtendedDescription), Esc/close guarding after metadata failure, oversized-upload rejection, and the stubbed live path (binary AddUsingPath with digest, ValidateUpdateListItem retry that never re-uploads)
```

Exit code is non-zero on any failure. Override endpoints with
`DCSPAD_URL` / `DCSPAD_FIXTURES` if you serve on different ports.

Known quirk: a failing `custom library loads` check almost always means the fixtures server (8643) isn't running. Public-CDN library presets can't be exercised from the Claude sandbox (egress proxy blocks CDNs) — the fixture covers the identical injection mechanism.
