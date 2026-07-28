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
node smoke.mjs      # 50 checks: capture, isolation, Fluent preview runtime, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave
node monaco.mjs     # 33 checks: models, diagnostics, PnPjs + Alpine + BMO + Fluent completion/hover, generated catalogs, isolated snippet undo/redo, assets and worker degradation
node config.mjs     #  8 checks: config URL resolution, framework/asset intelligence/runtime, and ordered framework fallback
node hosted.mjs     # 11 checks: early/delayed splash, exact boot/bundle/config path, hosted flag/base, versioned runtime/types/intelligence/Fluent bridge and same-origin worker
node darkmode.mjs   #  8 checks: preview theme toggle + user-CSS-wins layering
node splash.mjs     #  3 checks: boot splash lifecycle
node ux.mjs         # 18 checks: pane toggles (visibility/persistence/Ctrl+J via Monaco), editor text-size stepper, error count pills, REPL Eval button, add-framework footer validation, sidebar split persistence
node files.mjs      # 19 checks: local import confirmation plus explicit/Modern SP context, same-tenant site switching, ResourcePath browsing, cross-site import/digest upload, and overwrite confirmation
node workbench-hosted.mjs # 9 checks: the workbench's boot-workbench.js + bundle path on a fixture whose mount sits below the fold (as on a real page) — pinned in-viewport, painted on top, edit-mode suspension, and the edit-mode boot guard
node workbench.mjs  # 39 checks: SP Workbench — mock grids, list drilldown, security + site views, export formats, script generator, same-tenant site switching, and the stubbed live path (Accept header, paging, re-targeted requests)
```

Exit code is non-zero on any failure. Override endpoints with
`DCSPAD_URL` / `DCSPAD_FIXTURES` if you serve on different ports.

Known quirk: a failing `custom library loads` check almost always means the fixtures server (8643) isn't running. Public-CDN library presets can't be exercised from the Claude sandbox (egress proxy blocks CDNs) — the fixture covers the identical injection mechanism.
