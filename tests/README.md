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
node smoke.mjs      # 49 checks: capture, isolation, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave
node monaco.mjs     # 26 checks: models, diagnostics, PnPjs + Alpine + BMO completion/hover, legacy catalogs, isolated snippet undo/redo, assets and worker degradation
node config.mjs     #  7 checks: config URL resolution, framework/asset intelligence, and ordered framework fallback
node hosted.mjs     # 10 checks: early/delayed splash, exact boot/bundle/config path, hosted flag/base, versioned runtime/types/intelligence and same-origin worker
node darkmode.mjs   #  8 checks: preview theme toggle + user-CSS-wins layering
node splash.mjs     #  3 checks: boot splash lifecycle
```

`npm test` runs all 103 checks in that order. Exit code is non-zero on any
failure. Override endpoints with `DCSPAD_URL` / `DCSPAD_FIXTURES` if you serve
on different ports.

Known quirk: a failing `custom library loads` check almost always means the fixtures server (8643) isn't running. Public-CDN library presets can't be exercised from the Claude sandbox (egress proxy blocks CDNs) — the fixture covers the identical injection mechanism.
