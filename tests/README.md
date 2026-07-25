# DCSPad tests

Playwright-driven browser suites against a locally served copy of the app. They exist because this project's failure modes are timing- and frame-boundary-shaped — run them after any change to `runner.js`, `harness.js`, the console/network panels, or the inspector.

## Setup (once)

```bash
npm init -y && npm i playwright-core     # anywhere; or a global install
```

Chromium is resolved in this order: `CHROMIUM_PATH` env var → `/opt/pw-browsers/chromium` (Claude Code sandbox) → your locally installed Chrome (`channel: 'chrome'`). `HTTPS_PROXY` is honored automatically when set.

## Run

Two static servers, then the suites:

```bash
# terminal 1 — the app
cd devpad && python3 -m http.server 8642

# terminal 2 — test fixtures (serves tests/fixtures/ for library-injection checks)
cd devpad/tests && python3 -m http.server 8643

# terminal 3
cd devpad/tests
node smoke.mjs      # 49 checks: capture, isolation, rerun lifecycle, fragment links, inspector, network, REPL, filters, catalog + catalog files, snippets, project files, exports, storage errors, autosave
node darkmode.mjs   #  8 checks: preview theme toggle + user-CSS-wins layering
node splash.mjs     #  3 checks: boot splash lifecycle
```

Exit code is non-zero on any failure. Override endpoints with `DCSPAD_URL` / `DCSPAD_FIXTURES` if you serve on different ports.

Known quirk: a failing `custom library loads` check almost always means the fixtures server (8643) isn't running. Public-CDN library presets can't be exercised from the Claude sandbox (egress proxy blocks CDNs) — the fixture covers the identical injection mechanism.
