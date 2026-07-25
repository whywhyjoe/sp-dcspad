# DCSPad — SharePoint Developer Workbench

A JSFiddle-style workbench for SharePoint development that runs entirely in the browser: HTML/CSS/JS editors, live preview, built-in console + network monitor, an SP-aware object inspector, and a library manager. No backend, no build step, no SPFx — deploy by uploading this folder to a SharePoint library.

This is a from-scratch rebuild focused on the **execution environment**. The core guarantee: **code written in the pad runs unmodified on a real SharePoint page.**

## Why the old version broke, and how this one doesn't

The previous build ran user code in a `Blob` + `createObjectURL()` iframe. Blob URLs get an *opaque origin*: `window.location` becomes `blob:…`, cookies don't flow, and PnPjs/SP REST can't resolve `_api` URLs → `"Failed to parse URL from ''"`. Script injection into a live document also made load order a race → "sometimes it initializes, sometimes it doesn't."

This version:

1. **Same-origin `srcdoc` iframe.** The preview document inherits the host SharePoint page's origin — cookies, auth, and same-origin REST just work.
2. **Deterministic full-document assembly.** Every run builds one complete HTML document — harness → SP context + `<base>` → library CSS → user CSS → user HTML → ordered blocking library scripts → user JS last — and hands it to the browser's normal parser. Exactly how a real page loads; no injection races.
3. **Fresh iframe per run.** The old iframe is destroyed; no leaked globals, no half-initialized state.
4. **`postMessage`-only instrumentation.** The in-iframe harness never gets reached into from outside; all console/network/error/REPL traffic crosses the boundary as messages, tagged with a per-run token so stale frames are ignored.
5. **SP context bridge.** The host page's real `_spPageContextInfo` is re-captured at every run and injected, and `<base href>` points at the web. PnPjs v2 auto-resolves its base URL from that — `pnp.sp.web.get()` works with zero pad-specific setup. On classic pages the `__REQUESTDIGEST` form field is re-read each run, so the injected digest tracks the host's refresh timer; on modern pages the digest is whatever the host exposes — for raw REST writes in a long session, fetch a fresh one from `/_api/contextinfo` (PnPjs does this for you automatically).

## Local development

```bash
cd devpad
python3 -m http.server 8642
# open http://localhost:8642/index.html
```

Outside SharePoint you get a clearly-flagged **mock** `_spPageContextInfo` (correct shape, `SP: Mock` chip). `_api` calls will 404 locally — that's expected; deploy to SharePoint for live APIs.

## Deploying to SharePoint

1. Upload the whole `devpad/` folder to a document library on your site — e.g. **Site Assets** → `SiteAssets/devpad/`. Keep the folder structure (`src/`, `styles/`, `vendor/`).
2. Open `…/SiteAssets/devpad/index.html` in the browser.
   - If your tenant blocks rendering `.html` files (they download instead), rename `index.html` → `devpad.aspx` and open that. No other changes needed.
3. Confirm the top-right chip reads **SP: Live** with your web URL in the status bar.

> The page must be opened where `_spPageContextInfo` exists (any classic page context, or a page on a custom-script-enabled site). If the chip says Mock while on SharePoint, the host page didn't expose `_spPageContextInfo` — host the file on a site with custom script enabled.

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

## Using the pad

| Thing | How |
|---|---|
| Run | `Run` button or `Ctrl/Cmd+Enter` anywhere in an editor |
| Auto-run | Toggle in the toolbar; re-runs ~800 ms after you stop typing |
| Top-level `await` | Settings ⚙ → "Run JS as module" (strict mode; `var` won't become window globals) |
| REPL | Input line under the console — evaluates *inside the current run's iframe*; `↑`/`↓` history; promises are awaited |
| Stack traces | Frames pointing into your JS are clickable → jumps the editor to that line |
| Libraries | Left sidebar; checkbox = include on next run, ★ = pin to top; custom URLs (`.js`/`.css`) at the bottom |
| Network | `_api only` filter; click a row for the response body rendered through the SP inspector |
| Maximize | ⛶ on the preview or console panel; `Esc` restores |
| Preview dark mode | ☀/🌙 on the preview header (default dark). Pad-only canvas color injected *before* your CSS, so anything you style wins — flip to light to see how it renders on a typical SharePoint page |

Work-in-progress (editors, libraries, settings, layout) autosaves to `localStorage` and restores on load.

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
vendor/codemirror.mjs vendored CodeMirror 6 bundle (single file — see tools/)
tools/build-vendor.mjs one-liner to regenerate the vendor bundle (esbuild)
src/
  main.js             bootstrap/wiring
  layout.js           splitters, tabs, collapse/maximize (persisted)
  editors.js          CodeMirror 6 editors
  state.js            workspace state + debounced autosave
  runner.js           document assembly + iframe lifecycle
  libraries.js        preset catalog + custom URLs
  console-panel.js    console UI + REPL
  network-panel.js    network UI
  splash.js           boot splash
  inspect/tree-view.js   generic expandable trees + tables
  inspect/sp-shapes.js   SP/OData/PnPjs smart views
  bridge/harness.js   iframe-side instrumentation (injected per run)
  bridge/sp-context.js   real/mock _spPageContextInfo capture
```

CodeMirror is vendored as one ESM file to avoid CM6's duplicate-`@codemirror/state` pitfall and any CDN dependency; regenerate with the commands at the top of `tools/build-vendor.mjs`.

## Tests

Playwright browser suites live in `tests/` — see `tests/README.md` for the two-server setup. Run them after touching the runner, harness, console/network panels, or inspector.

## Roadmap (deliberately not built yet)

- **Site Inspector** — sidebar panel for the discovery every SP build starts with: lists/libraries with GUIDs + internal names, fields with display/internal names + types, security groups & members, content types — rendered through `src/inspect/` with "copy as PnPjs call" so discovery flows into code.
- **Snippets / projects / templates / shared team resources** on SharePoint-backed JSON storage (`DevPadData/{user}.json`), replacing localStorage wholesale — `state.js` already keeps the workspace as a single serializable blob for exactly this.
- **Console remote handles** — lazy live-object expansion instead of eager capped serialization.
