# Handoff — DCSPad

Read `CLAUDE.md` first for architecture invariants; this file covers **state**
and **what's next**. Last updated: 2026-07-25, the session that got the pad
running live inside the SharePoint web part.

## Where things stand

**The pad runs on the real SharePoint page.** Editors, run pipeline, console,
inspector, REPL and network capture all work inside the web part; a live
`_api/web` call from the REPL returns real site data. The standalone
`index.html` still works and remains what the test suites drive.

### Hosting model (decided and working)

- Host page: `https://nervedotnet.sharepoint.com/sites/NewNerve/SitePages/DCSpad.aspx`
  (⚠ the published copy currently lives at `DCSpad(1).aspx` — a duplicate got
  created when the original page wouldn't save without a title; reconcile).
- Web part: PnP **Modern Script Editor** (Mikael Svenson), in
  "Use script from an external URL" mode. Despite the field name "Script URL",
  it fetches an **HTML file** and injects it with script re-creation.
- The URL points at `dcspad.webpart.html` (repo root): a 2-line anchor +
  absolute `<script src=…/boot.js>`. `boot.js` fetches `index.html`, injects
  the app shell, loads `styles/app.css`, and `import()`s `src/main.js` from an
  absolute URL — after which the whole module tree resolves itself. Entry URLs
  are the only absolute URLs; `index.html` stays the single source of truth.
- **Deployment = file copy.** The doc library folder
  `…/SiteAssets/Code/dcspad-live/` is OneDrive-synced to
  `C:\Users\other\NERVE\NewNerve - Code\dcspad-live`; copying files there goes
  live in seconds. `deploy/Deploy-DcsPad.ps1` exists but is unused.
  **Deploy with `deploy/Sync-Live.ps1`** — it rebuilds `dcspad.app.js` (the
  single-file bundle the web part actually runs) and copies everything.
  ⚠ Never hand-copy `src/` changes without rebuilding the bundle: hosted
  mode won't see them. Why a bundle: SPO's `max-age=86400`, Chrome's
  separate module-request cache, and the host page ignoring late import
  maps make a multi-file graph un-bustable (full story: cache gotcha in
  CLAUDE.md). Bump `?v=` in `dcspad.webpart.html` when boot.js itself
  changes. SPO can also serve a just-uploaded file stale for ~15–30s —
  verify with a cache-busted fetch before debugging "my change didn't work".
- Visual seating (deliberate, Joe's call): the SharePoint **suite bar stays
  visible** (desaturated while the pad runs); hosted mode pins `.app` at
  `inset: 53px 5px 5px`, borderless over a darker surround so it reads as
  part of the page (see "Web-part hosting" in `app.css`, activated by
  `boot.js` adding `.dcspad-hosted` to `<html>`). The earlier `env=WebView`
  chrome-less plan is abandoned — the suite bar is wanted context.
- **Page-edit safety**: boot.js refuses to boot when the URL carries
  `Mode=Edit` (or under `/_layouts/`), and also watches SPA navigation
  (pushState/replaceState/popstate) to *suspend* an already-booted pad —
  `.dcspad-suspended` hides the mount and reverts the html/body overrides so
  the edit canvas looks and scrolls normally. Leaving edit mode restores it.

### The two web-part-hosting bugs that were found and fixed

1. **SharePoint serves `.mjs` as `application/octet-stream`** — browsers
   refuse ES modules with that MIME. `vendor/codemirror.mjs` is now
   `vendor/codemirror.js` (import updated in `src/editors.js`, build output in
   `tools/build-vendor.mjs`). This is permanent: never ship `.mjs` to SP.
2. **Modern pages ship a nonce-based `script-src` CSP that `about:srcdoc`
   inherits.** Preview frames rendered markup but silently executed nothing.
   `runner.js` now stamps the host page's nonce on every assembled script tag
   (`hostNonce()`); standalone pages have no nonce → no-op. If preview
   execution ever "silently dies" on SP again, look here first.

### Spike results (deploy/webpart-spike.html, kept for re-verification)

1. classic injected script executes ✅
2. `type="module"` script executes ✅
3. relative import from an absolute module URL resolves ✅
4. `_spPageContextInfo` present ❌ — absent on the modern page by default

## Open items

- **SP chip reads Mock.** The web part has an "Enable classic
  _spPageContextInfo" toggle, still Disabled. Flipping it should make the chip
  read Live and give `sp-context.js` a real context (needed for request-digest
  POSTs; GET `_api` calls already work via session cookies). Verify after
  flipping — if the toggle injects context *after* boot captures it, capture
  timing in `sp-context.js` may need a retry.
- **Page naming**: reconcile `DCSpad.aspx` vs `DCSpad(1).aspx`, keep one.
- **CSS bleed, both directions**: `app.css` still styles `html`/`body`
  (darkens the host page behind the pad — currently invisible and arguably
  nice; the gap around the seated app shows it). SP styles also bleed into the
  pad; nothing visibly broken, but scope properly if oddities appear.
- Low-priority backlog lives in `REVIEW-LOG.md`.

## Feature backlog (Joe, 2026-07-25)

- **Per-pane import/export (file system + SharePoint)** — plan written, not
  started: `plans/file-sp-import-export.md` (REST `$value`/`Files/add` +
  contextinfo digest; new `src/sp-files.js` seam).
- **PnPjs autocomplete/IntelliSense** — plan written, not started:
  `plans/pnpjs-intellisense.md` (CM6 + TS worker recommended; CSP `worker-src`
  check gates it).
- **"VS Code editor" investigation** — i.e. Monaco (the pad currently uses
  CodeMirror 6). If a swap is decided, fold the IntelliSense plan into it.
- **Drag-and-drop reordering** for the frameworks list (replaces/augments ↑↓).
- **Framework/snippet row action icons only on hover** (currently always
  visible at 60% opacity).
- **Site Inspector / SP diagnostics tools** — already in the roadmap below;
  explicitly wanted.

## Roadmap (unchanged, seams reserved)

- **Site Inspector** sidebar section (renders through `src/inspect/`).
- **SharePoint JSON storage** replacing localStorage via the `state.js` seam.
- **Console remote handles** — lazy live-object expansion.
