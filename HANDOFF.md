# Handoff — DCSPad

Read `CLAUDE.md` first for architecture invariants; this file covers **state**
and **what's next**. Last updated: 2026-07-26, the Monaco migration.

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
- Enabling the PnPjs v2 runtime library loads the declarations and global
  `pnp` bridge; disabling it unloads them. Fluent completions, hover,
  signatures and JS diagnostics use Monaco's TypeScript worker.
- Runtime detection follows the enabled 2.15.0 script URL rather than a fixed
  catalog id, so the live custom `Pnpjs JSD` jsDelivr entry receives the same
  matching types.
- `tools/build-monaco.mjs` is the only way to regenerate vendor assets.
  `vendor/monaco/version.json` versions the set as a unit in hosted mode.
- `tests/monaco.mjs` covers the editor contract, completions, diagnostics,
  declaration lifecycle, asset policy and worker failure behavior.

## Open items

- **CSS bleed, both directions**: `app.css` still styles `html`/`body`
  (darkens the host page behind the pad — currently invisible and arguably
  nice; the gap around the seated app shows it). SP styles also bleed into the
  pad; nothing visibly broken, but scope properly if oddities appear.
- Low-priority backlog lives in `REVIEW-LOG.md`.

## Feature backlog (Joe, 2026-07-25)

- **UI design pass** — handoff package ready in `design/DESIGN-BRIEF.md` +
  `design/screenshots/` (captured via `tests/capture-design-shots.mjs`).
  Joe attaches his old-DCSPad reference screenshot as
  `design/reference-old-dcspad.png` when submitting. Claude Design returns
  a comp + token sheet; we implement (tokens are plain CSS vars + a Monaco
  theme, so implementation is mechanical).

- **Per-pane import/export (file system + SharePoint)** — plan written, not
  started: `plans/file-sp-import-export.md` (REST `$value`/`Files/add` +
  contextinfo digest; new `src/sp-files.js` seam).
- **Drag-and-drop reordering** for the frameworks list (replaces/augments ↑↓).
- **Framework/snippet row action icons only on hover** (currently always
  visible at 60% opacity).
- **Site Inspector / SP diagnostics tools** — already in the roadmap below;
  explicitly wanted.

## Roadmap (unchanged, seams reserved)

- **Site Inspector** sidebar section (renders through `src/inspect/`).
- **SharePoint JSON storage** replacing localStorage via the `state.js` seam.
- **Console remote handles** — lazy live-object expansion.
