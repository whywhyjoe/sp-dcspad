# Handoff — DCSPad

Written at the end of the session that split this repo out of `whywhyjoe/todo`.
Read `CLAUDE.md` first for architecture invariants; this file covers **state**
and **the next problem**.

## Where things stand

- Repo: `whywhyjoe/sp-dcspad`, single branch `main`. The old `whywhyjoe/todo`
  copy under `devpad/` is frozen — do not develop there.
- App works fully as a standalone page. Test suites: `smoke.mjs` 49 checks,
  `darkmode.mjs` 8, `splash.mjs` 3 — all passing. Two-server setup in
  `tests/README.md`.
- Three external code-review rounds triaged and landed; decisions, declines and
  corrections in `REVIEW-LOG.md`, which also carries an open low-priority
  backlog (history caps, run-scoped load timeout, build-vendor cleanup, etc.).
- `deploy/Deploy-DcsPad.ps1` exists but is **unexecuted and not syntax-checked**
  (no PowerShell in the authoring container). Run with `-WhatIf` first.

## The next problem: hosting inside a custom-script web part

The deployment target is **not** a standalone `.html` page. It is a modern
SharePoint page with a custom-script web part that fetches an HTML file and
injects it into a container div on that page. Everything below follows from
that, and none of it has been tested yet.

### What almost certainly breaks

1. **Relative asset URLs.** `index.html` references `styles/app.css` and
   `src/main.js` relatively. Injected into a page at
   `…/SitePages/Dev.aspx`, those resolve against *the page*, not the library
   folder — 404. Both entry points need absolute (or server-relative) URLs.
2. **Full-viewport layout** — *largely solved by the hosting plan, see below.*
   `.app` is `height: 100vh` (app.css:54) and `html, body { height: 100% }`
   (app.css:34). The plan is to render the host page chrome-less via
   SharePoint's `?env=WebView`, which gets the pad most of a viewport. Residual
   risk is the modern-page canvas: the web part still sits inside SharePoint's
   `CanvasZone` wrappers, which carry their own max-width, padding and
   overflow. Expect to either put the web part in a full-width section or pin
   the app with `position: fixed; inset: 0` (the splash already does exactly
   this at app.css:501). Watch for double scrollbars — the host document
   scrolling behind the app is the tell.
3. **Global CSS collides in both directions.** `app.css` sets rules on `html`,
   `body`, `:root` and a global `[hidden] { display: none !important }`
   (app.css:33), which restyle the SharePoint page itself — `env=WebView`
   shrinks the blast radius but doesn't remove it. The *reverse* is the more
   likely source of visual bugs and was missed in the first pass: SharePoint
   ships a large global stylesheet (resets, box-sizing, font stacks, line
   heights) that will bleed **into** the pad, which was written assuming a
   clean document. Budget time for both: scope the pad's rules under the app
   root, and defend its own layout against inherited styles.
4. **`<script>` tags in injected HTML do not execute.** Content set via
   `innerHTML` never runs its scripts. Script-editor web parts work around this
   by re-creating the script elements — whether the one in use does that *for
   `type="module"`* is the open question.

### What survives (don't panic-rewrite)

- **The whole execution core is unaffected.** The `srcdoc` iframe, per-run
  tokens, harness instrumentation, console/network panels and the inspector
  don't care where the pad's chrome lives. The iframe is still same-origin
  because the hosting page is on the SharePoint origin.
- **ES module imports resolve against the importing module's URL, not the
  page.** So if `src/main.js` is loaded from an absolute URL, every
  `./state.js` and `../vendor/codemirror.mjs` beneath it resolves correctly
  with no further changes. Only the two *entry* URLs need fixing — this is why
  the module question is narrower than it looks.
- **SP context may well be better here.** Modern pages generally expose
  `_spPageContextInfo`, so the chip may read **SP: Live** without the
  custom-script site setting the standalone deployment needs. Verify rather
  than assume.

### The hosting plan: chrome-less via `?env=WebView`

The decided approach for making the pad usable on a modern page: render the
host page with SharePoint's `?env=WebView`, which drops the site header,
navigation and suite bar, and have the injected HTML bounce the page to that
URL when it wasn't loaded that way.

Two things that redirect must get right:

```js
(function () {
  var params = new URLSearchParams(location.search);
  if (params.get('env') === 'WebView') return;
  // Never hijack page editing — without this you cannot edit the page the
  // pad is embedded on, because every load bounces to the chrome-less view.
  if (params.has('Mode') || location.pathname.indexOf('/_layouts/') > -1) return;
  // Loop guard: if SharePoint strips or rewrites the param, one attempt per
  // tab is all we get. Without this a rejected param reloads forever.
  if (sessionStorage.getItem('dcspad.webview.tried')) return;
  sessionStorage.setItem('dcspad.webview.tried', '1');
  params.set('env', 'WebView');
  // replace(), not href: otherwise Back bounces the user straight forward again.
  location.replace(location.pathname + '?' + params + location.hash);
})();
```

Note this redirect is itself script running in injected HTML — it only works if
spike test 1 passes. And confirm `_spPageContextInfo` is still exposed under
`env=WebView` (spike test 4, re-run with the param) rather than assuming the
chrome-less render carries the same page context.

### Do this first: a 20-minute spike

`deploy/webpart-spike.html` is a self-reporting probe. Deploy it to the library,
point the web part at it, and read the results on the page. It answers, in order:

1. Does a **classic** `<script>` run when injected?
2. Does a **`type="module"`** script run?
3. Does a module loaded from an **absolute URL** resolve its own relative
   imports correctly?
4. Is `_spPageContextInfo` visible from inside the web part?

The answers pick the path:

- **Modules run (likely with a decent script web part):** the work is
  URL rewriting + CSS scoping + container sizing. A day, not a rewrite. Best
  shape is a generated `dcspad.webpart.html` produced at deploy time, with the
  site/folder baked into absolute URLs — `Deploy-DcsPad.ps1` already knows both.
- **Modules do not run:** bundle `src/` into one classic IIFE with esbuild
  (`tools/build-vendor.mjs` is the existing precedent) and ship
  `dcspad.bundle.js`. This costs the "no build step" invariant in CLAUDE.md —
  update invariant 5 deliberately rather than letting it rot. Source stays
  modular; only the artifact changes.

Either way, **keep `index.html` working standalone.** It's what the test suites
drive, and losing it means losing the 49 checks.

### Open questions for the new session

- Which web part exactly? (Modern Script Editor, a custom SPFx one, something
  else.) Its script-handling implementation decides everything above.
- Fixed height, or should the pad offer a maximize-to-viewport mode? A
  workbench in a 400px web part box is unpleasant.
- Does the tenant block `.mjs`, or serve it with a non-JavaScript MIME type?
  See the third gotcha in `deploy/README.md` — it's the same risk in a new
  guise, and the fix is renaming to `.js` plus one import line.
