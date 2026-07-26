# Plan: PnPjs autocomplete / IntelliSense in the JS editor

Status: **planned, not started** (written 2026-07-25; requested because PnPjs
v2 classic is effectively always enabled).

## Answer to "is it difficult?"

Moderately. Real IntelliSense (chaining-aware: `pnp.sp.web.lists.getByTitle(…).items.select(…)`)
needs a TypeScript language service running against PnPjs's `.d.ts` files —
static keyword lists can't follow return types through a fluent API. That is
absolutely doable inside CodeMirror 6 without switching editors, but it means
one new vendored bundle and a worker, so it's a half-day-plus of careful work,
not an afternoon tweak.

## Options considered

1. **CM6 + TypeScript worker (recommended).** Keep CodeMirror; add
   [`@valtown/codemirror-ts`](https://github.com/val-town/codemirror-ts) +
   `typescript` running in a Web Worker, with PnPjs 2.x type declarations
   preloaded into the virtual FS. Gives completions, hover types, and
   signature help in the JS pane only. Footprint ≈ 3–4 MB minified worker,
   vendored like `vendor/codemirror.js` (same `tools/` esbuild precedent, so
   invariant 5's "no user build" stays true — it's a vendored artifact).
2. **Monaco migration** (the "VS Code editor" backlog item — Monaco *is* VS
   Code's editor; the pad currently uses CodeMirror 6, not Monaco). Monaco
   ships a TS worker natively; `addExtraLib(pnpjs.d.ts)` gives the same
   IntelliSense with less glue. But it replaces `editors.js` wholesale
   (Mod-Enter, gotoJsLine stack links, theme, layout persistence), costs
   ~5 MB+, and its AMD/worker packaging is awkward to vendor. Only worth it
   if the Monaco investigation decides to swap editors anyway — in which case
   fold this plan into that migration rather than doing both.
3. **Curated static completion source** (cheap interim). A hand-maintained
   JSON of `pnp.sp.*` members fed to CM6's `autocompletion()` override.
   Hours of work, zero new vendoring — but prefix-only, no chaining
   awareness, goes stale with PnPjs versions. Do this only if (1) stalls.

## Plan for option 1

1. **Vendor the worker.** New `tools/build-ts-worker.mjs` (esbuild, same
   pattern as `build-vendor.mjs`) producing `vendor/ts-worker.js` +
   `vendor/codemirror-ts.js`. Pin `typescript` and `@valtown/codemirror-ts`
   versions in the tool header.
2. **Bundle the types.** Script step that flattens `@pnp/sp@2.x` (+ `@pnp/odata`,
   `@pnp/common`, `@pnp/logging`) `.d.ts` files into `vendor/pnpjs-types.json`
   (path → contents map), plus a small `globals.d.ts` declaring
   `declare const pnp: typeof import('@pnp/sp/presets/all')`-style globals to
   match what the classic UMD bundle exposes as `window.pnp`.
3. **Wire into `src/editors.js`.** JS pane only: create the ts environment in
   the worker, load `pnpjs-types.json`, attach the completion/hover/lint
   extensions behind a feature check. HTML/CSS panes untouched.
4. **CSP check on SharePoint (do first — it gates everything).** The host
   page's CSP includes a `worker-src` directive; verify a same-origin worker
   script (`vendor/ts-worker.js`) is allowed — blob: workers may not be. If
   workers are blocked entirely, the language service can run on the main
   thread as a degraded fallback (typecheck on idle only), or option 3 kicks in.
5. **Settings toggle.** "PnPjs IntelliSense" checkbox in the ⚙ menu
   (default on); worth having because the worker costs ~4 MB on first load —
   cached by the browser afterward.
6. **Tests.** Extend `smoke.mjs`: type `pnp.sp.` in the JS pane, assert a
   completion tooltip lists `web`; assert the pad still boots with the worker
   file absent (graceful degradation).

## Open questions

- Exact PnPjs version to pin types to (catalog preset says v2 classic —
  confirm the tenant's UMD bundle version so types match).
- Whether to also type `_spPageContextInfo` and the harness globals while
  we're in there (cheap once the pipeline exists).
