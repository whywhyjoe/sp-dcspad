# PnPjs IntelliSense implementation record

Status: **implemented with the Monaco migration on 2026-07-26**.

## Decision

DCSPad moved from CodeMirror 6 to Monaco 0.55.1 because the goal was the
broader VS Code-style editing experience—suggestions, hover, signatures,
diagnostics, navigation and find/replace—not autocomplete alone. The rest of
the application still talks through the editor adapter in `src/editors.js`.

## Implementation

- `tools/build-monaco.mjs` pins and bundles Monaco, its CSS/font and four
  classic `.js` workers under `vendor/monaco/`.
- The same build gathers all 254 declaration files from the exact
  `@pnp/pnpjs` 2.15.0 dependency graph into `pnpjs-types.json`.
- A generated global declaration maps the module graph to the actual UMD
  runtime shape: `const pnp: typeof import("@pnp/pnpjs")`.
- The build validates representative `pnp.sp.web...`, list, destructuring,
  and namespace calls with TypeScript before writing the vendor set.
- `src/monaco-runtime.js` resolves standalone and hosted asset URLs and starts
  same-origin workers. It uses no CDN, blob worker, or `.mjs` file.
- `src/editors.js` loads the declaration graph only while the PnPjs v2 runtime
  library is enabled, and disposes it when disabled.
- If a language worker is blocked, the editor remains usable and a dedicated
  Monaco status item reports that language tools are limited; running user
  code cannot overwrite that infrastructure warning.

## Verification

`tests/monaco.mjs` exercises typed models, tab/model persistence, the run
shortcut, JavaScript diagnostics, chaining-aware `pnp.sp.web` completion,
declaration unloading, isolated snippet undo/redo, asset policy, clean page
execution and persistent blocked-worker degradation.

The remaining environment-specific gate is the live SharePoint page's
`worker-src` policy. `deploy/webpart-spike.html` includes a same-origin classic
worker check that can be rerun independently, and the hosted application
provides the definitive end-to-end validation.
