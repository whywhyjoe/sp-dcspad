// Bundles first-party src/main.js into one ESM file: ../dcspad.app.js — the
// artifact web-part hosting loads. Monaco remains an external, separately
// versioned vendor runtime because its workers and CSS/font assets must keep
// stable URLs. Why an app bundle at all: SharePoint serves library files with
// max-age=86400, Chrome caches module requests separately from fetch(),
// and the host page freezes import-map registration — so the ONLY reliable
// cache-busting unit is a single entry file behind a versioned URL
// (boot.js stamps it with its Last-Modified). One file = one version = no
// mixed-version graph. See the cache gotcha in CLAUDE.md.
//
// Source stays modular; standalone index.html and the test suites keep
// loading src/ unbundled. Rebuild after any src/ change:
//   cd tools && node build-app.mjs
// (deploy/Sync-Live.ps1 runs this for you.)
//
// src/bridge/harness.js is fetched as text at runtime, not imported, so it
// deploys as a separate file alongside the bundle.

import { build } from 'esbuild';

await build({
  entryPoints: ['../src/main.js'],
  bundle: true,
  format: 'esm',
  outfile: '../dcspad.app.js',
  logLevel: 'info',
});
console.log('dcspad.app.js rebuilt');
