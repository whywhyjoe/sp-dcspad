// Bundles src/workbench/main.js into one ESM file: ../dcspad.workbench.js —
// the artifact the workbench web-part page loads. Same rationale as
// build-app.mjs: SharePoint's caching makes a multi-file module graph
// un-bustable, so hosted mode gets ONE bundled entry behind a versioned URL
// (boot-workbench.js stamps it with its Last-Modified).
//
// Source stays modular; standalone workbench.html and the test suite load
// src/ unbundled. Rebuild after any src/ change touching workbench modules:
//   cd tools && node build-workbench.mjs
// (deploy/Sync-Live.ps1 runs this for you.)

import { build } from 'esbuild';

await build({
  entryPoints: ['../src/workbench/main.js'],
  bundle: true,
  format: 'esm',
  outfile: '../dcspad.workbench.js',
  logLevel: 'info',
});
console.log('dcspad.workbench.js rebuilt');
