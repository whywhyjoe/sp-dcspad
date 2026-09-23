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

import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

// Same build-identity stamping as build-app.mjs: the bundle logs and shows
// its build number so a stale cached bundle on a tenant is diagnosable at
// a glance. Unbundled standalone loads identify as Build #dev.
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function git(...args) {
  try {
    return execFileSync('git', args, {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
  } catch {
    return '';
  }
}

const requestedBuildNumber = String(process.env.DCSPAD_BUILD_NUMBER || '').trim();
// A shallow clone (cloud sessions, most CI checkouts) can only count the
// commits it fetched, so rev-list silently returns a number far below the
// real one and the build stamp goes BACKWARDS. Refuse to guess: unshallow,
// or pass DCSPAD_BUILD_NUMBER explicitly.
if (!requestedBuildNumber && git('rev-parse', '--is-shallow-repository') === 'true') {
  console.error(
    'Refusing to stamp a build number from a shallow clone — the count would be '
    + 'too low.\nRun `git fetch --unshallow`, or set DCSPAD_BUILD_NUMBER.',
  );
  process.exit(1);
}
const commitCount = git('rev-list', '--count', 'HEAD');
const shortRevision = git('rev-parse', '--short=8', 'HEAD');
// Build outputs never make a build 'dirty' — exclude the bundles.
const trackedChanges = git('status', '--porcelain', '--untracked-files=no', '--', '.', ':!dcspad.app.js', ':!dcspad.workbench.js');
const buildNumber = requestedBuildNumber
  || commitCount
  || new Date().toISOString().replace(/\D/g, '').slice(0, 14);
const buildLabel = `${buildNumber}${trackedChanges ? '-dirty' : ''}`;
const revisionLabel = `${shortRevision || 'unknown'}${trackedChanges ? '-dirty' : ''}`;

await build({
  entryPoints: ['../src/workbench/main.js'],
  bundle: true,
  format: 'esm',
  outfile: '../dcspad.workbench.js',
  define: {
    __DCSPAD_BUILD_NUMBER__: JSON.stringify(buildLabel),
    __DCSPAD_BUILD_REVISION__: JSON.stringify(revisionLabel),
  },
  logLevel: 'info',
});
console.log(`dcspad.workbench.js rebuilt — Build #${buildLabel} (${revisionLabel})`);
