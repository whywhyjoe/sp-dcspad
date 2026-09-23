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

import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';

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
  entryPoints: ['../src/main.js'],
  bundle: true,
  format: 'esm',
  outfile: '../dcspad.app.js',
  define: {
    __DCSPAD_BUILD_NUMBER__: JSON.stringify(buildLabel),
    __DCSPAD_BUILD_REVISION__: JSON.stringify(revisionLabel),
  },
  logLevel: 'info',
});
console.log(`dcspad.app.js rebuilt — Build #${buildLabel} (${revisionLabel})`);
