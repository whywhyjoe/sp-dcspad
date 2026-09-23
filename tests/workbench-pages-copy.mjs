// SP Workbench — page copy (design/PAGE-COPY.md).
//
// Own suite, own mock webs: /sites/pagesrc (source), /sites/pagesrc/team
// (same site collection, another web) and /sites/pagedst (another site
// collection), each with distinct site/web/list/file GUIDs so a retained
// source reference cannot pass for a copied one (mock-pagecopy.js). The
// stubbed-live sections route /_api/** themselves and are projection-strict.
//
// Sections live in tests/pages-copy/ — one module per concern, each
// exporting `run({ browser, check, WB_URL })` — and run in this order.

import { launchBrowser, check, exitWithResult, APP_URL } from './lib.mjs';

const WB_URL = process.env.DCSPAD_WORKBENCH_URL
  || APP_URL.replace(/index\.html.*$/, 'workbench.html');

const SECTIONS = [
  './pages-copy/pure.mjs',
  './pages-copy/runner.mjs',
  './pages-copy/live.mjs',
];

const browser = await launchBrowser();
for (const section of SECTIONS) {
  const mod = await import(section);
  await mod.run({ browser, check, WB_URL });
}
await browser.close();
exitWithResult();
