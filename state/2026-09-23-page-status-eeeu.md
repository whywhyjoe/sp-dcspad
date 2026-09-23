# State — SP Workbench page status + EEEU audit

Last touched: 2026-09-23
Mode: Joe
Branch: claude/dcspad-sp-utilities-availability-ylr25q, pushed
State: built, mock-tested and review-round-1 fixed on the branch; not yet deployed or tried on a
live tenant; no PR

## What this is

Two SP Workbench features. **Page status** (Pages view): a Scan Page Status button that adds
Published / Checked out / Inheritance columns, status chips and a Permissions tab in the page
drilldown, and a fixed-order Metadata tab — a rebuild of a feature built once on the work (bmo)
machine and lost. **EEEU audit** (Permissions view): finds where everyone in the organization
has access, across the site, its lists/libraries/items and chosen subsite trees. Design, Joe's
decisions and the live-tenant checklist: `HANDOFF.md` → "SP Workbench: page status + EEEU audit
(2026-09-23)".

## Done

- Code: `src/workbench/page-status.js`, `src/workbench/eeeu-audit.js` (new);
  `views/pages.js`, `views/security.js`, `grid.js` (`setColumns`), `field-editor.js` (`layout`),
  `page-export.js` (status lines), `sp-rest.js` (one shared request queue), `mock-data.js`
  (Site Pages status fixtures + the `/sites/eeeu` audit tree), `styles/workbench.css`.
- ChatGPT review round 1: all nine findings fixed (`4cfb92e`); summary in HANDOFF.md.
- Mock-tested in the cloud sandbox: `tests/workbench.mjs` 165/165 (15 new), workbench-edit
  26, workbench-schema 138, workbench-hosted 10 (against the rebuilt bundle, Build #208).
  The pad-side suites (smoke, monaco, config, hosted, files, ux…) were not run — nothing under
  `src/` outside `src/workbench/` changed.
- Docs: CLAUDE.md file map + test counts, tests/README.md counts, HANDOFF.md section.

## Next

- [ ] Local machine: `git fetch && git checkout claude/dcspad-sp-utilities-availability-ylr25q`,
      optionally run `tests/workbench.mjs` (both servers, see tests/README.md), then deploy to
      dev with `deploy\Sync-Live.ps1` (default env; rebuilds the bundle — no `?v=` bump needed).
- [ ] Run HANDOFF.md's "Live-tenant checklist" for this section on the dev FCUPortal site; tick
      items there with the evidence, as the markdown-export section did.
- [ ] Anything the checklist breaks: fix on this branch, re-run `tests/workbench.mjs`, redeploy.
- [ ] Prod (bmo): repeat the Metadata-row and EEEU checks — the bmo site columns
      (`bmocContentCategory`, `FolderType`, `Pillar`, `Org`, `Contact`) likely exist only there.
- [ ] Open a PR when Joe asks; on merge, promote per the project-state rules and delete this file.

## Open questions

For Joe:
- Name (FileLeafRef) stayed read-only in the Metadata tab (a rename breaks links to the page).
  Should it be editable?
- The EEEU audit lists only lists/subsites with their own permissions; inheriting ones are
  covered by the parent's row. Should every exposed list be listed individually instead? (One
  change in `scanWeb`, `src/workbench/eeeu-audit.js`.)

## Landmines

- **Unverified against SharePoint:** the status select (`File/MajorVersion`, `File/CheckOutType`,
  `CheckoutUser/Title` with `$expand=File,CheckoutUser`) and the per-version moderation shape of
  `items(id)/versions`. Both were designed from docs plus a ChatGPT review; the checklist has an
  item for each. `moderationOf()` in `page-status.js` is where a new payload shape goes.
- Cloud sandbox only: its HTTPS_PROXY intercepts `localhost` from Chromium (405s, blank
  workbench) — run suites there with `env -u HTTPS_PROXY -u https_proxy -u HTTP_PROXY -u http_proxy`.
  A local machine needs nothing special.
- `sp-rest.js` now shares one three-request queue across every client (module scope). A view
  that needs its own parallelism must not "fix" that by creating more clients — that was the bug.
- `Contact` (person) and `bmocContentCategory` (managed metadata) are read-only by design until
  the User/Taxonomy editor seam (CLAUDE.md Roadmap) is built.
