# State — SP Workbench page status + EEEU audit

Last touched: 2026-09-23
Mode: Joe
Branch: merged into `main` 2026-09-23 (from claude/dcspad-sp-utilities-availability-ylr25q);
  rollback point: branch `rollback/main-before-page-status-eeeu` (= `db95413`, Build #197)
State: on `main`, mock-tested and review-round-1 fixed; not yet deployed or tried on a live tenant

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

- [ ] Local machine: `git checkout main && git pull`, optionally run `tests/workbench.mjs`,
      then deploy to dev with `deploy\Sync-Live.ps1` (default env; rebuilds the bundle — no
      `?v=` bump). Before testing, confirm the served `dcspad.workbench.js` stamps a build newer
      than #197 (the pre-merge build) — a first local attempt on 2026-09-23 tested #197 and so
      tested none of this work. To roll back: `git checkout rollback/main-before-page-status-eeeu`
      and redeploy.
- [ ] Run HANDOFF.md's "Live-tenant checklist" for this section on the dev FCUPortal site; tick
      items there with the evidence, as the markdown-export section did.
- [ ] Anything the checklist breaks: fix on this branch, re-run `tests/workbench.mjs`, redeploy.
- [ ] Prod (bmo): repeat the Metadata-row and EEEU checks — the bmo site columns
      (`bmocContentCategory`, `FolderType`, `Pillar`, `Org`, `Contact`) likely exist only there.
- [ ] When the checklist passes (dev, then prod): promote per the project-state rules, delete this
      file and the `rollback/main-before-page-status-eeeu` branch.

## Open questions

For Joe:
- Name (FileLeafRef) stayed read-only in the Metadata tab (a rename breaks links to the page).
  Should it be editable?
- The EEEU audit lists only lists/subsites with their own permissions; inheriting ones are
  covered by the parent's row. Should every exposed list be listed individually instead? (One
  change in `scanWeb`, `src/workbench/eeeu-audit.js`.)

## Found live on the pre-merge `main` (2026-09-23), still present

A local session tested `main` (Build #197) on the dev tenant by mistake. It found four defects
in code this work did not change, so they are still on `main`. It queued separate fix tasks for
them — check those haven't landed before fixing here.
- `field-editor.js` `toFormValue` sends DateTime as ISO; live `ValidateUpdateListItem` rejects
  every ISO form and accepts only the site-locale form (`10/1/2026 9:30 AM`, site time zone), and
  one bad field fails the whole save. Hits the Files editor; the Pages Metadata tab edits no
  DateTime field. `list-data-apply.js`'s date calibration is the likely reusable fix.
- `canvas.js` `WEBPART_NAMES`: `1ef5ed11-…` is "Markdown", not "Code snippet".
- `canvas.js`: `controlType: 14` (section background image) is reported as a parse error.
- Pages grid lists folders (`SitePages/tools`) as pages, with MD/HTML export buttons. The Scan
  columns on such a row are blank; the fix (filter `FSObjType eq 1` rows out, or
  show them as folders) belongs in `views/pages.js` either way.

## Landmines

- **Unverified against SharePoint:** the status select (`File/MajorVersion`, `File/CheckOutType`,
  `CheckoutUser/Title` with `$expand=File,CheckoutUser`) and the per-version moderation shape of
  `items(id)/versions`. Both were designed from docs plus a ChatGPT review; the checklist has an
  item for each. `moderationOf()` in `page-status.js` is where a new payload shape goes.
- `sp-rest.js` now shares one three-request queue across every client (module scope). A view
  that needs its own parallelism must not "fix" that by creating more clients — that was the bug.
- `Contact` (person) and `bmocContentCategory` (managed metadata) are read-only by design until
  the User/Taxonomy editor seam (CLAUDE.md Roadmap) is built.
