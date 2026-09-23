# State — SP Workbench page status + EEEU audit

Last touched: 2026-09-23
Mode: Joe
Branch: feature merged into `main` 2026-09-23 (PR #21, `ab5a0f9`); live-test fixes on
  `claude/page-status-live-fixes` (PR open, not merged); rollback point: branch
  `rollback/main-before-page-status-eeeu` (= `db95413`, Build #197)
State: dev live-tenant checklist GREEN on the fix branch (Build #220); six defects found live and
  fixed there; waiting on the fix PR's merge, then prod (bmo)

## What this is

Two SP Workbench features. **Page status** (Pages view): a Scan Page Status button that adds
Published / Checked out / Inheritance columns, status chips and a Permissions tab in the page
drilldown, and a fixed-order Metadata tab — a rebuild of a feature built once on the work (bmo)
machine and lost. **EEEU audit** (Permissions view): finds where everyone in the organization
has access, across the site, its lists/libraries/items and chosen subsite trees. Design, Joe's
decisions, the live-tenant checklist (ticked, with evidence) and the six live fixes:
`HANDOFF.md` → "SP Workbench: page status + EEEU audit (2026-09-23)".

## Done

- Code: `src/workbench/page-status.js`, `src/workbench/eeeu-audit.js` (new);
  `views/pages.js`, `views/security.js`, `grid.js` (`setColumns`), `field-editor.js` (`layout`),
  `page-export.js` (status lines), `sp-rest.js` (one shared request queue), `mock-data.js`
  (Site Pages status fixtures + the `/sites/eeeu` audit tree), `styles/workbench.css`.
- ChatGPT review round 1: all nine findings fixed (`4cfb92e`); summary in HANDOFF.md.
- Merged to `main` as PR #21 (`ab5a0f9`, Build #215).
- Live dev-tenant checklist run 2026-09-23: first on `main` Build #215 (19/23 Pages checks),
  then all green on the fix branch Build #220 (Pages 29/29, EEEU 10/10). Six defects fixed on
  `claude/page-status-live-fixes` — version moderation key (`OData__x005f_ModerationStatus`) +
  `moderationApplies()`, person fields via the separate `/FieldValuesAsText`, hidden-field
  aliases, `getAll`'s large cap, `getAll` `shouldStop` for a prompt Cancel, the EEEU item read
  without `Title`. Mock fixtures now mirror the tenant; each fix has a check that fails without
  it. Suites: workbench 167, workbench-schema 138, workbench-edit 26, workbench-hosted 10.
- Docs: HANDOFF.md checklist ticked with evidence + the fixes list; issue #10 (new-tab links)
  root cause recorded under Open items; CLAUDE.md / tests/README.md counts (607 total).

## Next

- [ ] Merge the `claude/page-status-live-fixes` PR into `main`. **The dev tenant currently runs
      that branch's Build #220**, not `main` — after merging, redeploy `main` with
      `deploy\Sync-Live.ps1` and confirm the served stamp before any further live testing.
- [ ] Prod (bmo): repeat the Metadata-row and EEEU checks — the bmo site columns
      (`bmocContentCategory`, `FolderType`, `Pillar`, `Org`, `Contact`) exist only there; also a
      locked group for the EEEU Problems grid (not producible with a site-admin account on dev).
- [ ] When prod passes: promote per the project-state rules, delete this file and the
      `rollback/main-before-page-status-eeeu` branch.

## Open questions

For Joe:
- Name (FileLeafRef) stayed read-only in the Metadata tab (a rename breaks links to the page).
  Should it be editable?
- The EEEU audit lists only lists/subsites with their own permissions; inheriting ones are
  covered by the parent's row. Should every exposed list be listed individually instead? (One
  change in `scanWeb`, `src/workbench/eeeu-audit.js`.)
- Issue #10 (Workbench links open in the same tab): the live root cause points at a non-anchor
  control, which loses native middle/ctrl-click. Joe's call.

## Still present on `main`, queued separately (not this thread's to fix)

Found by the 2026-09-23 live test of Build #197; separate fix tasks were queued — check those
haven't landed before touching them here.
- `field-editor.js` `toFormValue` sends DateTime as ISO; live `ValidateUpdateListItem` accepts
  only the site-locale form (`10/1/2026 9:30 AM`, site time zone), and one bad field fails the
  whole save. `list-data-apply.js`'s date calibration is the likely reusable fix.
- `canvas.js` `WEBPART_NAMES`: `1ef5ed11-…` is "Markdown", not "Code snippet".
- `canvas.js`: `controlType: 14` (section background image) is reported as a parse error.
- Pages grid lists folders (`SitePages/tools`) as pages, with MD/HTML export buttons; the Scan
  columns on such a row are blank.

## Landmines

- **Versions ≠ items for moderation.** Items answer `OData__ModerationStatus`; versions omit it
  under that name (200, silently) and answer only `OData__x005f_ModerationStatus`. The
  `_ModerationStatus` field exists on every pages library, approval on or off — only
  `EnableModeration` says approval is on.
- **Inline `$expand=FieldValuesAsText` returns person-field ids**, not names; use the separate
  `items(id)/FieldValuesAsText` endpoint for display text.
- `sp-rest.js` shares one three-request queue across every client (module scope). A view that
  needs its own parallelism must not "fix" that by creating more clients — that was the bug.
  `getAll` with `allowLargeCap` and no `cap` now reads up to 100,000 items: pass `shouldStop`
  from anything cancellable.
- `Contact` (person) and `bmocContentCategory` (managed metadata) are read-only by design until
  the User/Taxonomy editor seam (CLAUDE.md Roadmap) is built.
- Dev-tenant fixtures: `zz-mod-pages` (119, content approval, `zz-mod-page.aspx` at approved 1.0
  + pending 1.1) is a KEEPER for approval checks. The EEEU fixtures were removed; recreating
  them grants org-wide access, so ask Joe first. The tenant's EEEU claim is
  `c:0-.f|rolemanager|spo-grid-all-users/29083078-4f8e-40bc-8b08-c5819bab3733` — not the Power
  Platform environment GUID.
