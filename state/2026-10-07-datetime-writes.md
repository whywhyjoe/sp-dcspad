# State — SP Workbench DateTime writes + REST getPage

Last touched: 2026-10-07
Mode: Joe
Branch: `main`, pushed (`c721329` merged as `7519c31`; review fixes `120ddbb`)
State: code done and on `main`, 778/778 tests green; NOT yet tested on a live tenant and NOT
  deployed (bundles not rebuilt)

## What this is

Two defects found live on the dev tenant (nervedotnet, NewNerve site) by sp-traffic-analytics,
which copies this repo's REST clients. **1.** The metadata editors (Pages metadata tab, Files
browser metadata panel) sent DateTime FieldValues to ValidateUpdateListItem as ISO 8601,
which SharePoint Online refuses; it accepts only the web's locale format on the web's clock.
`toFormValue` now formats from the web's RegionalSettings and its UTC offset at that instant
(`src/workbench/web-dates.js`). **2.** `getAll` treats `top` as a page size, so a `top: 1`
lookup walked ~5,000 one-row pages; `getPage` (one request) was added and `getAll` left as is.
The rules are in CLAUDE.md → Gotchas ("ValidateUpdateListItem refuses ISO 8601 dates…" and
"The 5,000-row threshold applies to the LEADING filter condition").

## Done

- Two Codex review rounds via xo (sp-dcspad turns 18 and 19). Everything raised is fixed in
  `120ddbb` except one accepted risk: the repeated-hour check's ±3h window can miss a
  rollback longer than 3h (none in use today; comment in `web-dates.js`).
- sp-traffic-analytics re-copied `sp-rest.js` (its `af4d895`). It doesn't copy
  `field-editor.js` or `web-dates.js`, so the date changes don't reach it.

## Next

- [ ] Rebuild and deploy: `deploy\Sync-Live.ps1` rebuilds `dcspad.app.js` and
      `dcspad.workbench.js` from `src/`, which this work changed. Confirm the served build
      stamp before testing.
- [ ] Live test on the dev tenant, Files browser and Pages metadata tab:
  - [ ] Save a DateTime field; check the stored value in SharePoint matches the time picked.
        The posted FieldValue should be like `10/6/2026 9:00 AM`, never ISO.
  - [ ] `/_api/web/RegionalSettings` (no `$select`) returns `TimeMarkerPosition`. If it
        doesn't, every date save on a 12-hour web is now refused (`webDateFormatOf` in
        `web-dates.js` requires it); fall back to the 24-hour format or the locale data.
  - [ ] A date-only field (DisplayFormat 0) shows the stored day, including from a browser
        in a zone west of the web, and saves the picked day unchanged.
  - [ ] If a non-US test web is available (day-first, '.' separator, or a 12-hour locale
        with the marker before the time): SPO accepts the produced string. Untested: only the
        string building is covered by tests, not SPO's parser.
- [ ] When green: promote the gotchas (already in CLAUDE.md) and delete this file.

## Landmines

- **Never add a default to the date path.** `toFormValue` and `webDateFormatOf` throw on a
  missing format part or offset on purpose: a default writes a wrong date silently (a
  day-first web would store June 10 for October 6). The mock web's RegionalSettings
  (`mock-data.js` `REGIONAL_SETTINGS`) must stay complete or every mock date save fails.
- **Page copy cannot fill a required DateTime gap** (`page-copy.js` `computeCarrySet` marks it
  unfillable): the pure planner has no destination regional settings. Fixing that means the
  copy dialog resolving them first, not reviving `toFormValue` without `webDate`.
- **`getAll`'s `top` is a page size.** The EEEU item scan and the list-data capture pass
  `top: 5000` and need pages followed past 5,000 rows. Don't make it stop at `top`; use
  `getPage` or `cap` for a fixed number of rows.
