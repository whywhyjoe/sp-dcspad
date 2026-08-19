Results:
Found six concrete defects. The first two can display—and potentially edit—the wrong page. No fixes were made.

## Confirmed defects

1. **High — Reloading a secondary-library detail opens the same ID from the default library.**  
   [pages.js](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:378) stores only pageId and pageName in the route; [showDetail()](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:685) resolves that ID against the newly defaulted current library.

   Trigger:

   - Select classic Pages.
   - Open Policies.aspx (ID 1).
   - Reload while the detail route is saved.

   Reproduced result:

   - Before: title Policies.aspx, path /Pages/Policies.aspx.
   - After: title still Policies.aspx, but path /SitePages/TeamNews.aspx, kind modern canvas page.

   Opening Metadata after reload would target the modern item while the header still names the classic page. The selected library must be part of the route; keeping only closure-level current is not safe.

2. **High — A completed request for the old library can overwrite the new grid.**  
   [loadPages()](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:321) has no run/generation guard. After awaiting the request at [line 389](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:389), it writes through the shared grid variable at line 390. switchLibrary() removes/reassigns that same variable.

   Trigger:

   - Start loading modern Site Pages.
   - Switch to classic Pages while the modern /items request is pending.
   - Let classic finish first, then modern.

   Reproduced result: the chip continued to say classic publishing Pages library, while its table changed from Classic.aspx to the stale StaleModern.aspx.

   A stale rejection can likewise call setError() on the replacement grid. Rapid switches during field probing have the same generation problem.

3. **Medium — A failed schema probe reintroduces the exact live-only 400 it was meant to prevent.**  
   [queryPlan()](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:410) catches every field-probe failure and uses pageQueryPlan(null, 'modern') at line 414.

   Trigger: a legacy BaseTemplate 119 library lacks PromotedState, while /fields fails because of a transient/network/permission error. The fallback assumes modern schema and requests PromotedState, causing the subsequent item query to fail instead of degrading safely.

4. **Medium — Breadcrumb clipping is not merely cosmetic.**  
   [renderCrumbs()](/C:/dev/repos/sp-dcspad/src/workbench/views/browser.js:174) computes scrollLeft and is-clipped only at lines 207–208. Adoption and resizing can change its width afterward.

   Reproduced by opening /Shared Documents/Reports at 1400 px and resizing to 720 px:

   - clientWidth: 210 → 0
   - scrollWidth: 210
   - scrollLeft: remained 0
   - is-clipped: remained false
   - deepest Reports crumb: no longer visible

   The hidden scrollbar means pointer users lose the breadcrumb navigation until another render or non-obvious horizontal/keyboard scrolling.

5. **Low — Every Pages library switch leaks grid-level document listeners.**  
   Each grid menu installs an anonymous document click listener at [grid.js:125](/C:/dev/repos/sp-dcspad/src/workbench/grid.js:125). switchLibrary() removes the grid element at [pages.js:296](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:296), but there is no grid disposal API to remove those listeners.

   Each rebuild adds two stale listeners—Copy as and Export—and keeps their detached menus reachable. The adopted strip itself is safe: it has one parent, and the persistent Open-link listener is not duplicated.

6. **Low — An old rejected detail request can evict the new library’s cache entry.**  
   The rejection handler at [pages.js:423](/C:/dev/repos/sp-dcspad/src/workbench/views/pages.js:423) unconditionally deletes by page ID.

   Trigger: start library A item ID 1, go back, switch, start library B item ID 1, then let A reject. A’s handler deletes B’s promise. detailRun prevents stale rendering, but subsequent visits refetch and can create duplicate concurrent requests. Clearing the map is therefore not fully race-safe.

## Test-quality audit

Baseline at exact commit 3a73905:

- workbench.mjs: 113/113
- workbench-edit.mjs: 19/19
- workbench-hosted.mjs: 10/10

Useful mutation results:

| Test | Deliberate break | Result |
|---|---|---|
| both: switching library re-probes the query plan | Removed planPromise = null | Failed, along with the grid-shape check |
| both: the drilldown follows the switched library | Made drilldown use the default library | Failed |
| both: switching library rebuilds the grid for the other shape | Always rendered the Promoted column | Failed |
| both: picker replaces… | Removed template numbers from options | Failed |
| info chip… | Removed .wb-info-chip from the library chip | Failed |
| Files toolbar-row check | Removed toolbarExtras: bar | Failed |

Three checks are incomplete:

- **Ranking test:** I moved visible title-only Pages below hidden 119/850 libraries. This violates “visibility outranks template,” but every one of the 113 checks still passed. The uncovered cases are visible generic Pages versus hidden 119 or hidden 850.
- **Info-chip test:** I removed the required hairline border. the library kind is an info chip… still passed because it never asserts the border despite saying it does.
- **Toolbar-row test:** I enabled toolbar wrapping. All 19 edit checks still passed at their 1400 px viewport; at 520 px the filter, location bar, and actions occupied three separate rows. Its geometry assertion is meaningful only at that one generous width.

None of the existing tests delays the old library request past the new one, so the confirmed stale-grid race remains entirely uncovered.

## Ranking verification

The rewrite’s actual ordering is:

| Shape | New rank |
|---|---:|
| Visible 119 | 1 |
| Visible 850 | 2 |
| Visible title-only Pages | 3 |
| Hidden 119 | 4 |
| Hidden 850 | 5 |
| Hidden title-only 101 | not a candidate |

Input order is preserved within a rank, and deduplication is safe for real SharePoint list IDs. Compared with the old ladder, differences are limited to the intended visibility cases:

- visible 850 now beats hidden 119;
- visible title-only Pages now beats hidden 119;
- visible title-only Pages now beats hidden 850.

The implementation matches the stated ranking; its tests do not cover the last two comparisons.

## Other conclusions

- webInfoPromise is library-independent. A site switch creates a new view instance/client, so it does not need clearing on a library switch.
- The webPartCache key contains the server-relative file path and is effectively library-qualified.
- The late webPartCache/planPromise declarations are not currently a TDZ failure because switchLibrary() cannot execute until after construction returns. I would still move all state declarations above callbacks; the present ordering obscures the lifecycle.
- The info/status register split is a sound design decision.
- I found no unsanitized new HTML-to-DOM path: classic body/web-part fragments reach sanitizeHtml() before innerHTML; raw HTML is placed in text/JSON output.
- At 760–1400 px, the global nonshrinking action rules did not clip actions in Lists, Permissions, Site, Advanced, Query, Pages, or Files. Panels has no grid. The confirmed narrow-layout problem is the Files breadcrumb lifecycle above.
- Validation was offline Playwright/mock/custom-client testing, not live-tenant proof. The SharePoint 400 scenarios are code-path analysis consistent with the supplied live behavior, not a tenant run.