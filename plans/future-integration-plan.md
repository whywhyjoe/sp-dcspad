# Integrated DCSpad / SPWorkbench Mode Switcher

## Summary

Create a shared, full-screen mode stage used by both existing entry points.
DCSpad remains the initial screen on its URL; SPWorkbench remains initial on
its URL. The opposite screen lazy-loads on first switch, then stays mounted
for instant switching and state preservation.

Use a paired 220 ms horizontal slide: DCSpad occupies the left slot and
SPWorkbench the right. Respect `prefers-reduced-motion` by switching
immediately.

Implementation starts on a new `codex/dcspad-workbench-switcher` branch. This
planning document does not introduce implementation changes.

## Key Changes

### Shared mode host and loading

- Introduce one mode-host controller, shared across the separately bundled
  applications through a `window.__DCSPAD_MODE_HOST__` singleton.
- Give the controller responsibility for:
  - Registering the DCSpad and SPWorkbench roots.
  - Deduplicating concurrent lazy-load requests.
  - Maintaining `dcspad`, `workbench`, `loading`, and `failed` states.
  - Applying transforms, `inert`, `aria-hidden`, pointer-event suppression,
    and focus transfer.
  - Retaining both initialized roots without reloading or disposing them.
- Refactor hosted positioning so a shared stage owns the fixed SharePoint
  inset. `.app` and `.wb` become absolute, full-stage panels rather than
  competing fixed viewport owners.
- Keep both boot entry points and establish the same hosted asset-resolver
  contract in each:
  - From DCSpad, the first switch fetches `workbench.html`, versioned
    `workbench.css`, and the separate `dcspad.workbench.js` bundle.
  - From SPWorkbench, the first switch fetches the DCSpad shell, configuration,
    versioned app bundle, Monaco manifest/runtime, and intelligence manifest.
  - Standalone pages use the corresponding source modules and relative assets.
  - Shell HTML remains fetched `no-store`; hosted CSS and bundles retain the
    existing Last-Modified cache-busting protocol.
- Show a branded destination loading panel as the first paired slide begins.
  On failure, replace it with a concise error plus Retry and Back actions;
  retries use a fresh module URL and never duplicate roots or listeners.
- Make Workbench initialization await its saved-site restoration and initial
  view readiness so the lazy-load promise represents a usable screen.

### Header controls and transitions

- Move the DCSpad `#btn-sp` control to immediately after the SP chip in
  `.topbar-right`; keep it visible in both live and disconnected modes.
- Add a compact D-logo switch button at the far left of the SPWorkbench header,
  before its wordmark.
- Create one shared raised mode-button component:
  - 28 by 28 px, fixed-size, icon-only, with an accessible name and tooltip.
  - Visible border, highlight, and compact control shadow at rest.
  - Stronger border/glow and a 1 px lift on hover; include pressed and
    keyboard-focus states.
  - The SP button uses the SP live green family when connected and a
    neutral/warn disconnected variant otherwise.
  - The D button consistently uses the DCSpad accent family because DCSpad
    remains functional offline.
- Keep the switch controls from shrinking. Let the project-name area and
  Workbench site field truncate first so headers remain usable at narrower
  desktop widths.
- After a completed switch, focus the destination screen's mode button. The
  inactive root remains mounted but becomes inert and hidden from assistive
  technology.
- Do not modify the URL, browser history, or persisted settings when switching.
  Refreshing either entry URL restores that entry's native initial mode.

### Shared SharePoint chip and chrome behavior

- Change the Workbench chip from a passive `<span>` to the same button structure
  used by DCSpad, including dot, text, chevron, title, ARIA state, live/mock
  variants, and focus behavior.
- Replace per-screen toolbar toggles with one shared SharePoint-chrome
  controller that registers both chips and synchronizes:
  - `SP` versus `SP: Mock` text.
  - Live/mock color and connection description.
  - Show/Hide SharePoint toolbar title, `aria-label`, and `aria-expanded`.
  - The root `dcspad-chrome-hidden` state and shared stage inset.
- Preserve Workbench's detailed status-bar text and inspected-site information
  separately from the host-connection chip.
- Update suspension rules so either hosted mount and both loaded panels
  disappear during SharePoint page-edit mode, while the suite bar and host
  scrolling are restored.

### Disconnected Workbench behavior

- Remove `mockResolver` and fabricated SharePoint records from the production
  Workbench initialization path.
- In disconnected mode:
  - Keep the header, navigation rail, and section choices visible.
  - Render a reusable "Open through SharePoint to inspect site data" empty
    state within each selected view.
  - Disable the site selector, Inspect button, exports, writes, and other
    data-dependent actions.
  - Make no `/_api` requests.
  - Keep the DCSpad/SPWorkbench mode switches available.
- Retain mock-data fixtures only for isolated tests; they must not be imported
  into the production Workbench bundle.

## Interfaces and Build Outputs

- No user-facing API or stored-document schema changes.
- Add internal contracts for `registerMode`, `loadMode`, `switchMode`, and
  hosted asset resolution through the mode-host singleton.
- Make the SharePoint chrome controller register multiple chip elements rather
  than assuming a single `#sp-chip`.
- Continue generating `dcspad.app.js` and `dcspad.workbench.js` as separate
  cache-bustable bundles.
- Rebuild both bundles, update boot-script query versions where required, and
  document the dual-entry/lazy-secondary architecture in the repository
  guidance.

## Test Plan

- Add mode-switch tests starting from both `index.html` and `workbench.html`:
  - Correct initial mode.
  - Correct slide direction in both directions.
  - Secondary assets load only after the first switch and only once.
  - Monaco/workspace, Workbench route, filters, and inspected-site state survive
    repeated switching.
  - No URL or history mutation.
  - Loading failure, Retry, rapid repeated clicks, and partial-load cleanup.
- Verify accessibility:
  - The inactive screen is inert and `aria-hidden`.
  - Focus transfers after switching.
  - Both controls have stable accessible names and visible focus rings.
  - Reduced-motion mode performs an immediate switch.
- Replace production-mock expectations with disconnected empty-state
  assertions:
  - No fabricated rows.
  - No REST calls.
  - Site/data actions are disabled.
  - Switching remains available.
- Extend hosted fixtures to verify:
  - Shared SharePoint toolbar state remains synchronized across both chips.
  - Hiding the suite bar moves the shared stage once without exposing host-page
    backgrounds.
  - SPA entry into and exit from page-edit mode suspends/restores both modes.
  - Each entry point lazy-loads the opposite versioned bundle correctly.
- Perform visual checks at 1500 by 900 and 1024 by 768 for header crowding,
  clipping, panel shadows, transition edges, loading/error states, and both
  live/disconnected button variants.
- Complete live SharePoint validation from both deployed entry pages, including
  suite-bar hiding, first-load timing, REST context, state retention, and
  absence of duplicate app roots or console errors.

## Assumptions

- DCSpad is always the logical left panel and SPWorkbench the right panel,
  regardless of which entry URL loaded first.
- A screen is initialized at most once per page lifetime and retained
  thereafter.
- "Mock is essentially empty" applies to SPWorkbench data; DCSpad's existing
  offline preview context remains unchanged.
- The existing standalone pages and both SharePoint-hosted entry pages remain
  supported.
