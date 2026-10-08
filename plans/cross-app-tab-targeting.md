# Cross-app tab targeting — DCSPad ⇄ SP Workbench

**Status:** findings only, nothing implemented. Filed 2026-08-19 from a
question about the pad's SPWorkbench button landing in the same tab.

**Goal.** The pad's SPWorkbench button should open the workbench in a
*targeted* tab (reuse it if it exists, create it otherwise), and the workbench
topbar should carry the mirror-image button back to the pad.

Verdict: **possible**, and the pad half is already coded — the reported
same-tab behaviour is a `window.name` collision, not a missing feature.

---

## 1. What the pad does today

`src/main.js:135-147`, added by `087a922` ("Refine topbar UX and docs actions")
together with the `#btn-sp` markup in `index.html`:

```js
const btnSp = document.getElementById('btn-sp');
if (btnSp) {
  if (!initialSpContext.live) {
    btnSp.hidden = true;
  } else {
    btnSp.hidden = false;
    btnSp.addEventListener('click', () => {
      const url = configResult.config?.workbench?.url || initialSpContext.webAbsoluteUrl || '/';
      const opened = window.open(url, 'dcspad-sp');
      opened?.focus?.();
    });
  }
}
```

That is already a **named target**, not a same-tab navigation. The URL comes
from `workbench.url` in `dcspad.config.json` (`src/config.js:152-157`
`normalizeWorkbench`, resolved against `siteUrl`), which `deploy/Sync-Live.ps1`
rewrites per environment from `workbenchPageUrl` in `deploy.settings.json`
(lines 118-150, 389-402).

## 2. Why it lands in the same tab anyway

**`window.name` survives navigation.** The first click opens a new tab and
names it `dcspad-sp`. Navigate *that* tab back to the pad — bookmark,
breadcrumb, back button, SharePoint nav — and the tab keeps the name. From
then on the pad is itself running in a window named `dcspad-sp`, so
`window.open(url, 'dcspad-sp')` resolves to the current window. Same tab,
deterministically, for the life of that tab. Opening a fresh pad tab clears it,
which is exactly why the bug reads as intermittent.

Ruled out:

- **Stale hosted bundle.** The button and the named target shipped in the same
  commit, so a pre-`087a922` `dcspad.app.js` would hide the button entirely
  rather than change where it opens.
- **Popup blocking.** That returns `null` from `window.open` — "nothing
  happens", not "same tab".

Fix: claim a name for the pad's own window at boot (`window.name =
'dcspad-pad'`) so it can never be the target of `dcspad-sp`, and guard the
click path against a name collision it did not set.

## 3. What the workbench→pad direction needs

Three pieces, none of them large:

### 3.1 Button

Mirror `#btn-sp` into `workbench.html`'s `.wb-topbar-right`, beside `#wb-chip`.
Same ghost-button treatment. Note the register rule in `design/INFO-CHIP.md`:
this is a navigation control, not a chip — it must not borrow status colour.

### 3.2 A pad URL — the workbench loads no config today

`boot-workbench.js:87` versions exactly three files:

```js
var VERSIONED = [
  'styles/app.css',
  'styles/workbench.css',
  'dcspad.workbench.js',
];
```

and its header comment states the intent plainly: *"no intelligence, no config
— the workbench's asset set is three files."* Nothing in `src/workbench/` calls
`loadAppConfig()`. So a configured pad URL is a new dependency, not a lookup.

Two sources, and both are worth having:

- **Durable — config.** Add `pad.url` to `dcspad.config.json` and
  `padPageUrl` to `deploy/deploy.settings.json`, mirroring the existing
  `workbench` / `workbenchPageUrl` pair; `Sync-Live.ps1`'s rewrite and
  validation lists already handle that shape, so it is a symmetric edit.
  Cost: `dcspad.config.json` joins the workbench's `VERSIONED` list, and
  `boot-workbench.js` gains a config fetch — which means bumping `?v=` in
  `workbench.webpart.html` (boot sits below the versioning layer).
- **Free — opener handoff.** Have the pad pass its own `location.href` when
  it opens the workbench (query param, or a `sessionStorage` key). Exact even
  if config drifts, and it covers the case where the two pages live in
  different site collections — which they do today: the pad deploys under
  `…/FCUPortal/Dev/tools/dcspad`, while the configured workbench page is
  `…/sites/NewNerve/SitePages/tools/SPWorkbench.aspx`.

Config is the cold-start answer (someone opens the workbench directly);
the handoff is the accurate one. Use the handoff when present, config as
fallback.

### 3.3 Return semantics — the one real limitation

`window.open(padUrl, 'dcspad-pad')` reuses an existing named tab **only if the
two tabs are in the same browsing context group** — i.e. one opened the other,
and the link was not `noopener`. Consequences:

- A pad tab the user opened by hand (bookmark, typed URL) is **not** reachable
  by name from a separately-opened workbench tab. That click opens a second
  pad tab. There is no browser API that fixes this; it is the security model.
- So: prefer a live, same-origin `window.opener` when there is one, and fall
  back to the named open. Do **not** add `noopener` to either of these two
  links — that severs the group and permanently breaks the pairing. (Contrast
  `src/workbench/grid.js:34-43`, which deliberately *does* use `noopener` for
  untrusted content links; different job, keep it as is.)
- Cross-tab `focus()` from a click gesture is honoured in Chrome/Edge — the
  tenant browsers. Firefox blocks it by default
  (`dom.disable_window_flip`), so there the click re-navigates the pad tab
  without visually switching to it. Acceptable; worth knowing before someone
  files it as a bug.

## 4. Hosting and test cost

- Any `src/` change → rebuild **both** bundles (`dcspad.app.js`,
  `dcspad.workbench.js`); `deploy/Sync-Live.ps1` does both.
- A `boot-workbench.js` change → bump `?v=` in `workbench.webpart.html`.
  A `boot.js` change → bump `?v=` in `dcspad.webpart.html`.
- Test homes: `tests/hosted.mjs` (pad side, window naming) and
  `tests/workbench-hosted.mjs` (workbench side, button + URL resolution).
  Both already drive the hosted boot path.

## 5. Suggested shape, if this gets picked up

1. Name both windows at boot: `dcspad-pad` and `dcspad-sp`, each claimed by
   its own entry point, each refusing to target a name it did not set.
2. Pad button: keep the named open, add the `location.href` handoff.
3. Workbench button: opener-first, then handoff URL, then configured
   `pad.url`, then nothing (hide the button, same as `#btn-sp` hides when the
   SP context is not live).
4. Pin both directions in the two hosted suites, including the
   "pad window is misnamed `dcspad-sp`" regression that started this.
