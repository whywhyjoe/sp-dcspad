# DCSPad post-Monaco UI integration contract

This is the implementation handoff for applying a UI design system that was
created against the pre-Monaco workbench. It documents the changes between the
last pre-Monaco revision and the current `monaco-plus-design-system` work, with
special attention to DOM hooks, settings, persistence, preview/runtime
configuration, and code ownership.

## Baseline and scope

- **Exact pre-Monaco baseline:** commit `3414854` (`screenshots`)
- **First Monaco commit:** `c925e10` (`initial monaco integration`)
- **Current line of work:** branch `monaco-plus-design-system`
- **Comparison command:** `git diff 3414854 --`

Using `main` as the comparison point is misleading because that branch already
contains Monaco-related work. Commit `3414854` is the direct parent of the
initial Monaco commit and is therefore the authoritative visual and wiring
baseline.

This document is both:

1. a migration record of what changed after `3414854`; and
2. a UI integration contract describing which IDs, classes, states, and
   behaviors a workbench design system must preserve.

## Executive compatibility summary

The workbench was not re-architected around Monaco. The following shell
structure remains in place and can keep the existing design-system treatment:

- top bar, File menu, Settings menu, Auto-run toggle, Run button, and SP chip;
- Resources sidebar, framework/snippet sections, and row actions;
- HTML/CSS/JS tab strip;
- Preview and Console/Network panels;
- vertical/horizontal splitters, collapse controls, and maximize controls;
- status bar, save state, cursor state, and SharePoint context state;
- fresh `srcdoc` preview iframe, console/network bridge, and inspector;
- localStorage workspace/catalog/snippet persistence.

The important visual integration changes are limited and explicit:

1. three CodeMirror mount elements became one Monaco mount element;
2. the status bar gained a persistent Monaco health item;
3. the splash gained readiness text and hosted/standalone transition states;
4. snippet naming moved from `prompt()` to a native in-app `<dialog>`;
5. Monaco creates suggestion, hover, parameter-hint, find, and diagnostic UI
   that must remain legible above the workbench chrome.

## DOM and UI wiring changes

The shell source of truth remains [`index.html`](../index.html). Hosted mode
fetches that same file and injects its body through [`boot.js`](../boot.js).

### Editor mount: three elements became one

Pre-Monaco:

```html
<div class="editor-pane active" id="pane-html"></div>
<div class="editor-pane" id="pane-css"></div>
<div class="editor-pane" id="pane-js"></div>
```

Current:

```html
<div class="editor-pane active" id="pane-editor"></div>
```

The tab contract is unchanged:

```html
<div id="editor-tabs">
  <button class="tab" data-editor="html">HTML</button>
  <button class="tab" data-editor="css">CSS</button>
  <button class="tab" data-editor="js">JS</button>
</div>
```

[`src/layout.js`](../src/layout.js) still owns active tab styling and persists
`layout.editorTab`, but it no longer shows/hides three `.editor-pane` elements.
It calls the editor adapter's `activate(name)` callback instead.

[`src/editors.js`](../src/editors.js) owns:

- one Monaco editor attached to `#pane-editor`;
- three persistent models with URIs `file:///dcspad/index.html`,
  `file:///dcspad/styles.css`, and `file:///dcspad/script.js`;
- model switching when a tab is selected;
- separate undo history, selection, scroll/fold view state, and cursor state
  for each document;
- `ResizeObserver` plus Monaco `automaticLayout`.

**Design-system rule:** preserve `#pane-editor` as a stable, measurable element
that fills `.editor-host`. Do not recreate it during tab changes. Style Monaco
through its theme and focused integration selectors; do not restore `.cm-*`
selectors.

The current shell integration selectors are:

```css
.editor-host
.editor-pane
.editor-pane .monaco-editor
.editor-pane .monaco-editor .overflow-guard
```

Monaco's own stylesheet is loaded separately by
[`src/monaco-runtime.js`](../src/monaco-runtime.js) as
`#dcspad-monaco-style`. Syntax/editor colors are defined by the
`dcspad-dark` theme inside `initEditors()`, not solely by `styles/app.css`.

### Monaco status item

The status bar gained:

```html
<span id="status-editor"
      class="status-item status-editor">Monaco …</span>
```

[`src/monaco-runtime.js`](../src/monaco-runtime.js) owns its states:

| Text | Class/state | Meaning |
| --- | --- | --- |
| `Monaco …` | base | runtime is loading |
| `Monaco ✓` | `.ok` | editor runtime loaded |
| `Monaco ⚠` | `.warning` | a language worker failed; editing remains available but language tools are limited |
| `Monaco ✕` | `.warning` | the editor runtime itself failed |

Worker warnings deliberately remain in `#status-editor`; `#status-run` is
transient and changes every time code runs. A redesign may restyle or replace
the glyph treatment, but it must keep editor health separate from run status
and preserve the warning title/tooltip.

### Splash and first-paint contract

The splash now contains a live status element:

```html
<div id="splash" class="splash" hidden>
  <pre id="splash-logo" class="splash-logo"></pre>
  <div id="splash-status"
       class="splash-status"
       role="status"
       aria-live="polite">Loading workbench…</div>
</div>
```

Standalone lifecycle is implemented in
[`src/splash.js`](../src/splash.js). Hosted lifecycle starts earlier in
[`boot.js`](../boot.js), immediately after the SharePoint mount is available
and before the shell, config, or Monaco assets are fetched. `src/splash.js`
adopts the hosted controller through `window.__DCSPAD_BOOT_SPLASH__`.

State/class contract:

| State | Hook | Required effect |
| --- | --- | --- |
| initial | `[hidden]` or opacity `0` | not interactive; gives the browser a committed entrance state |
| entering/visible | `.visible` | opacity `1` |
| leaving | `.fading` | opacity `0`, no pointer events |
| boot failure | `.failed` | curtain remains and status becomes an error |
| hosted early curtain | `.dcspad-boot-splash` | dependency-free boot styling before `app.css` exists |
| crossfade window | `html.dcspad-crossfading` | marks the reveal interval for hosted integration/tests |

The entrance and exit opacity transition is **0.75 seconds on every visit**.
The readiness hold is separate: normally 700 ms on a first visit and 150 ms
after `settings.seenSplash` is true. Reduced-motion mode removes transitions
and minimum hold time.

The completed workbench is fully painted beneath the curtain before the exit
begins. Hosted mode also keeps a dark `#dcspad-mount::before` underlay. Do not
add a competing app-opacity transition: fading both the app and curtain is
what previously exposed a white SharePoint frame between them.

Preserve the global rule:

```css
[hidden] { display: none !important; }
```

Several menus/panels have their own display rules, so removing this guard can
make hidden overlays or tools remain interactive.

### Snippet naming dialog

Pre-Monaco snippet naming used native `prompt()`. Embedded SharePoint/browser
contexts can suppress native prompts, so the current shell has a real dialog:

```text
#snippet-name-dialog
  #snippet-name-form
    #snippet-name-title
    #snippet-name-close
    #snippet-name-context
    #snippet-name-input
    #snippet-name-cancel
    #snippet-name-save
```

[`src/snippets.js`](../src/snippets.js) owns `showModal()`, focus timing,
submit, cancel, and close behavior. The contextual copy distinguishes saving
the current selection from saving the whole active pane.

The visual classes are:

```text
.app-dialog
.app-dialog__panel
.app-dialog__head
.app-dialog__context
.app-dialog__field
.app-dialog__actions
```

A design system may restyle these classes. Preserve the IDs, native `<dialog>`
behavior, `<form>` submit semantics, label/input association, `required`
validation, and `aria-labelledby`/`aria-describedby` links.

## Editor adapter contract

The application still talks to the editor through the small adapter returned
by `initEditors()` in [`src/editors.js`](../src/editors.js):

| Method | Consumers/purpose |
| --- | --- |
| `activate(name)` | HTML/CSS/JS tab switching |
| `getDocs()` | run, save/export, snippets |
| `focus(name)` | return focus to a document |
| `setDocs(docs)` | project import |
| `getSelection(name)` | snippet creation |
| `insertAtCursor(name, text)` | snippet insertion |
| `gotoJsLine(line)` | clickable stack frames |
| `setJsAsModule(enabled)` | align Monaco diagnostics with execution mode |
| `setIntelligencePacks(ids)` | compose PnPjs, Alpine, BSP, and Fluent language data |
| `setPnpTypesEnabled(enabled)` | compatibility/direct type toggle seam |
| `dispose()` | release editor/models/providers |

This isolation is why runner, persistence, snippets, console, and preview did
not need an editor-specific rewrite.

Programmatic project replacement and snippet insertion explicitly bracket
Monaco undo stack elements. One Undo reverses a snippet insertion without also
removing the user's preceding typing.

`Ctrl/Cmd+Enter` is registered as Monaco action `dcspad.run`. The UI can display
the shortcut differently, but the command should remain available inside all
three models.

## Settings: complete functional reference

There are three distinct configuration surfaces:

1. **workspace settings** in `state.settings`, persisted for the current
   browser workspace;
2. **persisted layout** in `state.layout`, changed directly by tabs,
   splitters, collapse, and maximize-related controls;
3. **deployment/runtime configuration** in `dcspad.config.json`, edited by a
   maintainer rather than through the Settings menu.

Do not merge these concepts in the UI without also changing their persistence
and export semantics.

### User/workspace settings

Defaults and storage ownership are in [`src/state.js`](../src/state.js).
Wiring is in [`src/main.js`](../src/main.js), with execution behavior in
[`src/runner.js`](../src/runner.js) and splash behavior in
[`src/splash.js`](../src/splash.js).

| State key | UI hook | Default | Function | Persistence/export |
| --- | --- | ---: | --- | --- |
| `settings.autorun` | `#chk-autorun`, `#live-dot` | `false` | Runs about 800 ms after the last editor/framework change. Enabling it schedules an immediate debounced run. | Workspace localStorage; not included in project export |
| `settings.jsAsModule` | `#chk-module` | `false` | Runs user JS as `<script type="module">`, enabling top-level `await` and module strictness. Also changes Monaco module detection so diagnostics match the runner. | Workspace localStorage and project JSON (`jsAsModule`) |
| `settings.autoClearConsole` | `#chk-autoclear` | `true` | Clears console before a run. If false, previous output remains and a numbered run divider is inserted. | Workspace localStorage; not included in project export |
| `settings.diagFontSize` | `#btn-diag-font-dec`, `#diag-font-val`, `#btn-diag-font-inc` | `12` | Sets root CSS variable `--diag-fs`; range 10–18 px in 1 px steps for Console/Network text. | Workspace localStorage; not included in project export |
| `settings.previewDark` | `#btn-preview-theme`, `#preview-host.dark` | `true` | Toggles a pad-only preview canvas color and reruns an existing preview. The style is injected before library/user CSS, so user CSS wins. | Workspace localStorage; not included in project export |
| `settings.seenSplash` | no direct control | `false` | Internal visit marker. First visit gets a 700 ms minimum readiness hold; later visits get 150 ms. It does not change the 0.75 s fade duration. | Workspace localStorage; not included in project export |

The visible Settings dropdown itself is unchanged from the pre-Monaco shell:

- Run JS as module;
- Clear console on each run;
- Console text size.

The module setting gained an additional responsibility after Monaco:
`editorsApi.setJsAsModule()` updates JavaScript language-service behavior as
well as preview execution.

Auto-run remains in the top bar, and preview theme remains in the Preview
panel header. They are settings even though they are not inside the gear menu.

### Persisted layout settings

These are also stored in the workspace through [`src/state.js`](../src/state.js)
and applied by [`src/layout.js`](../src/layout.js):

| State key | Default | Changed by |
| --- | ---: | --- |
| `layout.sidebarW` | `230` px | left splitter; clamped to 140–420 px |
| `layout.sidebarCollapsed` | `false` | Resources collapse/expand buttons |
| `layout.editorsFr` | `1` | center splitter |
| `layout.runtimeFr` | `1` | center splitter |
| `layout.previewFr` | `1` | restored as `--preview-h`; currently not changed by the runtime splitter |
| `layout.diagH` | `260` px | Preview/diagnostics horizontal splitter |
| `layout.diagCollapsed` | `false` | diagnostics collapse button or tab reopen |
| `layout.editorTab` | `"js"` | HTML/CSS/JS tab selection |
| `layout.diagTab` | `"console"` | Console/Network tab selection |

Preview/diagnostics maximize states are transient classes on `#main`
(`.max-preview` and `.max-diag`) and are intentionally not persisted. Escape
removes both.

### Controls that look like settings but are transient

The following are not stored and reset on reload:

- Console text filter and log/warn/error level buttons;
- Network `_api only` checkbox;
- current maximized panel;
- open/closed state of File and Settings menus;
- snippet dialog contents.

Framework enabled IDs and pin IDs are persisted under `state.libraries`, not
`state.settings`. Framework catalog entries and snippets are separate stored
documents (`dcspad.v2.catalog` and `dcspad.v2.snippets`).

### Persistence and project-file boundaries

[`src/state.js`](../src/state.js) is still the only module allowed to touch
localStorage.

- workspace key: `dcspad.v2.workspace`;
- framework catalog key: `dcspad.v2.catalog`;
- snippet library key: `dcspad.v2.snippets`;
- workspace saves are debounced 600 ms and flushed on `pagehide`;
- catalog/snippet writes are synchronous because they follow explicit actions.

A saved project contains:

```json
{
  "docs": { "html": "...", "css": "...", "js": "..." },
  "libraries": {
    "enabled": ["..."],
    "dcsUrl": "..."
  },
  "jsAsModule": false
}
```

It does **not** contain visual layout, Auto-run, console clearing, preview
theme, diagnostic font size, or splash history. Loading a project updates the
module checkbox and Monaco compiler mode as well as the runner setting.

## Editable runtime configuration

[`dcspad.config.json`](../dcspad.config.json) was added after the baseline so
environment URLs can change without rebuilding the app. It is a deployment
configuration document, not a per-user Settings menu.

[`src/config.js`](../src/config.js) fetches it with `cache: "no-cache"`,
resolves relative URLs against the config file URL, normalizes the supported
shape, and returns warnings instead of preventing the workbench from starting.
Warnings are written into the DCSPad Console after it is ready.

### Framework source settings

Global keys:

| Path | Values/default | Meaning |
| --- | --- | --- |
| `frameworks.prefer` | `"local"` or `"cdn"`; default `"local"` | source order inherited by framework items |
| `frameworks.fallbackToCdn` | boolean; default `true` | whether an item may use its alternate source |
| `frameworks.items` | object keyed by catalog ID | overrides matching stored catalog entries without mutating the catalog |

Per-item keys:

| Path | Meaning |
| --- | --- |
| `localUrl` | self-hosted/organization-hosted script URL; blank is ignored |
| `cdnUrl` | public backup script URL; blank is ignored |
| `prefer` | optional item-level `"local"`/`"cdn"` override |
| `fallbackToCdn` | optional item-level fallback override |
| `probeGlobal` | dotted browser global expected after the primary script, for example `pnp` or `Alpine` |
| `intelligence` | editor pack IDs that travel with this runtime even when its URL is opaque |

Source selection rules:

1. use the preferred nonblank URL;
2. if it is blank, use the other nonblank URL as primary;
3. automatic fallback is supported for JavaScript, not CSS;
4. fallback requires two URLs, fallback enabled, and a valid `probeGlobal`;
5. if the primary does not expose the global, the runner inserts the alternate
   parser-blocking at the same catalog position so later dependencies keep
   their order.

The current supplied config maps:

- `pnpjs2` to a blank local URL plus
  `https://cdnjs.cloudflare.com/ajax/libs/pnp-pnpjs/2.15.0/pnp.js`, global
  `pnp`, and intelligence pack `pnpjs-2.15.0`;
- `alpine` to a blank local URL plus the Alpine 3 CDN build, global `Alpine`,
  and intelligence pack `alpine-3`.

To move either runtime to SharePoint hosting, fill in `localUrl`. No catalog
edit or source rebuild is necessary.

### Asset-group settings

Each entry under `assets` supports:

| Path | Meaning |
| --- | --- |
| `prefer` | `"local"` or `"hosted"` |
| `localBaseUrl` | source/review folder relative to the config, or absolute URL |
| `hostedBaseUrl` | eventual deployed folder, normally an absolute SharePoint URL |
| `intelligence` | language packs enabled independently of framework checkboxes |
| `files` | logical file names relative to the selected base |
| `runtime.enabled` | automatically inject this group's runtime assets on every Run |
| `runtime.cssFiles` | logical `files` keys or relative CSS paths to inject |
| `runtime.fluentIconElement` | for `fluentIcons`, inject the preview-only `<fluent-icon>` adapter |

`selectedAssetBase()` uses the preferred nonblank base, then falls back to the
other nonblank base. Folder URLs are normalized with a trailing slash.

Current behavior by group:

| Group | Intelligence | Automatic preview runtime |
| --- | --- | --- |
| `assets.designSystem` | `bsp-design` loads generated CSS-token and HTML-class metadata | **No.** The config has no enabled runtime block. Add `styles.css`/`editorial.css` through the Frameworks UI when a project should render them. Intelligence alone never styles the preview. |
| `assets.fluentIcons` | `fluent-icons` loads exact icon name, sprite ID, and font-class metadata | **Yes.** Regular/Filled/Light font CSS and the font-backed `<fluent-icon>` adapter are injected because `runtime.enabled` is true. Set it false when the consuming page supplies the assets. |

The BMO icon sprite is intentionally not integrated as a new runtime feature.
That UI icon path is deprecated. Fluent `<use>` intelligence remains because
some consuming projects may provide a combined Fluent symbol sprite, although
the imported Fluent package itself contains individual SVG files rather than a
combined sprite.

## Framework checkboxes and intelligence coupling

Framework rows keep their pre-Monaco shell/classes and interactions:

- checkbox enables runtime for the next Run;
- arrows change parser injection order;
- star pins the row visually;
- X removes the catalog entry;
- custom URL form adds and immediately enables a runtime.

What changed is the checkbox side effect. The library manager now reports both:

- `getEnabledLibraries()` for preview runtime injection; and
- `getEnabledIntelligence()` for Monaco language tools.

[`src/main.js`](../src/main.js) calls
`editorsApi.setIntelligencePacks(getEnabledIntelligence())` after initial
library setup, on framework changes, and after project import.

Pack activation:

| Pack | Activation |
| --- | --- |
| `pnpjs-2.15.0` | matching enabled runtime metadata/URL; loads exact PnPjs 2.15 declaration graph and global `pnp` bridge |
| `alpine-3` | matching enabled Alpine 3 runtime; adds JavaScript API plus HTML directives/magics |
| `bsp-design` | configured asset intelligence; independent of a framework checkbox |
| `fluent-icons` | configured asset intelligence; independent of a framework checkbox |

Stored catalogs created before explicit intelligence metadata are handled by
runtime URL/ID detection in [`src/libraries.js`](../src/libraries.js).

Loading is generation-guarded in `src/editors.js`: a slow response from a pack
that was disabled while loading cannot re-enable stale data. JavaScript extra
libraries are stored by pack, so disabling Alpine does not remove PnP types and
vice versa.

Useful nonvisual test hooks on `<html>`:

```text
data-monaco-ready
data-monaco-worker-error
data-pnp-types
data-alpine-intelligence
data-bsp-intelligence
data-fluent-icon-intelligence
```

Pack states use values such as `loading`, `ready`, `disabled`, and `error`.

## Preview/runtime changes

The core fresh same-origin `srcdoc` architecture is unchanged. The current
assembly order in [`src/runner.js`](../src/runner.js) is:

1. console/network/error harness;
2. SharePoint context and `<base>`;
3. pad-only preview chrome style;
4. external library/configured asset CSS;
5. configured inline library styles;
6. user CSS;
7. user HTML;
8. ordered blocking library/configured asset scripts, including primary
   runtime probes and fallback scripts;
9. user JavaScript last.

The design-system consequence is important:

- `styles/app.css` styles the **workbench chrome** in the parent page;
- framework/design-system CSS styles the **preview document** only when it is
  included in the runner;
- generated intelligence metadata changes editor help but never injects CSS;
- the preview-dark style is deliberately earlier than all library and user
  CSS, so it cannot overrule a tested design system.

Do not globally load a preview design system into the workbench page merely to
style sample content. That would create CSS bleed into Monaco and SharePoint.

All assembled scripts still receive the SharePoint host's CSP nonce. Fallback
URLs and configured asset URLs are attribute-escaped before insertion.

## Monaco/SharePoint asset contract

Monaco is fully vendored under [`vendor/monaco/`](../vendor/monaco/):

- `monaco.js`;
- `monaco.css` and codicon font;
- `editor.worker.js`, `css.worker.js`, `html.worker.js`, `ts.worker.js`;
- exact PnPjs 2.15 type payload;
- `version.json`.

Workers are same-origin classic `.js` files. Do not switch them to blobs, CDN
workers, or `.mjs`:

- SharePoint can block blob workers through CSP;
- SharePoint serves `.mjs` with the wrong MIME type in this environment;
- workers must resolve from the uploaded SiteAssets folder rather than the
  SharePoint page URL.

[`boot.js`](../boot.js) sets these hosted globals:

```text
__DCSPAD_SRC_BASE__
__DCSPAD_ASSET_BASE__
__DCSPAD_MONACO_VERSION__
__DCSPAD_INTELLIGENCE_VERSION__
__DCSPAD_CONFIG_URL__
```

It independently versions `styles/app.css`, `dcspad.app.js`,
`dcspad.config.json`, `vendor/intelligence/manifest.json`, and
`vendor/monaco/version.json`.

## Generated design-system and icon intelligence

[`tools/build-design-intelligence.mjs`](../tools/build-design-intelligence.mjs)
reads the local folders named by `dcspad.config.json` and emits:

```text
vendor/intelligence/bsp-design.json
vendor/intelligence/fluent-icons.json
vendor/intelligence/manifest.json
```

Current generated coverage:

- 164 BSP CSS custom properties;
- 375 canonical BSP classes, including Editorial scope/composition metadata;
- 2,655 normalized Fluent icon groups;
- 18,681 real Fluent SVG variants with font availability retained.

The imported `bsp-design-system/` and `bsp-fluent-icon-lib/` folders are source
and review inputs. Hosted mode consumes the compact generated artifacts and
configured runtime font files; the source folders are not automatically copied
by `deploy/Sync-Live.ps1`.

The Fluent source catalog/file corrections made during this work include:

- corrected `person_suport` to `person_support` in the SVG filename and JSON;
- synchronized the Light CSS, Light JSON, specimen HTML, and master catalog;
- separated Chart Person and Chat Person entries rather than collapsing them
  into one normalized record;
- regenerated the compact intelligence artifact after source correction.

If either source repository changes, run:

```powershell
cd tools
npm run build:intelligence
```

Do not hand-edit generated files under `vendor/intelligence/`.

## Build and deployment changes

The hosted web part loads the generated single-file
[`dcspad.app.js`](../dcspad.app.js), while standalone development/tests load
`src/main.js` directly.

After source changes:

```powershell
cd tools
npm run build:app
```

After Monaco/version/type builder changes:

```powershell
cd tools
npm run build:monaco
```

[`deploy/Sync-Live.ps1`](../deploy/Sync-Live.ps1) now:

1. verifies the Monaco vendor set;
2. regenerates design/icon intelligence;
3. rebuilds `dcspad.app.js`;
4. copies the shell, bundle, config, `src/`, `styles/`, and `vendor/`.

A `boot.js` change still requires bumping the `?v=` reference in
[`dcspad.webpart.html`](../dcspad.webpart.html), because boot is below its own
versioning layer.

## Design-system implementation checklist

When applying the workbench UI design system:

- [ ] Keep all functional IDs in `index.html`; prefer changing classes/tokens,
      not replacing control markup.
- [ ] Target `#pane-editor`, not the removed `#pane-html/#pane-css/#pane-js`.
- [ ] Replace the `dcspad-dark` Monaco theme values as part of the design pass;
      changing only `:root` CSS variables will not recolor syntax/editor widgets.
- [ ] Verify suggestions, parameter hints, hover cards, find/replace, diagnostics,
      and context menus against the new palette.
- [ ] Preserve `fixedOverflowWidgets: true` behavior and do not clip Monaco
      widgets with a new ancestor `overflow` or low stacking context.
- [ ] Preserve `#status-editor` and distinct ready/warning treatment.
- [ ] Preserve splash `.visible`, `.fading`, `.failed`,
      `.dcspad-boot-splash`, and the hosted dark underlay.
- [ ] Keep the splash opacity transition at the intentional configured timing
      unless the product timing is explicitly changed again.
- [ ] Restyle the snippet dialog through `.app-dialog*`; keep native dialog,
      form, focus, and accessibility semantics.
- [ ] Keep `[hidden] { display: none !important; }`.
- [ ] Keep workbench chrome CSS out of the preview iframe and preview framework
      CSS out of the parent workbench.
- [ ] Keep status/error dots, SP Live/Mock states, autosave states, and Monaco
      worker warnings visibly distinguishable.
- [ ] Test hosted mode under the SharePoint suite bar as well as standalone.

## Verification targets

The most relevant suites are documented in [`tests/README.md`](../tests/README.md):

- `tests/monaco.mjs`: editor/model contract, completions, hovers, diagnostics,
  undo isolation, and worker warning behavior;
- `tests/config.mjs`: local/CDN configuration, asset intelligence/runtime, and
  ordered fallback;
- `tests/hosted.mjs`: earliest splash, shell/bundle/config paths, versioned
  Monaco/intelligence assets, Fluent bridge, and same-origin workers;
- `tests/smoke.mjs`: shell wiring, frameworks, snippets/dialog, project files,
  persistence, runner, console/network, and preview lifecycle;
- `tests/darkmode.mjs` and `tests/splash.mjs`: preview cascade and reveal timing.

Manual UI regression checks:

1. switch HTML/CSS/JS repeatedly and confirm selection, scroll position, folds,
   and undo history stay per pane;
2. resize/collapse/maximize every panel and reload to confirm persisted layout;
3. open both dropdown menus and the snippet dialog at narrow and wide widths;
4. trigger Monaco suggestions/hover/find while panels are narrow;
5. confirm `Monaco ✓`, then simulate a blocked worker and confirm `Monaco ⚠`
   persists after Run;
6. observe first-visit and repeat-visit splash entrance and exit in standalone
   and SharePoint;
7. load BSP CSS as a framework and verify it styles only the preview;
8. render Regular/Filled/Light Fluent icons and verify workbench chrome remains
   unaffected.

## File-by-file post-baseline map

| Area | Primary files | Post-`3414854` responsibility |
| --- | --- | --- |
| Shell/UI hooks | `index.html`, `styles/app.css` | shared editor host, Monaco status, splash status, snippet dialog, Monaco integration CSS |
| Editor | `src/editors.js`, `src/monaco-runtime.js` | Monaco models/editor/theme/workers, adapter, language packs |
| Settings/wiring | `src/main.js`, `src/state.js`, `src/layout.js` | readiness-gated boot, module diagnostic alignment, intelligence toggles, existing persisted settings/layout |
| Runtime config | `dcspad.config.json`, `src/config.js` | editable local/CDN/hosted URLs and asset/runtime selection |
| Frameworks | `src/libraries.js` | configured sources, runtime probes/fallbacks, pack activation, configured assets |
| Preview | `src/runner.js`, `src/bridge/fluent-icon-font.js` | configured CSS/JS, ordered fallback, Fluent adapter, CSP-safe assembly |
| Language packs | `src/intelligence/*.js` | Alpine, BSP, and Fluent completion/hover/diagnostic providers |
| Splash/host | `boot.js`, `src/splash.js`, `dcspad.webpart.html` | earliest hosted curtain, readiness status, clean reveal, versioned assets |
| Generators | `tools/build-monaco.mjs`, `tools/build-design-intelligence.mjs`, `tools/build-app.mjs` | reproducible Monaco/types, compact design/icon data, hosted bundle |
| Deployment | `deploy/Sync-Live.ps1` | validate/build/regenerate/copy deployable set |
| Tests | `tests/monaco.mjs`, `tests/config.mjs`, `tests/hosted.mjs`, `tests/smoke.mjs` | regression coverage for the new contracts |

## Complete repository delta inventory

This inventory is intentionally explicit so a future integration does not
mistake generated or supporting changes for unrelated work. It includes the
tracked `git diff 3414854 --` surface plus the current Fluent source/runtime
work.

### Added

| Files | Purpose |
| --- | --- |
| `.gitignore` | excludes local dependency trees used by the new build/test flow |
| `dcspad.config.json`, `src/config.js` | editable framework and asset URL configuration |
| `src/monaco-runtime.js` | Monaco CSS/module/worker/type asset loading |
| `src/intelligence/alpine.js` | Alpine JavaScript and HTML language data/providers |
| `src/intelligence/bsp.js` | BSP token/class completions and hovers |
| `src/intelligence/fluent-icons.js` | Fluent name/sprite/font completions, hovers, and diagnostics |
| `src/bridge/fluent-icon-font.js` | preview-only font-backed `<fluent-icon>` adapter |
| `tools/build-monaco.mjs` | reproducible Monaco and PnPjs type vendor build |
| `tools/build-design-intelligence.mjs` | deterministic BSP/Fluent compact-data generator |
| `vendor/monaco/**` | Monaco ESM runtime, CSS/font, classic workers, version manifest, and PnP type payload |
| `vendor/intelligence/bsp-design.json`, `fluent-icons.json`, `manifest.json` | generated browser intelligence artifacts |
| `tests/monaco.mjs`, `tests/config.mjs`, `tests/hosted.mjs`, `tests/hosted-fixture.html` | new editor/config/hosted regression suites and fixture |
| `deploy/spike-worker.js` | same-origin worker/CSP deployment probe |
| `plans/design-system-intellisense.md` | design/intelligence implementation plan and decisions |
| `design/POST-MONACO-UI-INTEGRATION.md` | this migration/UI contract |
| `bsp-design-system/**`, `bsp-fluent-icon-lib/**` | imported source/review inputs used by the generator; not workbench chrome |

### Modified

| Files | Purpose |
| --- | --- |
| `index.html`, `styles/app.css` | shared Monaco host, editor status, splash status/states, snippet dialog, Monaco integration styling |
| `src/editors.js` | CodeMirror adapter implementation replaced by Monaco plus composable intelligence |
| `src/layout.js` | tabs now activate Monaco models rather than three DOM panes |
| `src/main.js` | readiness boot order, config loading, intelligence synchronization, module-mode synchronization |
| `src/libraries.js` | configured sources, ordered fallback metadata, intelligence packs, configured asset runtime |
| `src/runner.js` | configured CSS/JS support, safe fallback insertion, Fluent runtime, attribute escaping |
| `src/snippets.js` | native prompt replaced by in-app dialog; Monaco-aware insertion remains adapter-based |
| `src/splash.js`, `boot.js` | earliest hosted curtain, readiness text, repeat-visit behavior, clean 0.75 s reveal |
| `dcspad.webpart.html` | boot version bump for hosted cache invalidation |
| `dcspad.app.js` | regenerated hosted bundle containing the current `src/` implementation |
| `deploy/Sync-Live.ps1`, `deploy/README.md`, `deploy/webpart-spike.html`, `deploy/spike-module.js` | Monaco/intelligence build-copy validation and worker/CSP re-verification |
| `tools/build-app.mjs`, `tools/package.json`, `tools/package-lock.json`, `tools/node_modules/.package-lock.json` | Monaco/esbuild/type dependencies and hosted bundle build |
| `tests/smoke.mjs`, `tests/darkmode.mjs`, `tests/README.md` | updated selectors plus snippet, runtime, splash, and intelligence coverage/documentation |
| `README.md`, `CLAUDE.md`, `HANDOFF.md`, `REVIEW-LOG.md` | user, maintainer, handoff, and review documentation |
| `design/DESIGN-BRIEF.md` | Monaco terminology and pointer to this authoritative integration contract |
| `plans/pnpjs-intellisense.md` | worker/CSP conclusions and implemented Monaco/PnP direction |

### Removed/replaced

| Files | Replacement |
| --- | --- |
| `vendor/codemirror.js` | `vendor/monaco/**` |
| `tools/build-vendor.mjs` | `tools/build-monaco.mjs` |

Generated files must be regenerated from their owning source rather than
edited manually:

- `dcspad.app.js` ← `tools/build-app.mjs`;
- `vendor/monaco/**` ← `tools/build-monaco.mjs`;
- `vendor/intelligence/**` ← `tools/build-design-intelligence.mjs`.
