# DCSPad starter snippet library

[`dcspad-starter-snippets.json`](dcspad-starter-snippets.json) is an
import-ready library of PnPjs, Alpine, BSP design-system, Fluent icon, and
general browser helpers.

## Import

1. If you already have personal snippets, use the Snippets **download** button
   first. DCSPad's current import operation replaces the whole library; it
   does not merge.
2. Use the Snippets **upload** button.
3. Select `examples/dcspad-starter-snippets.json`.
4. Confirm the replacement.

Entries with `(HTML)` and `(JS)` in the name are paired. Insert each entry into
its corresponding pane.

## Runtime prerequisites

- Enable **PnPjs v2 (classic)** before running PnP snippets. Writes use obvious
  placeholders such as `DCSPad Sandbox` or `YOUR LIST TITLE`; review them
  before running.
- Enable **Alpine.js** before running Alpine snippets.
- BSP intelligence does not style the preview. Add the regular design-system
  CSS as a Framework; add `editorial.css` after it when using Editorial
  snippets.
- The configured Fluent runtime automatically loads the three icon fonts and
  preview `<fluent-icon>` adapter unless
  `assets.fluentIcons.runtime.enabled` is set to `false`.
- PnP snippets need the hosted SharePoint workbench. `_api` calls are expected
  to fail under the local mock context.

## Maintaining the pack

The JSON is generated and validated against the current BSP and Fluent
intelligence catalogs:

```powershell
cd tools
npm run build:snippets
```

Edit `tools/build-starter-snippets.mjs`, not the generated JSON.

