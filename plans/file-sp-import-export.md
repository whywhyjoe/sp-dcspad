# Implementation record: project files, local code files, and SharePoint transfer

Status: **implemented** (2026-07-27).

Automated browser coverage verifies local import/export, project naming and
`.dcspad.json` behavior, explicit/global/Modern SharePoint context, safe
ResourcePath browsing, same-tenant site switching, cross-site reads and digest
selection, upload payloads, and overwrite confirmation. Live-tenant browser
validation has confirmed current-site browsing and import/read behavior.
Live upload and another-site browsing remain deployment-checklist items.

This record covers three related additions:

1. inline project naming and required-name `.dcspad.json` saves;
2. one local **Import file…** action that accepts HTML, CSS, or JavaScript
   and replaces the matching editor pane; and
3. HTML/CSS/JS import and export against a SharePoint document library.

Project saves (`*.dcspad.json`) stay local-only. Catalog and snippet JSON files
also remain unchanged.

## 1. Local single-file import

### User flow

The File menu provides:

```text
Project
  Save project (.dcspad.json)
  Load project (.dcspad.json)…
Import
  Import HTML, CSS, or JS…
Export
  Export all non-empty
  Export HTML
  Export CSS
  Export JS
```

`Import HTML, CSS, or JS file…` opens one hidden picker with:
```html
accept=".html,.htm,.css,.js,text/html,text/css,text/javascript,application/javascript"
```

After selection:

- inspect the final extension case-insensitively;
- map `.html` and `.htm` to the HTML pane, `.css` to CSS, and `.js` to JS;
- reject renamed/unsupported files with an in-app error and make no edits;
- reject files above the existing 5 MB import ceiling before calling
  `file.text()`;
- show an in-app confirmation before changing the editor:

  > Replace CSS code?
  >
  > `site-theme.css` will replace all code in the CSS editor.

- on confirmation, call `editorsApi.setDocs({ [pane]: text })`, select that
  editor tab, mark it unsaved, and let the normal workspace autosave persist
  it;
- on cancel, leave editor content, selection, and active tab unchanged.

The confirmation is shown even when the target pane is empty. That keeps the
replacement behavior explicit, as requested, and avoids two subtly different
flows.

### Implementation seams

- `src/io.js`
  - exports the shared `MAX_IMPORT_BYTES` guard;
  - provides `wirePaneImport(inputId, onCandidate)`, returning
    `{ fileName, pane, text }`;
  - remains storage-free.
- `index.html`
  - owns the File menu item, one hidden file input, and a confirmation dialog
    using the existing `.app-dialog*` component family.
- `src/main.js`
  - owns confirmation, editor replacement, tab selection, unsaved state, and
    status text.
- `styles/app.css`
  - styles the compact file-name/type and dialog states.

### Local-import coverage

`tests/files.mjs` and `tests/smoke.mjs` cover:

1. each supported extension maps to the correct pane;
2. mixed-case extensions work;
3. the confirmation names the file and destination pane;
4. cancel preserves the current content;
5. confirm replaces only the mapped pane and makes the change undoable as one
   Monaco action;
6. unsupported and oversized files are rejected without mutation;
7. selecting the same file twice still fires because the input is reset.

## 2. SharePoint document-library transfer

### Scope

The SharePoint feature moves plain HTML, CSS, and JS text only:

- **Export to SharePoint** uploads/overwrites one selected pane.
- **Import from SharePoint** downloads one selected file, infers its pane from
  the extension, shows the same replacement confirmation, and replaces that
  pane.
- Project JSON, snippet JSON, and framework-catalog JSON never appear in this
  picker.
- Standalone/mock mode keeps these actions visible but disabled with a
  “Requires SP: Live” explanation.

### REST feasibility

This can use same-origin SharePoint REST with the signed-in user's existing
permissions; it needs no Graph application or separate login.

- Browse a folder with
  `GET /_api/web/GetFolderByServerRelativePath(decodedUrl='…')`
  and select its `Folders` and `Files` collections.
- Download bytes with
  `GET /_api/web/GetFileByServerRelativePath(decodedUrl='…')/$value`.
- Upload a small text file with
  `POST /_api/web/GetFolderByServerRelativePath(decodedUrl='…')/Files/AddUsingPath(decodedUrl='…',overwrite=…)`
  and the UTF-8 text as the request body.
- Every non-GET request carries a fresh `X-RequestDigest`. Prefer the digest
  already captured by `getSpContext({ refresh: true })`; if it is missing or
  stale, refresh it with `POST /_api/contextinfo`.

Read, browse, and upload all use ResourcePath-based endpoints so decoded paths
containing `#` or `%` are unambiguous. Upload names are still constrained to a
safe HTML/CSS/JS leaf name containing letters, numbers, dots, hyphens, and
underscores.

### SharePoint context acquisition

Do not make the file-transfer feature depend solely on
`window._spPageContextInfo`. The HTML/CSS/JS transfer needs only the current
web URL and a fresh request digest, not the complete legacy page-context
object.

Resolve context through this ordered, guarded ladder:

1. **Explicit DCSPad host context.** Accept a small
   `window.__DCSPAD_SP_CONTEXT__` object containing `webAbsoluteUrl` and
   optional display/user fields. This is the stable integration seam for a
   custom script editor or a future SPFx application customizer.
2. **Classic/global context.** Read `window._spPageContextInfo` as today.
3. **Modern-page legacy context.** Best-effort read
   `legacyPageContext` from the Modern page's internal `PageManager`:

   ```js
   window.spModuleLoader
     ?._bundledComponents
     ?.[MODERN_SITE_PAGES_FEATURE_ID]
     ?.PageManager
     ?._instance
     ?.pageContext
     ?.legacyPageContext
   ```

   Use feature ID `b6917cb1-93a0-4b97-a84d-7cf49975d4ec`. Probe the current
   window and then same-origin `parent`/`top` windows inside `try/catch`, since
   a custom script editor may render DCSPad inside an iframe. Treat this route
   as optional: `spModuleLoader`, `_bundledComponents`, `PageManager`, and
   `_instance` are undocumented implementation details and can change.
4. **REST bootstrap.** When a candidate SharePoint web URL is known but no
   usable legacy object or digest exists, call
   `POST {webUrl}/_api/contextinfo`. Its response supplies `WebFullUrl`,
   `SiteFullUrl`, `FormDigestValue`, and the digest timeout. Use this endpoint
   to refresh the digest before writes regardless of which page-context route
   initially succeeded.
5. **Explicit setup or Mock mode.** If no candidate web URL is available,
   remain in Mock mode and let the SharePoint-files dialog accept/configure a
   same-origin web URL rather than guessing a site boundary from arbitrary URL
   path segments. Validate it with `/_api/contextinfo` before enabling file
   actions.

Normalize all successful routes into the existing `getSpContext()` result
shape and record a `source` such as `host`, `global`, `modern-legacy`, or
`rest`. Only the explicit host object and REST response are treated as stable
contracts. A failure in the Modern internal lookup must silently fall through,
not mark SharePoint unavailable or prevent the REST bootstrap.

The custom script editor does **not** have to emit the complete
`_spPageContextInfo`. If it can run same-origin code, it can either expose only
`webAbsoluteUrl` through the DCSPad host object or allow the guarded parent-page
probe. If neither is possible, the configured web URL plus `/_api/contextinfo`
is sufficient for the file operations.

Microsoft references:

- [Working with folders and files with REST](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)
- [Supporting `%` and `#` with ResourcePath APIs](https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/supporting-and-in-file-and-folder-with-the-resourcepath-api)
- [Working with `__REQUESTDIGEST`](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/basics/working-with-requestdigest)
- [Navigate SharePoint REST data with `/_api/contextinfo`](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/navigate-the-sharepoint-data-structure-represented-in-the-rest-service)
- [SPFx `PageContext.legacyPageContext`](https://learn.microsoft.com/en-us/javascript/api/sp-page-context/pagecontext?view=sp-typescript-latest)

Community reference for the guarded Modern-page lookup:

- [Get `legacyPageContext` in the JS console](https://sharepoint.stackexchange.com/questions/267389/get-legacypagecontext-in-js-console)

### Picker implementation

The app uses one reusable **SharePoint files** dialog rather than nested
File-menu submenus:

- mode heading: “Export CSS to SharePoint” or “Import from SharePoint”;
- persisted **SharePoint site** URL field, validated with `/_api/contextinfo`;
- breadcrumb/path field beginning at the selected web;
- folder list with parent navigation;
- file list filtered to `.html`, `.htm`, `.css`, and `.js`;
- import mode: select one file, then Continue to the replacement confirmation;
- export mode: pane selector (HTML/CSS/JS), safe file-name field prefilled from
  the project slug, explicit **Overwrite existing file** confirmation when a
  matching file exists;
- loading, empty, permission-denied, stale-digest retry, and network-error
  states inside the dialog;
- remember the selected web and last folder path in
  `state.settings.spFilesWebUrl` and `state.settings.spFilesFolder`.

The dialog reuses the current dark surfaces, compact rows, file-type badges,
focus ring, and native `<dialog>` behavior. It has no framework dependency.

### Module implementation

- `src/sp-files.js`
  - no DOM and no storage;
  - `listFolder(serverRelativePath)`;
  - `readTextFile(serverRelativePath)`;
  - `writeTextFile(folderPath, fileName, text, { overwrite })`;
  - `getDigest()` with expiry-aware caching and a single retry after a
    digest-related 403;
  - OData literal escaping (`'` → `''`) and decoded ResourcePath handling;
  - normalized error objects for 401/403, 404, conflict, invalid path, and
    network failures.
- `src/main.js`
  - gate by live SharePoint context;
  - connect picker state to the editor adapter;
  - reuse the local replacement confirmation;
  - persist the last folder through `state.js`.
- `index.html` / `styles/app.css`
  - reusable picker dialog and its list/breadcrumb states.
- `src/io.js`
  - remains local-file-only.

### Delivery record

1. Shipped local single-file import independently.
2. Extended `sp-context.js` with the context ladder above and tested every
   route, including inaccessible cross-origin parents and missing/changing
   Modern internals.
3. Built `sp-files.js` plus URL/digest coverage with stubbed REST responses.
4. Built the picker and its loading/error/empty states.
5. Wired live context, import confirmation, upload overwrite confirmation, and
   last-folder persistence.
6. Ran the standalone Playwright suites with REST stubs.
7. Deployed to NewNerve and verified current-site browsing and import/read
   behavior. Live upload, another-site browsing, contribute-only access, and
   read-only access remain explicit deployment checks.

### Acceptance checks

- No SharePoint action is callable in Mock mode.
- SharePoint file actions work when `_spPageContextInfo` is absent but the
  explicit host object, guarded Modern context, or configured REST bootstrap
  succeeds.
- A change or absence in the private Modern-page lookup falls through safely.
- Folder browsing never leaves the currently selected web/site boundary.
- Same-origin sites elsewhere on the tenant can be selected, browsed, imported
  from, and exported to without reusing the host web's digest.
- Only HTML/CSS/JS files are shown or accepted.
- Import never mutates a pane before explicit confirmation.
- Export never silently overwrites an existing SharePoint file.
- A 403 explains whether the likely issue is permissions or digest refresh
  failure.
- Project `.dcspad.json` files remain local-only.
- `#` in a ResourcePath folder is covered by automated URL-building/browser
  coverage. Spaces, apostrophes, `%`, and those path cases on a live tenant
  remain deployment-checklist cases.
- The hosted bundle is rebuilt after source changes.
