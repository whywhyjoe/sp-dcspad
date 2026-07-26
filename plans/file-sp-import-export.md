# Plan: per-pane import/export — file system + SharePoint

Status: **planned, not started** (written 2026-07-25). Project JSON save/load
stays exactly as it is — this adds *per-pane* (HTML / CSS / JS) import and
export, each against two targets: the local file system and SharePoint
document libraries. Import replaces the matching pane's content.

## Answer to "SP should be doable with the HTTP API, no?"

Yes, cleanly. The pad runs same-origin on the tenant with session cookies, so:

- **Read** — `GET /_api/web/GetFileByServerRelativeUrl('<path>')/$value`
  returns the raw file bytes. No digest needed for GETs.
- **Write (new or overwrite)** —
  `POST /_api/web/GetFolderByServerRelativeUrl('<folder>')/Files/add(url='<name>',overwrite=true)`
  with the content as the POST body, plus an `X-RequestDigest` header from
  `POST /_api/contextinfo` (cache the digest, it carries an expiry).
- **Browse (picker)** —
  `GET /_api/web/GetFolderByServerRelativeUrl('<folder>')?$expand=Folders,Files&$select=…`
  gives everything a minimal folder picker needs.

No Graph, no app registration, no CORS. Permissions are simply the signed-in
user's own. The only precondition is live SP context (already true in the web
part; standalone-on-localhost falls back to file-system-only with the SP
options disabled — same policy as the SP chip).

## UI shape

Extend the existing File menu (`index.html` file-menu):

```
Save project (.json)          ← unchanged
Load project (.json)…         ← unchanged
──────────────
Export HTML ▸   to file · to SharePoint…
Export CSS ▸    to file · to SharePoint…
Export JS ▸     to file · to SharePoint…
──────────────
Import HTML ▸   from file… · from SharePoint…
Import CSS ▸    from file… · from SharePoint…
Import JS ▸     from file… · from SharePoint…
```

(Exact affordance — submenu vs. six flat rows vs. a small dialog with a
pane selector — decide at build time; the menu is already getting tall.)

- **to/from file** — reuse `io.js` download + a per-type file input
  (`accept=".html,.htm"` / `".css"` / `".js,.mjs"`). Import reads the file and
  replaces the pane via the editors API; autosave persists it.
- **to/from SharePoint…** — opens a small picker dialog: server-relative path
  input + a folder browser (REST `$expand` above), file list filtered by the
  pane's extensions. Export asks for target folder + file name (prefilled);
  import selects an existing file. Remember the last-used folder per pane in
  `state.settings` so round-tripping is one click after the first time.

## Module seams (respect CLAUDE.md invariant 6)

- `src/io.js` — stays local-only: grows `pickTextFile(accept)` and keeps
  `download()`. Moves bytes, stores nothing.
- **New `src/sp-files.js`** — all SharePoint HTTP: digest fetch + cache,
  `readFile(serverRelPath)`, `writeFile(folder, name, content)`,
  `listFolder(serverRelPath)`. Talks `fetch` only; no DOM, no storage. Uses
  `getSpContext()` for the web URL and refuses politely when context is mock.
- `src/main.js` — wires menu items: get/set pane content via the editors API,
  route to `io.js` or `sp-files.js`. Confirmation prompt before an import
  replaces a non-empty pane (same pattern as the framework-remove confirm).
- Picker dialog markup in `index.html` + styles in `app.css` (reuse
  `.settings-menu` / dialog conventions; no framework, plain DOM).

## Sequenced steps

1. `src/sp-files.js` with digest handling + the three REST calls; unit-test
   the URL building (escaping apostrophes in paths: `'` → `''`).
2. Local per-pane import/export (file system) — no SP dependency; testable in
   `smoke.mjs` today alongside the existing export checks.
3. Picker dialog + SP wiring behind a `spContext.live` gate.
4. Per-pane last-used-path memory in `state.settings`.
5. Tests: smoke covers local import/export + dialog opens/validates in mock
   mode (SP calls stubbed with a fetch shim); live SP round-trip goes on the
   tenant verification checklist in README/deploy docs (like PnPjs checks).

## Open questions / later

- **Linked panes** (phase 2?): after an SP export/import, keep the file
  "linked" and show a one-click re-export (and maybe a modified-on-server
  warning via the file's `TimeLastModified`). Skip for v1.
- Whether Export JS should offer `.mjs` — no: SharePoint serves `.mjs` as
  `application/octet-stream` (see CLAUDE.md gotcha), always default `.js`.
- Large-file guard on import (a pane full of megabytes will hurt the
  autosave debounce) — probably just a size confirm at ~1 MB.
