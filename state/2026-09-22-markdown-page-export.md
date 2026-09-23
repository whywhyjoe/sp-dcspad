# State — SP Workbench markdown page export

Last touched: 2026-09-22
Mode: Joe
Branch: claude/markdown-page-export-rebase (unpushed)
State: code rebased and green (581 checks across all suites); docs done; bundle rebuild, Codex review, and dev-tenant live verification still to run this session

## What this is

The Pages "Extract content" export wrote a `.md` whose web-part sections were the part's raw
sanitized HTML, even for text parts whose rich-text body markdown can carry exactly. This work
ports a real HTML→Markdown converter (vendored Turndown, `src/html-markdown.js`) behind two
named profiles, adds a `markdown`/`html` content-format choice to the page content export, adds
per-row MD/HTML export buttons to the Pages grid, and lets a grid column be marked `action` so
per-row controls don't leak into CSV/JSON/markdown exports. Originally built on
`origin/claude/workbench-markdown-extraction-tpuu10` (2026-08-26); rebased onto `main` today over
the List Schema stages and the list Tools tab. Full design detail lives in
`HANDOFF.md` → "SP Workbench: markdown page export (2026-09-22)" — this file tracks status, not
design.

## Done

- `src/html-markdown.js` (the one HTML→Markdown converter; `listField` and `pageContent`
  profiles; unknown profile throws), `vendor/turndown/` (vendored Turndown browser-ESM build),
  `tools/build-vendor-turndown.mjs` (copy step with provenance sha256 + bare-import guard +
  default-export check).
- `src/workbench/item-export.js` now calls `html-markdown.js`'s `listField` profile instead of
  its own hand-rolled converter.
- `src/workbench/page-export.js`: `buildContentExport(…, format)` — `'markdown'` (default) or
  `'html'`, unknown format throws, metadata frame byte-identical between the two,
  `contentFileName()` keeps `-content.md` / `-content-html.md` apart, `bundleEntryName()` reuses
  it so a bundled page is byte-identical to a solo export.
- `src/workbench/views/pages.js`: per-row Export MD / Export HTML buttons (export without
  opening the page; a denied read goes through `denied.js`), the detail pane's Export content
  became the same pair, the grid's Export ▾ menu offers both zip formats.
- `src/workbench/grid.js` columns may be marked `action`; `src/workbench/export.js` drops them
  from CSV/JSON/markdown exports.
- `tools/build-app.mjs` + `tools/build-workbench.mjs` refuse to stamp a build number from a
  shallow clone (`DCSPAD_BUILD_NUMBER` still wins).
- `CLAUDE.md` (file map for html-markdown.js/vendor/turndown/build-vendor-turndown.mjs, amended
  page-export.js/item-export.js/grid.js/views/pages.js entries, invariant 5's rebuild sentence,
  the Tests paragraph), `tests/README.md` (workbench.mjs line, suite total), `HANDOFF.md` (new
  "SP Workbench: markdown page export (2026-09-22)" section with an unchecked live-tenant
  checklist) all updated to match the rebased code and the 581-check total.
- Rebase itself: seven original commits collapsed to four source commits; only the test-count
  paragraphs in `CLAUDE.md`/`tests/README.md` conflicted, resolved in favour of `main`.
- `workbench.mjs` measured at 141 (was 133 on main) after the rebase; all suites together 581
  (was 573).

## Next

- [ ] Rebuild `dcspad.app.js` and `dcspad.workbench.js` from clean HEAD (this session's docs
      pass did not touch either bundle).
- [ ] Codex review (xo) of this branch, plus one re-review round.
- [ ] Deploy to the dev tenant (`deploy/Sync-Live.ps1 -Environment dev`) and run the live-tenant
      checklist in `HANDOFF.md`.
- [ ] Record live results in `HANDOFF.md`'s checklist.

## Open questions

- None outstanding from Joe as of 2026-09-22.

## Landmines

- Turndown is vendored, never bare-imported — `src/` must keep loading unbundled (standalone
  `index.html`, every test suite) as well as bundled; a bare `'turndown'` specifier would 404
  outside the esbuild bundle.
- `htmlToMarkdown(html, profile)` throws on an unknown profile, on purpose — no silent fallback
  to the wrong structural conventions.
- Both `-content.md` and `-content-html.md` are markdown files; the stem is the only thing that
  tells them apart. Don't let a future format add a third file that collides with either stem.
- Action columns (`grid.js` `action: true`) must stay out of every export — CSV, JSON, and
  markdown all go through `export.js`'s `dataColumns()` filter; a new export path that bypasses
  it will leak the column back in.
- The shallow-clone build guard means a cloud session must `git fetch --unshallow` (or set
  `DCSPAD_BUILD_NUMBER` explicitly) before running `build-app.mjs` / `build-workbench.mjs`, or
  the build refuses to stamp a number at all.

## Live-tenant checklist

Not yet run this session. See `HANDOFF.md` → "SP Workbench: markdown page export (2026-09-22)" →
Live-tenant checklist for the unchecked items (real tables/no raw HTML in the markdown format,
byte-identical metadata frames between formats, row buttons + denied-read register, and the
bulk-zip format/byte-identical-bundle checks).
