# Page copy — same-web duplicate and cross-site copy (plan, v2)

Status: **plan, not built.** v1 written 2026-09-23 against the tabs branch; v2 (same day)
re-anchored on `main` at `ab5a0f9` and revised after a Codex review (§10). Nothing in the
repo copies or publishes a page today; the Pages view is read-only apart from the metadata
editor.

Line references are to `main` at `ab5a0f9`. `PnP-pages` = `@pnp/sp/clientside-pages/types.js`
and `PnP-files` = `@pnp/sp/files/types.js`, both inside `lib-mirror/pnp2.bundle.js.map`
(`sourcesContent`); they are the only local ground truth for the site-page REST surface.

## 0. What the vendored PnPjs 2.15.0 source demonstrates

These are the request shapes PnPjs sends. They establish the callable form, not server
semantics; anything marked *spike* is a tenant question.

| Step | Request (as PnPjs sends it) | Source |
|---|---|---|
| Create | `POST {web}/_api/sitepages/pages` body `{ "__metadata":{"type":"SP.Publishing.SitePage"}, PageLayoutType, PromotedState }` → page JSON with `Id`. No folder parameter. The **first `savepage` `Title` becomes the file stem** (PnPjs: `title = pageName` → `save(false)` → `title = realTitle`). | PnP-pages 1066–1089 |
| Create in a folder (alternative) | `POST {web}/_api/web/GetFolderByServerRelativePath(decodedUrl='…')/Files/addTemplateFile(urloffile='/sites/x/SitePages/sub/Name.aspx',templatefiletype=3)` (`TemplateFileType.ClientSidePage = 3`) → file JSON. Callable form only; suitability is a *spike* question (§9). | PnP-files 129–147, 852 |
| Read | `GET {web}/_api/sitepages/pages({id})` → DTO with `Title, CanvasContent1, LayoutWebpartsContent, BannerImageUrl, BannerThumbnailUrl, Description, TopicHeader, AuthorByline, PageLayoutType, PromotedState, FileName, IsPageCheckedOutToCurrentUser, …`. PnPjs reads `CommentsDisabled` from the **list item** separately — the DTO and the item are two representations, not one. | PnP-pages 284–300 |
| Check out | `POST …/pages({id})/checkoutpage` (no body) — skipped when `IsPageCheckedOutToCurrentUser`. | PnP-pages 357 |
| Save | `POST …/pages({id})/savepage`, header `If-Match: *`, body `{ "__metadata":{"type":"SP.Publishing.SitePage"}, AuthorByline, CanvasContent1, Description, LayoutWebpartsContent, Title, TopicHeader, BannerImageUrl }`; when a custom thumbnail is set the `BannerImageUrl` sent is the thumbnail URL. PnPjs then **re-reads** the page to refresh identity and URLs. `savepageasdraft` is not demonstrated by this source. | PnP-pages 362–378, 393–397 |
| Publish | `POST …/pages({id})/publish` (no body). | PnP-pages 383 |
| Discard checkout | `POST …/pages({id})/discardPage` body `{ "__metadata":{"type":"SP.Publishing.SitePage"} }`. This undoes a checkout; it does **not** delete a page. | PnP-pages 415 |
| Delete a created page | `POST {web}/_api/web/GetFileByServerRelativePath(decodedUrl='…')/recycle` (standard file API; not page-specific). | — |
| Image import | `POST {web}/_api/sitepages/AddImageFromExternalUrl?imageFileName='..'&pageName='..'&externalUrl='..'&$select=ServerRelativeUrl` → `{ ServerRelativeUrl }` into `SiteAssets/SitePages/{pageName}/`. Demonstrates **image** import from a URL the server can fetch; authenticated same-tenant URLs and non-image files are *spike* questions. | PnP-pages 591–615 |
| Banner metadata | `GET {web}/_api/web/getFileByServerRelativePath(decodedUrl='..')?$select=ListId,WebId,UniqueId,SiteId,Name` → `BannerImageUrl = {web}/_layouts/15/getpreview.ashx?guidSite&guidWeb&guidFile`; header part gets `serverProcessedContent.imageSources.imageSource` (server-relative path) and `customMetadata.imageSource {siteId, webId, listId, uniqueId}`, plus the same four ids in `properties`. The copy path resolves an existing preview URL back to its file by those GUIDs. | PnP-pages 318–354, 506–538 |
| Author byline | `AuthorByline` in the DTO is an array of **UPN strings**; the list item's `_AuthorByline` is a person field. Never map one onto the other blindly. | PnP-pages 664 |
| Components | `GET {web}/_api/web/GetClientSideWebParts` → component definitions available on that web. Absence of an id means the target cannot render it; how `savepage` treats an unknown id is a *spike* question. | — |
| Whole-file copy | `POST {web}/_api/SP.MoveCopyUtil.CopyFileByPath(overwrite=@a1)?@a1=false` body `{ srcPath:{ "__metadata":{"type":"SP.ResourcePath"}, DecodedUrl:<absolute> }, destPath:{…absolute…}, options:{ "__metadata":{"type":"SP.MoveCopyOptions"}, KeepBoth, ResetAuthorAndCreatedOnCopy:true, ShouldBypassSharedLocks:true } }`. Copies the file and its item; author/created are reset by the option. What else is carried, and whether a copied published news page is publicly visible before any reset, are *spike* questions. `MoveFileByPath` has the same shape. | PnP-files 270–290 |
| Comments | Not in the vendored source. Hypothesis from PnP Core: `POST {web}/_api/web/lists(guid'…')/items({id})/SetCommentsDisabled` body `{ value: <bool> }`. *Spike.* | PnP Core `ListItem.cs` |

Rejected alternatives: Graph `POST /sites/{id}/pages` (documented web-part subset, needs a
Graph token); `Site.CreateCopyJobs` (async queue, verbatim copy); `Copy-SPOPersonalSitePage`
(admin cmdlet, not callable from the page — but its recipe, "assets folder → copy assets →
create page → latest version only → permissions dropped", is what §5 mirrors).

## 1. Scope and settled decisions

| | Same web ("Duplicate") | Another web or site collection, same tenant ("Copy to site") |
|---|---|---|
| Modern JSON-canvas pages, layout Article/Home/SingleWebPartAppPage | full fidelity per the §4 matrix | best effort with a preflight report |
| Legacy HTML-format canvas | `CopyFileByPath` only | unsupported (explicit message) |
| Repost, Spaces, Topic, HeaderlessSearchResults layouts; classic wiki/publishing | out of scope (v1) | out of scope |
| Version history, comments, likes, item permissions, scheduled publish, translations | never copied | never copied |

Decisions (Joe, 2026-09-23):

1. **Naming:** keep the source file name; `{name}-copy` (then `-copy-2`…) only when that name
   already exists at the destination. Same-folder duplicates therefore always get the suffix.
2. **Same-web engine: the sitepages API, created by path A** (decided 2026-09-23 from the
   spike, §11 Q1/Q5). `CopyFileByPath` lands a copy of a promoted page as a 0.1 draft that
   still carries `PromotedState` 2 and the source's `FirstPublishedDate`, and it is in the
   list's news query (`PromotedState eq 2`) within a second — a news-visible window no
   pre-copy step can close without writing to the source; whether readers see the draft
   depends on the library's `DraftVersionVisibility`, which a copy tool cannot assume. It
   also regenerates `Description` from the body. Path B (`addTemplateFile`) is out: the file
   it makes lacks the Site Page content type and every sitepages call on it is refused.
   `CopyFileByPath` stays only for legacy-HTML canvases on the same web (§1 table).
3. **Links into the source web:** keep by default; rewrite is opt-in.
4. **Unportable web parts:** copy with warnings; never block; drop only on the operator's
   explicit per-part choice. A server rejection of an unknown id must surface as a failure
   with the offending part named, never as a silent drop.
5. **Cross-site-collection target for the spike:** `https://nervedotnet.sharepoint.com/sites/TestSiteCollection` (provisioned). Cross-web: `/sites/NewNerve/sputils-test`.
6. **Classic and repost pages:** out of scope for v1.

## 2. Architecture

### 2.1 Modules

| File | Role |
|---|---|
| `src/workbench/page-copy.js` | **Pure.** `snapshotFromReads(...)` builds the source snapshot (§2.2); `inventoryReferences(snapshot)` (§5.1); `analyzeCopy({ snapshot, target, options })` → plan; `rewriteContent(snapshot, plan, transferResults)` → `{ CanvasContent1, LayoutWebpartsContent, BannerImageUrl }` strings. Operates on the **raw** `JSON.parse(CanvasContent1)` array and raw `LayoutWebpartsContent`, preserving unknown properties byte-for-byte except for patched values; `canvas.js` is used only to label controls. |
| `src/workbench/page-copy-analyzers.js` | Schema-aware per-web-part analyzers (§5.2): `{ id, refs(instance) → [...], patch(instance, mapping) }`. Data-only table plus small functions; no I/O. |
| `src/workbench/sp-pages.js` | Thin client over an injected `{ client, write }` pair: `getPage`, `createPage`, `addTemplateFile`, `checkoutPage`, `savePage`, `publishPage`, `discardPage`, `recycleFile`, `moveFileByPath`, `copyFileByPath`, `addImageFromExternalUrl`, `fileInfo`, `readFileBytes`, `clientSideWebParts`, `setCommentsDisabled`, `featureActive`. Exact request shapes from §0; no planning. |
| `src/workbench/page-copy-run.js` | `runCopy(frozen, deps, { onStep })` — executes the plan, keeps the run journal (§7), returns `{ outcome, journal, pageUrl }`. |
| `src/workbench/page-copy-dialog.js` | The dialog (§6), house style of `upload-metadata.js` / `list-schema-dialog.js`. |

`sp-write.js` gains nothing except what `sp-pages.js` needs and cannot express with
`postJson`/`mergeJson`: a raw `ArrayBuffer` GET is a read, so `readFileBytes` lives in
`sp-files.js` (beside `readTextFile`, taking `{ webUrl }` explicitly) and is re-exported
through `sp-pages.js`.

### 2.2 Source snapshot (the copy payload is never the display item)

`DETAIL_SELECT` (pages.js:185) lacks `PageLayoutType`, `AuthorByline`, `TopicHeader`,
`CommentsDisabled`, `BannerThumbnailUrl` and every custom column, so the dialog assembles a
dedicated snapshot from the **source** client (passed explicitly; never the default host web):

1. `getPage(id)` — the sitepages DTO (§0 Read), which is the content of record.
2. The list item: `items(id)?$select=*&$expand=FieldValuesAsText` (the metadata pane's read,
   pages.js:1220–1247) — `CommentsDisabled`, custom columns, `FileRef`/`FileDirRef`,
   `CheckOutType`, `_UIVersionString`.
3. The library field schema (the view's `FIELD_SELECT` probe) — types, required flags,
   `EntityPropertyName` aliases.
4. `pageStatus(...)` (pages.js:557) — published/checked-out state, so the dialog can say which
   version is being copied (the current DTO, i.e. the latest draft if one exists).

Eligibility gate (§2.4) runs on the snapshot, not on the grid row.

### 2.3 Target client pair and preflight binding

The shell's shared `client` is never retargeted (that resets every view). The Pages view
must accept and forward the `createClient` and `mockWriter` dependencies `main.js:68` already
injects (lists.js:175 does; `createPagesView` currently takes only `{ client, navigate,
updateRoute }`). The dialog follows `list-schema-dialog.js:568–620` exactly:

- a **fresh client per connect** (`createClient()` then `connectWeb(url)`), because
  `connectWeb` commits its target after an await and a shared client could be repointed by a
  slower older connect (`sp-rest.js:128`);
- a `connectSeq` guard so only the latest connect applies;
- `writerFor(candidate)` = `createSpWriteClient({ client: candidate, mockWriter })` — POSTs
  build from `candidate.webUrl()` and the digest is fetched for that same URL
  (`sp-write.js:104`; `sp-files.js` caches digests per web and reuses the page digest for the
  host web only);
- **"This site" is also a connect**: the pair is built from the inspected web's URL, never
  from an unconnected client (which defaults to the hosting web).

Preflight results are bound to `(destination, fileName, folder, options, snapshot.etag)`;
any input change invalidates them synchronously (the Copy button goes dark until re-checked).
On Copy, the plan and the client pair are **frozen** into the run; later dialog input cannot
touch them.

Testing caveat: `fetchImpl` injected into `sp-write` does not reach the singleton digest
client — stubbed-live tests use `page.route`, which covers both.

### 2.4 Eligibility

`contentKind === 'canvas'` (`classic-page.js:54`) only says the canvas field is non-empty.
The Copy button is shown for canvas pages, but the dialog refuses unless:

- `JSON.parse(CanvasContent1)` succeeds and yields an array (else: legacy HTML format →
  same-web `CopyFileByPath` only, cross-site refused);
- `PageLayoutType ∈ {Article, Home, SingleWebPartAppPage}`;
- the source library is BaseTemplate 119 (`pickPagesLibrary` ranks 850 and hidden libraries
  too, pages.js:226–252 — those are not copy sources or destinations);
- on the target: the Site Pages feature is active (`MODERN_SITE_PAGES_FEATURE_ID`,
  `sp-context.js:14`; `web/features/getbyid`), and a 119 library resolves.

A digest proves authentication, not Add/Edit/Publish rights; permission failures surface at
the write that hits them (create, save, publish, asset upload), each tested separately.

## 3. The copy sequence (target side)

```
[preflight]  connect · eligibility · name/folder probe · components · fields · reference inventory
[create]     A: POST sitepages/pages {layout, PromotedState:0}     → id            (root)
             B: addTemplateFile(folder/Name.aspx, 3)               → file, item id (in place)   ← spike picks A or B
[name]       A only: checkoutpage? → savepage {Title: <stem>} → GET pages(id) (confirms FileName)
[assets]     transfer per §5.3, collect transferResults
[content]    checkoutpage (if needed) → savepage {Title: <real>, CanvasContent1', LayoutWebpartsContent', BannerImageUrl', Description, TopicHeader, AuthorByline}
[move]       A only, folder ≠ root: MoveFileByPath(root/Name → folder/Name) → GET pages(id) (refresh identity)
[metadata]   VULI carry set (§5.5) → SetCommentsDisabled if source had it
[publish]    only when ticked: publish; PromotedState set before publish when "Promote as news"
[verify]     GET pages(id) read-back → compare against the plan's expectations; report drift
```

**Naming is two saves** under path A (the first `savepage` `Title` becomes the file stem;
PnP-pages 1066–1089). Under path B the file is named at creation; the spike must show that
later saves keep the chosen name and that `getPage`/`savepage` accept the item it produces.

**Collisions are checked where each step writes**: the root (path A's create/name step) *and*
the final folder, independently. Path A stages under a **collision-safe temporary stem**
(`{name}~copy-{shortId}`) in the root, then moves to the final name with a non-overwriting
`MoveFileByPath`, so a root-level `Report.aspx` never blocks copying into `Archive/`. Another
operator claiming a name between preflight and write is reported as a move failure; the
staged page is then cleaned up (§7). Path B has no staging step, which is its attraction.

## 4. Same-web duplicate

### 4.1 Preservation matrix

| Item | Sitepages API path | `CopyFileByPath` path |
|---|---|---|
| Canvas controls incl. `instanceId`, `dynamicDataPaths/Values`, `serverProcessedContent`, page-settings slice | carried verbatim | carried (file + item) |
| `LayoutWebpartsContent` (all entries), banner, custom thumbnail flag | carried verbatim (same-web ids stay valid) | carried |
| Title, Description, `TopicHeader`, `AuthorByline`, `PageLayoutType` | carried (create/save body) | carried; Title then re-set |
| Editable custom columns (Text, Note, Choice, Number, DateTime, Boolean, URL) | VULI carry set | carried |
| User / Lookup / Taxonomy custom columns | **not carried** (editor conventions unsupported; documented limitation) | carried — the only fidelity gap between engines |
| `CommentsDisabled` | `SetCommentsDisabled` | carried |
| `PromotedState`, `FirstPublishedDate` | reset (0 unless "Promote as news") | copied — must be reset; visibility window is the spike's key question |
| Author/Editor/Created/Modified, versions, likes, comments, permissions | not carried | reset by option / not carried |

"Full fidelity" therefore means the matrix above with the User/Lookup/Taxonomy line stated
in the dialog when the API path is chosen.

### 4.2 Deciding the engine (spike)

Run both on the same published, promoted, thumbnailed, custom-column page. Compare not only
the final list items but: (a) whether the `CopyFileByPath` copy is a published news item at
any moment before the reset lands (watch `PromotedState`/`_UIVersionString`/news query
between the two calls); (b) whether it lands as a draft in a minor-versioned library; (c)
whether the copied page renders and its assets resolve for a reader with read-only rights.
`CopyFileByPath` is chosen only if (a) shows no visible window or a pre-copy neutralisation
exists; otherwise the API path with the matrix's stated limitation.

## 5. Cross-site copy

### 5.1 Reference inventory (discover) vs analyzers (authorize)

Two layers, deliberately separated:

- **Generic scanner** — walks every string in the raw canvas controls, the layout parts and
  the DTO fields and reports *candidates*: GUIDs, server-relative paths under the source web,
  absolute URLs on the source host, `getpreview.ashx` URLs with GUIDs, and HTML attributes
  (`href`, `src`, `data-*`) inside `innerHTML` (`canvas.js:109` shows rich text lives there,
  not only in `serverProcessedContent`). The scanner **never rewrites**; it produces the
  "suspicious references" list, each with its JSON path and what it looks like.
- **Analyzers** (`page-copy-analyzers.js`) — selected by web part id, they read the instance's
  actual configuration, classify each reference (asset / list-data / site-config / nav-link /
  dynamic-data) and are the only thing allowed to **patch**, using verified
  source→destination mappings from the transfer results. Unrecognized references stay
  unchanged and flagged. Any id without an analyzer gets the `default` analyzer: no patches,
  scanner findings reported as "unverified".

`WEBPART_NAMES` (`canvas.js`) is a label table only; it explicitly marks some ids as
lower-confidence and must not drive classification.

### 5.2 Per-instance verdicts

Verdicts are per reference, rolled up per web part for the report:

| Verdict | Meaning | Examples |
|---|---|---|
| **asset → copied** | file in the source web that the analyzer knows how to transfer and patch | Page header banner & custom thumbnail; Image; Image gallery; Hero tiles; Quick links thumbnails/icons; **Call to action background**; **Countdown background**; File viewer document |
| **data → warn** | list/library/site-scoped data source; kept as-is with a warning that it still points at the source (or, for context-relative configs, may resolve against the destination) | List, Document library, Events, Highlighted content, News, Site activity, **Quick chart in list mode**, Group calendar, **Sites with selected sites / hub scope** |
| **config → warn** | site-bound configuration that is neither asset nor list | Hub-relative navigation, audience targeting ids |
| **dynamic** | `dynamicDataPaths` / `dynamicDataValues` naming another instance on the page | kept; if the operator drops the provider, every consumer is listed and the drop is confirmed per consumer |
| **link** | href/src into the source web (rich text, `serverProcessedContent.links`, property URLs) | kept by default (decision 3); opt-in rewrite uses **path-boundary matching** (`/sites/source/` never matches `/sites/source-other`) and structured attribute rewriting via `DOMParser`, never string replace |
| **unavailable** | id absent from the target's `GetClientSideWebParts` | kept with a warning (decision 4); per-part drop offered; server rejection surfaces as a named failure |
| **unverified** | default analyzer | kept; scanner findings listed |

Analyzers to write for v1, each with a fixture instance: Page header (layout part), Image,
Image gallery, Hero, Quick links, Call to action, Countdown, File viewer, List/Library, Events,
Highlighted content, News, Quick chart, Sites, Text (rich-text links only). Property paths
are recorded from real pages during the spike, one fixture per shape, **before** the
analyzer is written.

### 5.3 Assets — identity and transfer contract

- **Identity:** source `(webId, listId, uniqueId)` or, failing that, the server-relative
  path, resolved to a file via `fileInfo` on the source. Preview URLs are resolved to files by
  their GUIDs (PnP-pages 506–538). Two references to the same file transfer once
  (deduplication by source identity).
- **Destination:** `SiteAssets/SitePages/{targetPageStem}/` on the target; parent folders
  created with `createFolder`; **collision-safe names** (`logo.png`, `logo-2.png`) when two
  source folders contribute the same basename; the mapping `source identity → { path, ids }`
  is recorded per file and is what analyzers patch from.
- **Transfer path:** images may use `AddImageFromExternalUrl` if the spike proves it accepts
  authenticated same-tenant URLs; otherwise, and for every non-image (File viewer PDFs etc.),
  `readFileBytes` on the source + `uploadFile` on the target (`sp-write.js:274`: no overwrite,
  50 MB cap). Oversized or unreadable assets are **not** transferred: the reference is kept
  pointing at the source and reported as "retained (reason)". A SiteAssets library requiring
  checkout is handled with the existing `checkOutFile`/`checkInFile` primitives.
- **Ordering:** assets transfer **before** the content save so the rewritten content is
  saved once with real ids; a transfer failure after the page exists is a partial outcome
  (§7), never a silent retained reference.

### 5.4 Thumbnails, page settings, layout parts

The page-settings control (`pageSettingsSlice.isDefaultThumbnail / isDefaultDescription`)
travels with the canvas. `BannerImageUrl` and `BannerThumbnailUrl` are separate: a custom
thumbnail is an asset in its own right, transferred and patched like the banner, and the save
body's `BannerImageUrl` follows PnPjs's rule (thumbnail URL when a custom thumbnail exists;
PnP-pages 371–375). Every `LayoutWebpartsContent` entry is walked, not only the header.
`TopicHeader` (DTO), `_TopicHeader` (item field) and `properties.topicHeader` (header part) are
three representations; the DTO value is authoritative and the other two are derived.

### 5.5 Metadata carry set

The carry set is computed, never inferred from `isEditable` alone (`field-editor.js:39`
admits Title and Description):

1. start from the target's fields ∩ source's fields by `InternalName` and `TypeAsString`;
2. **remove operation-owned fields**: `Title, Description, BannerImageUrl, BannerThumbnailUrl,
   PromotedState, FirstPublishedDate, PageLayoutType, CanvasContent1, LayoutWebpartsContent,
   _TopicHeader, _AuthorByline, _SPSitePageFlags, ContentTypeId, FileLeafRef, FileRef,
   Author, Editor, Created, Modified, CommentsDisabled` and everything hidden/read-only;
3. keep only types the editor conventions can round-trip (`EDITABLE_TYPES`), converting with
   `fromItemValue` → `toFormValue`, aliases via `EntityPropertyName`;
4. Choice values absent from the target's choices, and User/Lookup/Taxonomy columns, are
   skipped and listed;
5. target **required** fields with no carried value are reported before any write; the
   dialog offers a value or refuses.

## 6. The dialog

Fields: Destination (This site / Another site, with favorites and recents from `favorites.js`
and a Connect button); File name (decision 1 default, live collision check for both the
staging root and the final folder); Title; Folder; Publish (default off); Promote as news
(only when the source is promoted; default off); Carry custom metadata (default on; the carry
set and the skipped list shown); Rewrite links into the source site (default off); the
preflight report (per-part verdict chips in the info register, assets to copy with sizes,
suspicious references, warnings, required-field gaps). Copy enables only with a current
preflight; any input change darkens it.

Phase two: step log with the journal (§7), "Open copy ↗", and on failure the cleanup offer.
Cancel before the first write produces zero writes. `busy` guards double runs; Escape and ✕
are blocked while a write is in flight (the `upload-metadata.js` recipe).

## 7. Failure handling and the run journal

The runner keeps a journal: `{ pageId, currentPath, createdBy: 'this run' | 'unknown',
assets: [{ path, confirmed }], steps: [{ name, status: 'done' | 'failed' | 'unknown', detail }] }`.

- **Confirmed** resources are only those whose creating request returned success.
- **Outcome unknown**: a request that was sent but whose response was lost (network error
  after send, or a `TypeError` from `fetch`) marks the step `unknown` and stops the run. v1
  does **not** restart creation, does not delete by an inferred name, and reports: what is
  confirmed, which operation is uncertain, and where to look (the target folder URL, the
  staging name). `sp-write.js post()` retries only on 403 before any success, so a retry never
  double-creates; that is the only automatic retry and the journal notes it.
- **Cleanup offer** ("Discard copy") recycles only confirmed resources created by this run —
  the page file (by its confirmed path, `recycleFile`) and confirmed assets; `discardPage`
  alone is never treated as deletion. Cleanup failures are reported with the leftovers list.
- **Metadata** VULI is per-field, not atomic (`sp-write.js:166`): field failures are reported
  per field with the page kept.
- v2 (not in this plan): retry-from-failed-step, reconciliation by re-listing the folder.

## 8. Tests and fixtures

Fixtures must be able to *catch* retained source references, so mock identities are distinct:

- a new mock web pair owned by the page-copy suite (like `workbench-schema.mjs` owns
  `/sites/schema` and `/sites/target`): `/sites/pagesrc` and `/sites/pagedst`, plus a target on
  a different site-collection path, each with **distinct** site/web/list/file GUIDs;
- a source page fixture with: banner + distinct custom thumbnail, Image, Hero, Quick links
  with thumbnails, CTA with background, List, Quick chart (list mode), Sites, a text part with
  a link into the source web, a dynamic provider/consumer pair, one site-scoped SPFx id, one
  unknown id; custom columns `PageCategory` (exists), plus new `ReviewDate`, `ShowInNav`,
  `RelatedLink`, one required column, one User column;
- a **stateful, strict** mock: `sitepages/pages` create/read/save/publish/discard, `addTemplateFile`,
  `MoveFileByPath`, `CopyFileByPath`, `recycle`, `AddImageFromExternalUrl`, `GetClientSideWebParts`,
  file-info, name probes (404 vs 200 driven by state). Unknown page endpoints return `null`
  (404), never the generic `web` echo; the mock writer's `{ ok: true }` fallback is not relied
  on for any page-copy assertion;
- stubbed-live routes that are **projection-sensitive** (a `$select` missing a field the
  plan needs must fail the test, mirroring the existing Author/Editor 400 stub).

Checks (house style: `check(...)`, `page.evaluate(import(...))`, `page.route`):

- planner/rewriter units: naming defaults; carry-set computation (Title never carried);
  inventory vs analyzer separation (scanner finds, nothing rewritten without a mapping);
  path-boundary link matching; every analyzer's patch on its fixture; dynamic consumer
  listing on a provider drop; thumbnail vs banner; unknown-id default analyzer;
- dialog (mock): eligibility refusals (legacy HTML, repost, 850 library); connect race
  (slow A, fast B — B's report wins); edited input after a check darkens Copy; double-click;
  Esc during a write; cancel = zero writes; same-web run write sequence; cross-web run writes
  hit `/sites/pagedst/_api/…` with distinct GUIDs in the saved canvas; asset duplication
  names; oversized asset retained; required field gap refused;
- stubbed-live: two webs' `contextinfo` with distinct digests and a target-only 403 refresh;
  path A and path B sequences; root vs final-folder collision; move failure cleanup; a lost
  response marking `unknown` with no re-create; publish only when ticked; read-back drift.

Counts go into `tests/README.md` (currently 605 across suites; `workbench.mjs` 165,
`workbench-edit.mjs` 26, `workbench-schema.mjs` 138) — a new `workbench-pages-copy.mjs` suite
keeps the page-copy mock webs out of the default fixtures.

## 9. Phases and the spike

| Phase | Work | Verification |
|---|---|---|
| **0 — Spike (dev tenant)** | Headless Playwright with `pw-profile`. Questions: (1) `addTemplateFile(…,3)` in a subfolder — placement, existing-name behaviour, item identity, `getPage`/`savepage` compatibility, whether later saves keep the name, draft visibility, layouts, promoted state; (2) path A's staging + `MoveFileByPath` semantics and collision/sanitisation rules; (3) `AddImageFromExternalUrl` with an authenticated same-tenant image URL and with a non-image; (4) `savepage` with an id absent from `GetClientSideWebParts`; (5) `CopyFileByPath` on a published promoted page — the visibility transition (§4.2), minor-version landing, what fields carry; (6) `SetCommentsDisabled` shape; (7) record real property shapes for every analyzer in §5.2 from pages built for the purpose. Targets: `/sites/NewNerve/sputils-test` and `/sites/TestSiteCollection`. | Each answer read back via REST by Claude; a complete rewritten copy into the separate site collection is then **opened, rendered and its assets fetched as a read-only reader**; everything created is recycled afterwards. Results and fixture captures are appended to this document before Phase 1 starts. |
| **1 — Same-web duplicate** | `sp-pages.js`, snapshot, runner + journal, dialog, Pages-view button + dependency forwarding, mock webs and writer branches, carry set, comments. | Phase-1 checks from §8; live check on dev. |
| **2 — Cross-site** | Scanner, analyzers, asset contract, thumbnails/layout parts, preflight report, link rewrite (opt-in), metadata intersection. | Phase-2 checks; live dev web → subweb → other site collection, rendered as a reader. |
| **3 — Polish** | Legacy-HTML same-web path; per-part drop; docs (HANDOFF, CLAUDE, README counts); bundle rebuild (`tools/build-workbench.mjs`, run by `Sync-Live.ps1`). | Full suites. |

## 10. Review log

- **v1 → v2 (Codex, GPT "Astra 6" via the app, 2026-09-23; xo turns 16/17 timed out or were
  stopped).** Accepted: two-save naming (§3); snapshot instead of display item (§2.2);
  request-contract corrections (`ShouldBypassSharedLocks`, `SP.MoveCopyOptions`/`SP.ResourcePath`
  metadata, absolute paths, `SitePage` metadata on save/discard, `SetCommentsDisabled` as a
  hypothesis) (§0); root-vs-folder collisions and staging names (§3); thumbnails, page-settings
  and all layout entries, three topic-header representations (§5.4); per-instance analyzers
  with a discover-only scanner, CTA/Countdown/Sites/Quick chart reclassified (§5.1–5.2);
  rich-text links, dynamic consumers, path-boundary matching (§5.2); asset identity/transfer
  contract with non-images on the bytes path (§5.3); operation-owned field exclusion and
  required-field gaps (§5.5); fresh-client-per-connect with a sequence guard and frozen
  plan (§2.3); run journal with an explicit outcome-unknown state and confirmed-only cleanup
  (§7); distinct-identity stateful fixtures and strict stubs (§8); explicit eligibility (§2.4);
  softened wording on "stable"/"what the editor calls" (§0). Pushed back and kept: v1 does
  not implement resumable retry or automatic reconciliation (§7 v2 note). Decision 2 stays
  open with the visibility-transition test as the deciding criterion (§4.2), and
  `addTemplateFile` is added as path B for the spike (§3, §9).

## 11. Spike results (dev tenant, 2026-09-23)

Driven headlessly (Playwright, `pw-profile`) with REST from a page on the dev web; captures in
the session scratchpad `spike/captures/`. Everything created was named `zz-pagecopy-*`.

- **Formats.** The sitepages DTO (`GET sitepages/pages(id)`) returns `CanvasContent1` as a
  JSON array on every modern page on all three webs; the **list item's** `CanvasContent1` on
  the same pages is the HTML storage format (`<div data-sp-canvascontrol …>`). The snapshot
  therefore reads the DTO, never the item field, for content.
- **Create.** `POST sitepages/pages` names the file `Page.aspx` (`Url: SitePages/Page.aspx`),
  checked out to the caller. The first `savepage` Title renames it; the verbose body
  (`Content-Type: application/json;odata=verbose`, `If-Match: *`) answers `{"odata.null":true}`.
- **Q1 — `addTemplateFile(…, 3)` in a subfolder (path B).** Creates `zz-pagecopy-b1.aspx` in
  place (0.1, not checked out; a second call with the same name → 500 "The file exists").
  But `GET/checkoutpage/savepage` on it all answer **406 "This page does not have the site
  page content type. Only site pages can be served with this API."** Unusable — rejected.
- **Q2 — path A.** A staging Title `zz-pagecopy-a1~copy-abc123` becomes exactly that file
  name (`~` survives). `MoveFileByPath(overwrite=false)` of the still-checked-out page into
  `SitePages/zz-pagecopy-sub/` succeeds; the page keeps its id and its checkout. Moving onto
  an existing name → **400** "The destination file already exists." (not 409). Title →
  file-name sanitising maps `: " # % /` and spaces to `-` (`&` kept); a Title whose name is
  taken gets `(1)` (`zz-pagecopy-shapes(1).aspx`).
- **Q3 — `AddImageFromExternalUrl`.** Same-tenant URLs (an image and a PDF, to the subweb
  and to the other site collection) all fail **400 wrapping "(403) Forbidden"** — the server's
  fetch carries no user auth. Every asset takes the bytes path (`readFileBytes` +
  `AddUsingPath`).
- **Q4 — unknown web-part id.** `savepage` with id `d15c…ff` answers 200 and stores the
  control verbatim (properties intact); publish succeeds. The server never rejects an
  unknown id, so decision 4's "named failure" cannot arise from `savepage`; the preflight's
  `GetClientSideWebParts` check is the only warning an operator gets.
- **Q5 — `CopyFileByPath` on a published, promoted, custom-thumbnail page.** The copy lands
  as **0.1 draft** (minor versions on; `DraftVersionVisibility` 1) with `PromotedState` 2 and
  the source's `FirstPublishedDate`, appears in `items?$filter=PromotedState eq 2` at +1.1 s,
  keeps the canvas/layout strings byte-identical and the thumbnail, but its `Description` is
  regenerated from the body. See decision 2.
- **Description.** `savepage` with a `Description` **blanks** it (DTO and item read `null`);
  `ValidateUpdateListItem` `Description` sets it, and later Title-only saves and publish keep
  it. The run writes Description with the metadata step.
- **Thumbnail.** On a custom-thumbnail page (`pageSettingsSlice.isDefaultThumbnail: false`)
  the DTO's `BannerImageUrl` **is the thumbnail's plain URL**; `BannerThumbnailUrl` is a
  tokened `_vti_bin/afdcache.ashx/authitem/…?_oat_=…` CDN URL that must never be copied; the
  banner is only in the header part's `imageSources`/`customMetadata`. The save sends the
  DTO `BannerImageUrl` as read (mapped across webs).
- **Q6 — comments.** `POST items(id)/SetCommentsDisabled` works with a `{ "value": true }`
  body and as `SetCommentsDisabled(true)`; the **DTO's `CommentsDisabled`** reflects it, the
  list item's `CommentsDisabled` column kept reading `false`. The snapshot reads the DTO.
- **Q7 — shapes.** First-party ids from `GetClientSideWebParts` on dev: Hero `c4bd7b2f…`,
  Image `d1d91016…`, Image gallery `af8be689…`, Quick links `c70391ea…`, Call to action
  `df8e44e7…`, Countdown `62cac389…`, File and media `b7dd04e1…`, List/Document library
  `f92bf067…`, Events `20745d7d…`, Highlighted content `daf0b71c…`, News `8c88f208…`,
  Quick chart `91a50c94…`, **Sites `7cba020c…`** (canvas.js's label table calls it "Document
  library") and Organization chart `e84a8ca2…` (labelled "Sites" there). Configured instance
  shapes: pending (`zz-pagecopy-shapes.aspx` built with default instances; configuration in
  the editor requested from Joe).

## 12. Progress log

- 2026-09-23 — branch `claude/page-copy` from `main` @ `ab5a0f9`. Tenant work (Phase 0) waits for
  the live-fix session on `claude/page-status-live-fixes` to finish (semaphore
  `.PAGE-STATUS-LIVE-FIX-SESSION-DONE`); local Phase 1 work proceeds meanwhile.
- 2026-09-23 — `sp-pages.js` (§0 shapes, one call per method) + `sp-files.js readFileBytes`
  landed after a clean blind review; `files.mjs` 63/63. Next: page-copy core, mock webs,
  runner, dialog.
- 2026-09-23 — core, runner (journal §7), dialog, Pages-view Copy button and mock webs landed
  with blind-review fixes; PR #22 merged to `main` (tags `pre-pr22-merge`, `pr22-merged`) and
  merged into this branch. Spike Q1–Q6 answered (§11); decision 2 = API + path A; path B
  removed. Suites: workbench 165 (pre-merge), schema 138, edit 26, pages-copy pure 21 +
  runner 13. Next: Q7 shapes (waiting on the editor configuration), dialog + live test
  sections, same-web live check.
