# Utilities

Console-first scripts for SharePoint administration. They are not apps in the
L1/L2 sense: no shell, no routing, no design-system markup. Each one is a single
self-hosted file that exposes a global object, writes progress to an on-page
status panel, and is driven from the browser console or from DCSPad's JS pane.

| File | Global | Purpose |
| --- | --- | --- |
| [`dcspad-sp-utilities.js`](dcspad-sp-utilities.js) | `window.SPUtils` | List inspection, schema export, batch item delete, column creation, Excel export, group membership from CSV, sharing links. Built on PnPjs 2. |

## dcspad-sp-utilities.js

### Requirements

- **PnPjs 2** on the page as the global `pnp2` (the tenant's self-hosted
  `pnp2.bundle.js`, 2.15.0). The script does not load it for you.
- **ExcelJS** only for `exportListToExcel`. Everything else works without it.
- No page markup. The script draws its own status panel on first use. If the
  page already has an element with `id="status"`, that element is reused and
  nothing is drawn. To control placement, put `data-sputils-host` on any
  element and the panel mounts inside it.

### Where it runs

**DCSPad (recommended).** The framework catalog ships a **DCSPad SP Utilities
(SPUtils)** entry directly below PnPjs. Tick both and the script is injected
into every run; the JS editor gains completion and hover documentation for
`SPUtils.*` from `src/intelligence/sp-utils.js`. Then, in the JS pane:

```js
SPUtils.setupContext("sites/YourSite");
await SPUtils.getAllLists();
```

**Script Editor web part.** Load `pnp2.bundle.js`, then this file, in one
script web part. The old status markup is no longer needed. Drive it from the
browser console.

**Browser console on any tenant page.** Paste the PnPjs bundle and this file
into the console, then call functions on `window.SPUtils`.

### Quick start

1. `SPUtils.setupContext("sites/YourSite")` — always first. It points PnPjs at
   the site every later call uses.
2. `SPUtils.help()` — prints every function to the console as a table and opens
   the usage panel on the page.
3. Most functions are `async`; prefix them with `await`.
4. Anything that writes has a `dryRun` option. Run with `dryRun: true` first.

### Functions

| Function | What it does |
| --- | --- |
| `setupContext(siteSubUrl)` | Point PnPjs at a site. Call first. |
| `help()` | List functions in the console and open the usage panel. |
| `clearStatus()` | Empty the on-page status log. |
| `getAllLists()` | Every list on the site (Id, Title). |
| `getListFields(listTitle)` | Visible fields with type, required, read-only. Returns the array. |
| `getListSchemaForMigration(listTitle)` | Settings, fields, views and content types as one object. |
| `previewListItems(listTitle, batchSize)` | All items (Id, Title) as a table. |
| `deleteListItems(listTitle, { dryRun, filterFn, batchSize })` | Batch-delete items, optionally filtered by a predicate. |
| `addFieldsToList(listServerRelativeUrl, specs, { dryRun })` | Create columns from a JSON spec. Skips existing columns. |
| `exportListToExcel(listTitle, { fileName })` | Download every item as `.xlsx`. Needs ExcelJS. |
| `getAllSecurityGroups()` | Site groups (Id, Title). |
| `getSiteMembersWithGroups()` | Each member and the groups they belong to. |
| `getSiteGroupMembers()` | Each group with its members, job title and department. |
| `addUsersFromCSVToGroups(filePath, { dryRun, concurrency })` | Add users to groups from a CSV in a library. |
| `syncUsersFromCSVToGroups(filePath, { dryRun, adminEmail, umbrellaGroup })` | Make membership match a master CSV. Adds **and removes**. `dryRun` defaults to `true`. |
| `getSharingLinkForItem(filePath, canEdit, expireInDays)` | Create a view or edit sharing link. `0` days means no expiry. |
| `removeAllSharingLinksForItem(filePath, areYouSure)` | Revoke every sharing link on a file. Second argument must be `true`. |
| `getListSchema(listTitle)` | **Copy step 1.** Settings, fields (SchemaXml, lookup targets, formatting), views, content types. No items. |
| `exportListData(listTitle, { includeAttachments, schema })` | **Copy step 2.** Every item as JSON with people and lookups resolved. Attachments opt-in. |
| `createListFromSchema(schema, { title, description, dryRun, lookupMap })` | **Copy step 3.** Create the list, or add missing columns and views to an existing one. |
| `importListData(data, { listTitle, dryRun, concurrency, preserveAuthorship, includeAttachments, idMap })` | **Copy step 4.** Create the items. Self-referencing lookups are filled in a second pass. |
| `copyList(sourceTitle, { targetTitle, targetSite, dryRun, … })` | Steps 1 to 4 chained, optionally across sites. |
| `downloadJson(object, fileName)` | Save a schema or data document as a `.json` download. |
| `readJsonFile(filePath)` | Read a `.json` document from a library on the current site. |

The table above mirrors the `USAGE` registry at the top of the script. When you
add a function, update four places: `USAGE`, the export block, the
`SPUtilsApi` declaration in `src/intelligence/sp-utils.js`, and
`SP_UTILS_FUNCTION_NAMES` in that same file. `tests/monaco.mjs` compares the
declared names with the keys the runtime exposes, so a missed one fails the
suite.

### Copying a list

SharePoint has no good built-in way to copy a list with its data. The four
primitives do it in steps that each support `dryRun`; any step can be rerun
on its own, and a partial import resumes from its report (see below):

```js
SPUtils.setupContext("sites/Source");
const schema = await SPUtils.getListSchema("Requests");
const data   = await SPUtils.exportListData("Requests", { includeAttachments: true });

SPUtils.setupContext("sites/Target");
await SPUtils.createListFromSchema(schema, { dryRun: true });   // read the plan
await SPUtils.createListFromSchema(schema);
await SPUtils.importListData(data, { dryRun: true });           // sample payloads
await SPUtils.importListData(data, { preserveAuthorship: true });
```

or, in one call, `SPUtils.copyList("Requests", { targetSite: "sites/Target" })`.
For a later session, `downloadJson(schema)` / `downloadJson(data)` and drop the
files in a library on the target site, then `readJsonFile()` them back.

What it handles and what it does not:

- **Columns.** Custom columns (added by the list author, not inherited) are
  recreated from their SchemaXml with source-specific attributes stripped, in
  the order plain → lookup → calculated. The Title column's display name and
  required flag are applied. Column and view formatting come across.
- **Lookups.** Targets are recorded by list title and re-resolved on the
  target site; pass `lookupMap` when the target list is named differently.
  Values are matched by the shown value; a value that is missing or appears
  more than once on the target is reported as a field error and left empty
  (an ambiguous match is accepted only when one candidate carries the source
  id, as on a same-site copy). Self-referencing lookups are written in a
  second import pass through the old-to-new id map, and a reference to an
  item that failed to import is reported. Dependent (secondary) lookup
  columns are rebound to their primary column's new id.
- **Folders.** Folder rows are recreated first and items are created inside
  their original folder path.
- **Existing targets.** A column that exists with a different type is
  reported, not silently reused; attachments and folder creation are switched
  on when the source needs them; other settings of an existing list are left
  alone.
- **Dry runs** create nothing and resolve people without `ensureUser`, so
  they do not touch the target site's user list either.
- **Resuming.** Rerunning `importListData` creates every row again. To finish
  a partial import, pass the previous report's `idMap`: source items already
  in it are skipped and the lookup and authorship passes still run for them.
  ```js
  const r1 = await SPUtils.importListData(data, { listTitle: "Requests" });
  const r2 = await SPUtils.importListData(data, { listTitle: "Requests", idMap: r1.idMap });
  ```
- **Folders** carry their metadata and authorship like items. A folder whose
  creation is rejected is reported as failed, not counted.
- **Dependent (secondary) lookup columns** are created through SharePoint's
  dependent-lookup endpoint against the primary column's new id.
- **Regional and time-zone reads** are retried and never silently defaulted:
  if they fail after retries the affected dates fail with a clear message.
- **People.** Exported with email and login, re-resolved with `ensureUser`.
- **Values.** Written with the `ValidateUpdateListItem` string conventions,
  so text, choice, multi-choice, number, date, yes/no, URL, person, lookup
  and managed-metadata values all go through one path and the server reports
  per-field errors in the returned report.
- **Authorship.** `preserveAuthorship: true` keeps Created, Modified, Author
  and Editor, written again after attachments and lookup passes so later
  writes do not overwrite them. An item whose source author is an app or
  system principal gets the importing user instead.
- **Order.** Items are created one at a time by default so new ids ascend
  like the old ones. `concurrency` above 1 is faster and scrambles that.
- **Dates.** `ValidateUpdateListItem` refuses ISO dates. Values are converted
  to the target web's time zone (via the server's `utcToLocalTime`, exact on
  DST-transition days) and written in the web's date order with a 24-hour
  clock. The order is learned from the server's own validation sample, which
  was verified on en-US, en-GB, en-CA, fr-FR and ja-JP webs, DateOnly columns
  included.
- **Throttling.** 429 and 503 responses are retried with backoff, honouring
  `Retry-After`; the request digest refresh is shared between workers and
  retried the same way.
- **Not handled.** Document libraries (that is a file copy), content types
  (exported for reference only), creating managed-metadata columns (needs a
  term-set binding), item-level permissions, version history.

### Deployment

`deploy/Sync-Live.ps1` stages the `utilities/` folder with the rest of the pad,
and the catalog preset resolves the script against the deployed folder, so
nothing else is needed. The file has no build step and no dependencies beyond
`pnp2`. It must not contain a tenant host name: the deploy's leak check
refuses any file naming another environment's SharePoint host.

### Known gaps

- The list-copy primitives were verified live on the dev tenant (same-site
  and cross-site into a subsite) with choice, multi-choice, person, date,
  number, note, URL, yes/no, lookup, self-lookup and calculated columns, a
  formatted column, a formatted view, an attachment, and a preserved Created
  date. Managed-metadata columns and document libraries remain untested.
