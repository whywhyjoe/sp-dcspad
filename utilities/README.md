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

The table above mirrors the `USAGE` registry at the top of the script. When you
add a function, update four places: `USAGE`, the export block, the
`SPUtilsApi` declaration in `src/intelligence/sp-utils.js`, and
`SP_UTILS_FUNCTION_NAMES` in that same file. `tests/monaco.mjs` compares the
declared names with the keys the runtime exposes, so a missed one fails the
suite.

### Deployment

`deploy/Sync-Live.ps1` stages the `utilities/` folder with the rest of the pad,
and the catalog preset resolves the script against the deployed folder, so
nothing else is needed. The file has no build step and no dependencies beyond
`pnp2`. It must not contain a tenant host name: the deploy's leak check
refuses any file naming another environment's SharePoint host.

### Known gaps

- `previewListItems` and `deleteListItems` page with `$skip`, which fails past
  the list view threshold. `exportListToExcel` already follows `@odata.nextLink`;
  the other two should do the same.
- Planned: list-copy primitives (`getListSchema`, `exportListData` to JSON,
  `createListFromSchema`, `importListData`) and a `copyList` wrapper.
