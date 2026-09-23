# Deploying DCSPad to SharePoint

DCSPad is pure client-side. Its supported deployment path is:

```text
repository → build → OneDrive-synced document-library folder → SharePoint
```

The current NewNerve target is configured as the default `dev` environment:

```powershell
.\deploy\Sync-Live.ps1
```

This reads [`deploy.settings.json`](deploy.settings.json) and copies the DCSPad
runtime to `C:\dev\fcuportal-dev\tools\dcspad`, the local
mirror for `/sites/NewNerve/FCUPortal/Dev/tools/dcspad`. It does not move shared
resources: BSP Design and Fluent Icons continue to resolve from `/Code`, and
other shared tools may continue to live under `/Code/tools`.

To deploy to BMO production, first set `prod.livePath` in
`deploy/deploy.settings.json` to that site's OneDrive-synced destination, then
run:

```powershell
.\deploy\Sync-Live.ps1 -Environment prod
```

You can override the configured destination without changing the settings file:

```powershell
.\deploy\Sync-Live.ps1 -Environment prod `
  -LivePath 'C:\alternate\synced-library\tools\dcspad' `
  -AllowLivePathOverride
```

`-Environment` is required whenever `-LivePath` is supplied. If the override
differs from the configured `livePath`, also pass `-AllowLivePathOverride`
after verifying the destination. Every environment must have a configured
`livePath` before any override is accepted, and the script always refuses to
deploy into the repository itself.

Each environment records four adjustable values:

- `livePath`: local OneDrive-synced destination;
- `siteUrl`: SharePoint web that owns `/Code` and supplies runtime context;
- `deployFolderUrl`: public URL of the deployed DCSPad folder; and
- `workbenchPageUrl`: public SP Workbench hosting page.

`deploy.settings.json` is committed, but `livePath` is a per-machine OneDrive
mirror. Put yours in `deploy/deploy.settings.local.json` — gitignored, same
shape, and merged one level deep over the committed values — instead of
committing a path that only works on your machine:

```json
{
  "environments": {
    "prod": { "livePath": "C:\\your\\synced-library\\tools\\dcspad" }
  }
}
```

The initial `prod` values assume BMO has a `/Dev/tools/dcspad` library folder,
that the Workbench page lives at `/SitePages/tools/SPWorkbench.aspx`, and that
`/sites/FCUPortal` owns the shared `/Code` dependencies. Confirm those three
assumptions before the first production deployment and adjust the values if the
production structure differs.

The repository files remain the `sourceEnvironment` version (`dev`). During
deployment, the script rewrites matching URLs only in a temporary package. It
refuses to copy that package if a source-environment URL or site path remains,
or if it contains a SharePoint host used only by another configured
environment. This covers same-tenant environments as well as different-tenant
deployments, including copied JS/HTML/JSON/CSS/Markdown/source-map/SVG/text
files.

The synced document-library mirror must already exist. If its `tools\dcspad`
destination does not exist yet, `Sync-Live.ps1` creates it. The script then:

1. validates the required Monaco runtime;
2. regenerates design-system intelligence;
3. rebuilds `dcspad.app.js` and `dcspad.workbench.js`;
4. stages `index.html`, `boot.js`, `dcspad.webpart.html`, `dcspad.app.js`,
   `dcspad.config.json`, `workbench.html`, `boot-workbench.js`,
   `workbench.webpart.html`, `dcspad.workbench.js`, `src/`, `styles/`,
   `examples/`, `vendor/`, `lib-mirror/`, and `utilities/` in a temporary package;
5. rewrites and validates the selected environment's URLs in that package
   (`vendor/` is generated and is leak-checked but not rewritten);
6. copies the validated package to the configured synced destination; and
7. leaves publication to the OneDrive sync client.

Re-running the command overwrites existing files. It does not generally remove
files deleted from the repository, so remove obsolete files from the synced
folder manually when necessary.

## Configuring another site

Add another entry under `environments` in `deploy/deploy.settings.json`, then
select it with `-Environment <name>`. There is no need to edit the checked-in
web-part HTML or boot files for each target.

For a site's one-time host-page setup, point the Modern Script Editor web part's
external Script URL at the deployed `dcspad.webpart.html`. Bump the `?v=` value
whenever `boot.js` itself changes because SharePoint may cache library files for
a day.

On NewNerve, the host pages stay outside the runtime library at
`/sites/NewNerve/SitePages/tools/DCSpad.aspx` and
`/sites/NewNerve/SitePages/tools/SPWorkbench.aspx`. Moving the runtime to
`/Dev/tools/dcspad` does not move or recreate either ASPX page.

## Hosting the SP Workbench (second page)

The SP Workbench (site inspector) is a second entry point in the same deployed
folder. One-time setup, mirroring the pad's own hosting:

1. create a second modern page (e.g. `SPWorkbench.aspx`), add a **Modern
   Script Editor** web part, and point its external Script URL at the deployed
   `workbench.webpart.html`;
2. bump that `?v=` value whenever `boot-workbench.js` itself changes — the
   same cache rule as `boot.js`; everything else the workbench loads
   (`workbench.html` no-store; `styles/app.css`, `styles/workbench.css`,
   `dcspad.workbench.js` Last-Modified-versioned) self-busts on deploy.

Verify after deploying: the workbench chip reads **SP: Live**, the Lists view
shows this web's lists including hidden ones, a known list's Fields tab
matches its real columns, Export ▸ CSV opens in Excel, a "Copy as PnPjs 2"
snippet pastes into the DCSPad JS pane and runs, and entering another
same-tenant site in the **Site** box (e.g. `/sites/ProjectName`) reloads every
view against that web (the status bar shows "inspecting …").

## Common first-deployment issues

### `.html` files may download instead of render

Tenants with strict browser file handling, or sites without custom script
enabled, may download an HTML file instead of displaying it. A site
administrator may need to enable custom scripting. For a standalone host,
renaming `index.html` to `dcspad.aspx` is another option.

DCSPad can receive its current site context through the explicit host-context
adapter documented in the root README; it does not require a complete
`_spPageContextInfo` object. **SP: Mock** means no usable SharePoint context was
found. The pad will still run, but SharePoint REST requests will fail.

### Monaco requires correct MIME types and an allowed same-origin worker

Every generated Monaco artifact uses `.js`, not `.mjs`, because the target
tenant serves `.mjs` as `application/octet-stream`. Workers use ordinary
same-origin URLs, not `blob:`. If the editor appears but language features do
not, inspect the editor status and the page's `worker-src` CSP.

If the editor itself does not appear, confirm these files deployed together:

- `vendor/monaco/monaco.js`
- `vendor/monaco/monaco.css`
- `vendor/monaco/assets/codicon-*.ttf`
- `vendor/monaco/version.json`

## Web-part hosting spike

For a new custom-script web-part environment, the files below can test whether
ES modules and a same-origin classic worker survive that hosting model:

- `deploy/webpart-spike.html`
- `deploy/spike-module.js`
- `deploy/spike-import.js`
- `deploy/spike-worker.js`

`deploy/` is not copied by `Sync-Live.ps1`. Copy those four files manually into
a `deploy/` folder under the synced DCSPad target. Before copying, set
`DCSPAD_SPIKE_BASE` at the top of `webpart-spike.html` to the absolute URL of
the hosted DCSPad folder, including the trailing slash. See `HANDOFF.md` for
how to interpret the spike.

## Verifying the deployment

After OneDrive completes the upload:

1. reload the SharePoint page and confirm the chip reads **SP: Live**;
2. confirm the status bar shows the expected web URL and user;
3. confirm the editor reports no unavailable language worker;
4. enable **PnPjs v2**, type `pnp.sp.w` in the JS editor, and confirm `web`
   appears in suggestions; and
5. run `pnp.sp.web.get().then(w => console.log(w))` and confirm the Network
   panel shows a successful `_api/web` request;
6. choose **File ▸ Import from SharePoint…**, browse the current web, select
   an HTML/CSS/JS file, and verify the replacement confirmation appears;
7. enter another site URL on the same tenant, choose **Open site**, and confirm
   its document-library boundary can be browsed; and
8. in a disposable folder, export one pane to SharePoint, then repeat with the
   same name and confirm overwrite requires a second explicit action.
