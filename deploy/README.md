# Deploying DCSPad to SharePoint

DCSPad is pure client-side. Its supported deployment path is:

```text
repository → build → OneDrive-synced document-library folder → SharePoint
```

The current NewNerve target is already configured as the default:

```powershell
.\deploy\Sync-Live.ps1
```

To deploy to another SharePoint site or tenant, first sync that site's document
library with OneDrive and pass the local destination folder:

```powershell
.\deploy\Sync-Live.ps1 `
  -LivePath 'C:\path\to\the-synced-library\dcspad-live'
```

The destination folder must already exist. `Sync-Live.ps1` then:

1. validates the required Monaco runtime;
2. regenerates design-system intelligence;
3. rebuilds `dcspad.app.js` and `dcspad.workbench.js`;
4. copies `index.html`, `boot.js`, `dcspad.webpart.html`,
   `dcspad.config.json`, `workbench.html`, `boot-workbench.js`,
   `workbench.webpart.html`, `dcspad.workbench.js`, `src/`, `styles/`,
   and `vendor/`; and
5. leaves publication to the OneDrive sync client.

Re-running the command overwrites existing files. It does not generally remove
files deleted from the repository, so remove obsolete files from the synced
folder manually when necessary.

## Configuring another site

Before the first deployment, update the one site-specific URL in
`dcspad.webpart.html` so its script points to the new site's hosted `boot.js`:

```html
<script src="https://tenant.sharepoint.com/sites/site/SiteAssets/Code/dcspad-live/boot.js?v=1"></script>
```

Then point the Modern Script Editor web part's external Script URL at the
deployed `dcspad.webpart.html`. Bump the `?v=` value whenever `boot.js` itself
changes because SharePoint may cache library files for a day.

## Hosting the SP Workbench (second page)

The SP Workbench (site inspector) is a second entry point in the same deployed
folder. One-time setup, mirroring the pad's own hosting:

1. update the site-specific URL inside `workbench.webpart.html` so its script
   points at the hosted `boot-workbench.js` (same folder as `boot.js`);
2. create a second modern page (e.g. `SPWorkbench.aspx`), add a **Modern
   Script Editor** web part, and point its external Script URL at the deployed
   `workbench.webpart.html`;
3. bump that `?v=` value whenever `boot-workbench.js` itself changes — the
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
