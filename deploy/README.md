# Deploying DCSPad to SharePoint

DCSPad is pure client-side — deployment is "copy these files into a document
library." `Deploy-DcsPad.ps1` does that with PnP.PowerShell.

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser        # once

cd deploy
./Deploy-DcsPad.ps1 -SiteUrl https://contoso.sharepoint.com/sites/dev -ClientId <app-id>
```

Add `-WhatIf` first to see exactly what would upload without touching the site.
Re-running overwrites in place, so it doubles as the redeploy command; add
`-Clean` to delete the target folder first when you want files removed from the
repo to also disappear from SharePoint.

What ships: `index.html`, `src/`, `styles/`, `vendor/`. What doesn't: `.git`,
`tests/`, `tools/`, `deploy/`, and markdown. The script walks the repo rather
than using a fixed file list, so new modules under `src/` deploy automatically.

## Three things that commonly bite on first deploy

**1. Interactive sign-in needs your own Entra app.** PnP.PowerShell 2.x removed
the shared multi-tenant app, so `-Interactive` without a client id fails with
`AADSTS700016`. Either register one once —

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP Rocks" -Tenant contoso.onmicrosoft.com
```

— which stores a default client id for later runs, or pass `-ClientId` from an
existing app registration (needs delegated `AllSites.FullControl`, admin
consented).

**2. `.html` may download instead of render.** Tenants with strict browser file
handling, or sites without custom script enabled, will hand you the file rather
than display it. Fix: rename `index.html` → `dcspad.aspx` after upload (nothing
else changes), and/or enable custom script for the site:

```powershell
Set-PnPSite -Identity https://contoso.sharepoint.com/sites/dev -NoScriptSite $false
```

That setting also matters for the chip: DCSPad reads the host page's
`_spPageContextInfo`, which classic/script-enabled pages expose and locked-down
modern pages may not. **SP: Mock** in the top-right means it wasn't found — the
pad still runs, but `_api` calls will fail.

**3. Monaco needs both correct MIME types and an allowed same-origin worker.**
Every generated Monaco artifact uses `.js`, not `.mjs`, because the target
tenant serves `.mjs` as `application/octet-stream`. Workers use ordinary
same-origin URLs, not `blob:`. If the editor appears but language features do
not, check the status bar and the page's `worker-src` CSP. Run the spike below
to isolate that policy. If the editor itself does not appear, verify that
`vendor/monaco/monaco.js`, `monaco.css`, `assets/codicon-*.ttf`, and
`version.json` were deployed together.

## Web-part hosting spike

If the target is a modern page with a custom-script web part rather than a
standalone `.html`, run the spike in this folder **before** deploying the app —
it answers whether ES modules and a same-origin classic worker survive that
hosting model. See `HANDOFF.md` for what each answer implies.

`deploy/` is excluded from normal deploys, so upload these four by hand
(same folder as `index.html`, keeping them under `deploy/`):

```powershell
Add-PnPFile -Path webpart-spike.html -Folder "SiteAssets/dcspad/deploy"
Add-PnPFile -Path spike-module.js    -Folder "SiteAssets/dcspad/deploy"
Add-PnPFile -Path spike-import.js    -Folder "SiteAssets/dcspad/deploy"
Add-PnPFile -Path spike-worker.js    -Folder "SiteAssets/dcspad/deploy"
```

Edit `DCSPAD_SPIKE_BASE` at the top of `webpart-spike.html` first — it needs the
absolute URL of the `dcspad/` folder, with a trailing slash. Then point the web
part at `webpart-spike.html` and read the five results it prints.

## Verifying the deployment

The tenant checklist in the root `README.md` is the full version. The short one:

1. Chip reads **SP: Live**, status bar shows your web URL and user name.
2. The editor status does not report that a language worker is unavailable.
3. Enable **PnPjs v2** in the Frameworks list; in the JS editor type
   `pnp.sp.w` and confirm `web` appears in Monaco suggestions. Then run
   `pnp.sp.web.get().then(w => console.log(w))` — you should get your web's
   properties, rendered through the SP-aware inspector.
4. The Network panel shows the `_api/web` request with a 200.
