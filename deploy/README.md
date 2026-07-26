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

**3. `vendor/codemirror.mjs` depends on the MIME type SharePoint returns.**
Browsers refuse an ES module served as `application/octet-stream`. If the editors
never appear and the console shows *"Failed to load module script: Expected a
JavaScript module script…"*, that's this. Fix is two steps:

```powershell
Rename-Item vendor/codemirror.mjs codemirror.js     # in the repo
```

then change the import at the top of `src/editors.js` to `'../vendor/codemirror.js'`
and redeploy. (`.js` is always served as JavaScript; `.mjs` is not universally
mapped.) Worth checking first — if the pad loads and you can type in the editors,
this isn't an issue on your tenant.

## Verifying the deployment

The tenant checklist in the root `README.md` is the full version. The short one:

1. Chip reads **SP: Live**, status bar shows your web URL and user name.
2. Enable **PnPjs v2** in the Frameworks list, then run
   `pnp.sp.web.get().then(w => console.log(w))` — you should get your web's
   properties, rendered through the SP-aware inspector.
3. The Network panel shows the `_api/web` request with a 200.
