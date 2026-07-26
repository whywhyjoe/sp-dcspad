# Builds the hosted bundle and copies all runtime files to the OneDrive-
# synced folder for …/SiteAssets/Code/dcspad-live/ (goes live in seconds).
# This is THE deploy command for the web-part hosting. See the cache gotcha
# in CLAUDE.md for why the bundle exists.

param(
    [string]$LivePath = "C:\Users\other\NERVE\NewNerve - Code\dcspad-live"
)

$repo = Split-Path $PSScriptRoot -Parent

Write-Host "Building dcspad.app.js…" -ForegroundColor Cyan
Push-Location (Join-Path $repo 'tools')
node build-app.mjs
if ($LASTEXITCODE -ne 0) { Pop-Location; throw "bundle build failed" }
Pop-Location

Write-Host "Copying runtime files to $LivePath" -ForegroundColor Cyan
Copy-Item (Join-Path $repo 'index.html')          $LivePath -Force
Copy-Item (Join-Path $repo 'boot.js')             $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.webpart.html') $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.app.js')       $LivePath -Force
Copy-Item (Join-Path $repo 'src')    (Join-Path $LivePath 'src')    -Recurse -Force
Copy-Item (Join-Path $repo 'styles') (Join-Path $LivePath 'styles') -Recurse -Force
Copy-Item (Join-Path $repo 'vendor') (Join-Path $LivePath 'vendor') -Recurse -Force

Write-Host "Done. OneDrive will sync momentarily; then reload the page." -ForegroundColor Green
