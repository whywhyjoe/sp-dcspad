# Builds the hosted bundle and copies all runtime files to the OneDrive-
# synced folder for …/SiteAssets/Code/dcspad-live/ (goes live in seconds).
# This is THE deploy command for the web-part hosting. See the cache gotcha
# in CLAUDE.md for why the bundle exists.

param(
    [string]$LivePath = "C:\Users\other\NERVE\NewNerve - Code\dcspad-live"
)

$repo = Split-Path $PSScriptRoot -Parent
$monaco = Join-Path $repo 'vendor\monaco'
$requiredMonaco = @(
    'version.json',
    'monaco.js',
    'monaco.css',
    'editor.worker.js',
    'css.worker.js',
    'html.worker.js',
    'ts.worker.js',
    'pnpjs-types.json'
)

foreach ($file in $requiredMonaco) {
    if (-not (Test-Path (Join-Path $monaco $file))) {
        throw "Missing vendor\monaco\$file. Run npm run build:monaco from tools first."
    }
}

if (-not (Get-ChildItem (Join-Path $monaco 'assets') -Filter 'codicon-*.ttf' -File -ErrorAction SilentlyContinue)) {
    throw "Missing vendor\monaco\assets\codicon-*.ttf. Run npm run build:monaco from tools first."
}

$liveRoot = (Resolve-Path -LiteralPath $LivePath).Path.TrimEnd('\')
$accidentalNestedFolders = @(
    (Join-Path $LivePath 'src\src'),
    (Join-Path $LivePath 'styles\styles'),
    (Join-Path $LivePath 'vendor\vendor')
)
foreach ($folder in $accidentalNestedFolders) {
    if (-not (Test-Path -LiteralPath $folder)) { continue }
    $resolved = (Resolve-Path -LiteralPath $folder).Path
    if (-not $resolved.StartsWith($liveRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove unexpected path outside the live folder: $resolved"
    }
    Remove-Item -LiteralPath $resolved -Recurse -Force
    Write-Host "Removed accidental nested folder $resolved" -ForegroundColor DarkGray
}

function Copy-DirectoryContents {
    param(
        [Parameter(Mandatory)] [string]$Source,
        [Parameter(Mandatory)] [string]$Destination
    )
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    Get-ChildItem -LiteralPath $Source | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $Destination -Recurse -Force
    }
}

Write-Host "Building design-system intelligence…" -ForegroundColor Cyan
Push-Location (Join-Path $repo 'tools')
node build-design-intelligence.mjs
if ($LASTEXITCODE -ne 0) { Pop-Location; throw "design-system intelligence build failed" }

Write-Host "Building dcspad.app.js…" -ForegroundColor Cyan
node build-app.mjs
if ($LASTEXITCODE -ne 0) { Pop-Location; throw "bundle build failed" }

Write-Host "Building dcspad.workbench.js…" -ForegroundColor Cyan
node build-workbench.mjs
if ($LASTEXITCODE -ne 0) { Pop-Location; throw "workbench bundle build failed" }
Pop-Location

Write-Host "Copying runtime files to $LivePath" -ForegroundColor Cyan
Copy-Item (Join-Path $repo 'index.html')            $LivePath -Force
Copy-Item (Join-Path $repo 'boot.js')               $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.webpart.html')   $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.app.js')         $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.config.json')    $LivePath -Force
Copy-Item (Join-Path $repo 'workbench.html')        $LivePath -Force
Copy-Item (Join-Path $repo 'boot-workbench.js')     $LivePath -Force
Copy-Item (Join-Path $repo 'workbench.webpart.html') $LivePath -Force
Copy-Item (Join-Path $repo 'dcspad.workbench.js')   $LivePath -Force
Copy-DirectoryContents (Join-Path $repo 'src')    (Join-Path $LivePath 'src')
Copy-DirectoryContents (Join-Path $repo 'styles') (Join-Path $LivePath 'styles')
Copy-DirectoryContents (Join-Path $repo 'vendor') (Join-Path $LivePath 'vendor')

$obsoleteCodeMirror = Join-Path $LivePath 'vendor\codemirror.js'
if (Test-Path -LiteralPath $obsoleteCodeMirror) {
    Remove-Item -LiteralPath $obsoleteCodeMirror -Force
    Write-Host "Removed obsolete vendor\codemirror.js from the live folder." -ForegroundColor DarkGray
}

Write-Host "Done. OneDrive will sync momentarily; then reload the page." -ForegroundColor Green
