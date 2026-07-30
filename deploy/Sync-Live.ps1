# Builds the hosted bundle and copies all runtime files to the OneDrive-
# synced folder for …/SiteAssets/Code/dcspad-live/ (goes live in seconds).
# This is THE deploy command for the web-part hosting. See the cache gotcha
# in CLAUDE.md for why the bundle exists.

param(
    [string]$LivePath = "C:\dev\fcuportal-code\tools\dcspad",
    [string]$DesignSystemSource = "",
    [string]$FluentIconsSource = ""
)

$repo = Split-Path $PSScriptRoot -Parent
$reposRoot = Split-Path $repo -Parent
if ([string]::IsNullOrWhiteSpace($DesignSystemSource)) {
    $DesignSystemSource = Join-Path $reposRoot 'bsp-design-system'
}
if ([string]::IsNullOrWhiteSpace($FluentIconsSource)) {
    $FluentIconsSource = Join-Path $reposRoot 'bsp-fluent-icon-lib'
}
if (-not (Test-Path -LiteralPath $DesignSystemSource -PathType Container)) {
    throw "BSP design-system source folder was not found: $DesignSystemSource. Pass -DesignSystemSource <path>."
}
if (-not (Test-Path -LiteralPath $FluentIconsSource -PathType Container)) {
    throw "Fluent icon source folder was not found: $FluentIconsSource. Pass -FluentIconsSource <path>."
}
$DesignSystemSource = (Resolve-Path -LiteralPath $DesignSystemSource).Path
$FluentIconsSource = (Resolve-Path -LiteralPath $FluentIconsSource).Path

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
    (Join-Path $LivePath 'examples\examples'),
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

function Ensure-LocalEsbuild {
    param(
        [Parameter(Mandatory)] [string]$ToolsPath
    )

    Push-Location $ToolsPath
    try {
        $nodeArch = (& node -p "process.arch").Trim()
        if ($LASTEXITCODE -ne 0 -or $nodeArch -notin @('x64', 'arm64')) {
            throw "Unsupported Node architecture '$nodeArch'. Expected x64 or arm64."
        }

        $platformPackage = "@esbuild/win32-$nodeArch"
        $platformBinary = Join-Path $ToolsPath "node_modules\$platformPackage\esbuild.exe"
        $esbuildReady = Test-Path -LiteralPath $platformBinary
        if ($esbuildReady) {
            & node -e "require('esbuild').version" *> $null
            $esbuildReady = $LASTEXITCODE -eq 0
        }
        if ($esbuildReady) {
            Write-Host "Using machine-local esbuild for Node $nodeArch." -ForegroundColor DarkGray
            return
        }

        Write-Host "Installing machine-local tools for Node $nodeArch ($platformPackage)…" -ForegroundColor Yellow
        $npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
        if (-not $npm) {
            throw "npm.cmd was not found. Install npm, then run 'npm install' from $ToolsPath."
        }
        $esbuildVersion = (& node -p "require('./package.json').dependencies.esbuild || require('./package.json').devDependencies.esbuild").Trim()
        $platformSpec = "$platformPackage@$esbuildVersion"
        & $npm.Source install $platformSpec --include=optional --no-save --no-package-lock --no-audit --no-fund
        if ($LASTEXITCODE -ne 0) {
            throw "Local tools install failed. Run 'npm install' from $ToolsPath and retry."
        }

        if (-not (Test-Path -LiteralPath $platformBinary)) {
            throw "npm did not install $platformPackage. Remove $ToolsPath\node_modules, run 'npm install', and retry."
        }
        & node -e "require('esbuild').version" *> $null
        if ($LASTEXITCODE -ne 0) {
            throw "esbuild still does not match Node $nodeArch. Remove $ToolsPath\node_modules, run 'npm install', and retry."
        }
    }
    finally {
        Pop-Location
    }
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
$tools = Join-Path $repo 'tools'
Ensure-LocalEsbuild -ToolsPath $tools
Push-Location $tools
node build-design-intelligence.mjs --design-root $DesignSystemSource --fluent-icons-root $FluentIconsSource
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
Copy-DirectoryContents (Join-Path $repo 'examples') (Join-Path $LivePath 'examples')
Copy-DirectoryContents (Join-Path $repo 'vendor') (Join-Path $LivePath 'vendor')
Copy-DirectoryContents (Join-Path $repo 'lib-mirror') (Join-Path $LivePath 'lib-mirror')

$obsoleteCodeMirror = Join-Path $LivePath 'vendor\codemirror.js'
if (Test-Path -LiteralPath $obsoleteCodeMirror) {
    Remove-Item -LiteralPath $obsoleteCodeMirror -Force
    Write-Host "Removed obsolete vendor\codemirror.js from the live folder." -ForegroundColor DarkGray
}

Write-Host "Done. OneDrive will sync momentarily; then reload the page." -ForegroundColor Green
