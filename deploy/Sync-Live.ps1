# Builds the hosted bundle and copies all runtime files to the configured
# OneDrive-synced SharePoint folder (goes live in seconds).
# This is THE deploy command for the web-part hosting. See the cache gotcha
# in CLAUDE.md for why the bundle exists.

param(
    [string]$Environment = "",
    [string]$SettingsPath = (Join-Path $PSScriptRoot 'deploy.settings.json'),
    [string]$LivePath = "",
    [switch]$AllowLivePathOverride,
    [string]$DesignSystemSource = "",
    [string]$FluentIconsSource = ""
)

$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
$reposRoot = Split-Path $repo -Parent
$environmentWasExplicit = $PSBoundParameters.ContainsKey('Environment')
$livePathWasExplicit = $PSBoundParameters.ContainsKey('LivePath')

if ($livePathWasExplicit -and -not $environmentWasExplicit) {
    throw "Pass -Environment <name> whenever you override -LivePath. This prevents deploying the default environment to another environment's folder."
}

if (-not (Test-Path -LiteralPath $SettingsPath -PathType Leaf)) {
    throw "Deployment settings were not found: $SettingsPath"
}

try {
    $deploySettings = Get-Content -Raw -LiteralPath $SettingsPath | ConvertFrom-Json
}
catch {
    throw "Deployment settings are not valid JSON: $SettingsPath. $($_.Exception.Message)"
}

# Machine-local overrides. deploy.settings.json is shared through git, but
# livePath is a per-machine OneDrive mirror; keeping yours in the gitignored
# deploy.settings.local.json avoids fighting over the committed value. Any
# environment key may be overridden, one shallow level deep.
$localSettingsPath = Join-Path (Split-Path $SettingsPath -Parent) 'deploy.settings.local.json'
if (Test-Path -LiteralPath $localSettingsPath -PathType Leaf) {
    try {
        $localSettings = Get-Content -Raw -LiteralPath $localSettingsPath | ConvertFrom-Json
    }
    catch {
        throw "Local deployment overrides are not valid JSON: $localSettingsPath. $($_.Exception.Message)"
    }
    $localEnvironments = if ($localSettings.PSObject.Properties.Name -contains 'environments') {
        $localSettings.environments
    } else { $null }
    foreach ($localEnvironment in @($localEnvironments.PSObject.Properties)) {
        $target = $deploySettings.environments.PSObject.Properties |
            Where-Object { $_.Name -eq $localEnvironment.Name } |
            Select-Object -First 1
        if (-not $target) {
            throw "$localSettingsPath overrides unknown environment '$($localEnvironment.Name)'."
        }
        foreach ($override in @($localEnvironment.Value.PSObject.Properties)) {
            $target.Value | Add-Member `
                -NotePropertyName $override.Name `
                -NotePropertyValue $override.Value `
                -Force
        }
    }
    if ($localEnvironments) {
        Write-Host "Applied machine-local overrides from $localSettingsPath" -ForegroundColor DarkGray
    }
}

if ([string]::IsNullOrWhiteSpace($Environment)) {
    $Environment = [string]$deploySettings.defaultEnvironment
}
if ([string]::IsNullOrWhiteSpace($Environment)) {
    throw "Set defaultEnvironment in $SettingsPath or pass -Environment <name>."
}

$environmentProperty = $deploySettings.environments.PSObject.Properties |
    Where-Object { $_.Name -eq $Environment } |
    Select-Object -First 1
if (-not $environmentProperty) {
    $available = ($deploySettings.environments.PSObject.Properties.Name -join ', ')
    throw "Unknown deployment environment '$Environment'. Available environments: $available"
}
$environmentConfig = $environmentProperty.Value

$sourceEnvironmentName = [string]$deploySettings.sourceEnvironment
if ([string]::IsNullOrWhiteSpace($sourceEnvironmentName)) {
    $sourceEnvironmentName = [string]$deploySettings.defaultEnvironment
}
$sourceEnvironmentProperty = $deploySettings.environments.PSObject.Properties |
    Where-Object { $_.Name -eq $sourceEnvironmentName } |
    Select-Object -First 1
if (-not $sourceEnvironmentProperty) {
    throw "The sourceEnvironment '$sourceEnvironmentName' does not exist in $SettingsPath."
}
$sourceEnvironmentConfig = $sourceEnvironmentProperty.Value

function Get-DeploymentUrl {
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$EnvironmentName
    )

    $value = [string]$Config.$Name
    $uri = $null
    if ([string]::IsNullOrWhiteSpace($value) -or
        -not [Uri]::TryCreate($value, [UriKind]::Absolute, [ref]$uri) -or
        $uri.Scheme -ne 'https') {
        throw "Environment '$EnvironmentName' must define $Name as an absolute https URL."
    }
    return $value.TrimEnd('/')
}

$sourceSiteUrl = Get-DeploymentUrl $sourceEnvironmentConfig 'siteUrl' $sourceEnvironmentName
$sourceDeployFolderUrl = Get-DeploymentUrl $sourceEnvironmentConfig 'deployFolderUrl' $sourceEnvironmentName
$sourceWorkbenchPageUrl = Get-DeploymentUrl $sourceEnvironmentConfig 'workbenchPageUrl' $sourceEnvironmentName
$siteUrl = Get-DeploymentUrl $environmentConfig 'siteUrl' $Environment
$deployFolderUrl = Get-DeploymentUrl $environmentConfig 'deployFolderUrl' $Environment
$workbenchPageUrl = Get-DeploymentUrl $environmentConfig 'workbenchPageUrl' $Environment

$configuredLivePath = [string]$environmentConfig.livePath
if ([string]::IsNullOrWhiteSpace($configuredLivePath)) {
    throw "Environment '$Environment' has no configured livePath. Set it in $SettingsPath before deploying; -LivePath is only an explicit override of an established environment destination."
}
if ([string]::IsNullOrWhiteSpace($LivePath)) {
    $LivePath = $configuredLivePath
}

$liveFullPath = [IO.Path]::GetFullPath($LivePath).TrimEnd([char[]]'\/')
$repoFullPath = [IO.Path]::GetFullPath($repo).TrimEnd([char[]]'\/')
if ([string]::Equals($liveFullPath, $repoFullPath, [StringComparison]::OrdinalIgnoreCase) -or
    $liveFullPath.StartsWith($repoFullPath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to deploy into the repository: $liveFullPath"
}

if ($livePathWasExplicit) {
    $configuredLiveFullPath = [IO.Path]::GetFullPath($configuredLivePath).TrimEnd([char[]]'\/')
    if (-not [string]::Equals($liveFullPath, $configuredLiveFullPath, [StringComparison]::OrdinalIgnoreCase) -and
        -not $AllowLivePathOverride) {
        throw "The -LivePath override does not match environment '$Environment'. Use its configured livePath or pass -AllowLivePathOverride after verifying the destination."
    }
}
$LivePath = $liveFullPath

$sourceSitePath = ([Uri]$sourceSiteUrl).AbsolutePath.TrimEnd('/')
$sitePath = ([Uri]$siteUrl).AbsolutePath.TrimEnd('/')
$urlReplacements = @(
    [pscustomobject]@{ From = $sourceWorkbenchPageUrl; To = $workbenchPageUrl },
    [pscustomobject]@{ From = $sourceDeployFolderUrl; To = $deployFolderUrl },
    [pscustomobject]@{ From = $sourceSiteUrl; To = $siteUrl },
    [pscustomobject]@{ From = $sourceSitePath; To = $sitePath }
) | Sort-Object { $_.From.Length } -Descending

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

# Turndown backs the HTML->Markdown exports. It is bundled into
# dcspad.workbench.js, so a missing copy fails the esbuild step below rather
# than the deploy - but standalone index.html and the test suites import the
# vendored file directly, so it must be present and shipped either way.
$turndown = Join-Path $repo 'vendor\turndown\turndown.js'
if (-not (Test-Path $turndown)) {
    throw "Missing vendor\turndown\turndown.js. Run npm run build:vendor from tools first."
}

$createLiveDestination = $false
if (Test-Path -LiteralPath $LivePath) {
    if (-not (Test-Path -LiteralPath $LivePath -PathType Container)) {
        throw "Live destination exists but is not a folder: $LivePath"
    }
    $LivePath = (Resolve-Path -LiteralPath $LivePath).Path
    if ([string]::Equals($LivePath, $repoFullPath, [StringComparison]::OrdinalIgnoreCase) -or
        $LivePath.StartsWith($repoFullPath + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to deploy into the repository: $LivePath"
    }
    # Path-string comparison stops at reparse points, and neither GetFullPath
    # nor Resolve-Path follows a junction. Checking for source-tree markers
    # catches any checkout — this one reached through a junction, or another
    # clone entirely — because the copy below would overwrite working files
    # with environment-rewritten ones.
    foreach ($marker in @('.git', 'deploy\Sync-Live.ps1', 'tools\build-app.mjs')) {
        if (Test-Path -LiteralPath (Join-Path $LivePath $marker)) {
            throw "Refusing to deploy over a source checkout (found $marker): $LivePath"
        }
    }
}
else {
    $liveMirrorRoot = Split-Path (Split-Path $LivePath -Parent) -Parent
    if (-not (Test-Path -LiteralPath $liveMirrorRoot -PathType Container)) {
        throw "The synced library mirror was not found: $liveMirrorRoot"
    }
    $createLiveDestination = $true
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
            # Deliberately unredirected. Windows PowerShell 5.1 turns
            # *redirected* native stderr into a terminating error under
            # $ErrorActionPreference = 'Stop', which would abort the deploy
            # before this exit-code probe could route to the repair install
            # below. Unredirected stderr does not. The probe prints nothing
            # when esbuild loads cleanly, so there is no noise to suppress.
            & node -e "require('esbuild').version"
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
        # Unredirected for the same 5.1 reason as the probe above — this one
        # exists purely to reach the actionable message below.
        & node -e "require('esbuild').version"
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

function Get-DeploymentValuePattern {
    param(
        [Parameter(Mandatory)] [string]$Value
    )

    # Match only at a URL/path boundary. In particular, do not rewrite a
    # sibling such as FCUPortalArchive or dcspad-legacy.
    return [regex]::Escape($Value) + '(?=$|[/?#&="'']|\s|[<`)\]},;])'
}

function Set-DeploymentUrls {
    param(
        [Parameter(Mandatory)] [string[]]$Files,
        [Parameter(Mandatory)] [object[]]$Replacements
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $changedFiles = New-Object System.Collections.Generic.List[string]
    foreach ($file in $Files) {
        $before = [IO.File]::ReadAllText($file)
        $after = $before
        foreach ($replacement in $Replacements) {
            if ([string]::Equals(
                [string]$replacement.From,
                [string]$replacement.To,
                [StringComparison]::OrdinalIgnoreCase
            )) { continue }
            # Ordinal pre-filter. The boundary pattern can only match where
            # the literal already occurs, and IndexOf is far cheaper than a
            # case-insensitive regex sweep of a multi-megabyte file.
            if ($after.IndexOf(
                [string]$replacement.From,
                [StringComparison]::OrdinalIgnoreCase
            ) -lt 0) { continue }
            $pattern = Get-DeploymentValuePattern ([string]$replacement.From)
            $replacementText = ([string]$replacement.To).Replace('$', '$$')
            $after = [regex]::Replace(
                $after,
                $pattern,
                $replacementText,
                [Text.RegularExpressions.RegexOptions]::IgnoreCase
            )
        }
        if ($after -ceq $before) { continue }
        [IO.File]::WriteAllText($file, $after, $utf8NoBom)
        $changedFiles.Add($file)
    }
    return $changedFiles.ToArray()
}

function Assert-DeploymentPackage {
    param(
        [Parameter(Mandatory)] [string]$StagingPath,
        [Parameter(Mandatory)] [string[]]$TopLevelFiles,
        [Parameter(Mandatory)] [string[]]$DirectoryNames
    )

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($relativePath in $TopLevelFiles) {
        if (-not (Test-Path -LiteralPath (Join-Path $StagingPath $relativePath) -PathType Leaf)) {
            $missing.Add($relativePath)
        }
    }
    foreach ($directoryName in $DirectoryNames) {
        $directoryPath = Join-Path $StagingPath $directoryName
        if (-not (Test-Path -LiteralPath $directoryPath -PathType Container) -or
            -not (Get-ChildItem -LiteralPath $directoryPath -Recurse -File | Select-Object -First 1)) {
            $missing.Add("$directoryName\ (missing or empty)")
        }
    }
    if ($missing.Count -gt 0) {
        throw "Deployment package is incomplete: $($missing -join ', ')"
    }
}

function Assert-NoEnvironmentUrlLeaks {
    param(
        [Parameter(Mandatory)] [string[]]$Files,
        [Parameter(Mandatory)] [object[]]$Replacements,
        [Parameter(Mandatory)] [object]$Settings,
        [Parameter(Mandatory)] [string]$SelectedEnvironment,
        [Parameter(Mandatory)] [object]$SelectedConfig
    )

    $sourcePatterns = $Replacements |
        Where-Object {
            -not [string]::Equals(
                [string]$_.From,
                [string]$_.To,
                [StringComparison]::OrdinalIgnoreCase
            )
        } |
        ForEach-Object {
            [pscustomobject]@{
                Label = [string]$_.From
                Pattern = Get-DeploymentValuePattern ([string]$_.From)
            }
        }

    $selectedHosts = @('siteUrl', 'deployFolderUrl', 'workbenchPageUrl') |
        ForEach-Object {
            $candidate = $null
            if ([Uri]::TryCreate([string]$SelectedConfig.$_, [UriKind]::Absolute, [ref]$candidate)) {
                $candidate.Host
            }
        } |
        Sort-Object -Unique

    $foreignHosts = $Settings.environments.PSObject.Properties |
        Where-Object { $_.Name -ne $SelectedEnvironment } |
        ForEach-Object {
            $otherEnvironment = $_.Value
            @('siteUrl', 'deployFolderUrl', 'workbenchPageUrl') | ForEach-Object {
                $candidate = $null
                if ([Uri]::TryCreate([string]$otherEnvironment.$_, [UriKind]::Absolute, [ref]$candidate)) {
                    $candidate.Host
                }
            }
        } |
        Where-Object { $_ -and $selectedHosts -notcontains $_ } |
        Sort-Object -Unique

    $leaks = New-Object System.Collections.Generic.List[string]
    foreach ($file in $Files) {
        $content = [IO.File]::ReadAllText($file)
        foreach ($sourcePattern in $sourcePatterns) {
            # Same ordinal pre-filter as the rewrite pass; vendor/ is scanned
            # here in full and is ~15 MB.
            if ($content.IndexOf(
                $sourcePattern.Label,
                [StringComparison]::OrdinalIgnoreCase
            ) -lt 0) { continue }
            if ([regex]::IsMatch(
                $content,
                $sourcePattern.Pattern,
                [Text.RegularExpressions.RegexOptions]::IgnoreCase
            )) {
                $leaks.Add("$file ($($sourcePattern.Label))")
            }
        }
        foreach ($hostName in $foreignHosts) {
            if ($content.IndexOf("https://$hostName", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $leaks.Add("$file ($hostName)")
            }
        }
    }
    if ($leaks.Count -gt 0) {
        throw "Deployment still contains URL or path values for another SharePoint environment: $($leaks -join ', ')"
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

$deploymentTextExtensions = @('.css', '.html', '.js', '.json', '.map', '.md', '.mjs', '.svg', '.ts', '.txt')
$topLevelDeploymentFiles = @(
    'index.html',
    'boot.js',
    'dcspad.webpart.html',
    'dcspad.app.js',
    'dcspad.config.json',
    'workbench.html',
    'boot-workbench.js',
    'workbench.webpart.html',
    'dcspad.workbench.js'
)
$deploymentDirectoryNames = @('src', 'styles', 'examples', 'vendor', 'lib-mirror')

$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$stagingPath = Join-Path $tempRoot ("dcspad-deploy-" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stagingPath -Force | Out-Null

try {
    Write-Host "Preparing $Environment deployment package…" -ForegroundColor Cyan
    foreach ($relativePath in $topLevelDeploymentFiles) {
        Copy-Item (Join-Path $repo $relativePath) $stagingPath -Force
    }
    foreach ($directoryName in $deploymentDirectoryNames) {
        Copy-DirectoryContents `
            (Join-Path $repo $directoryName) `
            (Join-Path $stagingPath $directoryName)
    }
    Assert-DeploymentPackage `
        -StagingPath $stagingPath `
        -TopLevelFiles $topLevelDeploymentFiles `
        -DirectoryNames $deploymentDirectoryNames

    $deploymentTextFiles = Get-ChildItem -LiteralPath $stagingPath -Recurse -File |
        Where-Object { $deploymentTextExtensions -contains $_.Extension.ToLowerInvariant() } |
        Select-Object -ExpandProperty FullName

    # vendor/ is wholly generated (Monaco runtime, PnPjs declarations, design
    # intelligence) and cannot carry an environment URL, but it is ~15 MB and
    # rewriting it dominated the deploy. It is still scanned by the leak check
    # below, so if one ever does appear the deploy fails loudly rather than
    # shipping a source-environment URL — at which point drop it from here.
    $rewriteExcludedRoots = @('vendor')
    $excludedPrefixes = @($rewriteExcludedRoots | ForEach-Object {
        (Join-Path $stagingPath $_) + [IO.Path]::DirectorySeparatorChar
    })
    $rewriteTextFiles = @($deploymentTextFiles | Where-Object {
        $candidate = $_
        -not ($excludedPrefixes | Where-Object {
            $candidate.StartsWith($_, [StringComparison]::OrdinalIgnoreCase)
        })
    })

    $urlAdjustedFiles = Set-DeploymentUrls -Files $rewriteTextFiles -Replacements $urlReplacements
    Assert-NoEnvironmentUrlLeaks `
        -Files $deploymentTextFiles `
        -Replacements $urlReplacements `
        -Settings $deploySettings `
        -SelectedEnvironment $Environment `
        -SelectedConfig $environmentConfig

    Write-Host "Environment: $Environment" -ForegroundColor DarkGray
    Write-Host "SharePoint site: $siteUrl" -ForegroundColor DarkGray
    Write-Host "Hosted folder: $deployFolderUrl" -ForegroundColor DarkGray
    Write-Host "Workbench page: $workbenchPageUrl" -ForegroundColor DarkGray
    if ($urlAdjustedFiles.Count -gt 0) {
        Write-Host "Adjusted environment URLs in $($urlAdjustedFiles.Count) packaged file(s)." -ForegroundColor DarkGray
    }

    if ($createLiveDestination) {
        New-Item -ItemType Directory -Path $LivePath -Force | Out-Null
        Write-Host "Created live destination $LivePath" -ForegroundColor DarkGray
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

    Write-Host "Copying validated runtime files to $LivePath" -ForegroundColor Cyan
    foreach ($relativePath in $topLevelDeploymentFiles) {
        Copy-Item (Join-Path $stagingPath $relativePath) $LivePath -Force
    }
    foreach ($directoryName in $deploymentDirectoryNames) {
        Copy-DirectoryContents `
            (Join-Path $stagingPath $directoryName) `
            (Join-Path $LivePath $directoryName)
    }
}
finally {
    if (Test-Path -LiteralPath $stagingPath) {
        $resolvedStagingPath = [IO.Path]::GetFullPath($stagingPath)
        $safeStageName = [IO.Path]::GetFileName($resolvedStagingPath).StartsWith('dcspad-deploy-')
        $resolvedStageParent = [IO.Path]::GetFullPath(
            [IO.Path]::GetDirectoryName($resolvedStagingPath) + [IO.Path]::DirectorySeparatorChar
        )
        $safeStageParent = [string]::Equals(
            $resolvedStageParent,
            $tempRoot,
            [StringComparison]::OrdinalIgnoreCase
        )
        if ($safeStageName -and $safeStageParent) {
            Remove-Item -LiteralPath $resolvedStagingPath -Recurse -Force
        }
        else {
            Write-Warning "Refusing to remove unexpected staging path: $resolvedStagingPath"
        }
    }
}

$obsoleteCodeMirror = Join-Path $LivePath 'vendor\codemirror.js'
if (Test-Path -LiteralPath $obsoleteCodeMirror) {
    Remove-Item -LiteralPath $obsoleteCodeMirror -Force
    Write-Host "Removed obsolete vendor\codemirror.js from the live folder." -ForegroundColor DarkGray
}

Write-Host "Done. OneDrive will sync momentarily; then reload the $Environment page." -ForegroundColor Green
