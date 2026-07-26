<#
.SYNOPSIS
    Uploads DCSPad's runtime files to a SharePoint document library with PnP.PowerShell.

.DESCRIPTION
    Walks the repository, skips everything that is development-only (.git, tests,
    tools, markdown), and uploads what remains — preserving folder structure —
    into <Library>/<FolderName> on the target site.

    Overwrites existing files by default, so re-running it is the normal way to
    redeploy. Files deleted from the repo are NOT removed from SharePoint unless
    you pass -Clean.

.PARAMETER SiteUrl
    Absolute URL of the target site, e.g. https://contoso.sharepoint.com/sites/dev

.PARAMETER Library
    Document library to deploy into. Defaults to SiteAssets, which exists on
    every modern site and is not indexed for search results.

.PARAMETER FolderName
    Folder inside the library. Defaults to dcspad.

.PARAMETER ClientId
    Entra ID application (client) ID for interactive sign-in. PnP.PowerShell 2.x
    no longer ships a shared multi-tenant app, so this is required unless you
    have already run Register-PnPEntraIDAppForInteractiveLogin (which stores a
    default client id for you).

.PARAMETER SourcePath
    Repository root. Defaults to the parent of this script's folder.

.PARAMETER Clean
    Delete the target folder before uploading, so removed files don't linger.
    Destructive — prompts for confirmation.

.PARAMETER SkipConnect
    Reuse an existing PnP connection instead of calling Connect-PnPOnline.

.EXAMPLE
    ./Deploy-DcsPad.ps1 -SiteUrl https://contoso.sharepoint.com/sites/dev -ClientId 00000000-1111-2222-3333-444444444444

.EXAMPLE
    ./Deploy-DcsPad.ps1 -SiteUrl https://contoso.sharepoint.com/sites/dev -Clean -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][string]$SiteUrl,
    [string]$Library = 'SiteAssets',
    [string]$FolderName = 'dcspad',
    [string]$ClientId,
    [string]$SourcePath = (Split-Path -Parent $PSScriptRoot),
    [switch]$Clean,
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw 'PnP.PowerShell is not installed. Install-Module PnP.PowerShell -Scope CurrentUser'
}

$root = (Resolve-Path $SourcePath).Path
if (-not (Test-Path (Join-Path $root 'index.html'))) {
    throw "No index.html under '$root' — point -SourcePath at the repository root."
}

# Development-only paths. Everything else is treated as runtime and uploaded,
# so a new src/ file is picked up automatically without editing this script.
$excludedDirs = @('.git', '.github', 'tests', 'tools', 'node_modules', 'deploy')
$excludedFiles = @('.gitignore', '.DS_Store')

$files = Get-ChildItem -Path $root -Recurse -File | Where-Object {
    $relative = $_.FullName.Substring($root.Length).TrimStart([IO.Path]::DirectorySeparatorChar)
    $segments = $relative -split '[\\/]'
    # Directory segments only. A root-level file has none — slicing
    # 0..-1 there would wrap around and test the file's own name.
    $dirSegments = if ($segments.Length -gt 1) { $segments[0..($segments.Length - 2)] } else { @() }
    $inExcludedDir = $dirSegments | Where-Object { $excludedDirs -contains $_ }
    -not $inExcludedDir -and
    $excludedFiles -notcontains $_.Name -and
    $_.Extension -ne '.md'
}

if (-not $files) { throw "Found no files to deploy under '$root'." }

$targetRoot = "$Library/$FolderName"
Write-Host "DCSPad deploy" -ForegroundColor Cyan
Write-Host "  source : $root"
Write-Host "  target : $SiteUrl/$targetRoot"
Write-Host "  files  : $($files.Count)"

if (-not $SkipConnect) {
    $connectArgs = @{ Url = $SiteUrl; Interactive = $true }
    if ($ClientId) { $connectArgs['ClientId'] = $ClientId }
    Connect-PnPOnline @connectArgs
}

if ($Clean) {
    $existing = Get-PnPFolder -Url $targetRoot -ErrorAction SilentlyContinue
    if ($existing -and $PSCmdlet.ShouldProcess("$targetRoot", 'Delete folder and all contents')) {
        Remove-PnPFolder -Name $FolderName -Folder $Library -Force
        Write-Host "  cleaned existing $targetRoot" -ForegroundColor Yellow
    }
}

# Resolve-PnPFolder creates missing folders; cache so nested paths cost one call each.
$ensured = [System.Collections.Generic.HashSet[string]]::new()
function Confirm-Folder([string]$siteRelativePath) {
    if ($ensured.Contains($siteRelativePath)) { return }
    Resolve-PnPFolder -SiteRelativePath $siteRelativePath | Out-Null
    [void]$ensured.Add($siteRelativePath)
}

$uploaded = 0
$failed = @()
$index = 0

foreach ($file in $files) {
    $index++
    $relative = $file.FullName.Substring($root.Length).TrimStart([IO.Path]::DirectorySeparatorChar)
    $relativeUrl = $relative -replace '\\', '/'
    $relativeDir = Split-Path $relativeUrl -Parent
    $spFolder = if ($relativeDir) { "$targetRoot/$($relativeDir -replace '\\', '/')" } else { $targetRoot }

    Write-Progress -Activity 'Uploading DCSPad' -Status $relativeUrl -PercentComplete (($index / $files.Count) * 100)

    if (-not $PSCmdlet.ShouldProcess("$spFolder/$($file.Name)", 'Upload')) { continue }

    try {
        Confirm-Folder $spFolder
        Add-PnPFile -Path $file.FullName -Folder $spFolder | Out-Null
        $uploaded++
        Write-Verbose "uploaded $relativeUrl"
    } catch {
        $failed += [pscustomobject]@{ File = $relativeUrl; Error = $_.Exception.Message }
        Write-Warning "failed: $relativeUrl — $($_.Exception.Message)"
    }
}
Write-Progress -Activity 'Uploading DCSPad' -Completed

Write-Host ''
Write-Host "  uploaded: $uploaded/$($files.Count)" -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Green' })

if ($failed) {
    Write-Host "  failed  : $($failed.Count)" -ForegroundColor Red
    $failed | Format-Table -AutoSize
    Write-Host 'A blocked file extension is the usual cause — check the tenant/site blocked-file-type list.' -ForegroundColor Yellow
}

$appUrl = "$($SiteUrl.TrimEnd('/'))/$targetRoot/index.html"
Write-Host ''
Write-Host "Open: $appUrl" -ForegroundColor Cyan
Write-Host @'

After it loads, confirm the top-right chip reads "SP: Live". If it says "Mock",
the host page did not expose _spPageContextInfo — see deploy/README.md
(custom-script setting, and the .aspx fallback when .html downloads instead of
rendering).
'@
