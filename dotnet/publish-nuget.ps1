<#
.SYNOPSIS
    Build the headless JS bundle, then build, pack, and push the
    WordCanvas.ClearScript project to NuGet.org (or any feed).

.DESCRIPTION
    Produces a self-contained NuGet package that carries everything the ClearScript
    V8 host needs at runtime:

      * the esbuild JS bundle  -> package assets\wordcanvas.clearscript.js
      * the bundled .ttf fonts -> package fonts\*.ttf
      * a build\WordCanvas.ClearScript.targets that copies both next to the
        consumer's output assembly (where WordCanvasEngineOptions.Default() looks)

    The JS bundle is generated (gitignored), so the script runs
    frontend/scripts/build-clearscript.mjs first unless -SkipJsBuild is given.

.PARAMETER ApiKey
    NuGet.org API key. Can also be set via the NUGET_API_KEY environment variable.
    Not required when -NoPush is used.

.PARAMETER Version
    Override the package version (e.g. "0.12.0"). If omitted, the version declared
    in the .csproj is used (unless -SyncNpmVersion is set).

.PARAMETER SyncNpmVersion
    Read the version from frontend/package.json instead of the .csproj, so the
    NuGet package tracks the @forevka/wordcanvas npm release. -Version still wins.

.PARAMETER Source
    NuGet feed URL. Defaults to https://api.nuget.org/v3/index.json

.PARAMETER SkipJsBuild
    Skip 'node build-clearscript.mjs' (use when the bundle is already fresh).

.PARAMETER SkipBuild
    Skip the dotnet build step (implies the project was already built in Release).

.PARAMETER NoPush
    Build and pack only; do not push. Handy for inspecting the .nupkg locally.

.EXAMPLE
    .\publish-nuget.ps1 -ApiKey "oy2abc..."

.EXAMPLE
    $env:NUGET_API_KEY = "oy2abc..."; .\publish-nuget.ps1 -SyncNpmVersion

.EXAMPLE
    .\publish-nuget.ps1 -Version 0.12.0 -NoPush      # local dry run, no key needed
#>
[CmdletBinding()]
param(
    [string] $ApiKey  = $env:NUGET_API_KEY,
    [string] $Version = "",
    [switch] $SyncNpmVersion,
    [string] $Source  = "https://api.nuget.org/v3/index.json",
    [switch] $SkipJsBuild,
    [switch] $SkipBuild,
    [switch] $NoPush
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── paths ─────────────────────────────────────────────────────────────────────
$dotnetDir = Split-Path -Parent $MyInvocation.MyCommand.Path      # <repo>/dotnet
$repoRoot  = Split-Path -Parent $dotnetDir                        # <repo>
$projectDir = Join-Path $dotnetDir "src/WordCanvas.ClearScript"
$csproj     = Join-Path $projectDir "WordCanvas.ClearScript.csproj"
$bundle     = Join-Path $projectDir "assets/wordcanvas.clearscript.js"
$jsBuild    = Join-Path $repoRoot "frontend/scripts/build-clearscript.mjs"
$nupkgDir   = Join-Path $projectDir "nupkg"

if (-not (Test-Path $csproj)) { throw "Project not found: $csproj" }

# ── api key guard (only when we intend to push) ───────────────────────────────
if (-not $NoPush -and [string]::IsNullOrWhiteSpace($ApiKey)) {
    Write-Error @"
No API key provided. Pass it with -ApiKey, set the NUGET_API_KEY environment
variable, or use -NoPush for a local build+pack:

    `$env:NUGET_API_KEY = "oy2..."
    .\publish-nuget.ps1
"@
}

# ── resolve version ───────────────────────────────────────────────────────────
$effectiveVersion = $Version
if ([string]::IsNullOrWhiteSpace($effectiveVersion) -and $SyncNpmVersion) {
    $pkgJsonPath = Join-Path $repoRoot "frontend/package.json"
    $pkgJson = Get-Content $pkgJsonPath -Raw | ConvertFrom-Json
    $effectiveVersion = $pkgJson.version
    if ([string]::IsNullOrWhiteSpace($effectiveVersion)) {
        throw "Could not read 'version' from $pkgJsonPath."
    }
}
if ([string]::IsNullOrWhiteSpace($effectiveVersion)) {
    [xml]$proj = Get-Content $csproj
    $effectiveVersion = $proj.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($effectiveVersion)) {
        throw "Could not read <Version> from $csproj. Pass -Version explicitly."
    }
}

$packageId = "WordCanvas.ClearScript"
$nupkgFile = Join-Path $nupkgDir "$packageId.$effectiveVersion.nupkg"

Write-Host ""
Write-Host "=== WordCanvas.ClearScript NuGet Publisher ===" -ForegroundColor Cyan
Write-Host "  Project  : $csproj"
Write-Host "  Package  : $packageId"
Write-Host "  Version  : $effectiveVersion"
Write-Host "  Feed     : $Source"
Write-Host "  Push     : $(if ($NoPush) { 'no (-NoPush)' } else { 'yes' })"
Write-Host ""

# ── [1/4] build the JS bundle ─────────────────────────────────────────────────
if (-not $SkipJsBuild) {
    Write-Host "[1/4] Building the ClearScript JS bundle (esbuild)..." -ForegroundColor Yellow
    Push-Location $repoRoot
    try {
        & node $jsBuild | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "JS bundle build failed (node $jsBuild)." }
    }
    finally { Pop-Location }
} else {
    Write-Host "[1/4] Skipping JS bundle build (-SkipJsBuild)." -ForegroundColor DarkGray
}

if (-not (Test-Path $bundle)) {
    throw "JS bundle not found at $bundle. Run without -SkipJsBuild, or run 'node $jsBuild' first."
}

# ── [2/4] build the .NET project in Release ───────────────────────────────────
if (-not $SkipBuild) {
    Write-Host "[2/4] Building WordCanvas.ClearScript in Release..." -ForegroundColor Yellow
    & dotnet build $csproj -c Release --nologo "/p:Version=$effectiveVersion" | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed for $packageId." }
} else {
    Write-Host "[2/4] Skipping dotnet build (-SkipBuild)." -ForegroundColor DarkGray
}

# ── [3/4] pack ────────────────────────────────────────────────────────────────
Write-Host "[3/4] Packing..." -ForegroundColor Yellow
if (Test-Path $nupkgDir) { Remove-Item $nupkgDir -Recurse -Force }
New-Item -ItemType Directory -Path $nupkgDir | Out-Null

$packArgs = @(
    "pack", $csproj,
    "-c", "Release",
    $(if ($SkipBuild) { "--no-build" } else { "--no-restore" }),
    "-o", $nupkgDir,
    "--nologo",
    "/p:Version=$effectiveVersion"
)
& dotnet @packArgs | Out-Host
if ($LASTEXITCODE -ne 0) { throw "dotnet pack failed for $packageId." }

if (-not (Test-Path $nupkgFile)) {
    $nupkgFile = Get-ChildItem $nupkgDir -Filter "*.nupkg" |
                 Where-Object { $_.Name -notlike "*.symbols.nupkg" } |
                 Select-Object -First 1 -ExpandProperty FullName
    if (-not $nupkgFile) { throw "No .nupkg produced in $nupkgDir." }
}
Write-Host "  Produced : $nupkgFile" -ForegroundColor DarkGray

# ── [4/4] push ────────────────────────────────────────────────────────────────
if ($NoPush) {
    Write-Host "[4/4] Skipping push (-NoPush)." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Done. Inspect the package at:" -ForegroundColor Green
    Write-Host "  $nupkgFile"
    Write-Host ""
    return
}

Write-Host "[4/4] Pushing to NuGet..." -ForegroundColor Yellow
& dotnet nuget push $nupkgFile `
    --api-key $ApiKey `
    --source $Source `
    --skip-duplicate | Out-Host
if ($LASTEXITCODE -ne 0) { throw "Push failed for $packageId." }

Write-Host ""
Write-Host "Done! Package will appear on nuget.org within ~15 minutes." -ForegroundColor Green
Write-Host "  https://www.nuget.org/packages/$packageId/$effectiveVersion"
Write-Host ""
