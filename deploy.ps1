<#
.SYNOPSIS
    Deploy canvas-word (full stack) to the VPS via tar + scp + docker compose.

.DESCRIPTION
    Packs the source (excluding node_modules, dist, local-db-data, .env, .git,
    big fixtures), uploads to the VPS, extracts into the remote project dir
    (preserving .env and ./local-db-data), creates .env on first run with a
    generated DB password, then runs `docker compose --profile app up -d --build`
    and a smoke test.

    The host Caddy must have a `doc-editor.<domain>` block pointing at
    127.0.0.1:$EdgePort (set up separately / by the deploy notes).

.PARAMETER RemoteHost
    SSH target. Default: root@your-server

.PARAMETER RemotePath
    Project directory on the VPS. Default: /opt/canvas-word

.PARAMETER Url
    URL to smoke-test after deploy. Default: https://doc-editor.example.com/openapi.json

.PARAMETER EdgePort
    Host port the web edge binds (127.0.0.1). Default: 3003

.EXAMPLE
    .\deploy.ps1
#>

[CmdletBinding()]
param(
    [string]$RemoteHost = 'root@your-server',
    [string]$RemotePath = '/opt/canvas-word',
    [string]$Url        = 'https://doc-editor.example.com/openapi.json',
    [int]   $EdgePort   = 3003
)

$ErrorActionPreference = 'Stop'

$projectRoot   = $PSScriptRoot
$tarballName   = "canvas-word-src-$(Get-Date -Format 'yyyyMMdd-HHmmss').tar.gz"
$localTarball  = Join-Path $env:TEMP $tarballName
$remoteTarball = "/tmp/$tarballName"

Write-Host "==> Project root : $projectRoot"
Write-Host "==> Remote target: $RemoteHost`:$RemotePath"
Write-Host "==> Tarball      : $localTarball"
Write-Host ""

# 1. Pack
Write-Host "==> [1/4] Packing source..."
# PowerShell's `tar` is bsdtar (libarchive) — handles Windows paths natively
# (no --force-local needed) and matches these glob excludes.
$excludes = @(
    '--exclude=./node_modules',
    '--exclude=./*/node_modules',
    '--exclude=./dist',
    '--exclude=./*/dist',
    '--exclude=./*/dist-lib',
    '--exclude=./local-db-data',
    '--exclude=./.git',
    '--exclude=./.omc',
    '--exclude=./.claude',
    '--exclude=./.playwright-mcp',
    '--exclude=./shots',
    '--exclude=./bugs',
    '--exclude=./.env',
    '--exclude=./*/*.tsbuildinfo',
    '--exclude=*RenderedReportsDocx_*',
    '--exclude=*.docx',
    '--exclude=*.zip'
)
& tar @excludes -czf $localTarball -C $projectRoot .
if ($LASTEXITCODE -ne 0) { throw "tar failed (exit $LASTEXITCODE)" }
$size = [math]::Round((Get-Item $localTarball).Length / 1KB, 1)
Write-Host "    packed $size KB"

# 2. Upload
Write-Host "==> [2/4] Uploading to $RemoteHost..."
& ssh $RemoteHost "mkdir -p $RemotePath"
if ($LASTEXITCODE -ne 0) { throw "remote mkdir failed (exit $LASTEXITCODE)" }
& scp $localTarball "$RemoteHost`:$remoteTarball"
if ($LASTEXITCODE -ne 0) { throw "scp failed (exit $LASTEXITCODE)" }

# 3. Extract + (first run) seed .env + rebuild on remote
Write-Host "==> [3/4] Extracting + docker compose up --build (this can take a few minutes)..."
# Double-quoted here-string: PowerShell interpolates $RemotePath/$remoteTarball/
# $EdgePort; bash's own $ are backtick-escaped (`$) so they reach the remote shell.
$remoteScript = @"
set -e
cd $RemotePath
tar -xzf $remoteTarball --exclude=.env
rm -f $remoteTarball
if [ ! -f .env ]; then
  echo '    .env not found -> generating with a random DB password'
  PW=`$(openssl rand -hex 24)
  cat > .env <<EOF
CW_DB_USER=postgres
CW_DB_PASSWORD=`$PW
CW_DB_NAME=canvasword
CW_DB_PORT=5544
EDGE_PORT=$EdgePort
EOF
fi
docker compose --profile app up -d --build 2>&1 | tail -25
sleep 5
docker compose --profile app ps
"@
& ssh $RemoteHost $remoteScript
if ($LASTEXITCODE -ne 0) { throw "remote build failed (exit $LASTEXITCODE)" }

# 4. Smoke test
Write-Host "==> [4/4] Smoke testing $Url..."
try {
    $resp = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 20
    if ($resp.StatusCode -eq 200) { Write-Host "    HTTP 200 OK" }
    else { Write-Warning "    HTTP $($resp.StatusCode)" }
} catch {
    Write-Warning "    smoke test failed: $($_.Exception.Message)"
    Write-Warning "    (DNS for doc-editor may still be propagating, or the host Caddy block isn't in place yet)"
}

Remove-Item $localTarball -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "==> Done."
