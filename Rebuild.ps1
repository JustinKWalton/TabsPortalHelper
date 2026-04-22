<#
.SYNOPSIS
  Stop the running TABS Portal Helper, rebuild from source, redeploy, and launch.

.DESCRIPTION
  Drop this at C:\TabsPortalHelper\Rebuild.ps1 and run from a PowerShell prompt.
  Default path does a fast dev-loop: stop -> publish -> copy -> launch.
  Use -Full to also run --install (refreshes startup/ARP registry entries);
  the script auto-detects first-install and runs it anyway in that case.

.EXAMPLES
  .\Rebuild.ps1            # fast iteration (no MessageBox)
  .\Rebuild.ps1 -Full      # include --install dance (with its MessageBox)
#>

param(
    [switch]$Full
)

$ErrorActionPreference = 'Stop'

$RepoRoot     = 'C:\TabsPortalHelper'
$ProjectDir   = Join-Path $RepoRoot   'TabsPortalHelper'
$PublishDir   = Join-Path $RepoRoot   'publish'
$InstallDir   = Join-Path $env:LOCALAPPDATA 'TabsPortalHelper'
$InstalledExe = Join-Path $InstallDir 'TabsPortalHelper.exe'
$StartupKey   = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'

Write-Host ""
Write-Host "=== TABS Portal Helper - Rebuild & Reinstall ===" -ForegroundColor Cyan
Write-Host "  Project:  $ProjectDir"
Write-Host "  Publish:  $PublishDir"
Write-Host "  Install:  $InstallDir"
Write-Host ""

# ── 1. Stop running tray instance ────────────────────────────────────────
Write-Host "[1/4] Stopping running helper..." -ForegroundColor Yellow
$procs = @(Get-Process -Name 'TabsPortalHelper' -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $procs | Stop-Process -Force
    Start-Sleep -Milliseconds 700   # let file handles drain before we overwrite the .exe
    Write-Host "      Stopped $($procs.Count) instance(s)." -ForegroundColor DarkGray
} else {
    Write-Host "      Not running." -ForegroundColor DarkGray
}

# ── 2. dotnet publish ────────────────────────────────────────────────────
Write-Host ""
Write-Host "[2/4] Publishing..." -ForegroundColor Yellow
Push-Location $ProjectDir
try {
    if (Test-Path $PublishDir) {
        Remove-Item "$PublishDir\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    dotnet publish -c Release -o $PublishDir
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed (exit $LASTEXITCODE)" }
}
finally { Pop-Location }

# ── 3. Deploy to %LOCALAPPDATA%\TabsPortalHelper ─────────────────────────
Write-Host ""
Write-Host "[3/4] Deploying to $InstallDir..." -ForegroundColor Yellow
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}
# /MIR mirrors (removes stale files from prior version). /R:5 /W:1 retries
# on transient file locks if the tray app didn't release handles fast enough.
$null = robocopy $PublishDir $InstallDir /MIR /NFL /NDL /NJH /NJS /NP /R:5 /W:1
if ($LASTEXITCODE -ge 8) { throw "robocopy failed (exit $LASTEXITCODE)" }
$LASTEXITCODE = 0   # robocopy uses 0-7 for success; normalize so later `$?` checks don't trip
Write-Host "      Copied." -ForegroundColor DarkGray

# ── 4. Launch (and --install if first-time or -Full) ─────────────────────
$firstInstall = -not (Get-ItemProperty -Path $StartupKey -Name 'TabsPortalHelper' -ErrorAction SilentlyContinue)
if ($firstInstall) {
    Write-Host ""
    Write-Host "[4/4] First install detected - running --install..." -ForegroundColor Yellow
    & $InstalledExe --install
} elseif ($Full) {
    Write-Host ""
    Write-Host "[4/4] Running --install (refresh registry entries)..." -ForegroundColor Yellow
    & $InstalledExe --install
} else {
    Write-Host ""
    Write-Host "[4/4] Launching..." -ForegroundColor Yellow
    Start-Process $InstalledExe
}

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "  Watch the system tray for the startup balloon." -ForegroundColor DarkGray
Write-Host "  Verify: curl http://localhost:52874/ping" -ForegroundColor DarkGray
Write-Host ""
