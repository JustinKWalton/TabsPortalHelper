# TABS Portal Helper — Installer
# Run this once to install. No admin rights required.

$ErrorActionPreference = "Stop"
$ExeName   = "TabsPortalHelper.exe"
$InstallDir = "$env:LOCALAPPDATA\TabsPortalHelper"
$ExePath    = Join-Path $InstallDir $ExeName
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceExe  = Join-Path $ScriptDir $ExeName

Write-Host ""
Write-Host "TABS Portal Helper — Installer" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan
Write-Host ""

# Verify exe exists next to this script
if (-not (Test-Path $SourceExe)) {
    Write-Host "ERROR: $ExeName not found next to Install.ps1" -ForegroundColor Red
    Write-Host "Make sure both files are in the same folder." -ForegroundColor Red
    exit 1
}

# Stop any running instance
Stop-Process -Name "TabsPortalHelper" -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# Create install directory and copy exe
if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir | Out-Null }
Copy-Item -Path $SourceExe -Destination $ExePath -Force
Write-Host "Installed to: $ExePath" -ForegroundColor Green

# Run install (registers startup + Add/Remove Programs)
& $ExePath --install

# Launch the tray app
Write-Host "Starting TABS Portal Helper..." -ForegroundColor Green
Start-Process $ExePath

Write-Host ""
Write-Host "Done! TABS Portal Helper is running in the system tray." -ForegroundColor Cyan
Write-Host "It will start automatically each time you log in to Windows." -ForegroundColor Cyan
Write-Host ""
