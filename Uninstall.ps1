# TABS Portal Helper — Uninstaller

$ExePath    = "$env:LOCALAPPDATA\TabsPortalHelper\TabsPortalHelper.exe"
$InstallDir = "$env:LOCALAPPDATA\TabsPortalHelper"

Write-Host "Uninstalling TABS Portal Helper..." -ForegroundColor Yellow

# Stop running instance
Stop-Process -Name "TabsPortalHelper" -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# Run uninstall (removes registry entries)
if (Test-Path $ExePath) { & $ExePath --uninstall }

# Remove install directory
if (Test-Path $InstallDir) { Remove-Item -Path $InstallDir -Recurse -Force }

Write-Host "TABS Portal Helper has been removed." -ForegroundColor Green
