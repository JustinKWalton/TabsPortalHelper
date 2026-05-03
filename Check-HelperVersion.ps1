# ╔═══════════════════════════════════════════════════════════════════════╗
# ║  Check-HelperVersion.ps1                                              ║
# ║                                                                       ║
# ║  Reports every TabsPortalHelper version on this machine and flags     ║
# ║  any disagreement between source / installed / running.               ║
# ║                                                                       ║
# ║  Usage from repo root:                                                ║
# ║      .\Check-HelperVersion.ps1                                        ║
# ║                                                                       ║
# ║  Usage from anywhere:                                                 ║
# ║      .\Check-HelperVersion.ps1 -ProjectDir 'C:\path\to\repo'          ║
# ╚═══════════════════════════════════════════════════════════════════════╝

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [int]   $Port       = 52874
)

# ─── Output helpers (match Apply-Patch script style) ────────────────────────
function Write-Step  ($m) { Write-Host "`n>> $m" -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host "   [OK]    $m" -ForegroundColor Green }
function Write-Warn2 ($m) { Write-Host "   [WARN]  $m" -ForegroundColor Yellow }
function Write-Bad   ($m) { Write-Host "   [FAIL]  $m" -ForegroundColor Red }
function Write-Info  ($m) { Write-Host "   $m" -ForegroundColor Gray }

# ─── Locate the source files ────────────────────────────────────────────────
# Try repo-root layout first (TabsPortalHelper\TabsPortalHelper.csproj),
# then fall back to running from inside the project folder itself.
$csproj     = Join-Path $ProjectDir "TabsPortalHelper\TabsPortalHelper.csproj"
$httpServer = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp    = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$installer  = Join-Path $ProjectDir "TabsPortalHelper\Installer.cs"

if (-not (Test-Path $csproj)) {
    $altCsproj = Join-Path $ProjectDir "TabsPortalHelper.csproj"
    if (Test-Path $altCsproj) {
        $csproj     = $altCsproj
        $httpServer = Join-Path $ProjectDir "HttpServer.cs"
        $trayApp    = Join-Path $ProjectDir "TrayApp.cs"
        $installer  = Join-Path $ProjectDir "Installer.cs"
    }
}

# Track everything we find for the final mismatch summary
$results = [ordered]@{}

# ─── 1. Source code versions ────────────────────────────────────────────────
Write-Step "Source code versions"

if (Test-Path $csproj) {
    $content = Get-Content $csproj -Raw
    if ($content -match '<Version>\s*([0-9.]+)\s*</Version>') {
        $results['.csproj <Version>']         = $matches[1]
        Write-Info (".csproj <Version>          : {0}" -f $matches[1])
    } else {
        Write-Warn2 ".csproj has no <Version> tag"
    }
} else {
    Write-Warn2 ".csproj not found at: $csproj"
    Write-Info  "(pass -ProjectDir if you're running this from outside the repo)"
}

if (Test-Path $httpServer) {
    $content = Get-Content $httpServer -Raw
    if ($content -match 'const\s+string\s+Version\s*=\s*"([0-9.]+)"') {
        $results['HttpServer.cs Version']      = $matches[1]
        Write-Info ("HttpServer.cs Version      : {0}" -f $matches[1])
    }
}

if (Test-Path $trayApp) {
    $content = Get-Content $trayApp -Raw
    if ($content -match 'const\s+string\s+Version\s*=\s*"([0-9.]+)"') {
        $results['TrayApp.cs Version']         = $matches[1]
        Write-Info ("TrayApp.cs Version         : {0}" -f $matches[1])
    }
}

if (Test-Path $installer) {
    $content = Get-Content $installer -Raw
    if ($content -match 'const\s+string\s+AppVersion\s*=\s*"([0-9.]+)"') {
        $results['Installer.cs AppVersion']    = $matches[1]
        Write-Info ("Installer.cs AppVersion    : {0}" -f $matches[1])
    }
}

# ─── 2. Installed exe ───────────────────────────────────────────────────────
Write-Step "Installed (deployed) helper"

$installedExe = Join-Path $env:LOCALAPPDATA "TabsPortalHelper\TabsPortalHelper.exe"
if (Test-Path $installedExe) {
    $info = (Get-Item $installedExe).VersionInfo
    $results['Installed FileVersion']     = $info.FileVersion
    $results['Installed ProductVersion']  = $info.ProductVersion
    Write-Info ("Path                       : {0}" -f $installedExe)
    Write-Info ("FileVersion                : {0}" -f $info.FileVersion)
    Write-Info ("ProductVersion             : {0}" -f $info.ProductVersion)
    Write-Info ("Last modified              : {0}" -f (Get-Item $installedExe).LastWriteTime)
} else {
    Write-Warn2 "Installed exe not found at $installedExe"
    Write-Info  "(helper may have never been installed on this user profile)"
}

# ─── 3. Registry version stamp ──────────────────────────────────────────────
$tabsKey = 'HKCU:\Software\TabsPortalHelper'
if (Test-Path $tabsKey) {
    $reg = Get-ItemProperty $tabsKey -ErrorAction SilentlyContinue
    if ($reg.InstalledVersion) {
        $results['Registry InstalledVersion'] = $reg.InstalledVersion
        Write-Info ("Registry InstalledVersion  : {0}  (installed {1})" -f $reg.InstalledVersion, $reg.InstalledAt)
    }
} else {
    Write-Warn2 "No registry entry at HKCU\Software\TabsPortalHelper"
}

# ─── 4. Running tray app ────────────────────────────────────────────────────
Write-Step "Running tray app (/ping on localhost:$Port)"

try {
    $ping = Invoke-RestMethod -Uri "http://localhost:$Port/ping" -TimeoutSec 2 -ErrorAction Stop
    $results['Running /ping']              = $ping.version
    Write-Ok "Tray app responding"
    Write-Info ("Version                    : {0}" -f $ping.version)
    Write-Info ("Drive root                 : {0}" -f $ping.driveRoot)
    Write-Info ("Bluebeam                   : {0}" -f $ping.bluebeam)
} catch {
    Write-Bad "Tray app NOT responding on localhost:$Port"
    Write-Info "(launch it from %LOCALAPPDATA%\TabsPortalHelper\TabsPortalHelper.exe"
    Write-Info " or check the system tray for a hidden/exited icon)"
}

# ─── 5. Mismatch summary ────────────────────────────────────────────────────
Write-Step "Summary"

$uniq = $results.Values | Where-Object { $_ } | Sort-Object -Unique

if ($uniq.Count -eq 0) {
    Write-Bad "No version information found anywhere."
}
elseif ($uniq.Count -eq 1) {
    Write-Ok ("All version sources agree: {0}" -f ($uniq -join ', '))
}
else {
    Write-Warn2 "VERSIONS DO NOT MATCH:"
    Write-Host ""
    foreach ($k in $results.Keys) {
        Write-Info ("  {0,-32}  {1}" -f $k, $results[$k])
    }
    Write-Host ""
    Write-Info "Common causes:"
    Write-Info "  - Source edited but Rebuild.ps1 not run yet                     (source > installed)"
    Write-Info "  - Rebuild.ps1 ran but old tray instance is still in memory      (installed > running)"
    Write-Info "    => exit it from the tray menu and re-launch the .exe"
    Write-Info "  - Different version constants drifted out of sync in source     (source files disagree)"
}

Write-Host ""
