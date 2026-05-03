# =============================================================================
#  Apply-Patch-v2.5.4.ps1
#
#  Upgrades TabsPortalHelper source from v2.5.3 to v2.5.4:
#    - Rewrites ExplorerHelper.cs to shell out to powershell.exe for the
#      open + activate operation. The PowerShell process is freshly spawned
#      and gets foreground-change rights that our long-running tray app
#      lacks, so SendKeys + AppActivate inside PS reliably brings Explorer
#      to the foreground even when the original HTTP request came from a
#      browser.
#
#  This drops a bunch of Win32 P/Invoke complexity. The new file is ~50 lines.
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.5.4.ps1
#      .\Apply-Patch-v2.5.4.ps1 -DryRun
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$FromVersion = '2.5.3'
$ToVersion   = '2.5.4'

function Write-Step  ($m) { Write-Host "`n>> $m" -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host "   [OK]    $m" -ForegroundColor Green }
function Write-Skip  ($m) { Write-Host "   [SKIP]  $m" -ForegroundColor DarkGray }
function Write-Warn2 ($m) { Write-Host "   [WARN]  $m" -ForegroundColor Yellow }
function Write-Bad   ($m) { Write-Host "   [FAIL]  $m" -ForegroundColor Red }
function Write-Info  ($m) { Write-Host "   $m" -ForegroundColor Gray }

function Save-FileContent {
    param([string]$Path, [string]$Content, [string]$Label)
    if ($DryRun) {
        Write-Warn2 "would save $Label"
    } else {
        [System.IO.File]::WriteAllText($Path, $Content)
        Write-Ok "saved $Label"
    }
}

# --- 1. Locate source files --------------------------------------------------
$csproj         = Join-Path $ProjectDir "TabsPortalHelper\TabsPortalHelper.csproj"
$httpServer     = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp        = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$installer      = Join-Path $ProjectDir "TabsPortalHelper\Installer.cs"
$explorerHelper = Join-Path $ProjectDir "TabsPortalHelper\ExplorerHelper.cs"

if (-not (Test-Path $csproj)) {
    $alt = Join-Path $ProjectDir "TabsPortalHelper.csproj"
    if (Test-Path $alt) {
        $csproj         = $alt
        $httpServer     = Join-Path $ProjectDir "HttpServer.cs"
        $trayApp        = Join-Path $ProjectDir "TrayApp.cs"
        $installer      = Join-Path $ProjectDir "Installer.cs"
        $explorerHelper = Join-Path $ProjectDir "ExplorerHelper.cs"
    }
}

Write-Step "Validating files exist"
$allFound = $true
foreach ($f in @($csproj, $httpServer, $trayApp, $installer, $explorerHelper)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        $allFound = $false
    }
}
if (-not $allFound) {
    Write-Bad "Source files not found."
    exit 1
}

# --- 2. Pre-flight version check --------------------------------------------
Write-Step "Pre-flight: confirm everything is at $FromVersion"
$canProceed = $true

function Check-Version {
    param([string]$Path, [string]$Label, [string]$Pattern)
    $content = Get-Content $Path -Raw
    if ($content -match $Pattern) {
        $found = $matches[1]
        if ($found -eq $script:ToVersion) {
            Write-Skip "${Label} already at $($script:ToVersion)"
            return @{ Action = 'skip' }
        } elseif ($found -eq $script:FromVersion) {
            Write-Ok "${Label} at $($script:FromVersion) (will bump)"
            return @{ Action = 'bump' }
        } else {
            Write-Bad "${Label} is at unexpected version: $found"
            return @{ Action = 'abort' }
        }
    } else {
        Write-Bad "${Label}: no version pattern matched"
        return @{ Action = 'abort' }
    }
}

$checks = @{
    csproj     = Check-Version $csproj     '.csproj'         '<Version>\s*([0-9.]+)\s*</Version>'
    httpServer = Check-Version $httpServer 'HttpServer.cs'   'const\s+string\s+Version\s*=\s*"([0-9.]+)"'
    trayApp    = Check-Version $trayApp    'TrayApp.cs'      'const\s+string\s+Version\s*=\s*"([0-9.]+)"'
    installer  = Check-Version $installer  'Installer.cs'    'const\s+string\s+AppVersion\s*=\s*"([0-9.]+)"'
}
foreach ($k in $checks.Keys) {
    if ($checks[$k].Action -eq 'abort') { $canProceed = $false }
}

if (-not $canProceed) {
    Write-Bad "Aborting: one or more files at unexpected version."
    exit 1
}

# --- 3. Replace ExplorerHelper.cs entirely ----------------------------------
Write-Step "Rewriting ExplorerHelper.cs (PowerShell shell-out approach)"
$content = Get-Content $explorerHelper -Raw

if ($content -match 'powershell\.exe') {
    Write-Skip "ExplorerHelper.cs already shells out to PowerShell (v2.5.4 behavior)"
} else {
    # Sanity check: file should still look like one of the Win32-based versions
    if ($content -notmatch 'AttachThreadInput' -and $content -notmatch 'BringWindowToTop') {
        Write-Bad "ExplorerHelper.cs doesn't look like any known prior version - inspect manually"
        exit 1
    }

    $newFileLines = @(
        'using System;'
        'using System.Diagnostics;'
        'using System.IO;'
        ''
        'namespace TabsPortalHelper'
        '{'
        '    static class ExplorerHelper'
        '    {'
        '        // ============================================================'
        '        // OpenFolder'
        '        //'
        '        // Opens a Windows Explorer window at folderPath and brings it'
        '        // to the foreground.'
        '        //'
        '        // Implementation note: rather than calling Process.Start +'
        '        // SetForegroundWindow directly from this long-running tray'
        '        // process (which Windows aggressively blocks from stealing'
        '        // focus when called from a busy browser), we shell out to a'
        '        // hidden powershell.exe. The PS process is *freshly spawned*'
        '        // by us, so for its first ~1s of life Windows grants it'
        '        // foreground-change rights that the tray app does not have.'
        '        // Inside PS we use SendKeys ''%'' (synthetic Alt) to top off'
        '        // the eligibility heuristic, then VB''s AppActivate to bring'
        '        // the Explorer window with matching title to the front.'
        '        //'
        '        // Why this works when raw Win32 didn''t: the foreground-grab'
        '        // restrictions that bit us across v2.5.0..2.5.3 apply per'
        '        // *process*. A new process gets a clean slate. Shelling out'
        '        // is essentially the "proxy process" pattern made cheap.'
        '        // ============================================================'
        '        public static bool OpenFolder(string folderPath)'
        '        {'
        '            try'
        '            {'
        '                var folderName = Path.GetFileName('
        '                    folderPath.TrimEnd(Path.DirectorySeparatorChar));'
        ''
        '                // Escape single quotes for PS single-quoted strings.'
        '                var pathArg = folderPath.Replace("''", "''''");'
        '                var nameArg = folderName.Replace("''", "''''");'
        ''
        '                // One-line PowerShell:'
        '                //   1. Start-Process Explorer at the target folder.'
        '                //   2. Load Forms + VB assemblies.'
        '                //   3. Loop up to ~1.2s polling the window into focus:'
        '                //        - SendKeys ''%'' synthesizes Alt for the PS'
        '                //          process (which is briefly foreground-eligible'
        '                //          as a freshly spawned process).'
        '                //        - AppActivate finds the window by partial title'
        '                //          match and activates it. Throws if not yet'
        '                //          present, so we retry.'
        '                var script = string.Join("; ", new[]'
        '                {'
        '                    $"Start-Process explorer.exe -ArgumentList ''\"{pathArg}\"''",'
        '                    "Add-Type -AssemblyName System.Windows.Forms",'
        '                    "Add-Type -AssemblyName Microsoft.VisualBasic",'
        '                    $"$folderName = ''{nameArg}''",'
        '                    "for ($i = 0; $i -lt 8; $i++) {" +'
        '                    "  Start-Sleep -Milliseconds 150;" +'
        '                    "  [System.Windows.Forms.SendKeys]::SendWait(''%'');" +'
        '                    "  try {" +'
        '                    "    [Microsoft.VisualBasic.Interaction]::AppActivate($folderName);" +'
        '                    "    break" +'
        '                    "  } catch {}" +'
        '                    "}"'
        '                });'
        ''
        '                Process.Start(new ProcessStartInfo'
        '                {'
        '                    FileName  = "powershell.exe",'
        '                    Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{script}\"",'
        '                    UseShellExecute = false,'
        '                    CreateNoWindow  = true,'
        '                });'
        '                return true;'
        '            }'
        '            catch'
        '            {'
        '                return false;'
        '            }'
        '        }'
        '    }'
        '}'
        ''
    )
    $newFileContent = $newFileLines -join "`r`n"
    Save-FileContent $explorerHelper $newFileContent "ExplorerHelper.cs (rewritten)"
}

# --- 4. Bump version constants ----------------------------------------------
Write-Step "Bumping .csproj"
$content = Get-Content $csproj -Raw
if ($content -match "<Version>\s*$([regex]::Escape($ToVersion))\s*</Version>") {
    Write-Skip "<Version> already at $ToVersion"
} else {
    $content = $content -replace "<Version>\s*$([regex]::Escape($FromVersion))\s*</Version>", "<Version>$ToVersion</Version>"
    Save-FileContent $csproj $content "TabsPortalHelper.csproj  ($FromVersion -> $ToVersion)"
}

Write-Step "Bumping HttpServer.cs Version"
$content = Get-Content $httpServer -Raw
if ($content -match "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "Version constant already at $ToVersion"
} else {
    $content = $content -replace "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($FromVersion))`"", "const string Version = `"$ToVersion`""
    Save-FileContent $httpServer $content "HttpServer.cs Version  ($FromVersion -> $ToVersion)"
}

Write-Step "Bumping TrayApp.cs Version"
$content = Get-Content $trayApp -Raw
if ($content -match "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "Version constant already at $ToVersion"
} else {
    $content = $content -replace "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($FromVersion))`"", "const string Version  = `"$ToVersion`""
    Save-FileContent $trayApp $content "TrayApp.cs  ($FromVersion -> $ToVersion)"
}

Write-Step "Bumping Installer.cs AppVersion"
$content = Get-Content $installer -Raw
if ($content -match "const\s+string\s+AppVersion\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "AppVersion constant already at $ToVersion"
} else {
    $content = $content -replace "const\s+string\s+AppVersion\s*=\s*`"$([regex]::Escape($FromVersion))`"", "const string AppVersion = `"$ToVersion`""
    Save-FileContent $installer $content "Installer.cs  ($FromVersion -> $ToVersion)"
}

# --- 5. Done ----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "Dry run - no files were actually modified."
} else {
    Write-Ok "All patches applied. Source is now at $ToVersion."
    Write-Host ""
    Write-Info "Next steps:"
    Write-Info "  1. .\Rebuild.ps1"
    Write-Info "  2. Click the Open Folder button in FlutterFlow."
    Write-Info "     Behavior should be: brief invisible PS spawn, then Explorer"
    Write-Info "     pops fully to the front."
}
Write-Host ""
