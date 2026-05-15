# =============================================================================
#  Apply-Patch-v2.5.0.ps1
#
#  Upgrades TabsPortalHelper source from v2.4.0 to v2.5.0:
#    - Bumps version constants in 4 source files
#    - Adds  /folder endpoint to HttpServer.cs (HandleFolder method
#      + route case + System.Diagnostics using directive)
#
#  Idempotent: safe to re-run. Aborts cleanly if source is at an
#  unexpected version.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.5.0.ps1
#      .\Apply-Patch-v2.5.0.ps1 -DryRun
#      .\Apply-Patch-v2.5.0.ps1 -ProjectDir 'C:\tabsportalhelper'
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$FromVersion = '2.4.0'
$ToVersion   = '2.5.0'

# --- Output helpers ----------------------------------------------------------
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
$csproj     = Join-Path $ProjectDir "TabsPortalHelper\TabsPortalHelper.csproj"
$httpServer = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp    = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$installer  = Join-Path $ProjectDir "TabsPortalHelper\Installer.cs"

if (-not (Test-Path $csproj)) {
    $alt = Join-Path $ProjectDir "TabsPortalHelper.csproj"
    if (Test-Path $alt) {
        $csproj     = $alt
        $httpServer = Join-Path $ProjectDir "HttpServer.cs"
        $trayApp    = Join-Path $ProjectDir "TrayApp.cs"
        $installer  = Join-Path $ProjectDir "Installer.cs"
    }
}

Write-Step "Validating files exist"
$allFound = $true
foreach ($f in @($csproj, $httpServer, $trayApp, $installer)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        $allFound = $false
    }
}
if (-not $allFound) {
    Write-Host ""
    Write-Bad "Source files not found. Pass -ProjectDir if running from outside the repo."
    exit 1
}

# --- 2. Pre-flight version check --------------------------------------------
Write-Step "Pre-flight: confirm everything is at $FromVersion"

$canProceed  = $true
$alreadyDone = $true

function Check-Version {
    param(
        [string]$Path,
        [string]$Label,
        [string]$Pattern  # regex group 1 = version string
    )
    $content = Get-Content $Path -Raw
    if ($content -match $Pattern) {
        $found = $matches[1]
        if ($found -eq $script:ToVersion) {
            Write-Skip "${Label} already at $($script:ToVersion)"
            return @{ Action = 'skip'; Found = $found }
        } elseif ($found -eq $script:FromVersion) {
            Write-Ok "${Label} at $($script:FromVersion) (will bump)"
            return @{ Action = 'bump'; Found = $found }
        } else {
            Write-Bad "${Label} is at unexpected version: $found"
            return @{ Action = 'abort'; Found = $found }
        }
    } else {
        Write-Bad "${Label}: no version pattern matched"
        return @{ Action = 'abort'; Found = $null }
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
    if ($checks[$k].Action -ne 'skip')  { $alreadyDone = $false }
}

if (-not $canProceed) {
    Write-Host ""
    Write-Bad "Aborting: one or more files is at an unexpected version."
    Write-Info "This script only knows how to upgrade $FromVersion -> $ToVersion."
    Write-Info "If your source is on a different version, patch manually or roll back first."
    exit 1
}

if ($alreadyDone) {
    Write-Warn2 "All version constants already at $ToVersion. Will still verify /folder patches."
}

# --- 3. Bump .csproj ---------------------------------------------------------
Write-Step "Patching TabsPortalHelper.csproj"
$content = Get-Content $csproj -Raw
if ($content -match "<Version>\s*$([regex]::Escape($ToVersion))\s*</Version>") {
    Write-Skip "<Version> already at $ToVersion"
} else {
    $content = $content -replace '<Version>\s*2\.4\.0\s*</Version>', "<Version>$ToVersion</Version>"
    Save-FileContent $csproj $content "TabsPortalHelper.csproj  ($FromVersion -> $ToVersion)"
}

# --- 4. Patch HttpServer.cs --------------------------------------------------
Write-Step "Patching HttpServer.cs"
$content = Get-Content $httpServer -Raw
$changed = $false

# 4a. Bump Version constant
if ($content -match "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "Version constant already at $ToVersion"
} else {
    $content = $content -replace 'const\s+string\s+Version\s*=\s*"2\.4\.0"', "const string Version = `"$ToVersion`""
    Write-Ok "bumped Version constant"
    $changed = $true
}

# 4b. Add 'using System.Diagnostics;' if missing
if ($content -match 'using\s+System\.Diagnostics\s*;') {
    Write-Skip "'using System.Diagnostics;' already present"
} else {
    $anchor = "using System.Collections.Generic;"
    $replacement = "using System.Collections.Generic;`r`nusing System.Diagnostics;"
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $replacement)
        Write-Ok "added 'using System.Diagnostics;'"
        $changed = $true
    } else {
        Write-Bad "could not find 'using System.Collections.Generic;' anchor"
        exit 1
    }
}

# 4c. Add /folder route case if missing
if ($content -match 'case\s*"/folder"\s*:') {
    Write-Skip "/folder route case already present"
} else {
    $anchor = 'case "/clipboard/bluebeam-markup": HandleClipboardMarkup(ctx);    break;'
    $replacement = $anchor + "`r`n                    case `"/folder`":                   HandleFolder(ctx, query);       break;"
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $replacement)
        Write-Ok "added /folder route case"
        $changed = $true
    } else {
        Write-Bad "could not find clipboard route anchor - switch table may have been modified"
        exit 1
    }
}

# 4d. Add HandleFolder method if missing
# Built as a string array -join (each line single-quoted, so no PS variable
# expansion or escaping issues on the C# code body).
if ($content -match 'void\s+HandleFolder\s*\(') {
    Write-Skip "HandleFolder method already present"
} else {
    $methodLines = @(
        ''
        '        void HandleFolder(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)'
        '        {'
        '            var folderId = query["folderId"];'
        '            if (string.IsNullOrWhiteSpace(folderId))'
        '            {'
        '                WriteJson(ctx, 400, new { error = "folderId parameter required" });'
        '                return;'
        '            }'
        ''
        '            var folderPath = DriveHelper.FindLocalPathByFileId(folderId);'
        '            if (folderPath == null)'
        '            {'
        '                WriteJson(ctx, 404, new'
        '                {'
        '                    error = "Folder not found locally. Make sure it is synced in Google Drive for Desktop.",'
        '                    folderId'
        '                });'
        '                return;'
        '            }'
        ''
        '            if (!Directory.Exists(folderPath))'
        '            {'
        '                WriteJson(ctx, 404, new'
        '                {'
        '                    error = "Path resolved but is not a directory. Verify the ID is for a folder, not a file.",'
        '                    folderPath'
        '                });'
        '                return;'
        '            }'
        ''
        '            try'
        '            {'
        '                // Quote the path so spaces/unicode are handled correctly.'
        '                // explorer.exe returns nonzero for benign cases - do not check exit code.'
        '                Process.Start(new ProcessStartInfo'
        '                {'
        '                    FileName  = "explorer.exe",'
        '                    Arguments = $"\"{folderPath}\"",'
        '                    UseShellExecute = false'
        '                });'
        '            }'
        '            catch (Exception ex)'
        '            {'
        '                WriteJson(ctx, 500, new { error = "Failed to launch Explorer: " + ex.Message, folderPath });'
        '                return;'
        '            }'
        ''
        '            WriteJson(ctx, 200, new { success = true, folderPath });'
        '        }'
        ''
    )
    $methodCode = ($methodLines -join "`r`n") + "`r`n"

    $anchor = "        void HandleFile(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)"
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $methodCode + $anchor)
        Write-Ok "added HandleFolder method"
        $changed = $true
    } else {
        Write-Bad "could not find HandleFile anchor for inserting HandleFolder"
        exit 1
    }
}

if ($changed) {
    Save-FileContent $httpServer $content "HttpServer.cs"
} else {
    Write-Skip "HttpServer.cs already fully patched - no save needed"
}

# --- 5. Bump TrayApp.cs ------------------------------------------------------
Write-Step "Patching TrayApp.cs"
$content = Get-Content $trayApp -Raw
if ($content -match "const\s+string\s+Version\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "Version constant already at $ToVersion"
} else {
    $content = $content -replace 'const\s+string\s+Version\s+=\s*"2\.4\.0"', "const string Version  = `"$ToVersion`""
    Save-FileContent $trayApp $content "TrayApp.cs  ($FromVersion -> $ToVersion)"
}

# --- 6. Bump Installer.cs ----------------------------------------------------
Write-Step "Patching Installer.cs"
$content = Get-Content $installer -Raw
if ($content -match "const\s+string\s+AppVersion\s*=\s*`"$([regex]::Escape($ToVersion))`"") {
    Write-Skip "AppVersion constant already at $ToVersion"
} else {
    $content = $content -replace 'const\s+string\s+AppVersion\s*=\s*"2\.4\.0"', "const string AppVersion = `"$ToVersion`""
    Save-FileContent $installer $content "Installer.cs  ($FromVersion -> $ToVersion)"
}

# --- 7. Done -----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "Dry run - no files were actually modified."
    Write-Info  "Re-run without -DryRun to apply."
} else {
    Write-Ok "All C# patches applied. Source is now at $ToVersion."
    Write-Host ""
    Write-Info "Next steps:"
    Write-Info "  1. Build & deploy:  .\Rebuild.ps1"
    Write-Info "  2. Verify:          .\Check-HelperVersion.ps1"
    Write-Info "  3. Smoke-test:      curl `"http://localhost:52874/ping`""
    Write-Info "                      curl `"http://localhost:52874/folder?folderId=<a-real-folder-id>`""
    Write-Host ""
    Write-Info "FlutterFlow side (manual via FF UI - not auto-patched):"
    Write-Info "  a. Add new custom action: open_project_folder.dart"
    Write-Info "  b. Add export line to    : lib/custom_code/actions/index.dart"
    Write-Info "  c. Bump requiredVersion  : 2.5.0 in check_tabs_helper.dart"
}
Write-Host ""
