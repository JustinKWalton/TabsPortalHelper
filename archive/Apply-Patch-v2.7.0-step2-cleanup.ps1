# =============================================================================
#  Apply-Patch-v2.7.0-step2-cleanup.ps1
#
#  Removes the active-file fix flow (close-scrub-reopen via keystroke
#  dispatch) that proved unreliable in testing. The fresh-PS foreground
#  trick that works for Explorer's title bar didn't carry over to Revu's
#  document area, so save+close keystrokes weren't reaching the right
#  window.
#
#  Removed:
#    - "Bluebeam > Fix Active File Columns..." tray menu item
#    - FixActiveFileFromTray() handler in TrayApp.cs
#    - /scrub-active HTTP endpoint (route case + HandleScrubActive method)
#    - BluebeamColumnsFixOrchestrator.cs (no longer referenced; deleted)
#
#  Kept:
#    - "Bluebeam > Scrub PDF File..." tray menu item (closed-file flow,
#      works reliably -- file picker -> PdfSharp scrub -> balloon)
#    - /scrub-file HTTP endpoint
#    - ActiveFileTracker.cs (potentially useful for step 3 / dropzone flow)
#    - ActiveFileTracker.SetLastFile() call in HandleOpen (harmless,
#      records the path in process memory)
#
#  Step 3 (dropzone scrub-on-upload) replaces the active-file flow as the
#  primary mechanism for handling poisoned files: column overrides get
#  stripped client-side before the bytes ever land in Drive.
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.7.0-step2-cleanup.ps1 -DryRun
#      .\Apply-Patch-v2.7.0-step2-cleanup.ps1
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

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
        [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
        Write-Ok "saved $Label"
    }
}

# --- 1. Locate source files --------------------------------------------------
$httpServer = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp    = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$orch       = Join-Path $ProjectDir "TabsPortalHelper\BluebeamColumnsFixOrchestrator.cs"

if (-not (Test-Path $httpServer)) {
    $alt = Join-Path $ProjectDir "HttpServer.cs"
    if (Test-Path $alt) {
        $httpServer = $alt
        $trayApp    = Join-Path $ProjectDir "TrayApp.cs"
        $orch       = Join-Path $ProjectDir "BluebeamColumnsFixOrchestrator.cs"
    }
}

Write-Step "Validating files exist"
foreach ($f in @($httpServer, $trayApp)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        exit 1
    }
}

# --- 2. Patch HttpServer.cs --------------------------------------------------
Write-Step "Patching HttpServer.cs (remove /scrub-active)"
$content = Get-Content $httpServer -Raw
$changed = $false

# 2a. Remove /scrub-active route case line.
if ($content -notmatch 'case\s*"/scrub-active"') {
    Write-Skip "/scrub-active route already absent"
} else {
    $routePattern = '(?m)^\s*case\s*"/scrub-active"\s*:\s*HandleScrubActive\(ctx\);\s*break;\s*\r?\n'
    $rmatch = [regex]::Match($content, $routePattern)
    if ($rmatch.Success) {
        $content = $content.Replace($rmatch.Value, "")
        Write-Ok "removed /scrub-active route case"
        $changed = $true
    } else {
        Write-Bad "regex failed to match /scrub-active route case line"
        exit 1
    }
}

# 2b. Remove HandleScrubActive method (with its leading comment block).
#     Anchor on the "// /scrub-active" comment line and consume forward to
#     the first 8-space-indented closing brace (the method's own brace).
if ($content -notmatch 'void\s+HandleScrubActive\s*\(') {
    Write-Skip "HandleScrubActive method already absent"
} else {
    # Try with comment block first
    $methodPattern = '(?ms)\r?\n        // -+\r?\n        // /scrub-active\r?\n(?:.*?\r?\n)        void HandleScrubActive\s*\([^)]*\)\s*\r?\n        \{.*?\r?\n        \}\r?\n'
    $mmatch = [regex]::Match($content, $methodPattern)
    if ($mmatch.Success) {
        $content = $content.Replace($mmatch.Value, "")
        Write-Ok "removed HandleScrubActive method (with comment block)"
        $changed = $true
    } else {
        # Fallback: just match the method body, no comment
        $methodPatternNoComment = '(?ms)\r?\n        void HandleScrubActive\s*\([^)]*\)\s*\r?\n        \{.*?\r?\n        \}\r?\n'
        $mmatch2 = [regex]::Match($content, $methodPatternNoComment)
        if ($mmatch2.Success) {
            $content = $content.Replace($mmatch2.Value, "")
            Write-Ok "removed HandleScrubActive method (no comment block matched)"
            $changed = $true
        } else {
            Write-Bad "could not match HandleScrubActive method body"
            exit 1
        }
    }
}

if ($changed) { Save-FileContent -Path $httpServer -Content $content -Label "HttpServer.cs" }
else          { Write-Skip "HttpServer.cs already cleaned -- no save needed" }

# --- 3. Patch TrayApp.cs -----------------------------------------------------
Write-Step "Patching TrayApp.cs (remove Fix Active File Columns... menu)"
$content = Get-Content $trayApp -Raw
$changed = $false

# 3a. Remove fixActiveItem 3-line block. Pattern: optional blank line above,
#     then var/click/Add lines.
if ($content -notmatch 'var\s+fixActiveItem\s*=') {
    Write-Skip "fixActiveItem menu block already absent"
} else {
    $blockPattern = '(?ms)\r?\n\s*var fixActiveItem = new ToolStripMenuItem\("Fix Active File Columns\.\.\."\);\s*\r?\n\s*fixActiveItem\.Click \+= \(s,\s*e\) => FixActiveFileFromTray\(\);\s*\r?\n\s*bluebeamSubmenu\.DropDownItems\.Add\(fixActiveItem\);\r?\n'
    $bmatch = [regex]::Match($content, $blockPattern)
    if ($bmatch.Success) {
        $content = $content.Replace($bmatch.Value, "`r`n")
        Write-Ok "removed 'Fix Active File Columns...' menu item"
        $changed = $true
    } else {
        Write-Bad "could not match fixActiveItem 3-line block"
        exit 1
    }
}

# 3b. Remove FixActiveFileFromTray method. No comment block, just the
#     method body. Anchor on "void FixActiveFileFromTray", consume to
#     8-space-indented closing brace.
if ($content -notmatch 'void\s+FixActiveFileFromTray\s*\(') {
    Write-Skip "FixActiveFileFromTray method already absent"
} else {
    $methodPattern = '(?ms)\r?\n        void FixActiveFileFromTray\s*\([^)]*\)\s*\r?\n        \{.*?\r?\n        \}\r?\n'
    $mmatch = [regex]::Match($content, $methodPattern)
    if ($mmatch.Success) {
        $content = $content.Replace($mmatch.Value, "")
        Write-Ok "removed FixActiveFileFromTray method"
        $changed = $true
    } else {
        Write-Bad "could not match FixActiveFileFromTray method body"
        exit 1
    }
}

if ($changed) { Save-FileContent -Path $trayApp -Content $content -Label "TrayApp.cs" }
else          { Write-Skip "TrayApp.cs already cleaned -- no save needed" }

# --- 4. Delete BluebeamColumnsFixOrchestrator.cs ----------------------------
Write-Step "Removing BluebeamColumnsFixOrchestrator.cs"
if (Test-Path $orch) {
    if ($DryRun) {
        Write-Warn2 "would delete BluebeamColumnsFixOrchestrator.cs"
    } else {
        Remove-Item $orch -Force
        Write-Ok "deleted BluebeamColumnsFixOrchestrator.cs"
    }
} else {
    Write-Skip "BluebeamColumnsFixOrchestrator.cs already absent"
}

# --- 5. Done -----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "DRY RUN -- no files were modified"
} else {
    Write-Ok "step 2 cleanup applied"
    Write-Info ""
    Write-Info "Next:"
    Write-Info "  cd .\TabsPortalHelper"
    Write-Info "  dotnet build"
    Write-Info ""
    Write-Info "Then redeploy with the working pattern:"
    Write-Info "  Stop-Process -Name TabsPortalHelper -ErrorAction SilentlyContinue"
    Write-Info "  dotnet publish -c Release"
    Write-Info "  Copy-Item -Force .\bin\Release\net9.0-windows\win-x64\publish\TabsPortalHelper.exe ${env:LOCALAPPDATA}\TabsPortalHelper\TabsPortalHelper.exe"
    Write-Info "  & ${env:LOCALAPPDATA}\TabsPortalHelper\TabsPortalHelper.exe"
}
