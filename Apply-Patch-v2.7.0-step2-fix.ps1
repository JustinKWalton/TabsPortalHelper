# =============================================================================
#  Apply-Patch-v2.7.0-step2-fix.ps1
#
#  Continuation of step 2 (active-file fix orchestrator). Picks up where
#  Apply-Patch-v2.7.0-step2.ps1 left off after the HttpServer.cs anchor
#  match failed in the previous run.
#
#  Differences from the original step2:
#    - Uses [regex]::Replace with \r?\n line-ending patterns and \s+ for
#      flexible whitespace, instead of exact-string .Contains matching.
#    - Skips file creation entirely (ActiveFileTracker.cs and
#      BluebeamColumnsFixOrchestrator.cs already exist on disk from the
#      previous partial run).
#    - All other behavior is identical: idempotent, dry-run support,
#      [OK]/[SKIP]/[WARN]/[FAIL] output.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.7.0-step2-fix.ps1 -DryRun
#      .\Apply-Patch-v2.7.0-step2-fix.ps1
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
$tracker    = Join-Path $ProjectDir "TabsPortalHelper\ActiveFileTracker.cs"
$orch       = Join-Path $ProjectDir "TabsPortalHelper\BluebeamColumnsFixOrchestrator.cs"

if (-not (Test-Path $httpServer)) {
    $alt = Join-Path $ProjectDir "HttpServer.cs"
    if (Test-Path $alt) {
        $httpServer = $alt
        $trayApp    = Join-Path $ProjectDir "TrayApp.cs"
        $tracker    = Join-Path $ProjectDir "ActiveFileTracker.cs"
        $orch       = Join-Path $ProjectDir "BluebeamColumnsFixOrchestrator.cs"
    }
}

Write-Step "Validating files exist"
$allFound = $true
foreach ($f in @($httpServer, $trayApp, $tracker, $orch)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        $allFound = $false
    }
}
if (-not $allFound) {
    Write-Host ""
    Write-Bad "Required files missing. Run Apply-Patch-v2.7.0-step2.ps1 first to create the new files."
    exit 1
}

# --- 2. Patch HttpServer.cs (regex-based) -----------------------------------
Write-Step "Patching HttpServer.cs (regex matching)"
$content = Get-Content $httpServer -Raw
$changed = $false

# 2a. Add tracker call right after the OpenFile line. We match just the
#     one line and append the tracker call after it; this is robust against
#     line ending and trailing-line-formatting variations.
if ($content -match 'ActiveFileTracker\.SetLastFile') {
    Write-Skip "ActiveFileTracker.SetLastFile call already present"
} else {
    # Capture leading whitespace so the inserted line matches surrounding indentation.
    $pattern = '(?m)^(\s*)var launched = BluebeamHelper\.OpenFile\(filePath\);\s*$'
    $matchInfo = [regex]::Match($content, $pattern)
    if (-not $matchInfo.Success) {
        Write-Bad "could not find 'var launched = BluebeamHelper.OpenFile(filePath);' in HandleOpen"
        Write-Info "Run this to inspect what's actually there:"
        Write-Info "  Get-Content $httpServer -Raw | Select-String -Pattern 'BluebeamHelper\.OpenFile' -Context 2,3"
        exit 1
    }

    $indent = $matchInfo.Groups[1].Value
    $insertion = $matchInfo.Value + "`r`n" + $indent + "ActiveFileTracker.SetLastFile(filePath);"
    $content = [regex]::Replace($content, $pattern, [System.Text.RegularExpressions.Regex]::Escape($insertion).Replace('\$', '$$$$'), 1)
    Write-Ok "added ActiveFileTracker.SetLastFile call to HandleOpen"
    $changed = $true
}

# 2b. Add /scrub-file and /scrub-active routes. Match the /folder route case
#     with flexible whitespace; insert two new cases right after.
if ($content -match 'case\s*"/scrub-file"\s*:') {
    Write-Skip "/scrub-file and /scrub-active routes already present"
} else {
    $routePattern = '(?m)^(\s*)case\s*"/folder"\s*:\s*HandleFolder\(ctx,\s*query\);\s*break;\s*$'
    $routeMatch = [regex]::Match($content, $routePattern)
    if (-not $routeMatch.Success) {
        Write-Bad "could not find /folder route case (route table may have been edited)"
        exit 1
    }
    $caseIndent = $routeMatch.Groups[1].Value
    $newCases = $routeMatch.Value + "`r`n" +
                $caseIndent + 'case "/scrub-file":               HandleScrubFile(ctx, query);   break;' + "`r`n" +
                $caseIndent + 'case "/scrub-active":             HandleScrubActive(ctx);        break;'
    # Use a plain substring replace to avoid regex substitution gotchas.
    $content = $content.Replace($routeMatch.Value, $newCases)
    Write-Ok "added /scrub-file and /scrub-active route cases"
    $changed = $true
}

# 2c. Add HandleScrubFile + HandleScrubActive method bodies. We need to
#     insert them inside the HttpServer class, before the closing brace.
#     Strategy: find the last "        }" (8-space indented closing brace --
#     the last method's closing brace) and insert our methods after it but
#     before the class brace "    }".
if ($content -match 'void\s+HandleScrubFile\s*\(') {
    Write-Skip "HandleScrubFile method already present"
} else {
    $methodLines = @(
        ''
        '        // --------------------------------------------------------'
        '        // /scrub-file?fileId=<drive_id>'
        '        //'
        '        // Resolves the Drive fileId to a local path (same pattern as'
        '        // /open and /file), then strips /BSIAnnotColumns from the'
        '        // PDF Catalog. Idempotent: a clean file returns AlreadyClean.'
        '        // The file does not need to be open in Bluebeam; this is the'
        '        // closed-file flow.'
        '        // --------------------------------------------------------'
        '        void HandleScrubFile(HttpListenerContext ctx, System.Collections.Specialized.NameValueCollection query)'
        '        {'
        '            var fileId = query["fileId"];'
        '            if (string.IsNullOrWhiteSpace(fileId))'
        '            {'
        '                WriteJson(ctx, 400, new { error = "fileId parameter required" });'
        '                return;'
        '            }'
        ''
        '            var filePath = DriveHelper.FindLocalPathByFileId(fileId);'
        '            if (filePath == null || !File.Exists(filePath))'
        '            {'
        '                WriteJson(ctx, 404, new { error = "File not found locally", fileId });'
        '                return;'
        '            }'
        ''
        '            var result = BsiAnnotColumnsScrubber.Scrub(filePath);'
        '            var statusCode = result.Success ? 200 : 500;'
        '            WriteJson(ctx, statusCode, new {'
        '                success  = result.Success,'
        '                status   = result.Status.ToString(),'
        '                message  = result.Message,'
        '                filePath = filePath,'
        '            });'
        '        }'
        ''
        '        // --------------------------------------------------------'
        '        // /scrub-active'
        '        //'
        '        // Runs the close-scrub-reopen orchestration on whatever file'
        '        // ActiveFileTracker thinks is currently active (most recently'
        '        // /opened). Non-interactive: returns NoFile rather than'
        '        // prompting if the tracker is empty.'
        '        // --------------------------------------------------------'
        '        void HandleScrubActive(HttpListenerContext ctx)'
        '        {'
        '            var result = BluebeamColumnsFixOrchestrator.FixActiveFile(interactive: false);'
        '            var statusCode = result.Status == BluebeamColumnsFixOrchestrator.Status.Failed ? 500 : 200;'
        '            WriteJson(ctx, statusCode, new {'
        '                status   = result.Status.ToString(),'
        '                message  = result.Message,'
        '                filePath = result.FilePath,'
        '            });'
        '        }'
    )
    $methodCode = ($methodLines -join "`r`n")

    # Find the last closing brace of a method (8-space indented "}" at line
    # start), and insert our methods right after it. The class's closing
    # brace is 4-space indented "    }" so it won't match.
    $closeMatches = [regex]::Matches($content, '(?m)^        \}\s*$')
    if ($closeMatches.Count -eq 0) {
        Write-Bad "could not find any 8-space-indented closing brace in HttpServer.cs"
        exit 1
    }
    $lastMethodClose = $closeMatches[$closeMatches.Count - 1]
    $insertAt = $lastMethodClose.Index + $lastMethodClose.Length
    $content = $content.Substring(0, $insertAt) + "`r`n" + $methodCode + $content.Substring($insertAt)
    Write-Ok "added HandleScrubFile + HandleScrubActive methods (inserted after last method close)"
    $changed = $true
}

if ($changed) { Save-FileContent -Path $httpServer -Content $content -Label "HttpServer.cs" }
else          { Write-Skip "HttpServer.cs already fully patched -- no save needed" }

# --- 3. Patch TrayApp.cs (regex-based) --------------------------------------
Write-Step "Patching TrayApp.cs (regex matching)"
$content = Get-Content $trayApp -Raw
$changed = $false

# 3a. Add 'using System.IO;' if missing.
if ($content -match 'using\s+System\.IO\s*;') {
    Write-Skip "'using System.IO;' already present"
} else {
    $usingPattern = '(?m)^(using\s+System\.Drawing\s*;)\s*$'
    $usingMatch = [regex]::Match($content, $usingPattern)
    if ($usingMatch.Success) {
        $content = $content.Replace($usingMatch.Value, $usingMatch.Value + "`r`n" + 'using System.IO;')
        Write-Ok "added 'using System.IO;'"
        $changed = $true
    } else {
        Write-Warn2 "could not find 'using System.Drawing;' anchor; skipping (other usings may cover Path)"
    }
}

# 3b. Replace the flat "Install Bluebeam Profile..." menu block with the
#     Bluebeam submenu. Use regex with flexible whitespace so it works
#     regardless of indentation/line-ending style.
if ($content -match 'var\s+bluebeamSubmenu\s*=') {
    Write-Skip "Bluebeam submenu already present in TrayApp"
} else {
    # The original block has 3 statements:
    #   var profileItem = new ToolStripMenuItem("Install Bluebeam Profile...");
    #   profileItem.Click += (s, e) => InstallBluebeamProfile();
    #   menu.Items.Add(profileItem);
    $oldPattern = '(?ms)^(\s*)var profileItem = new ToolStripMenuItem\("Install Bluebeam Profile\.\.\."\)\s*;\s*\r?\n' +
                  '\s*profileItem\.Click \+= \(s,\s*e\) => InstallBluebeamProfile\(\)\s*;\s*\r?\n' +
                  '\s*menu\.Items\.Add\(profileItem\)\s*;'
    $oldMatch = [regex]::Match($content, $oldPattern)
    if (-not $oldMatch.Success) {
        Write-Bad "could not find 'Install Bluebeam Profile...' menu block"
        Write-Info "Inspect with:"
        Write-Info "  Get-Content $trayApp -Raw | Select-String -Pattern 'Install Bluebeam Profile' -Context 2,3"
        exit 1
    }

    $indent = $oldMatch.Groups[1].Value
    $newBlock =
        $indent + 'var bluebeamSubmenu = new ToolStripMenuItem("Bluebeam");' + "`r`n" +
        "`r`n" +
        $indent + 'var profileItem = new ToolStripMenuItem("Install Profile...");' + "`r`n" +
        $indent + 'profileItem.Click += (s, e) => InstallBluebeamProfile();' + "`r`n" +
        $indent + 'bluebeamSubmenu.DropDownItems.Add(profileItem);' + "`r`n" +
        "`r`n" +
        $indent + 'var scrubFileItem = new ToolStripMenuItem("Scrub PDF File...");' + "`r`n" +
        $indent + 'scrubFileItem.Click += (s, e) => ScrubPdfFileFromTray();' + "`r`n" +
        $indent + 'bluebeamSubmenu.DropDownItems.Add(scrubFileItem);' + "`r`n" +
        "`r`n" +
        $indent + 'var fixActiveItem = new ToolStripMenuItem("Fix Active File Columns...");' + "`r`n" +
        $indent + 'fixActiveItem.Click += (s, e) => FixActiveFileFromTray();' + "`r`n" +
        $indent + 'bluebeamSubmenu.DropDownItems.Add(fixActiveItem);' + "`r`n" +
        "`r`n" +
        $indent + 'menu.Items.Add(bluebeamSubmenu);'

    $content = $content.Replace($oldMatch.Value, $newBlock)
    Write-Ok "replaced flat 'Install Bluebeam Profile' item with Bluebeam submenu"
    $changed = $true
}

# 3c. Add ScrubPdfFileFromTray + FixActiveFileFromTray methods.
if ($content -match 'void\s+ScrubPdfFileFromTray\s*\(') {
    Write-Skip "tray handler methods already present"
} else {
    $handlerLines = @(
        ''
        '        // ----------------------------------------------------------------'
        '        // Tray menu handlers for the Bluebeam submenu'
        '        // ----------------------------------------------------------------'
        '        void ScrubPdfFileFromTray()'
        '        {'
        '            using var ofd = new OpenFileDialog'
        '            {'
        '                Title           = "Select PDF to scrub",'
        '                Filter          = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",'
        '                CheckFileExists = true,'
        '                Multiselect     = false,'
        '            };'
        '            if (ofd.ShowDialog() != DialogResult.OK) return;'
        ''
        '            var result = BsiAnnotColumnsScrubber.Scrub(ofd.FileName);'
        '            var fileName = Path.GetFileName(ofd.FileName);'
        ''
        '            switch (result.Status)'
        '            {'
        '                case BsiAnnotColumnsScrubber.Status.Scrubbed:'
        '                    _trayIcon.ShowBalloonTip(3000, "Scrub complete",'
        '                        "Removed Bluebeam column overrides from " + fileName + ".",'
        '                        ToolTipIcon.Info);'
        '                    break;'
        '                case BsiAnnotColumnsScrubber.Status.AlreadyClean:'
        '                    _trayIcon.ShowBalloonTip(3000, "Already clean",'
        '                        fileName + " has no column overrides -- nothing to do.",'
        '                        ToolTipIcon.Info);'
        '                    break;'
        '                case BsiAnnotColumnsScrubber.Status.Failed:'
        '                default:'
        '                    _trayIcon.ShowBalloonTip(5000, "Scrub failed",'
        '                        result.Message ?? "Unknown error.",'
        '                        ToolTipIcon.Error);'
        '                    break;'
        '            }'
        '        }'
        ''
        '        void FixActiveFileFromTray()'
        '        {'
        '            var result = BluebeamColumnsFixOrchestrator.FixActiveFile(interactive: true);'
        ''
        '            string title; ToolTipIcon icon;'
        '            switch (result.Status)'
        '            {'
        '                case BluebeamColumnsFixOrchestrator.Status.Fixed:'
        '                    title = "Fixed"; icon = ToolTipIcon.Info; break;'
        '                case BluebeamColumnsFixOrchestrator.Status.AlreadyClean:'
        '                    title = "No fix needed"; icon = ToolTipIcon.Info; break;'
        '                case BluebeamColumnsFixOrchestrator.Status.UserCancelled:'
        '                    return; // no notification on user cancel'
        '                case BluebeamColumnsFixOrchestrator.Status.NoFile:'
        '                    title = "No active file"; icon = ToolTipIcon.Warning; break;'
        '                default:'
        '                    title = "Fix failed"; icon = ToolTipIcon.Error; break;'
        '            }'
        ''
        '            _trayIcon.ShowBalloonTip(4500, title,'
        '                result.Message ?? "",'
        '                icon);'
        '        }'
    )
    $handlerCode = ($handlerLines -join "`r`n")

    # Insert after the last 8-space-indented closing brace (last method close).
    $closeMatches = [regex]::Matches($content, '(?m)^        \}\s*$')
    if ($closeMatches.Count -eq 0) {
        Write-Bad "could not find any 8-space-indented closing brace in TrayApp.cs"
        exit 1
    }
    $lastMethodClose = $closeMatches[$closeMatches.Count - 1]
    $insertAt = $lastMethodClose.Index + $lastMethodClose.Length
    $content = $content.Substring(0, $insertAt) + "`r`n" + $handlerCode + $content.Substring($insertAt)
    Write-Ok "added ScrubPdfFileFromTray and FixActiveFileFromTray methods"
    $changed = $true
}

if ($changed) { Save-FileContent -Path $trayApp -Content $content -Label "TrayApp.cs" }
else          { Write-Skip "TrayApp.cs already fully patched -- no save needed" }

# --- 4. Done -----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "DRY RUN -- no files were modified"
} else {
    Write-Ok "step 2 fix-up applied successfully"
    Write-Info ""
    Write-Info "Next:"
    Write-Info "  cd .\TabsPortalHelper"
    Write-Info "  dotnet build"
}
