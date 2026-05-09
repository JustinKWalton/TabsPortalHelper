# =============================================================================
#  Apply-Patch-v2.7.0-step2.ps1
#
#  Step 2 of v2.7.0 -- the active-file fix orchestrator + tray menu.
#
#  Prerequisites:
#    Step 1 (Apply-Patch-v2.7.0.ps1) must already be applied. This script
#    requires BsiAnnotColumnsScrubber.cs to exist.
#
#  Step 2 changes:
#    - Creates TabsPortalHelper/ActiveFileTracker.cs
#    - Creates TabsPortalHelper/BluebeamColumnsFixOrchestrator.cs
#    - Patches HttpServer.cs:
#         * Adds ActiveFileTracker.SetLastFile() in HandleOpen
#         * Adds /scrub-file route + HandleScrubFile method
#         * Adds /scrub-active route + HandleScrubActive method
#    - Patches TrayApp.cs:
#         * Replaces "Install Bluebeam Profile..." flat item with a
#           "Bluebeam" submenu containing:
#             - Install Profile...
#             - Scrub PDF File...
#             - Fix Active File Columns...
#         * Adds ScrubPdfFileFromTray and FixActiveFileFromTray methods
#
#  After this patch, run:
#    dotnet build
#  ...and exercise via tray menu or HTTP:
#    curl "http://localhost:52874/scrub-file?fileId=<drive_file_id>"
#    curl "http://localhost:52874/scrub-active"
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.7.0-step2.ps1
#      .\Apply-Patch-v2.7.0-step2.ps1 -DryRun
#      .\Apply-Patch-v2.7.0-step2.ps1 -ProjectDir 'C:\path\to\repo'
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

function Save-NewFile {
    param([string]$Path, [string]$Content, [string]$Label)
    if (Test-Path $Path) {
        $existing = Get-Content $Path -Raw
        if ($existing.TrimEnd() -eq $Content.TrimEnd()) {
            Write-Skip "$Label already present and matches expected content"
            return
        } else {
            Write-Warn2 "$Label exists but differs from bundled content"
            Write-Info "(leaving on-disk version untouched -- review manually if needed)"
            return
        }
    }
    Save-FileContent -Path $Path -Content $Content -Label $Label
}

# --- 1. Locate source files --------------------------------------------------
$csproj     = Join-Path $ProjectDir "TabsPortalHelper\TabsPortalHelper.csproj"
$httpServer = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp    = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$scrubber   = Join-Path $ProjectDir "TabsPortalHelper\BsiAnnotColumnsScrubber.cs"
$projectDirInner = Join-Path $ProjectDir "TabsPortalHelper"

if (-not (Test-Path $csproj)) {
    $alt = Join-Path $ProjectDir "TabsPortalHelper.csproj"
    if (Test-Path $alt) {
        $csproj          = $alt
        $httpServer      = Join-Path $ProjectDir "HttpServer.cs"
        $trayApp         = Join-Path $ProjectDir "TrayApp.cs"
        $scrubber        = Join-Path $ProjectDir "BsiAnnotColumnsScrubber.cs"
        $projectDirInner = $ProjectDir
    }
}

Write-Step "Validating files exist (step 1 must be applied)"
$allFound = $true
foreach ($f in @($csproj, $httpServer, $trayApp, $scrubber)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        $allFound = $false
    }
}
if (-not $allFound) {
    Write-Host ""
    Write-Bad "Source files not found. Make sure step 1 (Apply-Patch-v2.7.0.ps1) ran first."
    exit 1
}

# --- 2. Create ActiveFileTracker.cs -----------------------------------------
Write-Step "Creating ActiveFileTracker.cs"
$activeTrackerCs = @'
// ActiveFileTracker.cs
//
// Records the most recent PDF path the user worked with through a
// TABSportal flow (currently: /open). Used by BluebeamColumnsFixOrchestrator
// to know which file to scrub when the user clicks "Fix Active File
// Columns..." without having to parse Bluebeam's window title.
//
// Why this approach (vs. parsing Revu's MainWindowTitle):
//   * Title parsing only gives a filename, not a full path. Resolving back
//     to a path requires a Drive search and is fragile when the same name
//     exists in multiple project folders.
//   * Title strings vary across Revu versions (v21 vs. v2019).
//   * The user's intent is to fix the file they JUST opened from TABSportal
//     -- which is exactly what /open tells us. Trust the most recent signal.
//
// Thread safety:
//   HttpServer dispatches requests on multiple threads. Writes (from
//   HandleOpen) and reads (from the orchestrator on the tray UI thread or
//   from HandleScrubActive) can race. Single object lock around two fields
//   is plenty.

using System;

namespace TabsPortalHelper
{
    public static class ActiveFileTracker
    {
        private static readonly object _lock = new();
        private static string? _lastFilePath;
        private static DateTime _lastUpdatedUtc;

        /// <summary>
        /// Records that this file path was most recently the focus of a
        /// TABSportal flow. Idempotent. Safe to call from any thread.
        /// Whitespace / null inputs are silently ignored so callers don't
        /// have to guard.
        /// </summary>
        public static void SetLastFile(string? filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;
            lock (_lock)
            {
                _lastFilePath  = filePath;
                _lastUpdatedUtc = DateTime.UtcNow;
            }
        }

        /// <summary>
        /// Returns the most recently tracked file path and when it was
        /// recorded, or (null, default) if no file has been tracked yet.
        /// The returned path is NOT guaranteed to still exist on disk;
        /// callers must validate before use.
        /// </summary>
        public static (string? FilePath, DateTime UtcWhen) GetLastFile()
        {
            lock (_lock)
            {
                return (_lastFilePath, _lastUpdatedUtc);
            }
        }

        /// <summary>
        /// Forgets the tracked file. Useful in tests, or if a caller wants
        /// to force the next "fix active file" operation to prompt the user
        /// for a file path.
        /// </summary>
        public static void Clear()
        {
            lock (_lock)
            {
                _lastFilePath   = null;
                _lastUpdatedUtc = default;
            }
        }
    }
}
'@
$activeTrackerPath = Join-Path $projectDirInner "ActiveFileTracker.cs"
Save-NewFile -Path $activeTrackerPath -Content $activeTrackerCs -Label "ActiveFileTracker.cs"

# --- 3. Create BluebeamColumnsFixOrchestrator.cs ----------------------------
Write-Step "Creating BluebeamColumnsFixOrchestrator.cs"
$orchestratorCs = @'
// BluebeamColumnsFixOrchestrator.cs
//
// Orchestrates the end-to-end "fix the active file's column overrides" flow:
//
//   1. Determine target file
//        * Primary: ActiveFileTracker.GetLastFile() -- the file most
//          recently opened through TABSportal's /open endpoint.
//        * Interactive fallback (tray menu): OpenFileDialog so the user
//          can pick a PDF manually.
//
//   2. Check if the file even has /BSIAnnotColumns (read-only via PdfSharp).
//      If not -> AlreadyClean, return immediately. No keystrokes, no UI
//      disruption, no risk to a clean file.
//
//   3. Try to acquire exclusive write access. If we get it, the file is
//      not currently open in Revu -> just scrub directly. Done.
//
//   4. If the file is locked (open in Revu), confirm with the user, then:
//        a. Bring Revu to foreground (fresh-PS pattern -- same trick as
//           ExplorerHelper.OpenFolder, since a long-running tray process
//           can't reliably grab focus on Win10/11).
//        b. Send Ctrl+S to save any pending markups.
//        c. Send Ctrl+W to close the active document tab.
//        d. Poll for file unlock.
//        e. Scrub.
//        f. Reopen in Revu via BluebeamHelper.OpenFile.
//
// Why a fresh PowerShell process for keystrokes:
//   Windows blocks foreground-stealing from background processes. A freshly
//   spawned process gets ~1 second of foreground eligibility, which is
//   enough to AppActivate the Revu window and dispatch SendKeys cleanly.
//   We send '%' (synthetic Alt) first to top up the eligibility heuristic.
//   This is the same pattern ExplorerHelper.OpenFolder uses and is
//   documented at length there.
//
// Modes:
//   FixActiveFile(interactive=true)  -- called from tray menu. Can show
//                                       MessageBox prompts, OpenFileDialog
//                                       fallback for missing tracker.
//   FixActiveFile(interactive=false) -- called from /scrub-active HTTP
//                                       endpoint. No UI; returns NoFile
//                                       if tracker is empty rather than
//                                       prompting.

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace TabsPortalHelper
{
    public static class BluebeamColumnsFixOrchestrator
    {
        public enum Status
        {
            /// <summary>Successfully scrubbed (file was closed, or got reopened).</summary>
            Fixed,

            /// <summary>File didn't have /BSIAnnotColumns -- nothing to do.</summary>
            AlreadyClean,

            /// <summary>No file could be determined (no tracker, user cancelled picker).</summary>
            NoFile,

            /// <summary>User declined the "save and close" prompt.</summary>
            UserCancelled,

            /// <summary>Save+close+wait+scrub or reopen failed.</summary>
            Failed,
        }

        public sealed class Result
        {
            public Status     Status    { get; init; }
            public string?    FilePath  { get; init; }
            public string?    Message   { get; init; }
            public Exception? Exception { get; init; }
        }

        // ---- Tunables -----------------------------------------------------

        // How long to wait for the file lock to release after sending Ctrl+W.
        private const int FileUnlockTimeoutMs = 8_000;
        private const int FileUnlockPollMs    = 250;

        // Inter-keystroke delays. Save needs time to flush to disk; close
        // gets a beat after that to let the document tab actually close.
        private const int DelayAfterSaveMs    = 800;

        // ---- Public API ---------------------------------------------------

        /// <summary>
        /// Main entry point. Resolves the active file, scrubs it, and
        /// reopens in Bluebeam if it was open before.
        /// </summary>
        /// <param name="interactive">
        /// True when called from a UI context (tray menu) -- enables
        /// MessageBox prompts and OpenFileDialog fallback. False when
        /// called from a non-interactive context (HTTP endpoint).
        /// </param>
        public static Result FixActiveFile(bool interactive)
        {
            // 1. Resolve target file
            var filePath = ResolveTargetFile(interactive);
            if (filePath == null)
            {
                return new Result
                {
                    Status  = Status.NoFile,
                    Message = interactive
                        ? "No file was selected."
                        : "No active file is tracked. Open a file from TABSportal first, then try again.",
                };
            }

            if (!File.Exists(filePath))
            {
                return new Result
                {
                    Status   = Status.NoFile,
                    FilePath = filePath,
                    Message  = "Tracked file is no longer present on disk.",
                };
            }

            // 2. Cheap read-only check: is the file even poisoned?
            if (!BsiAnnotColumnsScrubber.HasBsiAnnotColumns(filePath))
            {
                return new Result
                {
                    Status   = Status.AlreadyClean,
                    FilePath = filePath,
                    Message  = "No column overrides found in this file -- your TABS columns should already be visible. (If they are not, the issue is in your active Bluebeam profile, not the file.)",
                };
            }

            // 3. Is the file currently open / locked exclusive somewhere?
            bool fileIsLocked = !TryAcquireExclusive(filePath);

            if (!fileIsLocked)
            {
                // Easy path: file is closed, just scrub it.
                var directResult = BsiAnnotColumnsScrubber.Scrub(filePath);
                return new Result
                {
                    Status    = directResult.Success ? Status.Fixed : Status.Failed,
                    FilePath  = filePath,
                    Message   = directResult.Message,
                    Exception = directResult.Exception,
                };
            }

            // 4. File is locked. Confirm before save+close+reopen dance.
            if (interactive)
            {
                var fileName = Path.GetFileName(filePath);
                var answer = MessageBox.Show(
                    "This file is open in Bluebeam:\n\n" +
                    fileName + "\n\n" +
                    "To fix the column override, the helper needs to:\n" +
                    "  1. Save any pending markups (Ctrl+S)\n" +
                    "  2. Close the document in Bluebeam\n" +
                    "  3. Remove the embedded column overrides\n" +
                    "  4. Reopen the file in Bluebeam\n\n" +
                    "Continue?",
                    "Fix Bluebeam Column Overrides",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);

                if (answer != DialogResult.Yes)
                {
                    return new Result
                    {
                        Status   = Status.UserCancelled,
                        FilePath = filePath,
                        Message  = "Cancelled.",
                    };
                }
            }
            // Non-interactive: HTTP caller is implicitly authorizing. Proceed.

            // 5. Send Ctrl+S then Ctrl+W to Revu via fresh-PS pattern.
            if (!DispatchSaveAndCloseToRevu())
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = filePath,
                    Message  = "Couldn't dispatch save+close keystrokes to Bluebeam. The file may not be the active document. Save and close the file manually, then click Fix again.",
                };
            }

            // 6. Wait for file lock to release.
            if (!WaitForUnlock(filePath, FileUnlockTimeoutMs))
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = filePath,
                    Message  = "Timed out waiting for the file to close in Bluebeam. If the file has unsaved changes Bluebeam may be prompting you -- save and close manually, then try again.",
                };
            }

            // 7. Scrub.
            var scrubResult = BsiAnnotColumnsScrubber.Scrub(filePath);
            if (!scrubResult.Success)
            {
                // Best-effort reopen so the user isn't stranded with no document.
                try { BluebeamHelper.OpenFile(filePath); } catch { /* best effort */ }

                return new Result
                {
                    Status    = Status.Failed,
                    FilePath  = filePath,
                    Message   = "Scrub failed: " + (scrubResult.Message ?? "unknown error"),
                    Exception = scrubResult.Exception,
                };
            }

            // 8. Reopen in Revu.
            try
            {
                BluebeamHelper.OpenFile(filePath);
            }
            catch (Exception ex)
            {
                return new Result
                {
                    Status    = Status.Failed,
                    FilePath  = filePath,
                    Message   = "Scrubbed successfully, but reopen failed: " + ex.Message + ". Open the file manually.",
                    Exception = ex,
                };
            }

            return new Result
            {
                Status   = Status.Fixed,
                FilePath = filePath,
                Message  = "Removed column overrides and reopened.",
            };
        }

        // ---- Helpers ------------------------------------------------------

        private static string? ResolveTargetFile(bool interactive)
        {
            // Option 3 (preferred): tracker
            var (lastPath, _) = ActiveFileTracker.GetLastFile();
            if (!string.IsNullOrWhiteSpace(lastPath) && File.Exists(lastPath))
            {
                return lastPath;
            }

            // Interactive fallback: file picker
            if (interactive)
            {
                using var ofd = new OpenFileDialog
                {
                    Title           = "Select PDF to fix",
                    Filter          = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
                    CheckFileExists = true,
                    Multiselect     = false,
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    return ofd.FileName;
                }
            }

            return null;
        }

        /// <summary>
        /// True if we can briefly take exclusive access to the file --
        /// i.e. nothing else has it locked. We open with FileShare.None
        /// and dispose immediately; this is just a probe, not a hold.
        /// </summary>
        private static bool TryAcquireExclusive(string filePath)
        {
            try
            {
                using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                return true;
            }
            catch (IOException)   { return false; }
            catch (UnauthorizedAccessException) { return false; }
        }

        private static bool WaitForUnlock(string filePath, int timeoutMs)
        {
            var deadline = Environment.TickCount + timeoutMs;
            while (Environment.TickCount < deadline)
            {
                if (TryAcquireExclusive(filePath)) return true;
                Thread.Sleep(FileUnlockPollMs);
            }
            return false;
        }

        /// <summary>
        /// Spawns a fresh powershell.exe (which gets ~1s of foreground
        /// eligibility) to activate the Revu window and dispatch:
        ///   1. Synthetic Alt (top off foreground heuristic)
        ///   2. AppActivate('Revu') with retry
        ///   3. Ctrl+S
        ///   4. Brief sleep to let save flush
        ///   5. Ctrl+W (close active document)
        ///
        /// Returns true if the PS process exited cleanly within a reasonable
        /// timeout. False on launch failure or hang.
        /// </summary>
        private static bool DispatchSaveAndCloseToRevu()
        {
            // PS command. Each statement separated by ';'. Single line so
            // we don't fight ProcessStartInfo line-break handling.
            //
            // SendKeys syntax notes:
            //   '%'    = Alt
            //   '^s'   = Ctrl+S
            //   '^w'   = Ctrl+W
            //
            // The AppActivate retry loop accommodates Revu being momentarily
            // unfocusable (e.g., a save dialog briefly steals focus). We try
            // up to ~1.2s before giving up.
            var psCommand =
                "Add-Type -AssemblyName System.Windows.Forms; " +
                "Add-Type -AssemblyName Microsoft.VisualBasic; " +
                "[System.Windows.Forms.SendKeys]::SendWait('%'); " +
                "$activated = $false; " +
                "for ($i = 0; $i -lt 12; $i++) { " +
                "  try { [Microsoft.VisualBasic.Interaction]::AppActivate('Revu'); $activated = $true; break } " +
                "  catch { Start-Sleep -Milliseconds 100 } " +
                "}; " +
                "if (-not $activated) { exit 1 }; " +
                "Start-Sleep -Milliseconds 250; " +
                "[System.Windows.Forms.SendKeys]::SendWait('^s'); " +
                "Start-Sleep -Milliseconds " + DelayAfterSaveMs + "; " +
                "[System.Windows.Forms.SendKeys]::SendWait('^w');";

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName        = "powershell.exe",
                    UseShellExecute = false,
                    CreateNoWindow  = true,
                };
                // ArgumentList properly quotes each entry -- no manual escaping needed.
                psi.ArgumentList.Add("-NoProfile");
                psi.ArgumentList.Add("-ExecutionPolicy");
                psi.ArgumentList.Add("Bypass");
                psi.ArgumentList.Add("-Command");
                psi.ArgumentList.Add(psCommand);

                using var proc = Process.Start(psi);
                if (proc == null) return false;
                if (!proc.WaitForExit(5_000))
                {
                    try { proc.Kill(); } catch { }
                    return false;
                }
                return proc.ExitCode == 0;
            }
            catch
            {
                return false;
            }
        }
    }
}
'@
$orchestratorPath = Join-Path $projectDirInner "BluebeamColumnsFixOrchestrator.cs"
Save-NewFile -Path $orchestratorPath -Content $orchestratorCs -Label "BluebeamColumnsFixOrchestrator.cs"

# --- 4. Patch HttpServer.cs --------------------------------------------------
Write-Step "Patching HttpServer.cs"
$content = Get-Content $httpServer -Raw
$changed = $false

# 4a. Add tracker call inside HandleOpen, just before the WriteJson success line.
# Anchor on the existing successful response line so we slot the call in
# right after BluebeamHelper.OpenFile has dispatched.
if ($content -match 'ActiveFileTracker\.SetLastFile') {
    Write-Skip "ActiveFileTracker.SetLastFile call already present in HandleOpen"
} else {
    $anchor = '            var launched = BluebeamHelper.OpenFile(filePath);' + "`r`n" +
              '            WriteJson(ctx, 200, new { success = true, filePath, openedInBluebeam = launched });'
    $replacement = '            var launched = BluebeamHelper.OpenFile(filePath);' + "`r`n" +
                   '            ActiveFileTracker.SetLastFile(filePath);' + "`r`n" +
                   '            WriteJson(ctx, 200, new { success = true, filePath, openedInBluebeam = launched });'
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $replacement)
        Write-Ok "added ActiveFileTracker.SetLastFile call to HandleOpen"
        $changed = $true
    } else {
        Write-Bad "could not find HandleOpen success-line anchor (HandleOpen may have been edited)"
        exit 1
    }
}

# 4b. Add /scrub-file and /scrub-active routes
if ($content -match 'case\s*"/scrub-file"\s*:') {
    Write-Skip "/scrub-file and /scrub-active routes already present"
} else {
    $routeAnchor = 'case "/folder":                   HandleFolder(ctx, query);       break;'
    $routeReplacement = $routeAnchor + "`r`n" +
                        '                    case "/scrub-file":               HandleScrubFile(ctx, query);   break;' + "`r`n" +
                        '                    case "/scrub-active":             HandleScrubActive(ctx);        break;'
    if ($content.Contains($routeAnchor)) {
        $content = $content.Replace($routeAnchor, $routeReplacement)
        Write-Ok "added /scrub-file and /scrub-active route cases"
        $changed = $true
    } else {
        Write-Bad "could not find /folder route anchor (route table may have been edited)"
        exit 1
    }
}

# 4c. Add HandleScrubFile + HandleScrubActive method bodies.
# We insert them just before the final closing brace of the HttpServer
# class. Anchor on the trailing braces of the file (last "}\r\n}\r\n").
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
        ''
    )
    $methodCode = ($methodLines -join "`r`n") + "`r`n"

    # Anchor on the closing of the HttpServer class. The file ends with
    # something like:
    #         }   <-- last method's closing brace
    #     }       <-- class closing brace
    # }           <-- namespace closing brace
    # We slot our methods in just before the class closing brace. The
    # safest anchor is the literal "    }\r\n}\r\n" at end-of-file (class
    # close + namespace close).
    $endAnchor = "    }`r`n}`r`n"
    $endReplacement = $methodCode + "    }`r`n}`r`n"
    if ($content.EndsWith($endAnchor)) {
        $content = $content.Substring(0, $content.Length - $endAnchor.Length) + $endReplacement
        Write-Ok "added HandleScrubFile + HandleScrubActive methods"
        $changed = $true
    } else {
        # Fallback: try without trailing newline on namespace close
        $endAnchor2 = "    }`r`n}"
        if ($content.EndsWith($endAnchor2)) {
            $content = $content.Substring(0, $content.Length - $endAnchor2.Length) + $methodCode + "    }`r`n}"
            Write-Ok "added HandleScrubFile + HandleScrubActive methods (fallback anchor)"
            $changed = $true
        } else {
            Write-Bad "could not find HttpServer class closing brace anchor at EOF"
            exit 1
        }
    }
}

if ($changed) { Save-FileContent -Path $httpServer -Content $content -Label "HttpServer.cs" }
else          { Write-Skip "HttpServer.cs already fully patched -- no save needed" }

# --- 5. Patch TrayApp.cs -----------------------------------------------------
Write-Step "Patching TrayApp.cs"
$content = Get-Content $trayApp -Raw
$changed = $false

# 5a. Add `using System.IO;` if missing (needed for Path in scrub handler)
if ($content -match 'using\s+System\.IO\s*;') {
    Write-Skip "'using System.IO;' already present"
} else {
    $usingAnchor = 'using System.Drawing;'
    if ($content.Contains($usingAnchor)) {
        $content = $content.Replace($usingAnchor, $usingAnchor + "`r`n" + 'using System.IO;')
        Write-Ok "added 'using System.IO;'"
        $changed = $true
    } else {
        Write-Warn2 "could not find 'using System.Drawing;' anchor; will rely on existing usings"
    }
}

# 5b. Replace the flat "Install Bluebeam Profile..." menu item with a
#     "Bluebeam" submenu containing 3 items (Install Profile, Scrub PDF File,
#     Fix Active File Columns).
if ($content -match 'var\s+bluebeamSubmenu\s*=') {
    Write-Skip "Bluebeam submenu already present in TrayApp"
} else {
    $oldBlock = @"
            var profileItem = new ToolStripMenuItem("Install Bluebeam Profile...");
            profileItem.Click += (s, e) => InstallBluebeamProfile();
            menu.Items.Add(profileItem);
"@
    $newBlock = @"
            var bluebeamSubmenu = new ToolStripMenuItem("Bluebeam");

            var profileItem = new ToolStripMenuItem("Install Profile...");
            profileItem.Click += (s, e) => InstallBluebeamProfile();
            bluebeamSubmenu.DropDownItems.Add(profileItem);

            var scrubFileItem = new ToolStripMenuItem("Scrub PDF File...");
            scrubFileItem.Click += (s, e) => ScrubPdfFileFromTray();
            bluebeamSubmenu.DropDownItems.Add(scrubFileItem);

            var fixActiveItem = new ToolStripMenuItem("Fix Active File Columns...");
            fixActiveItem.Click += (s, e) => FixActiveFileFromTray();
            bluebeamSubmenu.DropDownItems.Add(fixActiveItem);

            menu.Items.Add(bluebeamSubmenu);
"@
    if ($content.Contains($oldBlock)) {
        $content = $content.Replace($oldBlock, $newBlock)
        Write-Ok "replaced flat 'Install Bluebeam Profile' item with Bluebeam submenu"
        $changed = $true
    } else {
        Write-Bad "could not find existing 'Install Bluebeam Profile...' menu block"
        Write-Info "(TrayApp.cs may have been hand-edited; menu setup needs manual update)"
        exit 1
    }
}

# 5c. Add ScrubPdfFileFromTray + FixActiveFileFromTray methods.
# Anchor before the existing InstallBluebeamProfile method (or whatever the
# next method is). Safest: insert before the namespace's class closing brace.
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
        ''
    )
    $handlerCode = ($handlerLines -join "`r`n") + "`r`n"

    # Anchor at end of TrayApp class. Same approach as HttpServer end-anchor.
    $endAnchor = "    }`r`n}`r`n"
    if ($content.EndsWith($endAnchor)) {
        $content = $content.Substring(0, $content.Length - $endAnchor.Length) + $handlerCode + $endAnchor
        Write-Ok "added ScrubPdfFileFromTray and FixActiveFileFromTray methods"
        $changed = $true
    } else {
        $endAnchor2 = "    }`r`n}"
        if ($content.EndsWith($endAnchor2)) {
            $content = $content.Substring(0, $content.Length - $endAnchor2.Length) + $handlerCode + $endAnchor2
            Write-Ok "added ScrubPdfFileFromTray and FixActiveFileFromTray methods (fallback anchor)"
            $changed = $true
        } else {
            Write-Bad "could not find TrayApp class closing brace anchor at EOF"
            exit 1
        }
    }
}

if ($changed) { Save-FileContent -Path $trayApp -Content $content -Label "TrayApp.cs" }
else          { Write-Skip "TrayApp.cs already fully patched -- no save needed" }

# --- 6. Done -----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "DRY RUN -- no files were modified"
} else {
    Write-Ok "v2.7.0 step 2 applied successfully"
    Write-Info ""
    Write-Info "Next:"
    Write-Info "  cd .\TabsPortalHelper"
    Write-Info "  dotnet build"
    Write-Info ""
    Write-Info "After build, exercise the new flow via:"
    Write-Info "  - Tray menu: right-click tray icon > Bluebeam > Scrub PDF File..."
    Write-Info "  - Tray menu: right-click tray icon > Bluebeam > Fix Active File Columns..."
    Write-Info "  - HTTP: GET http://localhost:52874/scrub-file?fileId=<drive_id>"
    Write-Info "  - HTTP: GET http://localhost:52874/scrub-active"
}
