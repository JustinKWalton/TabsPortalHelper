# =============================================================================
#  Apply-Patch-v2.5.1.ps1
#
#  Upgrades TabsPortalHelper source from v2.5.0 to v2.5.1:
#    - Bumps version constants in 4 source files
#    - Creates new ExplorerHelper.cs (launch Explorer + bring-to-front)
#    - Refactors HandleFolder in HttpServer.cs to delegate to ExplorerHelper
#      (fixes "Explorer opens but only flashes in taskbar" issue)
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.5.1.ps1
#      .\Apply-Patch-v2.5.1.ps1 -DryRun
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$FromVersion = '2.5.0'
$ToVersion   = '2.5.1'

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

$canProceed = $true

function Check-Version {
    param([string]$Path, [string]$Label, [string]$Pattern)
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
}

if (-not $canProceed) {
    Write-Host ""
    Write-Bad "Aborting: one or more files is at an unexpected version."
    Write-Info "This script upgrades $FromVersion -> $ToVersion."
    Write-Info "If your source is on 2.4.0, run Apply-Patch-v2.5.0.ps1 first."
    exit 1
}

# --- 3. Create ExplorerHelper.cs --------------------------------------------
Write-Step "Creating ExplorerHelper.cs"
if (Test-Path $explorerHelper) {
    Write-Skip "ExplorerHelper.cs already exists - leaving as-is"
} else {
    $explorerHelperLines = @(
        'using System;'
        'using System.Diagnostics;'
        'using System.IO;'
        'using System.Runtime.InteropServices;'
        'using System.Text;'
        'using System.Threading;'
        'using System.Threading.Tasks;'
        ''
        'namespace TabsPortalHelper'
        '{'
        '    static class ExplorerHelper'
        '    {'
        '        // ---- Win32 P/Invoke -----------------------------------------------'
        '        const int  SW_RESTORE      = 9;'
        '        const int  SW_SHOW         = 5;'
        '        const byte VK_MENU         = 0x12;   // Alt'
        '        const uint KEYEVENTF_KEYUP = 0x0002;'
        ''
        '        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);'
        ''
        '        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);'
        '        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);'
        '        [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);'
        '        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);'
        '        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hWnd);'
        '        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);'
        '        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);'
        '        [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hWnd);'
        '        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);'
        ''
        '        // ---- Public API ---------------------------------------------------'
        ''
        '        /// <summary>'
        '        /// Opens a folder in Windows Explorer and brings the resulting'
        '        /// window to the foreground. Returns true if explorer.exe was'
        '        /// launched. The bring-to-front step is best-effort and runs on'
        '        /// a background thread so the HTTP handler returns immediately.'
        '        /// </summary>'
        '        public static bool OpenFolder(string folderPath)'
        '        {'
        '            try'
        '            {'
        '                Process.Start(new ProcessStartInfo'
        '                {'
        '                    FileName  = "explorer.exe",'
        '                    Arguments = $"\"{folderPath}\"",'
        '                    UseShellExecute = false'
        '                });'
        '            }'
        '            catch'
        '            {'
        '                return false;'
        '            }'
        ''
        '            // Fire-and-forget: hunt for the new Explorer window and force'
        '            // it to the foreground.'
        '            Task.Run(() =>'
        '            {'
        '                Thread.Sleep(150); // let Explorer spawn the window'
        '                BringExplorerToFront(folderPath, timeoutMs: 3000);'
        '            });'
        ''
        '            return true;'
        '        }'
        ''
        '        // ---- Foreground machinery -----------------------------------------'
        ''
        '        static bool BringExplorerToFront(string folderPath, int timeoutMs)'
        '        {'
        '            // Explorer titles the window with just the leaf folder name.'
        '            var folderName = Path.GetFileName(folderPath.TrimEnd(Path.DirectorySeparatorChar));'
        '            var deadline = Environment.TickCount + timeoutMs;'
        ''
        '            while (Environment.TickCount < deadline)'
        '            {'
        '                var hWnd = FindExplorerWindowForFolder(folderName);'
        '                if (hWnd != IntPtr.Zero)'
        '                {'
        '                    ForceForeground(hWnd);'
        '                    return true;'
        '                }'
        '                Thread.Sleep(100);'
        '            }'
        '            return false;'
        '        }'
        ''
        '        static IntPtr FindExplorerWindowForFolder(string folderName)'
        '        {'
        '            IntPtr found = IntPtr.Zero;'
        '            EnumWindows((hWnd, _) =>'
        '            {'
        '                if (!IsWindowVisible(hWnd)) return true;'
        ''
        '                var cls = new StringBuilder(64);'
        '                GetClassName(hWnd, cls, cls.Capacity);'
        '                var c = cls.ToString();'
        '                // CabinetWClass = modern Win10/11 Explorer file window.'
        '                // ExploreWClass = older variant occasionally still seen.'
        '                if (c != "CabinetWClass" && c != "ExploreWClass") return true;'
        ''
        '                var title = new StringBuilder(512);'
        '                GetWindowText(hWnd, title, title.Capacity);'
        '                if (title.ToString().Equals(folderName, StringComparison.OrdinalIgnoreCase))'
        '                {'
        '                    found = hWnd;'
        '                    return false; // stop enumerating'
        '                }'
        '                return true;'
        '            }, IntPtr.Zero);'
        '            return found;'
        '        }'
        ''
        '        // Mirror of BluebeamHelper.ForceForeground - same Alt-key trick to'
        '        // grant our background thread foreground rights, then activate.'
        '        static void ForceForeground(IntPtr hWnd)'
        '        {'
        '            if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);'
        ''
        '            keybd_event(VK_MENU, 0, 0,               UIntPtr.Zero);'
        '            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);'
        ''
        '            BringWindowToTop(hWnd);'
        '            ShowWindow(hWnd, SW_SHOW);'
        '            SetForegroundWindow(hWnd);'
        '        }'
        '    }'
        '}'
        ''
    )
    $explorerHelperContent = ($explorerHelperLines -join "`r`n")
    Save-FileContent $explorerHelper $explorerHelperContent "ExplorerHelper.cs (new file)"
}

# --- 4. Refactor HandleFolder in HttpServer.cs ------------------------------
Write-Step "Refactoring HandleFolder in HttpServer.cs to use ExplorerHelper"
$content = Get-Content $httpServer -Raw

if ($content -match 'ExplorerHelper\.OpenFolder\s*\(') {
    Write-Skip "HandleFolder already delegates to ExplorerHelper"
} else {
    # Match the v2.5.0 inline Process.Start block exactly.
    $oldLines = @(
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
    )
    $oldBlock = $oldLines -join "`r`n"

    $newLines = @(
        '            if (!ExplorerHelper.OpenFolder(folderPath))'
        '            {'
        '                WriteJson(ctx, 500, new { error = "Failed to launch Explorer", folderPath });'
        '                return;'
        '            }'
    )
    $newBlock = $newLines -join "`r`n"

    if ($content.Contains($oldBlock)) {
        $content = $content.Replace($oldBlock, $newBlock)
        Save-FileContent $httpServer $content "HttpServer.cs (HandleFolder refactor)"
    } else {
        Write-Bad "could not find the v2.5.0 HandleFolder Process.Start block"
        Write-Info "HandleFolder may have been hand-edited - inspect the file and apply manually."
        exit 1
    }
}

# --- 5. Bump version constants ----------------------------------------------
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

# --- 6. Done ----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "Dry run - no files were actually modified."
    Write-Info  "Re-run without -DryRun to apply."
} else {
    Write-Ok "All patches applied. Source is now at $ToVersion."
    Write-Host ""
    Write-Info "Next steps:"
    Write-Info "  1. Build & deploy:  .\Rebuild.ps1"
    Write-Info "  2. Verify:          .\Check-HelperVersion.ps1"
    Write-Info "  3. Smoke-test:      curl `"http://localhost:52874/folder?folderId=1Sn6F7CKXFvNyiRYsf2f9BpqilL0Wt6Xz`""
    Write-Info "                      (Explorer should pop to the front, not just flash in taskbar)"
}
Write-Host ""
