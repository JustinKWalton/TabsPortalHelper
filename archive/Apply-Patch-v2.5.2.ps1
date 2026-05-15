# =============================================================================
#  Apply-Patch-v2.5.2.ps1
#
#  Upgrades TabsPortalHelper source from v2.5.1 to v2.5.2:
#    - Strengthens ExplorerHelper.ForceForeground using AttachThreadInput
#      + SwitchToThisWindow, fixing the "Explorer flashes red in taskbar
#      instead of activating" issue when /folder is called from a browser
#      (the Alt-key-only approach worked from PowerShell but not from FF).
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.5.2.ps1
#      .\Apply-Patch-v2.5.2.ps1 -DryRun
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$FromVersion = '2.5.1'
$ToVersion   = '2.5.2'

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
foreach ($f in @($csproj, $httpServer, $trayApp, $installer, $explorerHelper)) {
    if (Test-Path $f) {
        Write-Ok ("found {0}" -f (Split-Path $f -Leaf))
    } else {
        Write-Bad ("missing {0}" -f $f)
        $allFound = $false
    }
}
if (-not $allFound) {
    Write-Host ""
    Write-Bad "Source files not found. Run Apply-Patch-v2.5.1.ps1 first if ExplorerHelper.cs is missing."
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
    exit 1
}

# --- 3. Patch ExplorerHelper.cs ---------------------------------------------
Write-Step "Patching ExplorerHelper.cs (stronger ForceForeground)"
$content = Get-Content $explorerHelper -Raw

if ($content -match 'AttachThreadInput') {
    Write-Skip "ExplorerHelper.cs already has AttachThreadInput (v2.5.2 behavior)"
} else {
    if ($content -notmatch 'Mirror of BluebeamHelper\.ForceForeground') {
        Write-Bad "ExplorerHelper.cs doesn't look like the v2.5.1 version - inspect manually"
        exit 1
    }

    # 3a. Add new P/Invoke declarations after keybd_event line
    $oldPInvokes = '        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);'
    $newPInvokeLines = @(
        '        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);'
        '        [DllImport("user32.dll")] static extern bool   AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);'
        '        [DllImport("user32.dll")] static extern uint   GetWindowThreadProcessId(IntPtr hWnd, out uint pid);'
        '        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();'
        '        [DllImport("user32.dll")] static extern void   SwitchToThisWindow(IntPtr hWnd, bool fAltTab);'
        '        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();'
    )
    $newPInvokes = $newPInvokeLines -join "`r`n"

    if ($content.Contains($oldPInvokes)) {
        $content = $content.Replace($oldPInvokes, $newPInvokes)
        Write-Ok "added new P/Invoke declarations"
    } else {
        Write-Bad "could not find keybd_event P/Invoke anchor"
        exit 1
    }

    # 3b. Replace ForceForeground method body
    $oldMethodLines = @(
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
    )
    $oldMethod = $oldMethodLines -join "`r`n"

    $newMethodLines = @(
        '        // Aggressive foreground grab: when /folder is called from a browser'
        '        // (FlutterFlow), the browser keeps foreground throughout the fetch'
        '        // and SetForegroundWindow alone gets demoted to a taskbar flash. The'
        '        // sequence below uses AttachThreadInput to inherit the current'
        '        // foreground thread''s rights, then SwitchToThisWindow as the most'
        '        // lenient activation API (Windows treats it like Alt+Tab).'
        '        static void ForceForeground(IntPtr hWnd)'
        '        {'
        '            if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);'
        ''
        '            var foreHwnd     = GetForegroundWindow();'
        '            var foreThread   = GetWindowThreadProcessId(foreHwnd, out _);'
        '            var targetThread = GetWindowThreadProcessId(hWnd,     out _);'
        '            var myThread     = GetCurrentThreadId();'
        ''
        '            bool attachedFore   = false;'
        '            bool attachedTarget = false;'
        ''
        '            try'
        '            {'
        '                // Attach to the foreground thread so Windows treats us as'
        '                // eligible to change foreground.'
        '                if (foreThread != 0 && foreThread != myThread)'
        '                    attachedFore = AttachThreadInput(myThread, foreThread, true);'
        ''
        '                // Attach to target thread too so its message queue cleanly'
        '                // accepts the activation.'
        '                if (targetThread != 0 && targetThread != myThread && targetThread != foreThread)'
        '                    attachedTarget = AttachThreadInput(myThread, targetThread, true);'
        ''
        '                // Synthetic Alt-key, kept as belt-and-suspenders for the'
        '                // older "last input thread" eligibility heuristic.'
        '                keybd_event(VK_MENU, 0, 0,               UIntPtr.Zero);'
        '                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);'
        ''
        '                BringWindowToTop(hWnd);'
        '                ShowWindow(hWnd, SW_SHOW);'
        '                SetForegroundWindow(hWnd);'
        ''
        '                // Final escalation: undocumented but widely used. Windows'
        '                // treats this like a user-initiated Alt+Tab, bypassing the'
        '                // remaining foreground restrictions.'
        '                SwitchToThisWindow(hWnd, true);'
        '            }'
        '            finally'
        '            {'
        '                if (attachedFore)   AttachThreadInput(myThread, foreThread,   false);'
        '                if (attachedTarget) AttachThreadInput(myThread, targetThread, false);'
        '            }'
        '        }'
    )
    $newMethod = $newMethodLines -join "`r`n"

    if ($content.Contains($oldMethod)) {
        $content = $content.Replace($oldMethod, $newMethod)
        Write-Ok "replaced ForceForeground method with aggressive version"
    } else {
        Write-Bad "could not find v2.5.1 ForceForeground method block"
        Write-Info "ExplorerHelper.cs may have been hand-edited - inspect manually."
        exit 1
    }

    Save-FileContent $explorerHelper $content "ExplorerHelper.cs"
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
    Write-Info  "Re-run without -DryRun to apply."
} else {
    Write-Ok "All patches applied. Source is now at $ToVersion."
    Write-Host ""
    Write-Info "Next steps:"
    Write-Info "  1. Build & deploy:  .\Rebuild.ps1"
    Write-Info "  2. Verify:          .\Check-HelperVersion.ps1"
    Write-Info "  3. Test from FF:    click the Open Folder button - Explorer should"
    Write-Info "                      come fully to the front, not flash red."
}
Write-Host ""
