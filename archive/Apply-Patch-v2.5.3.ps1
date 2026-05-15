# =============================================================================
#  Apply-Patch-v2.5.3.ps1
#
#  Upgrades TabsPortalHelper source from v2.5.2 to v2.5.3:
#    - Wraps ExplorerHelper.ForceForeground in a SystemParametersInfo call
#      that temporarily zeroes the ForegroundLockTimeout system setting,
#      then restores it. This is the canonical fix used by AutoHotkey,
#      PowerToys, etc. for "background process needs to steal focus from
#      a busy foreground app (e.g. browser)."
#
#  Idempotent: safe to re-run.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.5.3.ps1
#      .\Apply-Patch-v2.5.3.ps1 -DryRun
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$FromVersion = '2.5.2'
$ToVersion   = '2.5.3'

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

# --- 3. Patch ExplorerHelper.cs ---------------------------------------------
Write-Step "Patching ExplorerHelper.cs (add ForegroundLockTimeout zeroing)"
$content = Get-Content $explorerHelper -Raw

if ($content -match 'SPI_SETFOREGROUNDLOCKTIMEOUT') {
    Write-Skip "ExplorerHelper.cs already has SPI_SETFOREGROUNDLOCKTIMEOUT (v2.5.3 behavior)"
} else {
    if ($content -notmatch 'AttachThreadInput') {
        Write-Bad "ExplorerHelper.cs doesn't look like the v2.5.2 version - inspect manually"
        exit 1
    }

    # 3a. Add SystemParametersInfo P/Invoke + constants after GetCurrentThreadId line
    $oldAnchor = '        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();'
    $newAnchorLines = @(
        '        [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();'
        '        [DllImport("user32.dll", SetLastError = true)] static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref uint pvParam, uint fWinIni);'
        '        [DllImport("user32.dll", SetLastError = true)] static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);'
        ''
        '        const uint SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;'
        '        const uint SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;'
    )
    $newAnchor = $newAnchorLines -join "`r`n"

    if ($content.Contains($oldAnchor)) {
        $content = $content.Replace($oldAnchor, $newAnchor)
        Write-Ok "added SystemParametersInfo P/Invokes and constants"
    } else {
        Write-Bad "could not find GetCurrentThreadId anchor"
        exit 1
    }

    # 3b. Replace ForceForeground method with ForegroundLockTimeout-wrapped version
    $oldMethodLines = @(
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
    $oldMethod = $oldMethodLines -join "`r`n"

    $newMethodLines = @(
        '        // Steals foreground reliably even when called from a busy browser.'
        '        // The key trick: zero out ForegroundLockTimeout (the system-wide'
        '        // setting that powers focus-stealing prevention) for the duration'
        '        // of the activation, then restore it. This is the canonical fix'
        '        // used by AutoHotkey, PowerToys, Voidtools Everything, etc.'
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
        '            uint origTimeout    = 0;'
        '            bool gotOrigTimeout = false;'
        ''
        '            try'
        '            {'
        '                // 1. Save and zero the system foreground-lock timeout.'
        '                gotOrigTimeout = SystemParametersInfo(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref origTimeout, 0);'
        '                SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, IntPtr.Zero, 0);'
        ''
        '                // 2. Attach our input queue to the foreground thread so we'
        '                //    inherit its foreground-change rights.'
        '                if (foreThread != 0 && foreThread != myThread)'
        '                    attachedFore = AttachThreadInput(myThread, foreThread, true);'
        '                if (targetThread != 0 && targetThread != myThread && targetThread != foreThread)'
        '                    attachedTarget = AttachThreadInput(myThread, targetThread, true);'
        ''
        '                // 3. Synthetic Alt-key (last-input-thread heuristic).'
        '                keybd_event(VK_MENU, 0, 0,               UIntPtr.Zero);'
        '                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);'
        ''
        '                // 4. Activate: BringToTop -> Show -> SetForeground -> SwitchTo.'
        '                BringWindowToTop(hWnd);'
        '                ShowWindow(hWnd, SW_SHOW);'
        '                SetForegroundWindow(hWnd);'
        '                SwitchToThisWindow(hWnd, true);'
        '            }'
        '            finally'
        '            {'
        '                // Restore everything we changed - never leave the system in a'
        '                // permanently-unprotected state.'
        '                if (gotOrigTimeout)'
        '                    SystemParametersInfo(SPI_SETFOREGROUNDLOCKTIMEOUT, origTimeout, IntPtr.Zero, 0);'
        '                if (attachedFore)   AttachThreadInput(myThread, foreThread,   false);'
        '                if (attachedTarget) AttachThreadInput(myThread, targetThread, false);'
        '            }'
        '        }'
    )
    $newMethod = $newMethodLines -join "`r`n"

    if ($content.Contains($oldMethod)) {
        $content = $content.Replace($oldMethod, $newMethod)
        Write-Ok "replaced ForceForeground with ForegroundLockTimeout-wrapped version"
    } else {
        Write-Bad "could not find v2.5.2 ForceForeground method block"
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
} else {
    Write-Ok "All patches applied. Source is now at $ToVersion."
    Write-Host ""
    Write-Info "Next steps:"
    Write-Info "  1. .\Rebuild.ps1"
    Write-Info "  2. Click the Open Folder button in FF - Explorer should now"
    Write-Info "     come fully to the front."
    Write-Host ""
    Write-Info "If it STILL flashes red, this is a Windows behavior we can't fully"
    Write-Info "fight. Last-resort options I can implement:"
    Write-Info "  - Have the helper minimize all other windows then activate Explorer"
    Write-Info "    (works 100% but is jarring UX)"
    Write-Info "  - Set HWND_TOPMOST briefly so the window visually appears on top"
    Write-Info "    even without keyboard focus transfer"
    Write-Info "  - Switch from explorer.exe to opening the folder via Shell.Application"
    Write-Info "    COM, which has different foreground semantics"
}
Write-Host ""
