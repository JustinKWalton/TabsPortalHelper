<#
.SYNOPSIS
    Apply the v2.2.0 patch: retry-loop dialog for Bluebeam column install.

.DESCRIPTION
    1. Copies updated .cs files (ColumnInstallDialog.cs, Installer.cs, TrayApp.cs)
       from $SourceDir into the project.
    2. Bumps version strings in HttpServer.cs (2.1.0 -> 2.2.0) and .csproj
       (<Version>2.1.0</Version> -> <Version>2.2.0</Version>).
    3. Removes template.bin.bak if present (git rm if tracked).
    4. Adds *.bak to .gitignore if missing.

    Idempotent: safe to run multiple times. Use -DryRun to preview.
    Does NOT build or commit anything — run .\Rebuild.ps1 afterwards,
    smoke test, then run .\Publish-Release-v2.2.0.ps1.

.PARAMETER SourceDir
    Folder holding the three updated .cs files. Default: Downloads.

.PARAMETER RepoDir
    Root of the TabsPortalHelper git repo. Default: C:\TabsPortalHelper

.PARAMETER DryRun
    Show what would happen, make no changes.

.EXAMPLE
    .\Apply-Patch-v2.2.0.ps1 -DryRun

.EXAMPLE
    .\Apply-Patch-v2.2.0.ps1

.EXAMPLE
    .\Apply-Patch-v2.2.0.ps1 -SourceDir "C:\Users\justi\Downloads\tabs-patch"
#>
[CmdletBinding()]
param(
    [string]$SourceDir = "$env:USERPROFILE\Downloads",
    [string]$RepoDir   = "C:\TabsPortalHelper",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# ─── Output helpers ──────────────────────────────────────────────────────────
function Write-Step { param($m) Write-Host "►  $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "✓  $m" -ForegroundColor Green }
function Write-Skip { param($m) Write-Host "—  $m" -ForegroundColor DarkGray }
function Write-Warn2{ param($m) Write-Host "⚠  $m" -ForegroundColor Yellow }
function Write-Fail { param($m) Write-Host "✗  $m" -ForegroundColor Red }
function Write-Hdr  {
    param($m)
    Write-Host ""
    Write-Host ("═" * 72) -ForegroundColor DarkCyan
    Write-Host "  $m" -ForegroundColor White
    Write-Host ("═" * 72) -ForegroundColor DarkCyan
}

Write-Hdr "TabsPortalHelper — Apply v2.2.0 Patch"
if ($DryRun) { Write-Warn2 "DRY RUN — no changes will be made." }
Write-Host "SourceDir: $SourceDir"
Write-Host "RepoDir:   $RepoDir"
Write-Host ""

# ─── 1. Validate paths ───────────────────────────────────────────────────────
Write-Step "Validating paths..."
$projectDir = Join-Path $RepoDir "TabsPortalHelper"

if (-not (Test-Path $RepoDir))    { Write-Fail "Repo not found: $RepoDir";    exit 1 }
if (-not (Test-Path $projectDir)) { Write-Fail "Project dir not found: $projectDir"; exit 1 }

$expectedFiles = @("ColumnInstallDialog.cs", "Installer.cs", "TrayApp.cs")
$missing = @()
foreach ($f in $expectedFiles) {
    if (-not (Test-Path (Join-Path $SourceDir $f))) { $missing += $f }
}
if ($missing.Count -gt 0) {
    Write-Fail "Missing source file(s) in ${SourceDir}: $($missing -join ', ')"
    Write-Host "  Download them from the chat and place them in $SourceDir, or pass -SourceDir explicitly."
    exit 1
}
Write-Ok "Paths look good."

# ─── 2. Copy .cs files into project ──────────────────────────────────────────
Write-Step "Copying .cs files into project..."
foreach ($f in $expectedFiles) {
    $src = Join-Path $SourceDir $f
    $dst = Join-Path $projectDir $f
    $exists = Test-Path $dst

    if ($exists) {
        $srcHash = (Get-FileHash $src).Hash
        $dstHash = (Get-FileHash $dst).Hash
        if ($srcHash -eq $dstHash) {
            Write-Skip "$f already up to date"
            continue
        }
    }

    if ($DryRun) {
        Write-Warn2 "would copy $f -> $dst"
    } else {
        Copy-Item -Force $src $dst
        if ($exists) { Write-Ok "$f replaced" } else { Write-Ok "$f added" }
    }
}

# ─── 3. Bump HttpServer.cs version ───────────────────────────────────────────
Write-Step "Bumping HttpServer.cs version..."
$httpServer = Join-Path $projectDir "HttpServer.cs"
if (-not (Test-Path $httpServer)) {
    Write-Warn2 "HttpServer.cs not found — bump its version manually."
} else {
    $content = [System.IO.File]::ReadAllText($httpServer)
    if ($content -match '"2\.2\.0"') {
        Write-Skip "HttpServer.cs already at 2.2.0"
    } elseif ($content -match '"2\.1\.0"') {
        if ($DryRun) {
            Write-Warn2 "would replace 2.1.0 -> 2.2.0 in HttpServer.cs"
        } else {
            $content = $content -replace '"2\.1\.0"', '"2.2.0"'
            [System.IO.File]::WriteAllText($httpServer, $content)
            Write-Ok "HttpServer.cs bumped to 2.2.0"
        }
    } else {
        Write-Warn2 "No 2.1.0 pattern found in HttpServer.cs — bump manually."
    }
}

# ─── 4. Bump .csproj version ─────────────────────────────────────────────────
Write-Step "Bumping TabsPortalHelper.csproj version..."
$csproj = Join-Path $projectDir "TabsPortalHelper.csproj"
if (-not (Test-Path $csproj)) {
    Write-Warn2 ".csproj not found — bump its <Version> manually."
} else {
    $content = [System.IO.File]::ReadAllText($csproj)
    if ($content -match '<Version>\s*2\.2\.0\s*</Version>') {
        Write-Skip ".csproj already at 2.2.0"
    } elseif ($content -match '<Version>\s*2\.1\.0\s*</Version>') {
        if ($DryRun) {
            Write-Warn2 "would replace <Version>2.1.0</Version> -> <Version>2.2.0</Version>"
        } else {
            $content = $content -replace '<Version>\s*2\.1\.0\s*</Version>', '<Version>2.2.0</Version>'
            [System.IO.File]::WriteAllText($csproj, $content)
            Write-Ok ".csproj bumped to 2.2.0"
        }
    } else {
        Write-Warn2 "No <Version>2.1.0</Version> tag in .csproj — bump manually."
    }
}

# ─── 5. Remove template.bin.bak ──────────────────────────────────────────────
Write-Step "Cleaning up template.bin.bak..."
$bakPath = Join-Path $projectDir "template.bin.bak"
if (-not (Test-Path $bakPath)) {
    Write-Skip "template.bin.bak not present"
} else {
    Push-Location $RepoDir
    try {
        $null = & git ls-files --error-unmatch "TabsPortalHelper/template.bin.bak" 2>&1
        $isTracked = ($LASTEXITCODE -eq 0)
        $global:LASTEXITCODE = 0

        if ($DryRun) {
            if ($isTracked) {
                Write-Warn2 "would 'git rm' template.bin.bak (tracked)"
            } else {
                Write-Warn2 "would delete untracked template.bin.bak"
            }
        } else {
            if ($isTracked) {
                & git rm "TabsPortalHelper/template.bin.bak" | Out-Null
                Write-Ok "template.bin.bak removed (git rm)"
            } else {
                Remove-Item $bakPath -Force
                Write-Ok "template.bin.bak removed (untracked)"
            }
        }
    } finally {
        Pop-Location
    }
}

# ─── 6. Add *.bak to .gitignore ──────────────────────────────────────────────
Write-Step "Ensuring *.bak is in .gitignore..."
$gitignorePath = Join-Path $RepoDir ".gitignore"
if (-not (Test-Path $gitignorePath)) {
    Write-Warn2 ".gitignore not found — skipping."
} else {
    $gitignore = [System.IO.File]::ReadAllText($gitignorePath)
    # Regex: *.bak as its own pattern (not part of a larger negation/exclude)
    if ($gitignore -match '(?m)^\s*\*\.bak\s*$') {
        Write-Skip "*.bak already in .gitignore"
    } else {
        if ($DryRun) {
            Write-Warn2 "would append '*.bak' to .gitignore"
        } else {
            $append = ""
            if (-not $gitignore.EndsWith("`n")) { $append += "`n" }
            $append += "`n# Backup files`n*.bak`n"
            [System.IO.File]::AppendAllText($gitignorePath, $append)
            Write-Ok "*.bak appended to .gitignore"
        }
    }
}

# ─── 7. Show summary + next steps ────────────────────────────────────────────
Write-Hdr "Patch Applied"

if ($DryRun) {
    Write-Host "This was a dry run. Re-run without -DryRun to apply." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Verify with:" -ForegroundColor White
Write-Host "  cd $RepoDir" -ForegroundColor Gray
Write-Host "  git status" -ForegroundColor Gray
Write-Host "  git diff --stat" -ForegroundColor Gray
Write-Host ""
Write-Host "Next: build the new .exe" -ForegroundColor White
Write-Host "  cd $RepoDir" -ForegroundColor Gray
Write-Host "  .\Rebuild.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "Then smoke test:" -ForegroundColor White
Write-Host "  1. Open Bluebeam Revu (at least one window)" -ForegroundColor Gray
Write-Host "  2. Run the freshly-built .exe (from bin\Release\...\publish\, or the one in Downloads)" -ForegroundColor Gray
Write-Host "  3. Install prompts -> dialog appears with install-success text + Bluebeam-running warning" -ForegroundColor Gray
Write-Host "  4. Verify buttons say [Retry] [Skip for now]" -ForegroundColor Gray
Write-Host "  5. Click Retry with Revu still open -> still-running message stays" -ForegroundColor Gray
Write-Host "  6. Close all Revu windows, click Retry -> 'Installing…' then success state with [OK]" -ForegroundColor Gray
Write-Host "  7. Open Revu, verify TABS columns are present in the Markups List column menu" -ForegroundColor Gray
Write-Host "  8. Right-click tray -> Install Bluebeam Columns... -> verify same dialog (smaller size, no preamble)" -ForegroundColor Gray
Write-Host ""
Write-Host "Finally, publish:" -ForegroundColor White
Write-Host "  .\Publish-Release-v2.2.0.ps1" -ForegroundColor Gray
Write-Host ""
