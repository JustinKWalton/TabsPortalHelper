<#
.SYNOPSIS
    Publish the v2.2.0 release: git commit + tag + push, and (optionally) create
    the GitHub release with the built .exe attached via the gh CLI.

.DESCRIPTION
    Run this AFTER Apply-Patch-v2.2.0.ps1 + .\Rebuild.ps1 + manual smoke test.

    Steps:
    1. Sanity-checks the built .exe exists and is newer than source files.
    2. Shows git status, prompts for confirmation.
    3. git add . && commit && tag v2.2.0 && push && push --tags
    4. If `gh` CLI is installed: creates the release and uploads the .exe.
       Otherwise, prints the URL and file path for manual upload.

.PARAMETER RepoDir
    Root of the TabsPortalHelper git repo. Default: C:\TabsPortalHelper

.PARAMETER SkipFreshness
    Skip the "is the .exe newer than source files?" check. Use if you know
    what you're doing and the check is being annoying.

.PARAMETER Force
    Skip the interactive confirmation prompt (fire-and-forget publish).

.EXAMPLE
    .\Publish-Release-v2.2.0.ps1

.EXAMPLE
    .\Publish-Release-v2.2.0.ps1 -Force
#>
[CmdletBinding()]
param(
    [string]$RepoDir = "C:\TabsPortalHelper",
    [switch]$SkipFreshness,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($m) Write-Host "►  $m" -ForegroundColor Cyan }
function Write-Ok   { param($m) Write-Host "✓  $m" -ForegroundColor Green }
function Write-Skip { param($m) Write-Host "—  $m" -ForegroundColor DarkGray }
function Write-Warn2{ param($m) Write-Host "⚠  $m" -ForegroundColor Yellow }
function Write-Fail { param($m) Write-Host "✗  $m" -ForegroundColor Red }
function Write-Hdr {
    param($m)
    Write-Host ""
    Write-Host ("═" * 72) -ForegroundColor DarkCyan
    Write-Host "  $m" -ForegroundColor White
    Write-Host ("═" * 72) -ForegroundColor DarkCyan
}

$Tag     = "v2.2.0"
$Version = "2.2.0"
$Commit  = "v${Version}: retry-loop for column install when Bluebeam is running"

Write-Hdr "TabsPortalHelper — Publish Release $Tag"

# ─── 1. Validate repo ────────────────────────────────────────────────────────
if (-not (Test-Path $RepoDir)) { Write-Fail "Repo not found: $RepoDir"; exit 1 }
$projectDir = Join-Path $RepoDir "TabsPortalHelper"
$exePath    = Join-Path $projectDir "bin\Release\net9.0-windows\win-x64\publish\TabsPortalHelper.exe"

Push-Location $RepoDir

try {
    # Confirm this is a git repo
    $null = & git rev-parse --is-inside-work-tree 2>&1
    if ($LASTEXITCODE -ne 0) { Write-Fail "$RepoDir is not a git repo."; exit 1 }
    $global:LASTEXITCODE = 0

    # ─── 2. Check .exe exists and is fresh ────────────────────────────────────
    Write-Step "Checking built .exe..."
    if (-not (Test-Path $exePath)) {
        Write-Fail "Built .exe not found: $exePath"
        Write-Host "  Run .\Rebuild.ps1 in the repo root first."
        exit 1
    }
    $exeInfo = Get-Item $exePath
    Write-Ok "Found: $exePath ($([math]::Round($exeInfo.Length / 1MB, 1)) MB, built $($exeInfo.LastWriteTime))"

    if (-not $SkipFreshness) {
        $sources = Get-ChildItem "$projectDir\*.cs", "$projectDir\*.csproj" -ErrorAction SilentlyContinue
        $newestSrc = ($sources | Measure-Object LastWriteTime -Maximum).Maximum
        if ($exeInfo.LastWriteTime -lt $newestSrc) {
            Write-Warn2 "Built .exe is older than source files!"
            Write-Host "    Newest source: $newestSrc"
            Write-Host "    .exe built:    $($exeInfo.LastWriteTime)"
            Write-Host "  Run .\Rebuild.ps1 first, or pass -SkipFreshness to override."
            exit 1
        }
        Write-Ok ".exe is newer than source files"
    }

    # ─── 3. Check tag doesn't already exist ───────────────────────────────────
    Write-Step "Checking tag $Tag..."
    $existingTag = & git tag --list $Tag
    if ($existingTag) {
        Write-Fail "Tag $Tag already exists locally."
        Write-Host "  If a previous run stopped partway: delete the tag with 'git tag -d $Tag'"
        Write-Host "  and (if it was pushed) 'git push origin :refs/tags/$Tag', then re-run."
        exit 1
    }
    Write-Ok "Tag $Tag is free"

    # ─── 4. Show git status ───────────────────────────────────────────────────
    Write-Step "Current git status:"
    & git status --short
    Write-Host ""

    if (-not $Force) {
        $confirm = Read-Host "Commit the above as '$Commit' and push $Tag? (yes/no)"
        if ($confirm -ne 'yes') {
            Write-Warn2 "Aborted by user."
            exit 0
        }
    }

    # ─── 5. Commit + tag + push ──────────────────────────────────────────────
    Write-Step "git add ."
    & git add .
    if ($LASTEXITCODE -ne 0) { Write-Fail "git add failed."; exit 1 }

    # Only commit if there's something to commit (idempotent re-run after partial failure)
    $porcelain = & git status --porcelain
    if ([string]::IsNullOrWhiteSpace($porcelain)) {
        Write-Skip "Nothing to commit — working tree clean"
    } else {
        Write-Step "git commit"
        & git commit -m $Commit
        if ($LASTEXITCODE -ne 0) { Write-Fail "git commit failed."; exit 1 }
        Write-Ok "Committed"
    }

    Write-Step "git tag $Tag"
    & git tag $Tag
    if ($LASTEXITCODE -ne 0) { Write-Fail "git tag failed."; exit 1 }
    Write-Ok "Tagged $Tag"

    Write-Step "git push"
    & git push
    if ($LASTEXITCODE -ne 0) { Write-Fail "git push failed."; exit 1 }

    Write-Step "git push --tags"
    & git push --tags
    if ($LASTEXITCODE -ne 0) { Write-Fail "git push --tags failed."; exit 1 }
    Write-Ok "Pushed commit + tag to origin"

    # ─── 6. GitHub release (gh CLI if available) ─────────────────────────────
    Write-Hdr "GitHub Release"

    $ghAvailable = $null -ne (Get-Command gh -ErrorAction SilentlyContinue)

    $notes = @"
## What's new in v$Version

**Retry dialog for Bluebeam column install.** When the installer or tray menu tries to install TABS column presets while Bluebeam Revu is running, it now shows a dialog with **Retry** and **Skip for now** buttons instead of warning and moving on. Close Bluebeam, click Retry, and the dialog rolls through to a success state in place.

### Also

- Version strings aligned across Installer.cs, TrayApp.cs, HttpServer.cs, and .csproj (was drifted at 2.0.0 / 2.1.0 / 2.1.0).
- Removed tracked ``template.bin.bak``; ``*.bak`` added to ``.gitignore``.

### Install

Download ``TabsPortalHelper.exe`` below and run it. If SmartScreen warns, click **More info** → **Run anyway**. The helper installs itself to ``%LOCALAPPDATA%\TabsPortalHelper\`` and starts automatically with Windows.
"@

    if ($ghAvailable) {
        Write-Step "Creating release via gh CLI..."
        $notesFile = New-TemporaryFile
        try {
            [System.IO.File]::WriteAllText($notesFile.FullName, $notes)
            & gh release create $Tag $exePath `
                --title "v$Version — Retry dialog for Bluebeam column install" `
                --notes-file $notesFile.FullName
            if ($LASTEXITCODE -ne 0) {
                Write-Fail "gh release create failed."
                Write-Host "  You can still publish manually — see instructions below."
                $ghAvailable = $false
            } else {
                Write-Ok "Release published"
                Write-Host ""
                Write-Host "  https://github.com/JustinKWalton/TabsPortalHelper/releases/tag/$Tag" -ForegroundColor White
            }
        } finally {
            Remove-Item $notesFile.FullName -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not $ghAvailable) {
        Write-Warn2 "gh CLI not available — publish the release manually:"
        Write-Host ""
        Write-Host "  1. Open: https://github.com/JustinKWalton/TabsPortalHelper/releases/new?tag=$Tag" -ForegroundColor White
        Write-Host "  2. Title: v$Version — Retry dialog for Bluebeam column install" -ForegroundColor Gray
        Write-Host "  3. Drag in the .exe from:" -ForegroundColor Gray
        Write-Host "       $exePath" -ForegroundColor Gray
        Write-Host "  4. Paste these release notes:" -ForegroundColor Gray
        Write-Host ""
        Write-Host $notes -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "  (Tip: install gh CLI once — ``winget install GitHub.cli`` — and this step becomes automatic.)" -ForegroundColor DarkGray
    }

    Write-Hdr "Done"
    Write-Host "The PWA's ``checkTabsHelper`` action will start serving v$Version to 'outdated' clients" -ForegroundColor White
    Write-Host "as soon as the release asset finishes uploading — no FlutterFlow change needed." -ForegroundColor White
    Write-Host ""
} finally {
    Pop-Location
}
