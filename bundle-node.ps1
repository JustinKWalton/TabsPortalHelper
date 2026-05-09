# bundle-node.ps1
# Build script: download portable Node.js + place cropper.js + zip it all up
# into a single cropper-bundle.zip that gets embedded in TabsPortalHelper.exe
# as a resource. Run this BEFORE building the helper.
#
# Output:
#   installer-payload\node\node.exe
#   installer-payload\cropper\cropper.js + node_modules
#   installer-payload\cropper-bundle.zip   <- this is what gets embedded
#
# The bundle zip is what the helper extracts to %LOCALAPPDATA%\TabsPortalHelper\
# on first run. CropperBundle.cs handles extraction.
#
# Tested with PowerShell 5.1+ on Windows 11.

param(
    # Where this script stages the bundle. The zip file ends up here too.
    [string]$InstallerPayloadDir = ".\installer-payload",

    # Path to cropper.js in your repo. Default assumes you've added it under
    # tools\cropper\ in the helper repo.
    [string]$CropperSourceDir = ".\tools\cropper",

    # Node.js version to bundle. Pin a specific LTS so the build is reproducible.
    [string]$NodeVersion = "20.18.0",

    # x64 only - TabsPortalHelper is Windows-only, no ARM yet.
    [string]$NodeArch = "x64"
)

$ErrorActionPreference = "Stop"

# ------------------------------------------------------------------
# Resolve all paths to absolute up-front. PowerShell's Push-Location and
# the .NET process's [Environment]::CurrentDirectory can disagree, so any
# relative path passed to a .NET API (like ZipFile.CreateFromDirectory)
# will resolve against the .NET cwd (likely the user's home dir) instead
# of the PowerShell location.
# ------------------------------------------------------------------
if (-not (Test-Path $InstallerPayloadDir)) {
    New-Item -ItemType Directory -Path $InstallerPayloadDir | Out-Null
}
$InstallerPayloadDir = (Resolve-Path $InstallerPayloadDir).Path

if (-not (Test-Path $CropperSourceDir)) {
    throw "CropperSourceDir not found: $CropperSourceDir"
}
$CropperSourceDir = (Resolve-Path $CropperSourceDir).Path

# Belt-and-suspenders: also sync .NET's idea of cwd to PowerShell's, so
# any other relative-path .NET call inside this script behaves as expected.
[Environment]::CurrentDirectory = (Get-Location).Path

Write-Host "Bundling Node.js $NodeVersion-$NodeArch + cropper into $InstallerPayloadDir" -ForegroundColor Cyan
Write-Host "  CropperSourceDir:    $CropperSourceDir"
Write-Host "  InstallerPayloadDir: $InstallerPayloadDir"

# ------------------------------------------------------------------
# 1. Download portable Node.js
# ------------------------------------------------------------------
$nodeZipUrl = "https://nodejs.org/dist/v$NodeVersion/node-v$NodeVersion-win-$NodeArch.zip"
$nodeZipPath = Join-Path $env:TEMP "node-v$NodeVersion-win-$NodeArch.zip"
$nodeExtractRoot = Join-Path $env:TEMP "node-extract-$NodeVersion"

if (-not (Test-Path $nodeZipPath)) {
    Write-Host "Downloading $nodeZipUrl..."
    Invoke-WebRequest -Uri $nodeZipUrl -OutFile $nodeZipPath
}

if (Test-Path $nodeExtractRoot) {
    Remove-Item $nodeExtractRoot -Recurse -Force
}
Write-Host "Extracting Node..."
Expand-Archive -Path $nodeZipPath -DestinationPath $nodeExtractRoot

$extractedNodeDir = Get-ChildItem -Path $nodeExtractRoot -Directory | Select-Object -First 1

# ------------------------------------------------------------------
# 2. Stage just node.exe (no npm or other tooling at runtime)
# ------------------------------------------------------------------
$nodeOutDir = Join-Path $InstallerPayloadDir "node"
if (Test-Path $nodeOutDir) {
    Remove-Item $nodeOutDir -Recurse -Force
}
New-Item -ItemType Directory -Path $nodeOutDir | Out-Null
Copy-Item -Path (Join-Path $extractedNodeDir.FullName "node.exe") -Destination $nodeOutDir

Write-Host "Staged node.exe at $nodeOutDir\node.exe"

# ------------------------------------------------------------------
# 3. Install cropper dependencies and stage the script + node_modules
# ------------------------------------------------------------------
$cropperOutDir = Join-Path $InstallerPayloadDir "cropper"

# Copy pre-installed cropper tree from source. node_modules MUST exist in
# tools\cropper\ before running this script (one-time setup):
#
#   cd tools\cropper
#   npm.cmd install --omit=dev --no-audit --no-fund
#
# Why not run npm install here: the portable Node zip we download is
# node-only (no npm bundled), and PowerShells & npm install against the
# system npm.ps1 mangles the first argument (install -> pm). Doing the
# install once into the source tree side-steps both problems.
$sourceModules = Join-Path $CropperSourceDir "node_modules"
if (-not (Test-Path $sourceModules)) {
    throw "tools\cropper\node_modules is missing. Run once before bundling: cd $CropperSourceDir; npm.cmd install --omit=dev --no-audit --no-fund"
}

if (Test-Path $cropperOutDir) {
    Remove-Item $cropperOutDir -Recurse -Force
}
New-Item -ItemType Directory -Path $cropperOutDir | Out-Null

Copy-Item -Path (Join-Path $CropperSourceDir "cropper.js")   -Destination $cropperOutDir
Copy-Item -Path (Join-Path $CropperSourceDir "package.json") -Destination $cropperOutDir
if (Test-Path (Join-Path $CropperSourceDir "package-lock.json")) {
    Copy-Item -Path (Join-Path $CropperSourceDir "package-lock.json") -Destination $cropperOutDir
}
Copy-Item -Path $sourceModules -Destination $cropperOutDir -Recurse

Write-Host "Staged cropper at $cropperOutDir"

# ------------------------------------------------------------------
# 4. Build cropper-bundle.zip - this is what gets embedded in the .exe
# ------------------------------------------------------------------
$bundleZipPath = Join-Path $InstallerPayloadDir "cropper-bundle.zip"
if (Test-Path $bundleZipPath) {
    Remove-Item $bundleZipPath -Force
}

Write-Host ""
Write-Host "Building cropper-bundle.zip..." -ForegroundColor Cyan

# Use the .NET ZipFile API directly - PowerShell 5.1's Compress-Archive
# can choke or be very slow on the 1,700+ files inside node_modules.
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zipStagingDir = Join-Path $env:TEMP "tabs-cropper-bundle-staging-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $zipStagingDir | Out-Null
try {
    Copy-Item -Path $nodeOutDir    -Destination (Join-Path $zipStagingDir "node")    -Recurse
    Copy-Item -Path $cropperOutDir -Destination (Join-Path $zipStagingDir "cropper") -Recurse

    # Both args are absolute paths now - safe regardless of .NET cwd.
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $zipStagingDir,
        $bundleZipPath,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false  # don't include base directory
    )
}
finally {
    Remove-Item -Path $zipStagingDir -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path $bundleZipPath)) {
    throw "Failed to create $bundleZipPath"
}

$zipSizeMB = [math]::Round((Get-Item $bundleZipPath).Length / 1MB, 1)
Write-Host "Bundle zip: $bundleZipPath ($zipSizeMB MB)" -ForegroundColor Green

# ------------------------------------------------------------------
# 5. Summary
# ------------------------------------------------------------------
Write-Host ""
Write-Host "Installer payload contents:" -ForegroundColor Cyan
Get-ChildItem -Path $InstallerPayloadDir -Recurse -File |
    Measure-Object -Property Length -Sum |
    ForEach-Object {
        $sizeMB = [math]::Round($_.Sum / 1MB, 1)
        Write-Host "  Total: $($_.Count) files, $sizeMB MB"
    }
Write-Host ""
Write-Host "Next: rebuild TabsPortalHelper. The .csproj EmbeddedResource entry"
Write-Host "for cropper-bundle.zip means it gets baked into the single-file .exe."
