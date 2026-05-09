# =============================================================================
#  Apply-Patch-v2.7.0.ps1
#
#  Upgrades TabsPortalHelper from any 2.5.x or 2.6.x to 2.7.0 (step 1 of
#  the v2.7.0 feature work -- Bluebeam custom columns scrubber).
#
#  Step 1 changes:
#    - Bumps version constants in csproj, TrayApp.cs, Installer.cs,
#      HttpServer.cs to 2.7.0
#    - Adds PdfSharp 6.1.1 NuGet package reference
#    - Adds Custom_Columns.xml as embedded resource
#    - Creates TabsPortalHelper/Custom_Columns.xml
#    - Creates TabsPortalHelper/BsiAnnotColumnsScrubber.cs
#
#  After this patch, run:
#    dotnet restore
#    dotnet build
#  ...to confirm compilation. The scrubber is then callable but not yet
#  wired into any tray menu / HTTP endpoint -- that comes in step 2.
#
#  Idempotent: safe to re-run. Aborts cleanly if source is at an
#  unexpected state.
#
#  Usage from repo root:
#      .\Apply-Patch-v2.7.0.ps1
#      .\Apply-Patch-v2.7.0.ps1 -DryRun
#      .\Apply-Patch-v2.7.0.ps1 -ProjectDir 'C:\path\to\repo'
# =============================================================================

[CmdletBinding()]
param(
    [string]$ProjectDir = (Get-Location).Path,
    [switch]$DryRun
)

$ToVersion       = '2.7.0'
$PdfSharpVersion = '6.1.1'

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
        # UTF-8 without BOM, matching the predominant style in the repo.
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
            return $false
        } else {
            Write-Warn2 "$Label exists but differs from bundled content"
            Write-Info "(leaving on-disk version untouched -- review manually if needed)"
            return $false
        }
    }
    Save-FileContent -Path $Path -Content $Content -Label $Label
    return $true
}

# --- 1. Locate source files --------------------------------------------------
$csproj     = Join-Path $ProjectDir "TabsPortalHelper\TabsPortalHelper.csproj"
$httpServer = Join-Path $ProjectDir "TabsPortalHelper\HttpServer.cs"
$trayApp    = Join-Path $ProjectDir "TabsPortalHelper\TrayApp.cs"
$installer  = Join-Path $ProjectDir "TabsPortalHelper\Installer.cs"
$projectDirInner = Join-Path $ProjectDir "TabsPortalHelper"

if (-not (Test-Path $csproj)) {
    $alt = Join-Path $ProjectDir "TabsPortalHelper.csproj"
    if (Test-Path $alt) {
        $csproj          = $alt
        $httpServer      = Join-Path $ProjectDir "HttpServer.cs"
        $trayApp         = Join-Path $ProjectDir "TrayApp.cs"
        $installer       = Join-Path $ProjectDir "Installer.cs"
        $projectDirInner = $ProjectDir
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
    Write-Bad "Source files not found. Pass -ProjectDir if running from outside the repo root."
    exit 1
}

# --- 2. Patch .csproj --------------------------------------------------------
Write-Step "Patching TabsPortalHelper.csproj"
$content = Get-Content $csproj -Raw
$changed = $false

# 2a. Bump <Version>
if ($content -match "<Version>\s*$([regex]::Escape($ToVersion))\s*</Version>") {
    Write-Skip "<Version> already at $ToVersion"
} else {
    $newContent = [regex]::Replace($content, '<Version>\s*[0-9]+\.[0-9]+\.[0-9]+\s*</Version>', "<Version>$ToVersion</Version>")
    if ($newContent -ne $content) {
        $content = $newContent
        Write-Ok "bumped <Version> to $ToVersion"
        $changed = $true
    } else {
        Write-Bad "could not find <Version> tag in csproj"
        exit 1
    }
}

# 2b. Add PdfSharp PackageReference
if ($content -match 'PackageReference\s+Include="PdfSharp"') {
    Write-Skip "PdfSharp PackageReference already present"
} else {
    $anchor = '<PackageReference Include="Microsoft.Data.Sqlite" Version="9.0.4" />'
    $replacement = $anchor + "`r`n    <PackageReference Include=`"PdfSharp`" Version=`"$PdfSharpVersion`" />"
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $replacement)
        Write-Ok "added PdfSharp $PdfSharpVersion PackageReference"
        $changed = $true
    } else {
        Write-Bad "could not find Microsoft.Data.Sqlite anchor in csproj"
        exit 1
    }
}

# 2c. Add Custom_Columns.xml EmbeddedResource
if ($content -match '<EmbeddedResource\s+Include="Custom_Columns\.xml"') {
    Write-Skip "Custom_Columns.xml EmbeddedResource already present"
} else {
    $anchor = @"
    <EmbeddedResource Include="TABSportal.bpx">
      <LogicalName>TabsPortalHelper.TABSportal.bpx</LogicalName>
    </EmbeddedResource>
"@
    $newBlock = @"
    <EmbeddedResource Include="TABSportal.bpx">
      <LogicalName>TabsPortalHelper.TABSportal.bpx</LogicalName>
    </EmbeddedResource>
    <EmbeddedResource Include="Custom_Columns.xml">
      <LogicalName>TabsPortalHelper.Custom_Columns.xml</LogicalName>
    </EmbeddedResource>
"@
    if ($content.Contains($anchor)) {
        $content = $content.Replace($anchor, $newBlock)
        Write-Ok "added Custom_Columns.xml EmbeddedResource"
        $changed = $true
    } else {
        Write-Bad "could not find TABSportal.bpx EmbeddedResource anchor in csproj"
        exit 1
    }
}

if ($changed) { Save-FileContent -Path $csproj -Content $content -Label "TabsPortalHelper.csproj" }

# --- 3. Bump version constants in source files ------------------------------
function Update-VersionConstant {
    param([string]$Path, [string]$ConstantName)

    Write-Step "Bumping $ConstantName in $(Split-Path $Path -Leaf)"
    $c = Get-Content $Path -Raw

    # Match: const string <Name>(spaces)= "X.Y.Z";
    $pattern = 'const\s+string\s+' + [regex]::Escape($ConstantName) + '\s*=\s*"([0-9]+\.[0-9]+\.[0-9]+)"'
    $match = [regex]::Match($c, $pattern)
    if (-not $match.Success) {
        Write-Warn2 "no $ConstantName constant found -- skipping"
        return
    }

    $current = $match.Groups[1].Value
    if ($current -eq $ToVersion) {
        Write-Skip "$ConstantName already at $ToVersion"
        return
    }

    # Replace just the version literal, keeping whitespace alignment intact.
    $c2 = [regex]::Replace($c, $pattern, {
        param($m)
        $m.Value -replace [regex]::Escape($current), $ToVersion
    })

Save-FileContent -Path $Path -Content $c2 -Label "$(Split-Path $Path -Leaf) (${ConstantName}: $current -> $ToVersion)"}

Update-VersionConstant -Path $trayApp    -ConstantName "Version"
Update-VersionConstant -Path $httpServer -ConstantName "Version"
Update-VersionConstant -Path $installer  -ConstantName "AppVersion"

# --- 4. Create Custom_Columns.xml -------------------------------------------
Write-Step "Creating Custom_Columns.xml"
$customColumnsXml = @'
<?xml version="1.0" encoding="utf-8"?>
<BluebeamUserDefinedColumns>
  <BSIColumnItem Index="0" Subtype="Text">
    <Name>Note</Name>
    <DisplayOrder>4</DisplayOrder>
    <Deleted>False</Deleted>
    <Multiline>False</Multiline>
  </BSIColumnItem>
  <BSIColumnItem Index="1" Subtype="Text">
    <Name>Comment Status</Name>
    <DisplayOrder>3</DisplayOrder>
    <DefaultValue>Not Acceptable</DefaultValue>
    <Deleted>False</Deleted>
    <Multiline>False</Multiline>
  </BSIColumnItem>
  <BSIColumnItem Index="2" Subtype="Text">
    <Name>Element</Name>
    <DisplayOrder>1</DisplayOrder>
    <Deleted>False</Deleted>
    <Multiline>False</Multiline>
  </BSIColumnItem>
  <BSIColumnItem Index="3" Subtype="Text">
    <Name>Location</Name>
    <DisplayOrder>0</DisplayOrder>
    <Deleted>False</Deleted>
    <Multiline>False</Multiline>
  </BSIColumnItem>
  <BSIColumnItem Index="4" Subtype="Text">
    <Name>Issue</Name>
    <DisplayOrder>2</DisplayOrder>
    <Deleted>False</Deleted>
    <Multiline>False</Multiline>
  </BSIColumnItem>
</BluebeamUserDefinedColumns>
'@
$customColumnsPath = Join-Path $projectDirInner "Custom_Columns.xml"
[void](Save-NewFile -Path $customColumnsPath -Content $customColumnsXml -Label "Custom_Columns.xml")

# --- 5. Create BsiAnnotColumnsScrubber.cs -----------------------------------
Write-Step "Creating BsiAnnotColumnsScrubber.cs"
$scrubberCs = @'
// BsiAnnotColumnsScrubber.cs
//
// Strips the /BSIAnnotColumns key from a PDF's document Catalog.
//
// Background
// ----------
// Bluebeam Revu stores PDF-level custom column definitions in the document
// Catalog under the /BSIAnnotColumns key. Per Bluebeam's own docs, when a
// PDF carries its own column definitions they OVERRIDE the active Profile's
// columns in the Markups List. This is the root cause of the "my TABS
// columns disappeared" symptom users hit when opening PDFs marked up by
// other firms (engineering drawings from LAN, plan reviews from outside
// reviewers, corrective mods from owner's design firms, etc.).
//
// The fix
// -------
// Remove the /BSIAnnotColumns key. With it absent, Bluebeam falls back to
// the active Profile's column definitions -- which is exactly what we want
// for the TABSportal profile. No other PDF content is touched: page content
// streams, annotations, signatures, metadata, bookmarks, and form fields
// are all preserved.
//
// Used by
// -------
//   - "Bluebeam > Scrub PDF File..." tray menu item (closed-file scrub)
//   - Active-file fix orchestrator (close -> scrub -> reopen flow)
//   - /scrub-file and /scrub-active HTTP endpoints
//
// Idempotency
// -----------
// All operations are idempotent. Scrubbing a file that is already clean
// is a no-op (returns Status.AlreadyClean) and never touches the file.

using System;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace TabsPortalHelper
{
    public static class BsiAnnotColumnsScrubber
    {
        /// <summary>
        /// The PDF Catalog key that holds Bluebeam's PDF-scoped custom
        /// column definitions. Presence of this key causes Revu to ignore
        /// the active Profile's columns in the Markups List for this file.
        /// </summary>
        public const string CatalogKey = "/BSIAnnotColumns";

        public enum Status
        {
            /// <summary>Removed /BSIAnnotColumns and saved the file.</summary>
            Scrubbed,

            /// <summary>File did not contain /BSIAnnotColumns. No write occurred.</summary>
            AlreadyClean,

            /// <summary>An error occurred while reading or writing.</summary>
            Failed,
        }

        public sealed class Result
        {
            public Status Status { get; init; }
            public string FilePath { get; init; } = "";
            public string? Message { get; init; }
            public Exception? Exception { get; init; }

            public bool Success => Status != Status.Failed;
        }

        /// <summary>
        /// Strips /BSIAnnotColumns from <paramref name="pdfPath"/> in place.
        ///
        /// Atomic: writes to a sibling .scrub.tmp file first, then
        /// File.Replace swaps it in. If anything fails partway, the
        /// original file is untouched.
        ///
        /// Idempotent: scrubbing a clean file returns AlreadyClean without
        /// writing.
        /// </summary>
        /// <param name="pdfPath">Absolute path to the PDF.</param>
        public static Result Scrub(string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(pdfPath))
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = pdfPath ?? "",
                    Message  = "PDF path is empty.",
                };
            }

            if (!File.Exists(pdfPath))
            {
                return new Result
                {
                    Status   = Status.Failed,
                    FilePath = pdfPath,
                    Message  = "File not found: " + pdfPath,
                };
            }

            string tmpPath = pdfPath + ".scrub.tmp";

            try
            {
                // Open in Modify mode so PdfSharp gives us a writable Catalog.
                using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
                {
                    var catalog = doc.Internals.Catalog;

                    if (!catalog.Elements.ContainsKey(CatalogKey))
                    {
                        // Already clean. Nothing to write.
                        return new Result
                        {
                            Status   = Status.AlreadyClean,
                            FilePath = pdfPath,
                            Message  = "/BSIAnnotColumns not present.",
                        };
                    }

                    catalog.Elements.Remove(CatalogKey);

                    // Save to .tmp first for atomic replace.
                    doc.Save(tmpPath);
                } // doc disposed here, file handle on tmpPath released

                // Atomic replace on NTFS: original goes away, tmp becomes
                // original. We don't keep .bak files (clutters Drive folders).
                File.Replace(tmpPath, pdfPath, destinationBackupFileName: null);

                return new Result
                {
                    Status   = Status.Scrubbed,
                    FilePath = pdfPath,
                    Message  = "Removed /BSIAnnotColumns and saved.",
                };
            }
            catch (Exception ex)
            {
                // Best-effort cleanup of the .tmp file.
                try { if (File.Exists(tmpPath)) File.Delete(tmpPath); } catch { }

                return new Result
                {
                    Status    = Status.Failed,
                    FilePath  = pdfPath,
                    Message   = "Scrub failed: " + ex.Message,
                    Exception = ex,
                };
            }
        }

        /// <summary>
        /// Returns true if the PDF currently contains /BSIAnnotColumns
        /// (i.e. would be modified by Scrub). Read-only; does not modify
        /// the file.
        ///
        /// Used by ambient detection -- e.g. checking the active Revu file
        /// before showing the "Fix it" toast on clipboard markup send.
        ///
        /// Returns false on any read failure (corrupt PDF, locked file,
        /// encrypted, etc.) -- better to skip a toast than to spam one on
        /// a file we cant actually read.
        /// </summary>
        public static bool HasBsiAnnotColumns(string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath))
                return false;

            try
            {
                using var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly);
                return doc.Internals.Catalog.Elements.ContainsKey(CatalogKey);
            }
            catch
            {
                return false;
            }
        }
    }
}
'@
$scrubberPath = Join-Path $projectDirInner "BsiAnnotColumnsScrubber.cs"
[void](Save-NewFile -Path $scrubberPath -Content $scrubberCs -Label "BsiAnnotColumnsScrubber.cs")

# --- 6. Done -----------------------------------------------------------------
Write-Step "Done"
if ($DryRun) {
    Write-Warn2 "DRY RUN -- no files were modified"
} else {
    Write-Ok "v$ToVersion step 1 applied successfully"
    Write-Info ""
    Write-Info "Next:"
    Write-Info "  cd `"$ProjectDir`""
    Write-Info "  dotnet restore"
    Write-Info "  dotnet build"
    Write-Info ""
    Write-Info "After build succeeds, BsiAnnotColumnsScrubber is callable from anywhere."
    Write-Info "Step 2 (toast helper + active-file orchestrator) and step 3 (tray menu"
    Write-Info "+ HTTP endpoints) come next."
}