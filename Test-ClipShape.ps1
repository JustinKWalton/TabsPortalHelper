# ==========================================================================
# Test-ClipShape.ps1 -- paste a resized Bluebeam cloud markup to the clipboard
# ==========================================================================
# Paste this whole block into PowerShell (powershell.exe, not pwsh -- see
# the STA note at the bottom). Then Alt+Tab to Bluebeam and Ctrl+V on a PDF.
# Edit the values below to iterate on dimensions and cloud size.
# Flip $SaveAsTemplate to $true to bake the new shape into template.bin
# permanently (re-run Rebuild.ps1 afterwards to pick it up).
# ==========================================================================

# ---- Edit these ----------------------------------------------------------
$WidthIn        = 1.5
$HeightIn       = 2.0
$CloudI         = 1       # cloud bump size -- 1 = small, 2 = medium, 3 = large
$TemplatePath   = 'C:\TabsPortalHelper\TabsPortalHelper\template.bin'
$SaveAsTemplate = $true  # $true = also overwrite template.bin (makes .bak first)
# --------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

# PDF coordinates are in points -- 72 per inch.
$w = [Math]::Round($WidthIn  * 72, 4)
$h = [Math]::Round($HeightIn * 72, 4)

# ---- Load the existing template -----------------------------------------
if (-not (Test-Path $TemplatePath)) { throw "Template not found at $TemplatePath" }
$template = [IO.File]::ReadAllBytes($TemplatePath)

# Find the PDF annotation dict string start.
$marker = [Text.Encoding]::ASCII.GetBytes('<</IT/PolygonCloud')
$markerPos = -1
for ($i = 0; $i -le $template.Length - $marker.Length; $i++) {
    $match = $true
    for ($j = 0; $j -lt $marker.Length; $j++) {
        if ($template[$i + $j] -ne $marker[$j]) { $match = $false; break }
    }
    if ($match) { $markerPos = $i; break }
}
if ($markerPos -lt 0) { throw "Marker not found in template (damaged?)" }

# Walk backwards to find the BinaryFormatter 7-bit LEB128 length prefix.
$prefixStart = $markerPos - 1
while ($prefixStart -gt 0 -and ($template[$prefixStart - 1] -band 0x80) -ne 0) {
    $prefixStart--
}

# Decode the length.
$value = 0; $shift = 0; $off = $prefixStart
while ($true) {
    $b = [int]$template[$off]; $off++
    $value = $value -bor (($b -band 0x7F) -shl $shift)
    if (($b -band 0x80) -eq 0) { break }
    $shift += 7
}
$oldLen = $value

# Extract the PDF dict string (UTF-8 on disk).
$oldData = [Text.Encoding]::UTF8.GetString($template, $markerPos, $oldLen)

# ---- Rewrite /Vertices, /Rect, /I ---------------------------------------
# Rectangle corners starting at (0,0), going CCW.
$newVerts = "[0 0 $w 0 $w $h 0 $h]"
$newRect  = "[0 0 $w $h]"

$newData = $oldData -replace '/Vertices\[[^\]]*\]',    "/Vertices$newVerts"
$newData = $newData -replace '/Rect\[[^\]]*\]',        "/Rect$newRect"
$newData = $newData -replace '/BE<</S/C/I\s+\d+>>',    "/BE<</S/C/I $CloudI>>"

# ---- Re-encode length prefix and splice ---------------------------------
$newBytes = [Text.Encoding]::UTF8.GetBytes($newData)
$newLen   = $newBytes.Length
$prefList = New-Object System.Collections.Generic.List[byte]
$v = $newLen
while ($v -ge 0x80) {
    $prefList.Add([byte](($v -band 0x7F) -bor 0x80))
    $v = $v -shr 7
}
$prefList.Add([byte]$v)
$prefix = [byte[]]$prefList

$suffixStart = $markerPos + $oldLen
$suffixLen   = $template.Length - $suffixStart
$total       = $prefixStart + $prefix.Length + $newBytes.Length + $suffixLen
$out         = New-Object byte[] $total
[Array]::Copy($template, 0,            $out, 0,                                                $prefixStart)
[Array]::Copy($prefix,   0,            $out, $prefixStart,                                     $prefix.Length)
[Array]::Copy($newBytes, 0,            $out, $prefixStart + $prefix.Length,                    $newBytes.Length)
[Array]::Copy($template, $suffixStart, $out, $prefixStart + $prefix.Length + $newBytes.Length, $suffixLen)

# Optionally bake into template.bin.
if ($SaveAsTemplate) {
    Copy-Item $TemplatePath "$TemplatePath.bak" -Force
    [IO.File]::WriteAllBytes($TemplatePath, $out)
    Write-Host "Saved new template to $TemplatePath (backup: $TemplatePath.bak)" -ForegroundColor Green
    Write-Host "Run Rebuild.ps1 to rebuild the helper with the new default shape." -ForegroundColor Gray
    Write-Host ""
}

# ---- Win32 clipboard write for the custom BBCopyItem format -------------
Add-Type -Namespace W32 -Name Clip -ErrorAction SilentlyContinue -MemberDefinition @'
    [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr h);
    [DllImport("user32.dll")] public static extern bool CloseClipboard();
    [DllImport("user32.dll")] public static extern bool EmptyClipboard();
    [DllImport("user32.dll")] public static extern IntPtr SetClipboardData(uint f, IntPtr h);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern uint RegisterClipboardFormatW(string s);
    [DllImport("kernel32.dll")] public static extern IntPtr GlobalAlloc(uint f, UIntPtr s);
    [DllImport("kernel32.dll")] public static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] public static extern bool GlobalUnlock(IntPtr h);
'@

# Clipboard API requires an STA thread. Windows PowerShell 5.1 is STA by
# default; PowerShell 7+ (pwsh) is MTA unless launched with -STA.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "PowerShell is in MTA mode -- clipboard write skipped."
    Write-Warning "Use Windows PowerShell (powershell.exe), or pwsh -sta, and re-run."
    return
}

$fmt = [W32.Clip]::RegisterClipboardFormatW('Bluebeam.Windows.View.Input.BBCopyItem')
if (-not [W32.Clip]::OpenClipboard([IntPtr]::Zero)) { throw "OpenClipboard failed" }
try {
    [void][W32.Clip]::EmptyClipboard()
    $sz  = New-Object UIntPtr([uint64]$out.Length)
    $hG  = [W32.Clip]::GlobalAlloc(2, $sz)                    # 2 = GMEM_MOVEABLE
    if ($hG -eq [IntPtr]::Zero) { throw "GlobalAlloc failed" }
    $ptr = [W32.Clip]::GlobalLock($hG)
    [Runtime.InteropServices.Marshal]::Copy($out, 0, $ptr, $out.Length)
    [void][W32.Clip]::GlobalUnlock($hG)
    if ([W32.Clip]::SetClipboardData($fmt, $hG) -eq [IntPtr]::Zero) { throw "SetClipboardData failed" }
} finally {
    [void][W32.Clip]::CloseClipboard()
}

Write-Host ""
Write-Host ("OK -- Markup on clipboard: {0}in x {1}in, cloud I={2}, {3:N0} bytes total" -f $WidthIn, $HeightIn, $CloudI, $out.Length) -ForegroundColor Green
Write-Host "     Alt+Tab to Bluebeam and Ctrl+V on a PDF." -ForegroundColor DarkGray
Write-Host ""
