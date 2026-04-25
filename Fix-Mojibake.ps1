# Fix-Mojibake.ps1
#
# Repairs cp1252-mangled UTF-8 in TabsPortalHelper .cs source files.
#
# When a UTF-8 file gets read as cp1252 and re-saved, multi-byte chars
# like em-dash ("\u2014") show up as "\u00e2\u0080\u201d" (visually "\u00e2\u20ac\u201d")
# in the bytes. This script reverses that by reading the file's bytes
# AS cp1252, which correctly reconstructs the original UTF-8 string,
# then writes it back as UTF-8 (no BOM).
#
# Idempotent on already-clean files: a clean UTF-8 em-dash decoded as
# cp1252 produces gibberish, but only if the file was already mangled.
# Files that are pristine UTF-8 will be re-mangled by this script, so
# we sniff for the mojibake signature first and skip clean files.

$ErrorActionPreference = 'Stop'

$files = Get-ChildItem 'C:\TabsPortalHelper\TabsPortalHelper\*.cs'

foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)

    # Mojibake signature: byte 0xE2 followed by 0x80 followed by something.
    # Real UTF-8 em-dash IS that byte sequence too — but in a UTF-8 file the
    # sequence forms a valid char and the FILE is already declared UTF-8, so
    # this script reads the file's existing encoding and decides.
    #
    # Heuristic: if reading as UTF-8 produces "\u00e2\u20ac" anywhere, it's mojibake.
    $utf8Text = [System.Text.Encoding]::UTF8.GetString($bytes)

    if ($utf8Text -match '\u00e2\u20ac') {
        # Mojibake detected. Decode the bytes as cp1252 to recover the original UTF-8.
        $repaired = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)
        # Write back as UTF-8 without BOM.
        [System.IO.File]::WriteAllText(
            $f.FullName,
            $repaired,
            [System.Text.UTF8Encoding]::new($false))
        Write-Host "Fixed: $($f.Name)" -ForegroundColor Green
    }
    else {
        Write-Host "Clean: $($f.Name)" -ForegroundColor DarkGray
    }
}

Write-Host ''
Write-Host 'Verifying...' -ForegroundColor Cyan
$still = Get-ChildItem 'C:\TabsPortalHelper\TabsPortalHelper\*.cs' |
    Select-String -Pattern '\u00e2\u20ac'

if ($still) {
    Write-Host 'Mojibake still present:' -ForegroundColor Red
    $still
}
else {
    Write-Host 'All .cs files clean.' -ForegroundColor Green
}
