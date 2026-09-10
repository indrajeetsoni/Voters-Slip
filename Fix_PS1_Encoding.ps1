# =============================================================================
# Voter Suvidha - Fix .ps1 file encoding
#
# WHY: Windows PowerShell 5.1 reads .ps1 files without a byte-order mark using
# the legacy ANSI codepage. server.ps1 / extractor.ps1 contain Devanagari text,
# so ANSI misreads those bytes and can produce stray quote or backtick
# characters, which breaks the parser. The dot-source then fails and functions
# like Extract-VotersFromPdf are never defined.
#
# This script re-reads each .ps1 as UTF-8 and re-saves it as UTF-8 WITH BOM,
# which PowerShell 5.1 reads correctly.
#
# HOW TO RUN (from the folder that contains Voter_Suvidha.bat):
#   powershell -ExecutionPolicy Bypass -NoProfile -File .\Fix_PS1_Encoding.ps1
# =============================================================================

$ErrorActionPreference = 'Stop'

$root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$psDir = Join-Path $root 'voter_suvidha'

if (-not (Test-Path $psDir)) {
    Write-Host "Could not find: $psDir" -ForegroundColor Red
    Write-Host "Put this script in the folder that contains Voter_Suvidha.bat." -ForegroundColor Yellow
    exit 1
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$utf8Bom   = New-Object System.Text.UTF8Encoding($true)

$files = Get-ChildItem -Path $psDir -Filter '*.ps1' -File

if ($files.Count -eq 0) {
    Write-Host "No .ps1 files found in $psDir" -ForegroundColor Yellow
    exit 1
}

foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)

    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

    if ($hasBom) {
        Write-Host ("SKIP   {0}  (already UTF-8 with BOM)" -f $f.Name) -ForegroundColor DarkGray
        continue
    }

    # Back up before touching anything
    $backup = "$($f.FullName).bak"
    if (-not (Test-Path $backup)) {
        [System.IO.File]::WriteAllBytes($backup, $bytes)
    }

    # Read the raw bytes as UTF-8 (not ANSI), then write back with a BOM
    $text = $utf8NoBom.GetString($bytes)
    [System.IO.File]::WriteAllText($f.FullName, $text, $utf8Bom)

    Write-Host ("FIXED  {0}  (backup: {1})" -f $f.Name, (Split-Path $backup -Leaf)) -ForegroundColor Green
}

Write-Host ""
Write-Host "Now verifying that each module loads cleanly..." -ForegroundColor Cyan
Write-Host ""

foreach ($f in $files) {
    if ($f.Name -eq 'server.ps1') { continue }
    try {
        . $f.FullName
        Write-Host ("LOAD OK   {0}" -f $f.Name) -ForegroundColor Green
    } catch {
        Write-Host ("LOAD FAIL {0}" -f $f.Name) -ForegroundColor Red
        Write-Host ("          {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
}

Write-Host ""
foreach ($fn in @('Extract-VotersFromPdf', 'Export-VotersToExcel', 'Export-VoterSlipsPdf', 'Get-InstalledBrowserPath')) {
    if (Get-Command $fn -ErrorAction SilentlyContinue) {
        Write-Host ("DEFINED   {0}" -f $fn) -ForegroundColor Green
    } else {
        Write-Host ("MISSING   {0}" -f $fn) -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
