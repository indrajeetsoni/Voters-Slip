# Voter Suvidha — Fix for `Extract-VotersFromPdf is not recognized`

Target folder: `C:\Users\sonit\Downloads\Voters-Slip-main (2)\Voters-Slip-main`

---

## What actually went wrong

1. `Voter_Suvidha.bat` checks for `py`, then `python`. Neither was on your PATH,
   so it fell through to the PowerShell backend, `voter_suvidha\server.ps1`.
2. On upload, `server.ps1` tries to hand the PDF to `extractor.py` first. No
   Python, so that step was skipped.
3. It then fell back to its own `Extract-VotersFromPdf`. That function is
   defined in `extractor.ps1`, but the `. (Join-Path $scriptDir "extractor.ps1")`
   dot-source at the top of `server.ps1` had already failed. PowerShell prints a
   dot-source failure and keeps running, so nothing broke until upload time.

Two independent problems, so two fixes. Fix 1 is the one that makes the app
work. Fix 2 and 3 stop the failure from being silent next time.

---

## Fix 1 — Get the Python path working (do this first)

The Python extractor (`extractor.py`, using PyMuPDF) is the good
implementation. The PowerShell extractor is a raw PDF-stream parser and is only
a fallback. You want Python.

1. Install Python 3 from <https://www.python.org/downloads/windows/>
   — on the first installer screen, **tick "Add python.exe to PATH"**.
2. Copy `SETUP_FIRST_TIME.bat` and the new `Voter_Suvidha.bat` into
   `Voters-Slip-main\` (next to the existing `Voter_Suvidha.bat` — overwrite it).
3. Double-click **`SETUP_FIRST_TIME.bat`** once. It installs `pymupdf` and
   `openpyxl`. This is the only step that needs internet.
4. Double-click **`Voter_Suvidha.bat`**. Upload should now work.

The new launcher also finds Python in common install folders even when it isn't
on PATH, verifies both packages before starting, and prints a plain-language
message instead of falling back to a broken backend.

Note: use `Voter_Suvidha.bat`, not `VoterSuvidha.exe`. The `.exe` launches
`server.ps1` directly and skips all of the above.

---

## Fix 2 — Repair the `.ps1` encoding

Most likely reason the dot-source failed: `extractor.ps1` and `server.ps1`
contain Devanagari text but are saved as UTF-8 **without** a byte-order mark.
Windows PowerShell 5.1 reads BOM-less files using the legacy ANSI codepage, so
those bytes get misread and can produce a stray quote or backtick that breaks
the parser.

From the `Voters-Slip-main` folder, open PowerShell and run:

```powershell
powershell -ExecutionPolicy Bypass -NoProfile -File .\Fix_PS1_Encoding.ps1
```

It backs up each file as `*.ps1.bak`, re-saves as UTF-8 with BOM, then reports
which modules load and which functions got defined. Read that output — if a
module still says `LOAD FAIL`, the message next to it is the real parse error.

To see the raw error yourself at any time:

```powershell
powershell -NoProfile -Command ". 'C:\Users\sonit\Downloads\Voters-Slip-main (2)\Voters-Slip-main\voter_suvidha\extractor.ps1'"
```

---

## Fix 3 — Make `server.ps1` fail loudly instead of silently

Two small edits to `voter_suvidha\server.ps1`.

### Edit A — verify modules at startup

**Find:**

```powershell
# Load sub-modules
. (Join-Path $scriptDir "extractor.ps1")
. (Join-Path $scriptDir "excel_builder.ps1")
. (Join-Path $scriptDir "pdf_builder.ps1")
```

**Replace with:**

```powershell
# Load sub-modules - fail loudly, do not start a half-broken server
$moduleLoadErrors = @()
foreach ($m in @("extractor.ps1", "excel_builder.ps1", "pdf_builder.ps1")) {
    $mPath = Join-Path $scriptDir $m
    if (-not (Test-Path $mPath)) {
        $moduleLoadErrors += "$m : file not found at $mPath"
        continue
    }
    try {
        . $mPath
        Write-Host "Loaded module: $m" -ForegroundColor DarkGray
    } catch {
        $moduleLoadErrors += "$m : $($_.Exception.Message)"
    }
}

# Confirm the functions the server actually depends on are now defined
foreach ($needed in @("Extract-VotersFromPdf", "Export-VotersToExcel", "Export-VoterSlipsPdf")) {
    if (-not (Get-Command $needed -ErrorAction SilentlyContinue)) {
        $moduleLoadErrors += "function '$needed' was not defined after loading modules"
    }
}

if ($moduleLoadErrors.Count -gt 0) {
    Write-Host "" 
    Write-Host "=======================================================================" -ForegroundColor Red
    Write-Host "  MODULE LOAD FAILED - server not started" -ForegroundColor Red
    Write-Host "=======================================================================" -ForegroundColor Red
    foreach ($e in $moduleLoadErrors) { Write-Host "  - $e" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "  Likely fix: run Fix_PS1_Encoding.ps1, then try again." -ForegroundColor Cyan
    Write-Host "  Better fix: install Python 3 and run Voter_Suvidha.bat," -ForegroundColor Cyan
    Write-Host "              which uses the Python extractor instead." -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
```

### Edit B — clear message when there's no working extractor

**Find:**

```powershell
                        if (-not $extracted) {
                            $extracted = Extract-VotersFromPdf $targetPath
                        }
```

**Replace with:**

```powershell
                        if (-not $extracted) {
                            if (-not (Get-Command Extract-VotersFromPdf -ErrorAction SilentlyContinue)) {
                                throw "No working PDF extractor. Python was not found (so extractor.py could not be used) and extractor.ps1 did not load. Install Python 3, run SETUP_FIRST_TIME.bat, then start the app with Voter_Suvidha.bat."
                            }
                            $extracted = Extract-VotersFromPdf $targetPath
                        }
```

Edit B is what turns the message you saw into something you can act on.

---

## After it runs — one bug worth knowing about

In `voter_suvidha\extractor.py`, `extract_pdf_elector_data` returns a hardcoded
part number:

```python
"part": "1",
```

So if you upload two भाग (part) files for one ward, both get labelled भाग 1 in
the Excel and on every slip. Not your current blocker, but it will bite you on a
multi-part ward. Tell me when you get there and I'll write the part-number
detection.

---

## Quick checklist

- [ ] Python 3 installed, "Add to PATH" ticked
- [ ] `SETUP_FIRST_TIME.bat` run once, ended with "SETUP COMPLETE"
- [ ] New `Voter_Suvidha.bat` copied in, old one overwritten
- [ ] `Fix_PS1_Encoding.ps1` run, output checked
- [ ] `server.ps1` Edits A and B applied
- [ ] Launch with `Voter_Suvidha.bat` (not `VoterSuvidha.exe`)
- [ ] Upload `BADNOR-Ward No-001.pdf` and confirm the stats panel fills in
