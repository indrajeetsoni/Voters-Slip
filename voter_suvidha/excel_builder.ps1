# ==============================================================================
# VOTER SUVIDHA - EXCEL EXPORT MODULE
# ==============================================================================
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Export-VotersToExcel($votersList, $outputPath, $templatePath = "") {
    if ([string]::IsNullOrWhiteSpace($templatePath)) {
        $localTemplate = Join-Path $PSScriptRoot "template.xlsx"
        if (Test-Path $localTemplate) {
            $templatePath = $localTemplate
        } else {
            $workspaceDir = "c:\Users\sonit\Downloads\serene-nobel\serene-nobel"
            $templatePath = Join-Path $workspaceDir "beawar_ward_001_part_001.xlsx"
        }
    }

    if (-not (Test-Path $templatePath)) {
        throw "Template Excel file not found: $templatePath"
    }

    # Copy template to target output path
    Copy-Item $templatePath $outputPath -Force

    # Update sheet1.xml in zip archive
    $zip = [System.IO.Compression.ZipFile]::Open($outputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    $existingEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    if ($existingEntry) { $existingEntry.Delete() }
    
    $newEntry = $zip.CreateEntry('xl/worksheets/sheet1.xml', [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)

    try {
        $writer.WriteLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        $writer.WriteLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

        $totalRows = $votersList.Count + 1
        $writer.WriteLine("<dimension ref=`"A1:L$totalRows`"/>")
        $writer.WriteLine('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>')
        $writer.WriteLine('<sheetFormatPr defaultRowHeight="15"/>')
        $writer.WriteLine('<cols>')
        $writer.WriteLine('<col min="1" max="1" width="8" customWidth="1"/>')
        $writer.WriteLine('<col min="2" max="2" width="12" customWidth="1"/>')
        $writer.WriteLine('<col min="3" max="3" width="12" customWidth="1"/>')
        $writer.WriteLine('<col min="4" max="4" width="45" customWidth="1"/>')
        $writer.WriteLine('<col min="5" max="5" width="12" customWidth="1"/>')
        $writer.WriteLine('<col min="6" max="6" width="24" customWidth="1"/>')
        $writer.WriteLine('<col min="7" max="7" width="24" customWidth="1"/>')
        $writer.WriteLine('<col min="8" max="8" width="14" customWidth="1"/>')
        $writer.WriteLine('<col min="9" max="9" width="8" customWidth="1"/>')
        $writer.WriteLine('<col min="10" max="10" width="8" customWidth="1"/>')
        $writer.WriteLine('<col min="11" max="11" width="18" customWidth="1"/>')
        $writer.WriteLine('<col min="12" max="12" width="18" customWidth="1"/>')
        $writer.WriteLine('</cols>')
        $writer.WriteLine('<sheetData>')

        # Header Row (Exact 12 columns)
        $headers = @(
            "Sr No", "वार्ड संख्या", "भाग संख्या", "मतदान केंद्र की संख्या व पता",
            "क्रम संख्या", "निर्वाचक का नाम", "पिता/पति का नाम", "मकान संख्या",
            "आयु", "लिंग", "EPIC No", "ग्राम पंचायत"
        )
        $colLetters = @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L')

        $writer.WriteLine('<row r="1" spans="1:12">')
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $col = $colLetters[$c]
            $val = [System.Security.SecurityElement]::Escape($headers[$c])
            $writer.WriteLine("<c r=`"$col`1`" t=`"inlineStr`" s=`"1`"><is><t>$val</t></is></c>")
        }
        $writer.WriteLine('</row>')

        # Data Rows
        $srNo = 1
        foreach ($v in $votersList) {
            $r = $srNo + 1
            $w = [System.Security.SecurityElement]::Escape("$($v.Ward)")
            $p = [System.Security.SecurityElement]::Escape("$($v.Part)")
            $b = [System.Security.SecurityElement]::Escape("$($v.Booth)")
            $sn = [System.Security.SecurityElement]::Escape("$($v.SerialNo)")
            $nm = [System.Security.SecurityElement]::Escape("$($v.VoterName)")
            $rel = [System.Security.SecurityElement]::Escape("$($v.RelativeName)")
            $h = [System.Security.SecurityElement]::Escape("$($v.HouseNo)")
            $age = [System.Security.SecurityElement]::Escape("$($v.Age)")
            $g = [System.Security.SecurityElement]::Escape("$($v.Gender)")
            $epic = [System.Security.SecurityElement]::Escape("$($v.EPIC)")
            $gp = [System.Security.SecurityElement]::Escape("$($v.GramPanchayat)")

            $writer.WriteLine("<row r=`"$r`" spans=`"1:12`">" +
                "<c r=`"A$r`"><v>$srNo</v></c>" +
                "<c r=`"B$r`" t=`"inlineStr`"><is><t>$w</t></is></c>" +
                "<c r=`"C$r`" t=`"inlineStr`"><is><t>$p</t></is></c>" +
                "<c r=`"D$r`" t=`"inlineStr`"><is><t>$b</t></is></c>" +
                "<c r=`"E$r`"><v>$sn</v></c>" +
                "<c r=`"F$r`" t=`"inlineStr`"><is><t>$nm</t></is></c>" +
                "<c r=`"G$r`" t=`"inlineStr`"><is><t>$rel</t></is></c>" +
                "<c r=`"H$r`" t=`"inlineStr`"><is><t>$h</t></is></c>" +
                "<c r=`"I$r`"><v>$age</v></c>" +
                "<c r=`"J$r`" t=`"inlineStr`"><is><t>$g</t></is></c>" +
                "<c r=`"K$r`" t=`"inlineStr`"><is><t>$epic</t></is></c>" +
                "<c r=`"L$r`" t=`"inlineStr`"><is><t>$gp</t></is></c>" +
                "</row>")
            $srNo++
        }

        $writer.WriteLine('</sheetData>')
        $writer.WriteLine('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        $writer.WriteLine('</worksheet>')
    } finally {
        $writer.Flush()
        $writer.Close()
        $zip.Dispose()
    }

    return $outputPath
}
