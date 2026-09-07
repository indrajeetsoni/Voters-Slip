# ==============================================================================
# VOTER SUVIDHA - PDF & HTML SLIP GENERATOR MODULE
# ==============================================================================
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-LayoutConfig($slipsPerPage) {
    switch ($slipsPerPage) {
        4 {
            return @{
                Rows = 2
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "4mm 5mm"
                TitleSize = "16px"
                MetaSize = "13px"
                NameSize = "18px"
                DetailSize = "13.8px"
                BoothSize = "12.5px"
                CNameSize = "16px"
                CPartySize = "12.5px"
                CAppealSize = "11px"
                ImgW = "68px"
                ImgH = "76px"
                SymW = "52px"
                SymH = "52px"
            }
        }
        6 {
            return @{
                Rows = 3
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "3.5mm 4.5mm"
                TitleSize = "15px"
                MetaSize = "12px"
                NameSize = "16.5px"
                DetailSize = "12.5px"
                BoothSize = "11px"
                CNameSize = "14px"
                CPartySize = "11px"
                CAppealSize = "9.2px"
                ImgW = "58px"
                ImgH = "64px"
                SymW = "42px"
                SymH = "42px"
            }
        }
        8 {
            return @{
                Rows = 4
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "2.8mm 4mm"
                TitleSize = "13.5px"
                MetaSize = "10.5px"
                NameSize = "14.5px"
                DetailSize = "11px"
                BoothSize = "9.8px"
                CNameSize = "12.2px"
                CPartySize = "9.5px"
                CAppealSize = "8px"
                ImgW = "46px"
                ImgH = "50px"
                SymW = "34px"
                SymH = "34px"
            }
        }
        10 {
            return @{
                Rows = 5
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "2.2mm 3.5mm"
                TitleSize = "12px"
                MetaSize = "9.5px"
                NameSize = "13px"
                DetailSize = "9.8px"
                BoothSize = "8.8px"
                CNameSize = "11px"
                CPartySize = "8.5px"
                CAppealSize = "7.2px"
                ImgW = "40px"
                ImgH = "44px"
                SymW = "28px"
                SymH = "28px"
            }
        }
        12 {
            return @{
                Rows = 6
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "1.8mm 3mm"
                TitleSize = "10.5px"
                MetaSize = "8.8px"
                NameSize = "11.8px"
                DetailSize = "8.8px"
                BoothSize = "8px"
                CNameSize = "9.8px"
                CPartySize = "7.8px"
                CAppealSize = "6.8px"
                ImgW = "34px"
                ImgH = "38px"
                SymW = "24px"
                SymH = "24px"
            }
        }
        default {
            return @{
                Rows = 4
                Cols = 2
                PagePadding = "5mm 5mm"
                GridGap = "2.8mm 4mm"
                TitleSize = "13.5px"
                MetaSize = "10.5px"
                NameSize = "14.5px"
                DetailSize = "11px"
                BoothSize = "9.8px"
                CNameSize = "12.2px"
                CPartySize = "9.5px"
                CAppealSize = "8px"
                ImgW = "46px"
                ImgH = "50px"
                SymW = "34px"
                SymH = "34px"
            }
        }
    }
}

function Generate-VoterSlipsHtml($votersList, $config, $isSampleOnly = $false) {
    $slipsPerPage = if ($config.SlipsPerPage) { [int]$config.SlipsPerPage } else { 8 }
    $layout = Get-LayoutConfig $slipsPerPage

    $candidate = if ($config.CandidateName) { $config.CandidateName } else { "मनोज बाबेल" }
    $party = if ($config.PartyName) { $config.PartyName } else { "भारतीय जनता पार्टी (BJP)" }
    $message = if ($config.BottomMessage) { $config.BottomMessage } else { "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!" }
    $candidatePhoto = $config.CandidatePhoto
    $partySymbol = $config.PartySymbol

    $photoHtml = if ($candidatePhoto) { "<img class=`"cand-photo`" src=`"$candidatePhoto`" alt=`"Candidate`">" } else { '<div class="cand-avatar">&#128100;</div>' }
    $symbolHtml = if ($partySymbol) { "<div class=`"cand-symbol-frame`"><img class=`"cand-symbol-img`" src=`"$partySymbol`" alt=`"चुनाव चिन्ह`"></div>" } else { "" }

    $totalVoters = $votersList.Count
    $sliceCount = if ($isSampleOnly) { [Math]::Min($slipsPerPage, $totalVoters) } else { $totalVoters }
    $pageCount = [Math]::Ceiling($sliceCount / $slipsPerPage)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="hi">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('  <meta charset="UTF-8">')
    [void]$sb.AppendLine('  <title>&#2357;&#2379;&#2335;&#2352; &#2360;&#2369;&#2357;&#2367;&#2343;&#2366; &#2360;&#2381;&#2354;&#2367;&#2346;</title>')
    [void]$sb.AppendLine('  <style>')
    [void]$sb.AppendLine('    @page { size: A4 portrait; margin: 0; }')
    [void]$sb.AppendLine('    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }')
    [void]$sb.AppendLine('    body { font-family: "Nirmala UI", "Mangal", "Segoe UI", Arial, sans-serif; background: #fff; color: #111; }')
    [void]$sb.AppendLine("    .a4-page { width: 210mm; height: 297mm; padding: $($layout.PagePadding); page-break-after: always; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }")
    [void]$sb.AppendLine("    .slips-grid { display: grid; grid-template-columns: repeat($($layout.Cols), 1fr); grid-template-rows: repeat($($layout.Rows), 1fr); gap: $($layout.GridGap); width: 100%; height: 100%; }")
    [void]$sb.AppendLine('    .slip { border: 1.5px solid #111; border-radius: 5px; padding: 4px 5px; display: flex; flex-direction: row; justify-content: space-between; align-items: stretch; gap: 6px; background: #fff; overflow: hidden; }')
    [void]$sb.AppendLine('    .slip-left { width: 63%; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }')
    [void]$sb.AppendLine('    .slip-header-block { border-bottom: 1.2px dashed #444; padding-bottom: 2px; margin-bottom: 1.5px; }')
    [void]$sb.AppendLine("    .slip-title { text-align: center; font-weight: 900; font-size: $($layout.TitleSize); letter-spacing: 0.6px; color: #000; margin-bottom: 1.5px; }")
    [void]$sb.AppendLine("    .meta-row { display: flex; justify-content: space-between; align-items: center; font-size: $($layout.MetaSize); margin-bottom: 1.5px; }")
    [void]$sb.AppendLine('    .meta-row-serial { display: flex; justify-content: flex-start; align-items: center; }')
    [void]$sb.AppendLine('    .badge-item { color: #111; }')
    [void]$sb.AppendLine('    .badge-item strong { color: #000; }')
    [void]$sb.AppendLine("    .serial-badge { background: #f4f5f7; border: 1.3px solid #111; border-radius: 3px; padding: 0.5px 5px; font-size: $($layout.MetaSize); font-weight: 800; display: inline-block; }")
    [void]$sb.AppendLine('    .voter-body { flex: 1; display: flex; flex-direction: column; justify-content: space-evenly; padding: 1px 0; }')
    [void]$sb.AppendLine("    .voter-line { font-size: $($layout.DetailSize); color: #111; line-height: 1.25; }")
    [void]$sb.AppendLine("    .voter-name { font-size: $($layout.NameSize); font-weight: 900; color: #000; }")
    [void]$sb.AppendLine('    .voter-line strong { color: #000; }')
    [void]$sb.AppendLine("    .booth-line { font-size: $($layout.BoothSize); line-height: 1.22; border-top: 1.2px dashed #666; padding-top: 2px; margin-top: 1px; color: #000; background: #fafafa; border-radius: 2px; padding-left: 2px; }")
    [void]$sb.AppendLine('    .booth-label { font-weight: 900; color: #000; }')
    [void]$sb.AppendLine('    .slip-right-box { width: 36%; min-width: 36%; max-width: 37%; border: 1.8px solid #1a237e; border-radius: 6px; background: #fbfbfd; padding: 3px 3px; display: flex; flex-direction: column; align-items: center; justify-content: space-between; text-align: center; box-sizing: border-box; }')
    [void]$sb.AppendLine("    .cand-photo-frame { width: $($layout.ImgW); height: $($layout.ImgH); border-radius: 4px; border: 1.2px solid #7986cb; background: #e8eaf6; display: flex; align-items: center; justify-content: center; overflow: hidden; }")
    [void]$sb.AppendLine('    .cand-photo { width: 100%; height: 100%; object-fit: cover; display: block; }')
    [void]$sb.AppendLine('    .cand-avatar { font-size: 28px; line-height: 1; }')
    [void]$sb.AppendLine("    .cand-name { font-size: $($layout.CNameSize); font-weight: 900; color: #0d47a1; line-height: 1.2; margin-top: 1px; }")
    [void]$sb.AppendLine("    .cand-symbol-frame { width: $($layout.SymW); height: $($layout.SymH); display: flex; align-items: center; justify-content: center; margin: 1px 0; }")
    [void]$sb.AppendLine('    .cand-symbol-img { max-width: 100%; max-height: 100%; object-fit: contain; }')
    [void]$sb.AppendLine("    .cand-party { font-size: $($layout.CPartySize); font-weight: 800; color: #2e7d32; line-height: 1.2; }")
    [void]$sb.AppendLine("    .cand-appeal { font-size: $($layout.CAppealSize); font-weight: 800; color: #b71c1c; line-height: 1.2; margin-top: 1px; }")
    [void]$sb.AppendLine('  </style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')

    $voterIdx = 0
    for ($p = 0; $p -lt $pageCount; $p++) {
        [void]$sb.AppendLine('  <div class="a4-page">')
        [void]$sb.AppendLine('    <div class="slips-grid">')

        for ($s = 0; $s -lt $slipsPerPage; $s++) {
            if ($voterIdx -lt $sliceCount) {
                $v = $votersList[$voterIdx]
                $w = [System.Security.SecurityElement]::Escape("$($v.Ward)")
                $part = [System.Security.SecurityElement]::Escape("$($v.Part)")
                $sn = [System.Security.SecurityElement]::Escape("$($v.SerialNo)")
                $nm = [System.Security.SecurityElement]::Escape("$($v.VoterName)")
                $rel = [System.Security.SecurityElement]::Escape("$($v.RelativeName)")
                $h = [System.Security.SecurityElement]::Escape("$($v.HouseNo)")
                $age = [System.Security.SecurityElement]::Escape("$($v.Age)")
                $g = [System.Security.SecurityElement]::Escape("$($v.Gender)")
                $epic = [System.Security.SecurityElement]::Escape("$($v.EPIC)")
                $booth = [System.Security.SecurityElement]::Escape("$($v.Booth)")

                $cHtml = [System.Security.SecurityElement]::Escape("$candidate")
                $pHtml = [System.Security.SecurityElement]::Escape("$party")
                $mHtml = [System.Security.SecurityElement]::Escape("$message")

                [void]$sb.AppendLine('      <div class="slip">')
                [void]$sb.AppendLine('        <div class="slip-left">')
                [void]$sb.AppendLine('          <div class="slip-header-block">')
                [void]$sb.AppendLine('            <div class="slip-title">&#2357;&#2379;&#2335;&#2352; &#2360;&#2369;&#2357;&#2367;&#2343;&#2366; &#2360;&#2381;&#2354;&#2367;&#2346;</div>')
                [void]$sb.AppendLine('            <div class="meta-row">')
                [void]$sb.AppendLine("              <span class=`"badge-item`">&#2357;&#2366;&#2352;&#2381;&#2337; &#2344;&#2306; : <strong>[ $w ]</strong></span>")
                [void]$sb.AppendLine("              <span class=`"badge-item`">&#2349;&#2366;&#2327; : <strong>[ $part ]</strong></span>")
                [void]$sb.AppendLine('            </div>')
                [void]$sb.AppendLine('            <div class="meta-row-serial">')
                [void]$sb.AppendLine("              <span class=`"badge-item serial-badge`">&#2325;&#2381;&#2352;&#2350; &#2360;&#2306;&#2326;&#2381;&#2351;&#2366; : <strong>[ $sn ]</strong></span>")
                [void]$sb.AppendLine('            </div>')
                [void]$sb.AppendLine('          </div>')
                [void]$sb.AppendLine('          <div class="voter-body">')
                [void]$sb.AppendLine("            <div class=`"voter-line voter-name`">&#2350;&#2340;&#2342;&#2366;&#2340;&#2366; &#2325;&#2366; &#2344;&#2366;&#2350; : <strong>$nm</strong></div>")
                [void]$sb.AppendLine("            <div class=`"voter-line`">&#2346;&#2367;&#2340;&#2366;/&#2346;&#2340;&#2367; &#2325;&#2366; &#2344;&#2366;&#2350; : <span>$rel</span></div>")
                [void]$sb.AppendLine("            <div class=`"voter-line`">&#2350;&#2325;&#2366;&#2344; &#2344;&#2306;. : <strong>$h</strong> &nbsp;|&nbsp; &#2313;&#2350;&#2381;&#2352; : <strong>$age</strong> &nbsp;|&nbsp; &#2354;&#2367;&#2306;&#2327; : <strong>$g</strong></div>")
                [void]$sb.AppendLine("            <div class=`"voter-line`">&#2346;&#2361;&#2331;&#2366;&#2344; &#2346;&#2340;&#2381;&#2352; &#2325;&#2381;&#2352;. (EPIC) : <strong>$epic</strong></div>")
                [void]$sb.AppendLine('          </div>')
                [void]$sb.AppendLine("          <div class=`"booth-line`"><span class=`"booth-label`">&#2350;&#2340;&#2342;&#2366;&#2344; &#2325;&#2375;&#2306;&#2342;&#2381;&#2352; :</span> $booth</div>")
                [void]$sb.AppendLine('        </div>')
                [void]$sb.AppendLine('        <div class="slip-right-box">')
                [void]$sb.AppendLine("          <div class=`"cand-photo-frame`">$photoHtml</div>")
                [void]$sb.AppendLine("          <div class=`"cand-name`">$cHtml</div>")
                if ($symbolHtml) {
                    [void]$sb.AppendLine("          $symbolHtml")
                }
                [void]$sb.AppendLine("          <div class=`"cand-party`">($pHtml)</div>")
                [void]$sb.AppendLine("          <div class=`"cand-appeal`">$mHtml</div>")
                [void]$sb.AppendLine('        </div>')
                [void]$sb.AppendLine('      </div>')

                $voterIdx++
            } else {
                [void]$sb.AppendLine('      <div class="slip" style="visibility:hidden;"></div>')
            }
        }

        [void]$sb.AppendLine('    </div>')
        [void]$sb.AppendLine('  </div>')
    }

    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')

    return $sb.ToString()
}

function Export-VoterSlipsPdf($votersList, $config, $outPdfPath) {
    $browserCandidates = @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
    )
    $browserExe = $null
    foreach ($cand in $browserCandidates) {
        if ($cand -and (Test-Path $cand)) {
            $browserExe = $cand
            break
        }
    }

    if (-not $browserExe) {
        throw "Neither Google Chrome nor Microsoft Edge was found for offline PDF generation."
    }

    $tempHtml = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "voter_slips_print.html")
    $htmlContent = Generate-VoterSlipsHtml $votersList $config $false
    [System.IO.File]::WriteAllText($tempHtml, $htmlContent, [System.Text.Encoding]::UTF8)

    $procArgs = @(
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        "--print-to-pdf=`"$outPdfPath`"",
        "`"$tempHtml`""
    )

    $p = Start-Process -FilePath $browserExe -ArgumentList $procArgs -Wait -NoNewWindow -PassThru
    if ($p.ExitCode -ne 0 -and -not (Test-Path $outPdfPath)) {
        throw "Browser PDF printing exited with code $($p.ExitCode)"
    }

    if (Test-Path $tempHtml) {
        Remove-Item $tempHtml -Force -ErrorAction SilentlyContinue
    }
}
