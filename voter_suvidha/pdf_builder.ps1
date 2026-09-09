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
                CPostSize = "14px"
                CNameSize = "17.5px"
                CPartySize = "13.5px"
                CAppealSize = "11.5px"
                ImgW = "95px"
                ImgH = "100px"
                SymW = "82px"
                SymH = "82px"
                AvatarSize = "60px"
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
                CPostSize = "12.5px"
                CNameSize = "15px"
                CPartySize = "12px"
                CAppealSize = "10px"
                ImgW = "82px"
                ImgH = "86px"
                SymW = "68px"
                SymH = "68px"
                AvatarSize = "52px"
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
                CPostSize = "11.5px"
                CNameSize = "13px"
                CPartySize = "10.5px"
                CAppealSize = "8.8px"
                ImgW = "68px"
                ImgH = "72px"
                SymW = "56px"
                SymH = "56px"
                AvatarSize = "42px"
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
                CPostSize = "10px"
                CNameSize = "11.5px"
                CPartySize = "9.5px"
                CAppealSize = "7.8px"
                ImgW = "56px"
                ImgH = "58px"
                SymW = "44px"
                SymH = "44px"
                AvatarSize = "34px"
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
                CPostSize = "9px"
                CNameSize = "10.5px"
                CPartySize = "8.5px"
                CAppealSize = "7.2px"
                ImgW = "48px"
                ImgH = "50px"
                SymW = "38px"
                SymH = "38px"
                AvatarSize = "30px"
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
                CPostSize = "11.5px"
                CNameSize = "13px"
                CPartySize = "10.5px"
                CAppealSize = "8.8px"
                ImgW = "68px"
                ImgH = "72px"
                SymW = "56px"
                SymH = "56px"
                AvatarSize = "42px"
            }
        }
    }
}

function Save-Base64ImageToTempFile($base64DataUri, $prefix) {
    if ([string]::IsNullOrWhiteSpace($base64DataUri)) { return $null }
    if ($base64DataUri.StartsWith("data:image/")) {
        $commaIdx = $base64DataUri.IndexOf(",")
        if ($commaIdx -ge 0) {
            $header = $base64DataUri.Substring(0, $commaIdx)
            $rawB64 = $base64DataUri.Substring($commaIdx + 1)
            $ext = ".jpg"
            if ($header -like "*png*") { $ext = ".png" }
            elseif ($header -like "*webp*") { $ext = ".webp" }
            elseif ($header -like "*svg*") { $ext = ".svg" }
            
            $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$prefix$ext")
            try {
                $bytes = [System.Convert]::FromBase64String($rawB64)
                [System.IO.File]::WriteAllBytes($tempPath, $bytes)
                return ([System.Uri]::new($tempPath)).AbsoluteUri
            } catch {
                Write-Host "Warning: Failed to save temp image: $_" -ForegroundColor Yellow
                return $base64DataUri
            }
        }
    }
    return $base64DataUri
}

function Stream-VoterSlipsHtml($votersList, $config, [System.IO.TextWriter]$writer, $isSampleOnly = $false) {
    $slipsPerPage = if ($config.SlipsPerPage) { [int]$config.SlipsPerPage } else { 8 }
    $layout = Get-LayoutConfig $slipsPerPage

    $candidatePost = if ($config.CandidatePost) { $config.CandidatePost } else { "सरपंच" }
    $candidate = if ($config.CandidateName) { $config.CandidateName } else { "मनोज बाबेल" }
    $party = if ($config.PartyName) { $config.PartyName } else { "भारतीय जनता पार्टी (BJP)" }
    $message = if ($config.BottomMessage) { $config.BottomMessage } else { "को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!" }
    
    # Save base64 image data once to temp files to avoid duplicating megabytes across 1200 slips
    $candidatePhotoUri = Save-Base64ImageToTempFile $config.CandidatePhoto "voter_cand_photo"
    $partySymbolUri = Save-Base64ImageToTempFile $config.PartySymbol "voter_party_symbol"

    $photoHtml = if ($candidatePhotoUri) { "<img class=`"cand-photo`" src=`"$candidatePhotoUri`" alt=`"Candidate`">" } else { '<div class="cand-avatar">&#128100;</div>' }
    $symbolHtml = if ($partySymbolUri) { "<div class=`"cand-symbol-frame`"><img class=`"cand-symbol-img`" src=`"$partySymbolUri`" alt=`"चुनाव चिन्ह`"></div>" } else { "" }

    $totalVoters = $votersList.Count
    $sliceCount = if ($isSampleOnly) { [Math]::Min($slipsPerPage, $totalVoters) } else { $totalVoters }
    $pageCount = [Math]::Ceiling($sliceCount / $slipsPerPage)

    $writer.WriteLine('<!DOCTYPE html>')
    $writer.WriteLine('<html lang="hi">')
    $writer.WriteLine('<head>')
    $writer.WriteLine('  <meta charset="UTF-8">')
    $writer.WriteLine('  <title>&#2357;&#2379;&#2335;&#2352; &#2360;&#2369;&#2357;&#2367;&#2343;&#2366; &#2360;&#2381;&#2354;&#2367;&#2346;</title>')
    $writer.WriteLine('  <style>')
    $writer.WriteLine('    @page { size: A4 portrait; margin: 0; }')
    $writer.WriteLine('    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }')
    $writer.WriteLine('    body { font-family: "Nirmala UI", "Mangal", "Segoe UI", Arial, sans-serif; background: #fff; color: #111; }')
    $writer.WriteLine("    .a4-page { width: 210mm; height: 297mm; padding: $($layout.PagePadding); page-break-after: always; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }")
    $writer.WriteLine("    .slips-grid { display: grid; grid-template-columns: repeat($($layout.Cols), 1fr); grid-template-rows: repeat($($layout.Rows), 1fr); gap: $($layout.GridGap); width: 100%; height: 100%; }")
    $writer.WriteLine('    .slip { border: 1.5px solid #111; border-radius: 5px; padding: 4px 5px; display: flex; flex-direction: row; justify-content: space-between; align-items: stretch; gap: 6px; background: #fff; overflow: hidden; }')
    $writer.WriteLine('    .slip-left { width: 63%; display: flex; flex-direction: column; justify-content: space-between; overflow: hidden; }')
    $writer.WriteLine('    .slip-header-block { border-bottom: 1.2px dashed #444; padding-bottom: 2px; margin-bottom: 1.5px; }')
    $writer.WriteLine("    .slip-title { text-align: center; font-weight: 900; font-size: $($layout.TitleSize); letter-spacing: 0.6px; color: #000; margin-bottom: 1.5px; }")
    $writer.WriteLine("    .meta-row { display: flex; justify-content: space-between; align-items: center; font-size: $($layout.MetaSize); margin-bottom: 1.5px; }")
    $writer.WriteLine('    .meta-row-serial { display: flex; justify-content: flex-start; align-items: center; }')
    $writer.WriteLine('    .badge-item { color: #111; }')
    $writer.WriteLine('    .badge-item strong { color: #000; }')
    $writer.WriteLine("    .serial-badge { background: #f4f5f7; border: 1.3px solid #111; border-radius: 3px; padding: 0.5px 5px; font-size: $($layout.MetaSize); font-weight: 800; display: inline-block; }")
    $writer.WriteLine('    .voter-body { flex: 1; display: flex; flex-direction: column; justify-content: space-evenly; padding: 1px 0; }')
    $writer.WriteLine("    .voter-line { font-size: $($layout.DetailSize); color: #111; line-height: 1.25; }")
    $writer.WriteLine("    .voter-name { font-size: $($layout.NameSize); font-weight: 900; color: #000; }")
    $writer.WriteLine('    .voter-line strong { color: #000; }')
    $writer.WriteLine("    .booth-line { font-size: $($layout.BoothSize); line-height: 1.22; border-top: 1.2px dashed #666; padding-top: 2px; margin-top: 1px; color: #000; background: #fafafa; border-radius: 2px; padding-left: 2px; }")
    $writer.WriteLine('    .booth-label { font-weight: 900; color: #000; }')
    $writer.WriteLine('    .slip-right-box { width: 36%; min-width: 36%; max-width: 37%; border: 1.8px solid #1a237e; border-radius: 6px; background: #fbfbfd; padding: 3px 2px; display: flex; flex-direction: column; align-items: center; justify-content: space-around; text-align: center; box-sizing: border-box; }')
    $writer.WriteLine("    .cand-post { font-size: $($layout.CPostSize); font-weight: 900; color: #b71c1c; line-height: 1.15; margin-bottom: 2px; text-align: center; width: 100%; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; letter-spacing: 0.2px; }")
    $writer.WriteLine("    .cand-photo-frame { width: $($layout.ImgW); height: $($layout.ImgH); border-radius: 4px; border: 1.2px solid #7986cb; background: #e8eaf6; display: flex; align-items: center; justify-content: center; overflow: hidden; flex-shrink: 0; }")
    $writer.WriteLine('    .cand-photo { width: 100%; height: 100%; object-fit: cover; display: block; }')
    $writer.WriteLine("    .cand-avatar { font-size: $($layout.AvatarSize); line-height: 1; }")
    $writer.WriteLine("    .cand-name { font-size: $($layout.CNameSize); font-weight: 900; color: #0d47a1; line-height: 1.15; margin: 1px 0; }")
    $writer.WriteLine("    .cand-symbol-frame { width: $($layout.SymW); height: $($layout.SymH); display: flex; align-items: center; justify-content: center; margin: 1px 0; flex-shrink: 0; }")
    $writer.WriteLine('    .cand-symbol-img { width: 100%; height: 100%; object-fit: contain; }')
    $writer.WriteLine("    .cand-party { font-size: $($layout.CPartySize); font-weight: 800; color: #2e7d32; line-height: 1.15; }")
    $writer.WriteLine("    .cand-appeal { font-size: $($layout.CAppealSize); font-weight: 800; color: #b71c1c; line-height: 1.15; margin-top: 1px; }")
    $writer.WriteLine('  </style>')
    $writer.WriteLine('</head>')
    $writer.WriteLine('<body>')

    $voterIdx = 0
    for ($p = 0; $p -lt $pageCount; $p++) {
        $writer.WriteLine('  <div class="a4-page">')
        $writer.WriteLine('    <div class="slips-grid">')

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

                $postHtml = [System.Security.SecurityElement]::Escape("$candidatePost")
                $cHtml = [System.Security.SecurityElement]::Escape("$candidate")
                $pHtml = [System.Security.SecurityElement]::Escape("$party")
                $mHtml = [System.Security.SecurityElement]::Escape("$message")

                $writer.WriteLine('      <div class="slip">')
                $writer.WriteLine('        <div class="slip-left">')
                $writer.WriteLine('          <div class="slip-header-block">')
                $writer.WriteLine('            <div class="slip-title">&#2357;&#2379;&#2335;&#2352; &#2360;&#2369;&#2357;&#2367;&#2343;&#2366; &#2360;&#2381;&#2354;&#2367;&#2346;</div>')
                $writer.WriteLine('            <div class="meta-row">')
                $writer.WriteLine("              <span class=`"badge-item`">&#2357;&#2366;&#2352;&#2381;&#2337; &#2344;&#2306; : <strong>[ $w ]</strong></span>")
                $writer.WriteLine("              <span class=`"badge-item`">&#2349;&#2366;&#2327; : <strong>[ $part ]</strong></span>")
                $writer.WriteLine('            </div>')
                $writer.WriteLine('            <div class="meta-row-serial">')
                $writer.WriteLine("              <span class=`"badge-item serial-badge`">&#2325;&#2381;&#2352;&#2350; &#2360;&#2306;&#2326;&#2381;&#2351;&#2366; : <strong>[ $sn ]</strong></span>")
                $writer.WriteLine('            </div>')
                $writer.WriteLine('          </div>')
                $writer.WriteLine('          <div class="voter-body">')
                $writer.WriteLine("            <div class=`"voter-line voter-name`">&#2350;&#2340;&#2342;&#2366;&#2340;&#2366; &#2325;&#2366; &#2344;&#2366;&#2350; : <strong>$nm</strong></div>")
                $writer.WriteLine("            <div class=`"voter-line`">&#2346;&#2367;&#2340;&#2366;/&#2346;&#2340;&#2367; &#2325;&#2366; &#2344;&#2366;&#2350; : <span>$rel</span></div>")
                $writer.WriteLine("            <div class=`"voter-line`">&#2350;&#2325;&#2366;&#2344; &#2344;&#2306;. : <strong>$h</strong> &nbsp;|&nbsp; &#2313;&#2350;&#2381;&#2352; : <strong>$age</strong> &nbsp;|&nbsp; &#2354;&#2367;&#2306;&#2327; : <strong>$g</strong></div>")
                $writer.WriteLine("            <div class=`"voter-line`">&#2346;&#2361;&#2331;&#2366;&#2344; &#2346;&#2340;&#2381;&#2352; &#2325;&#2381;&#2352;. (EPIC) : <strong>$epic</strong></div>")
                $writer.WriteLine('          </div>')
                $writer.WriteLine("          <div class=`"booth-line`"><span class=`"booth-label`">&#2350;&#2340;&#2342;&#2366;&#2344; &#2325;&#2375;&#2306;&#2342;&#2381;&#2352; :</span> $booth</div>")
                $writer.WriteLine('        </div>')
                $writer.WriteLine('        <div class="slip-right-box">')
                $writer.WriteLine("          <div class=`"cand-post`"><strong>$postHtml &#2346;&#2342; &#2361;&#2375;&#2340;&#2369;</strong></div>")
                $writer.WriteLine("          <div class=`"cand-photo-frame`">$photoHtml</div>")
                $writer.WriteLine("          <div class=`"cand-name`">$cHtml</div>")
                if ($symbolHtml) {
                    $writer.WriteLine("          $symbolHtml")
                }
                $writer.WriteLine("          <div class=`"cand-party`">($pHtml)</div>")
                $writer.WriteLine("          <div class=`"cand-appeal`">$mHtml</div>")
                $writer.WriteLine('        </div>')
              $writer.WriteLine('      </div>')

                $voterIdx++
            } else {
                $writer.WriteLine('      <div class="slip" style="visibility:hidden;"></div>')
            }
        }

        $writer.WriteLine('    </div>')
        $writer.WriteLine('  </div>')
    }

    $writer.WriteLine('</body>')
    $writer.WriteLine('</html>')
}

function Generate-VoterSlipsHtml($votersList, $config, $isSampleOnly = $false) {
    $sw = New-Object System.IO.StringWriter
    Stream-VoterSlipsHtml $votersList $config $sw $isSampleOnly
    return $sw.ToString()
}

function Get-InstalledBrowserPath {
    # 1. Check Windows Registry App Paths (most accurate across all Windows installations)
    $regKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"
    )
    foreach ($rk in $regKeys) {
        if (Test-Path $rk) {
            $val = (Get-ItemProperty -Path $rk -ErrorAction SilentlyContinue).'(default)'
            if ($val -and (Test-Path $val)) { return $val }
        }
    }

    # 2. Check standard file system paths (32-bit & 64-bit)
    $progFiles = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)
    $progFilesX86 = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFilesX86)
    $localAppData = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)

    $candidates = @(
        "$progFiles\Google\Chrome\Application\chrome.exe",
        "$progFilesX86\Google\Chrome\Application\chrome.exe",
        "$localAppData\Google\Chrome\Application\chrome.exe",
        "$progFiles\Microsoft\Edge\Application\msedge.exe",
        "$progFilesX86\Microsoft\Edge\Application\msedge.exe",
        "$localAppData\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    )
    foreach ($cand in $candidates) {
        if ($cand -and (Test-Path $cand)) {
            return $cand
        }
    }

    # 3. Check PATH environment variable
    $pathCmd = Get-Command "chrome.exe", "msedge.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($pathCmd -and $pathCmd.Source) {
        return $pathCmd.Source
    }

    return $null
}

function Export-VoterSlipsPdf($votersList, $config, $outPdfPath) {
    $browserExe = Get-InstalledBrowserPath

    if (-not $browserExe) {
        throw "Neither Google Chrome nor Microsoft Edge was found for offline PDF generation."
    }

    $outPdfPath = [System.IO.Path]::GetFullPath($outPdfPath)
    $tempHtml = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "voter_slips_print.html")
    
    # Stream HTML directly to disk - zero multi-gigabyte memory buffer
    $fs = [System.IO.File]::Create($tempHtml)
    $writer = New-Object System.IO.StreamWriter($fs, [System.Text.Encoding]::UTF8)
    try {
        Stream-VoterSlipsHtml $votersList $config $writer $false
    } finally {
        $writer.Flush()
        $writer.Close()
        $fs.Close()
    }

    $procArgs = @(
        "--headless",
        "--disable-gpu",
        "--allow-file-access-from-files",
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

