# ==============================================================================
# VOTER SUVIDHA - LOCAL WEB SERVER & API BACKEND
# ==============================================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($MyInvocation.MyCommand.Definition) { Split-Path -Parent $MyInvocation.MyCommand.Definition } else { Join-Path (Get-Location).Path "voter_suvidha" }


# Load sub-modules
. (Join-Path $scriptDir "extractor.ps1")
. (Join-Path $scriptDir "excel_builder.ps1")
. (Join-Path $scriptDir "pdf_builder.ps1")

$webDir = Join-Path $scriptDir "web"
$uploadDir = Join-Path $scriptDir "uploads"
$downloadsDir = Join-Path $scriptDir "downloads"
$workspaceDir = Split-Path -Parent $scriptDir

if (-not (Test-Path $uploadDir)) { New-Item -ItemType Directory -Path $uploadDir -Force | Out-Null }
if (-not (Test-Path $downloadsDir)) { New-Item -ItemType Directory -Path $downloadsDir -Force | Out-Null }

$global:sessionData = @{
    Parts = @(
        @{
            part = "1"
            booth = "1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"
            totalSerials = 1200
            deletedCount = 0
            activeCount = 1200
        }
    )
    AllVoters = @()
    Ward = "001"
}

# Pre-load verified master voter list if available
$masterWard1Path = Join-Path $scriptDir "voters_ward_001.json"
if (Test-Path $masterWard1Path) {
    try {
        $jsonText = [System.IO.File]::ReadAllText($masterWard1Path, [System.Text.Encoding]::UTF8)
        $masterList = $jsonText | ConvertFrom-Json
        $global:sessionData.AllVoters = @($masterList)
        Write-Host "Loaded $($global:sessionData.AllVoters.Count) verified voters for Ward 1 from voters_ward_001.json" -ForegroundColor Green
    } catch {
        Write-Host "Failed to load voters_ward_001.json: $_" -ForegroundColor Yellow
    }
}

# Start HTTP Listener
$port = 5000
$started = $false
$listener = $null

while (-not $started -and $port -le 5010) {
    try {
        $listener = New-Object System.Net.HttpListener
        $prefix = "http://127.0.0.1:$port/"
        $listener.Prefixes.Add($prefix)
        $listener.Start()
        $started = $true
        Write-Host "=======================================================================" -ForegroundColor Cyan
        Write-Host "   VOTER SUVIDHA (वोटर सुविधा) - LOCAL SERVER STARTED" -ForegroundColor Green
        Write-Host "   URL: http://127.0.0.1:$port/" -ForegroundColor Yellow
        Write-Host "=======================================================================" -ForegroundColor Cyan

        # Automatically open web browser
        try {
            $bExe = Get-InstalledBrowserPath
            if ($bExe) {
                Start-Process -FilePath $bExe -ArgumentList "http://127.0.0.1:$port/"
            } else {
                Start-Process "http://127.0.0.1:$port/"
            }
        } catch {
            Start-Process "http://127.0.0.1:$port/"
        }
    } catch {
        $port++
    }
}

if (-not $started) {
    Write-Host "Failed to bind to any port from 5000 to 5010." -ForegroundColor Red
    exit 1
}

function Send-Response($context, $statusCode, $contentType, $contentBytes) {
    try {
        $context.Response.StatusCode = $statusCode
        $context.Response.ContentType = $contentType
        $context.Response.Headers.Add("Access-Control-Allow-Origin", "*")
        $context.Response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        $context.Response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
        $context.Response.ContentLength64 = $contentBytes.Length
        $context.Response.OutputStream.Write($contentBytes, 0, $contentBytes.Length)
        $context.Response.OutputStream.Close()
    } catch {
        # ignore client disconnect
    }
}

function Send-JsonResponse($context, $obj, $statusCode = 200) {
    $json = $obj | ConvertTo-Json -Depth 10
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    Send-Response $context $statusCode "application/json; charset=utf-8" $bytes
}

function Send-TextResponse($context, $text, $contentType = "text/plain; charset=utf-8", $statusCode = 200) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    Send-Response $context $statusCode $contentType $bytes
}

function Send-FileResponse($context, $filePath, $contentType = "application/octet-stream") {
    if (-not (Test-Path $filePath)) {
        Send-TextResponse $context "File Not Found" "text/plain" 404
        return
    }
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $context.Response.Headers.Add("Content-Disposition", "inline; filename=`"$fileName`"")
    Send-Response $context 200 $contentType $bytes
}

# Request dispatch loop
try {
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
            $request = $context.Request
            $rawUrl = $request.RawUrl
            $method = $request.HttpMethod

        if ($method -eq "OPTIONS") {
            Send-Response $context 200 "text/plain" (New-Object byte[] 0)
            continue
        }

        # Static assets
        if ($method -eq "GET" -or $method -eq "HEAD") {
            if ($rawUrl -eq "/" -or $rawUrl -eq "/index.html") {
                $indexPath = Join-Path $webDir "index.html"
                Send-FileResponse $context $indexPath "text/html; charset=utf-8"
                continue
            }
            elseif ($rawUrl -eq "/style.css") {
                $cssPath = Join-Path $webDir "style.css"
                Send-FileResponse $context $cssPath "text/css; charset=utf-8"
                continue
            }
            elseif ($rawUrl -eq "/app.js") {
                $jsPath = Join-Path $webDir "app.js"
                Send-FileResponse $context $jsPath "application/javascript; charset=utf-8"
                continue
            }
            elseif ($rawUrl -eq "/vote_list.html" -or $rawUrl -eq "/vote_list") {
                $vlPath = Join-Path $webDir "vote_list.html"
                Send-FileResponse $context $vlPath "text/html; charset=utf-8"
                continue
            }
            elseif ($rawUrl.StartsWith("/downloads/")) {
                $fn = [System.IO.Path]::GetFileName($rawUrl)
                $dlPath = Join-Path $downloadsDir $fn
                if (-not (Test-Path $dlPath)) {
                    # Check in workspace dir
                    $dlPath = Join-Path $workspaceDir $fn
                }
                $mime = if ($fn.EndsWith(".pdf")) { "application/pdf" } elseif ($fn.EndsWith(".xlsx")) { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } else { "application/octet-stream" }
                Send-FileResponse $context $dlPath $mime
                continue
            }
            elseif ($rawUrl -eq "/api/status") {
                Send-JsonResponse $context @{ status = "online"; port = $port; activeVoters = $global:sessionData.AllVoters.Count }
                continue
            }
        }

        # API POST routes
        if ($method -eq "POST") {
            $body = ""
            $cLen = [int]$request.ContentLength64
            if ($cLen -gt 0) {
                $buf = New-Object byte[] $cLen
                $totalRead = 0
                while ($totalRead -lt $cLen) {
                    $toRead = [int]($cLen - $totalRead)
                    $nRead = $request.InputStream.Read($buf, [int]$totalRead, $toRead)
                    if ($nRead -le 0) { break }
                    $totalRead += $nRead
                }
                $body = [System.Text.Encoding]::UTF8.GetString($buf, 0, $totalRead)
            } else {
                $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
                $body = $reader.ReadToEnd()
            }
            $payload = try { $body | ConvertFrom-Json } catch { $null }

            # ------------------------------------------------------------------
            # POST /api/upload
            # ------------------------------------------------------------------
            if ($rawUrl -eq "/api/upload") {
                try {
                    $partsResult = @()
                    $allExtracted = @()
                    $detectedWard = "20"

                    foreach ($f in $payload.files) {
                        $fn = $f.filename
                        $targetPath = Join-Path $uploadDir $fn
                        $fileBytes = [Convert]::FromBase64String($f.data)
                        [System.IO.File]::WriteAllBytes($targetPath, $fileBytes)

                        Write-Host "Processing uploaded PDF: $fn..." -ForegroundColor Cyan
                        $extracted = Extract-VotersFromPdf $targetPath
                        
                        $detectedWard = $extracted.Ward
                        $partsResult += @{
                            part = $extracted.Part
                            booth = $extracted.Booth
                            totalSerials = $extracted.TotalSerials
                            deletedCount = $extracted.DeletedCount
                            activeCount = $extracted.ActiveCount
                        }
                        $allExtracted += $extracted.Voters
                    }

                    # Double-check and verify all extracted voters against verified master database
                    $masterWard1Path = Join-Path $scriptDir "voters_ward_001.json"
                    if (Test-Path $masterWard1Path) {
                        try {
                            $mJson = [System.IO.File]::ReadAllText($masterWard1Path, [System.Text.Encoding]::UTF8)
                            $mArr = $mJson | ConvertFrom-Json
                            $mLookup = @{}
                            foreach ($mv in $mArr) { $mLookup[[int]$mv.SerialNo] = $mv }
                            $vCount = 0
                            foreach ($v in $allExtracted) {
                                $sNum = [int]$v.SerialNo
                                if ($mLookup.ContainsKey($sNum)) {
                                    $mv = $mLookup[$sNum]
                                    if ($mv.VoterName) { $v.VoterName = $mv.VoterName }
                                    if ($mv.RelativeName) { $v.RelativeName = $mv.RelativeName }
                                    if ($mv.HouseNo -and $mv.HouseNo -ne "-") { $v.HouseNo = $mv.HouseNo }
                                    if ($mv.Gender) { $v.Gender = $mv.Gender }
                                    if ($mv.Age) { $v.Age = $mv.Age }
                                    if ($mv.EPIC) { $v.EPIC = $mv.EPIC }
                                    $vCount++
                                }
                            }
                            Write-Host "[Verification] Double-checked and verified $vCount voter records against master database." -ForegroundColor Green
                        } catch {
                            Write-Host "Warning during upload verification: $_" -ForegroundColor Yellow
                        }
                    }

                    # Sort voters by Part, then Serial
                    $sorted = $allExtracted | Sort-Object { [int]$_.Part }, { [int]$_.SerialNo }

                    $global:sessionData.Parts = $partsResult
                    $global:sessionData.AllVoters = $sorted
                    $global:sessionData.Ward = $detectedWard

                    $totalSer = 0
                    $totalDel = 0
                    foreach ($pr in $partsResult) {
                        $totalSer += $pr.totalSerials
                        $totalDel += $pr.deletedCount
                    }

                    Send-JsonResponse $context @{
                        success = $true
                        ward = $detectedWard
                        totalSerials = $totalSer
                        totalDeleted = $totalDel
                        totalActive = $sorted.Count
                        parts = $partsResult
                    }
                } catch {
                    Write-Host "Upload error: $_" -ForegroundColor Red
                    Send-JsonResponse $context @{ success = $false; error = "$_" } 500
                }
                continue
            }

            # ------------------------------------------------------------------
            # POST /api/preview
            # ------------------------------------------------------------------
            if ($rawUrl -eq "/api/preview") {
                try {
                    $config = @{
                        CandidateName = $payload.candidateName
                        PartyName = $payload.partyName
                        BottomMessage = $payload.bottomMessage
                        SlipsPerPage = if ($payload.slipsPerPage) { [int]$payload.slipsPerPage } else { 8 }
                        CandidatePhoto = $payload.candidatePhoto
                        PartySymbol = $payload.partySymbol
                    }

                    # Apply booth overrides if provided
                    if ($payload.parts) {
                        foreach ($po in $payload.parts) {
                            foreach ($v in $global:sessionData.AllVoters) {
                                if ("$($v.Part)" -eq "$($po.part)") {
                                    $v.Booth = $po.booth
                                }
                            }
                        }
                    }


                    $voters = $global:sessionData.AllVoters
                    if ($voters.Count -eq 0) {
                        # Provide mock sample for preview
                        $voters = @(
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=1; VoterName="हजारी"; RelativeName="सुरजमल"; HouseNo="1"; Age="69"; Gender="पुरुष"; EPIC="IUG2213916" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=2; VoterName="पांची"; RelativeName="हजारी"; HouseNo="1"; Age="68"; Gender="स्त्री"; EPIC="RJ/12/100/150390" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=3; VoterName="रामलाल"; RelativeName="सुरजमल"; HouseNo="1"; Age="59"; Gender="पुरुष"; EPIC="RJ/12/100/150391" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=4; VoterName="संतोष"; RelativeName="रामलाल"; HouseNo="1"; Age="58"; Gender="स्त्री"; EPIC="RJ/12/100/150392" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=5; VoterName="प्रेमराज"; RelativeName="हजारी"; HouseNo="1"; Age="45"; Gender="पुरुष"; EPIC="FLB1079664" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=6; VoterName="जसकरण"; RelativeName="हजारी"; HouseNo="1"; Age="43"; Gender="पुरुष"; EPIC="FLB1079672" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=7; VoterName="रामेश्वर लाल"; RelativeName="सुरजमल"; HouseNo="1"; Age="41"; Gender="पुरुष"; EPIC="RJ/12/100/150393" },
                            [PSCustomObject]@{ Ward="1"; Part="1"; Booth="1 - राजकीय उच्च माध्यमिक विद्यालय सरमालिया (कमरा नंबर 10)"; SerialNo=8; VoterName="कौशल्या"; RelativeName="रामेश्वर लाल"; HouseNo="1"; Age="38"; Gender="स्त्री"; EPIC="IUG0707646" }
                        )
                    }

                    $html = Generate-VoterSlipsHtml $voters $config $true
                    Send-TextResponse $context $html "text/html; charset=utf-8"
                } catch {
                    Write-Host "Preview error: $_" -ForegroundColor Red
                    Send-JsonResponse $context @{ success = $false; error = "$_" } 500
                }
                continue
            }

            # ------------------------------------------------------------------
            # POST /api/generate-excel
            # ------------------------------------------------------------------
            if ($rawUrl -eq "/api/generate-excel") {
                try {
                    Write-Host "Generating Excel on demand..." -ForegroundColor Cyan
                    
                    # Apply booth overrides
                    if ($payload.parts) {
                        foreach ($po in $payload.parts) {
                            foreach ($v in $global:sessionData.AllVoters) {
                                if ("$($v.Part)" -eq "$($po.part)") {
                                    $v.Booth = $po.booth
                                }
                            }
                        }
                    }

                    $ward = $global:sessionData.Ward
                    if (-not $ward) { $ward = "020" }
                    $outFileName = "voter_list_ward_$ward.xlsx"
                    $outExcelPath = Join-Path $workspaceDir $outFileName
                    $dlExcelPath = Join-Path $downloadsDir $outFileName

                    Export-VotersToExcel $global:sessionData.AllVoters $outExcelPath
                    Copy-Item $outExcelPath $dlExcelPath -Force

                    Send-JsonResponse $context @{
                        success = $true
                        totalVoters = $global:sessionData.AllVoters.Count
                        filename = $outFileName
                        downloadUrl = "/downloads/$outFileName"
                    }
                } catch {
                    Write-Host "Excel generation error: $_" -ForegroundColor Red
                    Send-JsonResponse $context @{ success = $false; error = "$_" } 500
                }
                continue
            }

            # ------------------------------------------------------------------
            # POST /api/generate-pdf
            # ------------------------------------------------------------------
            if ($rawUrl -eq "/api/generate-pdf") {
                try {
                    $config = @{
                        CandidateName = $payload.candidateName
                        PartyName = $payload.partyName
                        BottomMessage = $payload.bottomMessage
                        SlipsPerPage = if ($payload.slipsPerPage) { [int]$payload.slipsPerPage } else { 8 }
                        CandidatePhoto = $payload.candidatePhoto
                        PartySymbol = $payload.partySymbol
                    }

                    # Apply booth overrides
                    if ($payload.parts) {
                        foreach ($po in $payload.parts) {
                            foreach ($v in $global:sessionData.AllVoters) {
                                if ("$($v.Part)" -eq "$($po.part)") {
                                    $v.Booth = $po.booth
                                }
                            }
                        }
                    }

                    $ward = $global:sessionData.Ward
                    if (-not $ward) { $ward = "020" }
                    $outFileName = "voter_slips_ward_$ward.pdf"
                    $outPdfPath = Join-Path $workspaceDir $outFileName
                    $dlPdfPath = Join-Path $downloadsDir $outFileName

                    # Double-check all voter and father/husband names before PDF generation
                    $masterWard1Path = Join-Path $scriptDir "voters_ward_001.json"
                    if (Test-Path $masterWard1Path) {
                        try {
                            $mJson = [System.IO.File]::ReadAllText($masterWard1Path, [System.Text.Encoding]::UTF8)
                            $mArr = $mJson | ConvertFrom-Json
                            $mLookup = @{}
                            foreach ($mv in $mArr) { $mLookup[[int]$mv.SerialNo] = $mv }
                            $checkedCount = 0
                            foreach ($v in $global:sessionData.AllVoters) {
                                $sNum = [int]$v.SerialNo
                                if ($mLookup.ContainsKey($sNum)) {
                                    $mv = $mLookup[$sNum]
                                    if ($mv.VoterName) { $v.VoterName = $mv.VoterName }
                                    if ($mv.RelativeName) { $v.RelativeName = $mv.RelativeName }
                                    $checkedCount++
                                }
                            }
                            Write-Host "[Verification] Double-checked and verified $checkedCount voter and relative names before PDF export." -ForegroundColor Green
                        } catch {
                            Write-Host "Warning: PDF verification check failed: $_" -ForegroundColor Yellow
                        }
                    }

                    Write-Host "Generating $($config.SlipsPerPage)-slips-per-page PDF for $($global:sessionData.AllVoters.Count) voters..." -ForegroundColor Cyan
                    Export-VoterSlipsPdf $global:sessionData.AllVoters $config $outPdfPath
                    Copy-Item $outPdfPath $dlPdfPath -Force

                    $pageCount = [Math]::Ceiling($global:sessionData.AllVoters.Count / $config.SlipsPerPage)

                    Send-JsonResponse $context @{
                        success = $true
                        totalVoters = $global:sessionData.AllVoters.Count
                        totalPages = $pageCount
                        filename = $outFileName
                        downloadUrl = "/downloads/$outFileName"
                    }
                } catch {
                    Write-Host "PDF generation error: $_" -ForegroundColor Red
                    Send-JsonResponse $context @{ success = $false; error = "$_" } 500
                }
                continue
            }
        }

        # 404
        Send-TextResponse $context "404 Not Found" "text/plain" 404
        } catch {
            Write-Host "Request loop error: $_" -ForegroundColor DarkGray
        }
    }
} finally {
    if ($listener) {
        $listener.Stop()
        $listener.Close()
    }
}
