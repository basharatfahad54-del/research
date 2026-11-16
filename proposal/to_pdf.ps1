$ErrorActionPreference = "Stop"
$src = "C:\Users\Muhammad Fahad\Desktop\Research paper\docs\SETILL_PhD_Research_Proposal_UJ_REFINED.docx"
$pdf = "C:\Users\Muhammad Fahad\Desktop\Research paper\docs\SETILL_PhD_Research_Proposal_UJ_REFINED.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Open($src, $false, $false)

# update all fields
try { $doc.Fields.Update() | Out-Null } catch {}
foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
foreach ($tof in $doc.TablesOfFigures) { $tof.Update() }

# refresh headers/footers page counts
try { $doc.Repaginate() } catch {}

$doc.Save()
Start-Sleep -Milliseconds 500

$doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 0, 0, 0, $true, $true, 0, $true, $true, $false)

$pages = $doc.ComputeStatistics(2)
$words = $doc.ComputeStatistics(0)
Write-Output "PAGES=$pages WORDS=$words"

$doc.Close($false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Write-Output "PDF=$pdf"
