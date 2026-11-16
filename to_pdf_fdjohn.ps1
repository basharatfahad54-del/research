$ErrorActionPreference = 'Stop'
$log = Join-Path $PSScriptRoot 'pdf_convert.log'
Set-Location $PSScriptRoot
$src = Join-Path $PSScriptRoot 'fd john\Client_Requirements_Specification.docx'
$dst = Join-Path $PSScriptRoot 'fd john\Client_Requirements_Specification.pdf'

function Log($m) { "$(Get-Date -Format HH:mm:ss)  $m" | Tee-Object -FilePath $log -Append }

Log "start"
$word = New-Object -ComObject Word.Application
Log "word created"
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3   # force-disable macros
try {
  Log "opening doc"
  $doc = $word.Documents.Open($src, $false, $true, $false)
  Log "doc opened; updating fields"
  $doc.Fields.Update() | Out-Null
  Log "fields updated; exporting"
  $doc.ExportAsFixedFormat($dst, 17, $false, 0, 0, 0, 0, 1, $true, $false, 0, $false, $false, $false)
  Log "export done"
  $doc.Close($false)
  Log "doc closed"
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  Log "word quit"
}
Log "SUCCESS"
