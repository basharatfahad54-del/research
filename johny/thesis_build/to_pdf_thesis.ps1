# Exports every DOCX in Final_Thesis to PDF, updating fields/TOCs first.
$dir = "C:\Users\Muhammad Fahad\Desktop\research-main\research-main\johny\Final_Thesis"
$files = Get-ChildItem -Path $dir -Filter *.docx | Where-Object { $_.Name -ne "~$*" }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  foreach ($f in $files) {
    $path = $f.FullName
    Write-Host ("PDF: " + $f.Name)
    $doc = $word.Documents.Open($path)
    try {
      $null = $doc.Fields.Update()
      foreach ($toc in @($doc.TablesOfContents)) { $null = $toc.Update() }
      $doc.Save()
    } catch {
      Write-Host ("  field update warning: " + $_.Exception.Message)
    }
    $pdf = [System.IO.Path]::ChangeExtension($path, ".pdf")
    $doc.ExportAsFixedFormat($pdf, 17)
    Write-Host ("  exported: " + (Split-Path $pdf -Leaf))
    $doc.Close(0)
  }
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host "DONE"
