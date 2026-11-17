# Dumps TOC entries (level, text, page) from the compiled thesis.
$src = "C:\Users\Muhammad Fahad\Desktop\research-main\research-main\johny\Final_Thesis\Complete_Thesis.docx"
$out = "C:\Users\Muhammad Fahad\Desktop\research-main\research-main\johny\thesis_build\toc_entries.txt"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($src)
  $null = $doc.Fields.Update()
  foreach ($toc in @($doc.TablesOfContents)) { $null = $toc.Update() }
  $doc.Save()
  $lines = @()
  foreach ($p in $doc.TablesOfContents.Item(1).Range.Paragraphs) {
    $t = $p.Range.Text.Trim([char]13, [char]7)
    if (-not $t.Trim()) { continue }
    $style = ""
    try { $style = $p.Style.NameLocal } catch { $style = "TOC" }
    $parts = $t -split "`t"
    $page = $parts[$parts.Count - 1].Trim()
    $label = ""
    if ($parts.Count -gt 1) { $label = ($parts[0..($parts.Count - 2)] -join " ").Trim() } else { $label = $t.Trim() }
    $lines += ($style + "<<>>" + $label + "<<>>" + $page)
  }
  [System.IO.File]::WriteAllLines($out, $lines)
  $doc.Close(0)
} finally {
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host ("ENTRIES: " + $lines.Count)
