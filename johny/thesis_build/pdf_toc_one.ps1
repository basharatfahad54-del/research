$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open("C:\Users\Muhammad Fahad\Desktop\research-main\research-main\johny\Final_Thesis\Table_of_Contents.docx")
  $doc.ExportAsFixedFormat("C:\Users\Muhammad Fahad\Desktop\research-main\research-main\johny\Final_Thesis\Table_of_Contents.pdf", 17)
  $doc.Close(0)
} finally {
  $word.Quit()
}
Write-Host "TOC PDF OK"
