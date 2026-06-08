$sourcePath = 'c:\Users\Administrator\resume.md'
$outputPath = 'c:\Users\Administrator\resume.pdf'
$content = Get-Content -Path $sourcePath -Raw
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$range = $doc.Content
$range.Text = $content
$doc.SaveAs2($outputPath, 17)
$doc.Close($false)
$word.Quit()
Write-Host "PDF exported to $outputPath"
