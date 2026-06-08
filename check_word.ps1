try {
    $w = New-Object -ComObject Word.Application
    $w.Quit()
    Write-Host 'WORD_OK'
} catch {
    Write-Host 'WORD_FAIL'
}
