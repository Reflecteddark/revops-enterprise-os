$w = New-Object -ComObject Word.Application
$v = $w.Version
$w.Quit()
Write-Host "MS Word is ready! Version: $v"
