
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\full_manual_temp.docx')
    $doc.SaveAs([ref]'C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\full_manual_temp.pdf', [ref]17)
    $doc.Close()
    Write-Host "SUCCESS_CONVERT"
} catch {
    Write-Host "ERROR: $_"
} finally {
    $word.Quit()
}
