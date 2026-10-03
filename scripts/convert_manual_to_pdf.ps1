
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open("C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Руководство_Клиента_RevOps_OS_V18.docx")
    $doc.SaveAs([ref]"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Руководство_Клиента_RevOps_OS_V18.pdf", [ref]17)
    $doc.Close()
    Write-Host "SUCCESS: PDF created at C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Руководство_Клиента_RevOps_OS_V18.pdf"
} catch {
    Write-Host "ERROR: $_"
} finally {
    $word.Quit()
}
