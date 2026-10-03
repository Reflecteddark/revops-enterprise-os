$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $docxPath = 'C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Полное_Постраничное_Руководство_RevOps_OS_V18.docx'
    $pdfPath = 'C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\Полное_Постраничное_Руководство_RevOps_OS_V18.pdf'
    $doc = $word.Documents.Open($docxPath)
    $doc.SaveAs([ref]$pdfPath, [ref]17)
    $doc.Close()
    Write-Host "SUCCESS_CONVERT_PDF"
    
    $presPdf = 'C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\presentation\Полное_Постраничное_Руководство_RevOps_OS_V18.pdf'
    Copy-Item $pdfPath $presPdf -Force
    Write-Host "SUCCESS_COPIED_TO_PRESENTATION"
} catch {
    Write-Host "ERROR: $_"
} finally {
    $word.Quit()
}
