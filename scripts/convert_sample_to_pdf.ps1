$docxPath = "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\samples\RevOps_Sample_Audit_Report.docx"
$pdfPath = "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\docs\samples\RevOps_Sample_Audit_Report.pdf"
$pdfPathRoot = "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os\samples\RevOps_Sample_Audit_Report.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docxPath)
$wdFormatPDF = 17
$doc.SaveAs($pdfPath, $wdFormatPDF)
$doc.Close()
$word.Quit()

Copy-Item $pdfPath $pdfPathRoot -Force
Write-Host "PDF generated successfully at $pdfPath"
