import os
import shutil
import subprocess

docx_orig = os.path.abspath("docs/Каталог_Кастомных_Опций_Интеграции_amoCRM.docx")
temp_docx = os.path.abspath("docs/custom_features_temp.docx")
temp_pdf = os.path.abspath("docs/custom_features_temp.pdf")
final_pdf = os.path.abspath("docs/Каталог_Кастомных_Опций_Интеграции_amoCRM.pdf")
pres_pdf = os.path.abspath("presentation/Каталог_Кастомных_Опций_Интеграции_amoCRM.pdf")

shutil.copyfile(docx_orig, temp_docx)

ps_code = """
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open('""" + temp_docx + """')
    $doc.SaveAs([ref]'""" + temp_pdf + """', [ref]17)
    $doc.Close()
    Write-Host "SUCCESS_CONVERT"
} catch {
    Write-Host "ERROR: $_"
} finally {
    $word.Quit()
}
"""

ps_path = os.path.abspath("scripts/run_custom_pdf_conv.ps1")
with open(ps_path, "w", encoding="utf-8") as f:
    f.write(ps_code)

res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_path], capture_output=True, text=True)
print("STDOUT:", res.stdout.strip())
print("STDERR:", res.stderr.strip())

if os.path.exists(temp_pdf):
    shutil.copyfile(temp_pdf, final_pdf)
    shutil.copyfile(temp_pdf, pres_pdf)
    if os.path.exists(temp_docx):
        os.remove(temp_docx)
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
    print(f"SUCCESS: Generated {final_pdf} ({os.path.getsize(final_pdf):,} bytes)")
    print(f"SUCCESS: Copied to {pres_pdf}")
else:
    print("FAILED to generate PDF")
