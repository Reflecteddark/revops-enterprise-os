import subprocess
import os
import pymupdf

browser = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
html_path = os.path.abspath("pilot-roadmap.html")
pdf_path = os.path.abspath("test_roadmap.pdf")
html_url = "file:///" + html_path.replace("\\", "/")

cmd = [
    browser,
    "--headless=new",
    "--disable-gpu",
    "--no-pdf-header-footer",
    "--allow-file-access-from-files",
    "--run-all-compositor-stages-before-draw",
    "--virtual-time-budget=5000",
    f"--print-to-pdf={pdf_path}",
    html_url
]

print("Running command...")
res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
print("Return code:", res.returncode)

if os.path.exists(pdf_path):
    print(f"PDF created: {pdf_path} ({os.path.getsize(pdf_path)} bytes)")
    doc = pymupdf.open(pdf_path)
    print("Pages:", len(doc))
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    pix.save("test_roadmap.png")
    print("Saved test_roadmap.png")
else:
    print("PDF NOT created. Stderr:", res.stderr)
