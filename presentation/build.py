import os
import sys
import json
import time
import subprocess
import jinja2
import pypdf

# Include scratch root for audit_log import
PRESENTATION_DIR = os.path.dirname(os.path.abspath(__file__))
SCRATCH_ROOT = os.path.dirname(PRESENTATION_DIR)
if SCRATCH_ROOT not in sys.path:
    sys.path.insert(0, SCRATCH_ROOT)

from audit_log import log_generation
DATA_FILE = os.path.join(PRESENTATION_DIR, 'data.json')
TEMPLATE_FILE = os.path.join(PRESENTATION_DIR, 'template.html')
OUT_HTML = os.path.join(PRESENTATION_DIR, 'presentation.html')
OUT_PDF = os.path.join(PRESENTATION_DIR, 'RevOps_Executive_Summary_2026-09-30.pdf')
ARTIFACT_PDF = r'C:\Users\strel\.gemini\antigravity\brain\6a2f6569-0174-4758-bbb9-311ec84e9bf3\RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf'

CHROME_PATHS = [
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe'),
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
]

def find_browser():
    for p in CHROME_PATHS:
        if os.path.exists(p):
            return p
    raise FileNotFoundError("Chrome or Edge browser executable not found!")

def main():
    start_time = time.time()
    try:
        print("=== Step 0: Updating data from Excel & re-rendering charts ===")
        scratch_dir = os.path.dirname(PRESENTATION_DIR)
        try:
            subprocess.run([sys.executable, os.path.join(PRESENTATION_DIR, 'build_data.py')], cwd=scratch_dir, check=True)
            subprocess.run([sys.executable, os.path.join(PRESENTATION_DIR, 'generate_charts.py')], cwd=scratch_dir, check=True)
        except Exception as e:
            print(f"Warning: could not run auto-update step: {e}")

        print("=== Step 1: Loading data.json ===")
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        print(f"Loaded data keys: {list(data.keys())}")

        print("=== Step 2: Rendering Jinja2 template ===")
        with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
            tmpl_content = f.read()

        template = jinja2.Template(tmpl_content)
        rendered_html = template.render(**data)

        with open(OUT_HTML, 'w', encoding='utf-8') as f:
            f.write(rendered_html)
        print(f"Rendered HTML saved to: {OUT_HTML} ({len(rendered_html)} chars)")

        print("=== Step 3: Printing to PDF via Chrome Headless ===")
        browser_bin = find_browser()
        print(f"Using browser: {browser_bin}")

        html_url = f"file:///{OUT_HTML.replace(os.sep, '/')}"
        cmd = [
            browser_bin,
            '--headless=new',
            '--disable-gpu',
            '--no-pdf-header-footer',
            '--allow-file-access-from-files',
            '--run-all-compositor-stages-before-draw',
            '--virtual-time-budget=5000',
            f'--print-to-pdf={OUT_PDF}',
            html_url
        ]
        print(f"Running command: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            print(f"Browser stderr: {result.stderr}")
            raise RuntimeError(f"Browser exited with code {result.returncode}")

        if not os.path.exists(OUT_PDF):
            raise FileNotFoundError(f"PDF was not generated at: {OUT_PDF}")

        pdf_size = os.path.getsize(OUT_PDF)
        print(f"Generated PDF size: {pdf_size:,} bytes")

        print("=== Step 4: Verifying PDF with PyPDF ===")
        reader = pypdf.PdfReader(OUT_PDF)
        page_count = len(reader.pages)
        print(f"Total pages in PDF: {page_count}")
        for idx, page in enumerate(reader.pages, 1):
            box = page.mediabox
            print(f"  Page {idx}: width={box.width:.1f}pt, height={box.height:.1f}pt")

        if page_count != 10:
            print(f"WARNING: Expected 10 pages, got {page_count} pages!")
        else:
            print("PERFECT: Exactly 10 pages generated!")

        print(f"=== Step 5: Copying to Brain Artifacts & Desktop ===")
        import shutil
        shutil.copyfile(OUT_PDF, ARTIFACT_PDF)
        print(f"Copied PDF to: {ARTIFACT_PDF}")

        desktop_pdf = os.path.expanduser(r'~\Desktop\RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf')
        try:
            shutil.copyfile(OUT_PDF, desktop_pdf)
            print(f"SUCCESS: Copied PDF to Desktop -> {desktop_pdf}")
        except Exception as e:
            print(f"Warning: could not copy to Desktop: {e}")

        # Фиксация в audit_log.csv
        duration = time.time() - start_time
        size_kb = pdf_size / 1024.0
        log_generation(
            status="success",
            duration_sec=duration,
            role="ceo",
            report_type="Executive_Summary",
            pdf_size_kb=size_kb
        )

    except Exception as e:
        duration = time.time() - start_time
        log_generation(
            status="error",
            duration_sec=duration,
            role="ceo",
            report_type="Executive_Summary",
            error_message=str(e)
        )
        raise

if __name__ == '__main__':
    main()
