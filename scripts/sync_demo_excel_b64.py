"""
Syncs the latest RevOps_Platform_Demo_Sample.xlsx into the embedded Base64 constant
in presentation/landing.html, web/index.html, and Desktop landing page.
"""
import base64
import re
from pathlib import Path

def sync_base64():
    xlsx_path = Path("web/RevOps_Platform_Demo_Sample.xlsx")
    with open(xlsx_path, "rb") as f:
        b64_str = base64.b64encode(f.read()).decode("utf-8")

    print(f"Read {xlsx_path} ({xlsx_path.stat().st_size} bytes), Base64 length: {len(b64_str)}")

    targets = [
        Path("presentation/landing.html"),
        Path("web/index.html"),
        Path(r"C:\Users\strel\Desktop\RevOps Platform\Сайт\RevOps_AI_Landing_Page.html")
    ]

    pattern = re.compile(r'const DEMO_EXCEL_B64 = "([^"]+)";')

    for target in targets:
        if not target.exists():
            print(f"Skipping non-existent: {target}")
            continue

        content = target.read_text(encoding="utf-8")
        if pattern.search(content):
            new_content = pattern.sub(f'const DEMO_EXCEL_B64 = "{b64_str}";', content)
            target.write_text(new_content, encoding="utf-8")
            print(f"Updated {target}")
        else:
            print(f"WARNING: DEMO_EXCEL_B64 constant not found in {target}")

if __name__ == "__main__":
    sync_base64()
