import os
import glob
import shutil
import subprocess
import sys
import time
from datetime import datetime
import openpyxl
from audit_log import log_generation
from data_validator import validate_data

start_time = time.time()
now = datetime.now()
gen_date_str = now.strftime("%d.%m.%Y")
gen_time_str = now.strftime("%H:%M")
timestamp_file_str = now.strftime("%Y-%m-%d_%H%M")

sys.stdout.reconfigure(encoding='utf-8')

candidates = [f for f in glob.glob('RevOps Platform V17*.xlsx') if not os.path.basename(f).startswith('~$')]
if not candidates:
    print("[ОШИБКА] Не найден файл RevOps Platform V17*.xlsx в папке")
    sys.exit(1)
wb_path = max(candidates, key=os.path.getmtime)
print(f"Используем файл: {wb_path}")

errors = validate_data(wb_path)
if errors:
    print("[ОШИБКА ВАЛИДАЦИИ] Данные не прошли проверку:")
    for err in errors:
        print(f"  • {err}")
    log_generation(
        status="error",
        duration_sec=time.time() - start_time,
        role="full",
        report_type="Role_Report",
        error_message="Validation failed: " + "; ".join(errors[:3])
    )
    sys.exit(1)

wb = openpyxl.load_workbook(wb_path, data_only=True)

# 15 Showcase Sheets to include in Full Report (Strictly client-facing)
showcase_sheets = [
    '⚙️ Настройки',
    '📄 Executive_OnePager',
    '⚡ Пульс_Компании',
    '📋 Пульт_РОПа_15_Минут',
    '💸 Диагностика_Утечек_ОП',
    '⚡ Экспресс_Калькулятор_3_Цифры',
    '🎙️ ИИ_Аудит',
    '🌐 Мультиканальная_Атрибуция',
    '💳 Финансы_и_AI_Дожим',
    '👥 Мотивация_ОП',
    '🔮 Симулятор_Роста',
    '🎯 Воронка_и_SLA',
    '📊 Юнит_Экономика',
    '🚨 Радар_Алертов',
    '👥 Ресурсный_План'
]

print(f"Generating Full Report for {len(showcase_sheets)} showcase sheets...")

# Helper to format numbers cleanly
def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if 0 < v < 1:
            return f"{v*100:.1f}%"
        if v.is_integer():
            return f"{int(v):,}".replace(",", " ")
        return f"{v:,.2f}".replace(",", " ")
    if isinstance(v, int):
        return f"{v:,}".replace(",", " ")
    return str(v)

slides_html = []

for sname in showcase_sheets:
    if sname not in wb.sheetnames:
        print(f"Warning: {sname} not in workbook")
        continue
    ws = wb[sname]
    
    # Read rows
    max_r = min(ws.max_row, 35)
    max_c = min(ws.max_column, 16)
    
    rows_data = []
    for r in range(1, max_r + 1):
        row_cells = [ws.cell(r, c).value for c in range(1, max_c + 1)]
        # Skip empty rows at the end
        if any(c is not None for c in row_cells):
            rows_data.append(row_cells)
            
    # Title from row 1 or sheet name
    title = sname
    if rows_data and rows_data[0][0]:
        title = str(rows_data[0][0])
        
    table_rows = []
    # Render from row 3 onwards (or row 2 if row 2 has data)
    start_r = 1 if len(rows_data) > 1 and rows_data[1][0] else 2
    for r_idx in range(start_r, min(len(rows_data), 26)):
        r = rows_data[r_idx]
        cells_html = []
        is_header = r_idx in [1, 2, 6, 7]
        for c in r[:10]: # limit to 10 columns for print width
            val_str = fmt_val(c)
            # styling
            bg = "#F1F5F9" if is_header else ("#FFFFFF" if r_idx % 2 == 0 else "#F8FAFC")
            fw = "700" if is_header else "400"
            cells_html.append(f'<td style="background:{bg}; font-weight:{fw}; padding: 4px 6px; border-bottom: 1px solid #E2E8F0; font-size: 8.5px; font-variant-numeric: tabular-nums;">{val_str}</td>')
        table_rows.append(f'<tr>{"".join(cells_html)}</tr>')

    render_duration = round(time.time() - start_time, 1)
    slide_html = f"""
    <section class="slide">
      <div class="slide-header">
        <div class="slide-header-left">
          <span class="brand-tag">RevOps Enterprise OS</span>
          <span class="version-pill">V17.6 FULL SUITE</span>
          <h2 class="slide-title">{title}</h2>
        </div>
        <div class="slide-meta">Лист: {sname} • Статус: Enterprise Validated</div>
      </div>
      <div class="table-container" style="flex: 1; overflow: hidden; background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px;">
        <table style="width: 100%; border-collapse: collapse;">
          <tbody>
            {"".join(table_rows)}
          </tbody>
        </table>
      </div>
      <div class="slide-footer">
        <span>Данные на {gen_date_str} {gen_time_str} MSK • Сгенерировано за {render_duration} сек</span>
        <span>Раздел витрин (Лист: {sname}) • RevOps Platform V17.6 Production Suite</span>
      </div>
    </section>
    """
    slides_html.append(slide_html)

full_html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <title>RevOps Enterprise OS V17.6 — Full Report</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:ital,wght@0,600;0,700;1,600&display=swap" rel="stylesheet">
  <style>
    @page {{
      size: 297mm 210mm;
      margin: 0;
    }}
    *, *::before, *::after {{
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
    }}
    body {{
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: #F8FAFC;
      color: #0F172A;
      font-size: 10px;
    }}
    .slide {{
      width: 297mm;
      height: 210mm;
      max-width: 297mm;
      max-height: 210mm;
      page-break-after: always;
      page-break-inside: avoid;
      position: relative;
      overflow: hidden;
      background: #FFFFFF;
      display: flex;
      flex-direction: column;
      padding: 14mm 16mm;
    }}
    .slide-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      border-bottom: 2px solid #E2E8F0;
      padding-bottom: 8px;
      margin-bottom: 10px;
    }}
    .slide-header-left {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .brand-tag {{
      font-family: 'Playfair Display', serif;
      font-size: 14px;
      font-weight: 700;
      color: #0F172A;
    }}
    .version-pill {{
      background: #EEF2FF;
      color: #4F46E5;
      font-weight: 700;
      font-size: 9.5px;
      padding: 2px 7px;
      border-radius: 9999px;
      border: 1px solid #C7D2FE;
    }}
    .slide-title {{
      font-family: 'Playfair Display', serif;
      font-size: 15px;
      font-weight: 700;
      color: #0F172A;
    }}
    .slide-meta {{
      font-size: 9.5px;
      color: #64748B;
    }}
    .slide-footer {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-top: 8px;
      padding-top: 5px;
      border-top: 1px solid #E2E8F0;
      font-size: 10px;
      color: #64748B;
    }}
  </style>
</head>
<body>
  {"".join(slides_html)}
</body>
</html>
"""

full_html_path = os.path.abspath("RevOps_Enterprise_Full_Report.html")
with open(full_html_path, "w", encoding="utf-8") as f:
    f.write(full_html)
print(f"Generated {full_html_path}")

def find_chrome() -> str:
    if os.environ.get("MOCK_NO_CHROME") == "1":
        print("[ОШИБКА] Chrome не найден по стандартным путям")
        sys.exit(1)
    for cmd_name in ['chrome', 'google-chrome']:
        p = shutil.which(cmd_name)
        if p and os.path.exists(p):
            return p
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    localappdata = os.environ.get('LOCALAPPDATA')
    if localappdata:
        candidates.append(os.path.join(localappdata, r"Google\Chrome\Application\chrome.exe"))
    for p in candidates:
        if os.path.exists(p):
            return p
    print("[ОШИБКА] Chrome не найден по стандартным путям")
    sys.exit(1)

pdf_filename = f"RevOps_Enterprise_OS_V17.6_Full_Report_{timestamp_file_str}.pdf"
pdf_path = os.path.abspath(pdf_filename)
brain_pdf = os.path.join(r"C:\Users\strel\.gemini\antigravity\brain\6a2f6569-0174-4758-bbb9-311ec84e9bf3", pdf_filename)
desktop_pdf = os.path.expanduser(f"~\\Desktop\\{pdf_filename}")

chrome_path = find_chrome()
cmd = [
    chrome_path,
    "--headless=new",
    "--disable-gpu",
    f"--print-to-pdf={pdf_path}",
    "--no-pdf-header-footer",
    f"file:///{full_html_path.replace(os.sep, '/')}"
]
try:
    print("Rendering Full Report PDF via Chrome headless...")
    res = subprocess.run(cmd, capture_output=True, text=True)
    print("Return code:", res.returncode)
    if os.path.exists(pdf_path) and res.returncode == 0:
        size = os.path.getsize(pdf_path)
        size_kb = size / 1024.0
        print(f"SUCCESS: Generated {pdf_path} ({size} bytes)")
        shutil.copy2(pdf_path, brain_pdf)
        print(f"Copied to brain artifact: {brain_pdf}")
        try:
            shutil.copy2(pdf_path, desktop_pdf)
            print(f"✓ Copied to Desktop: {desktop_pdf}")
        except Exception as e:
            print(f"Warning copying to desktop: {e}")

        # Also update static alias filenames
        try:
            shutil.copy2(pdf_path, os.path.abspath("RevOps_Enterprise_OS_V17.6_Full_Report.pdf"))
            shutil.copy2(pdf_path, os.path.expanduser(r"~\Desktop\RevOps_Enterprise_OS_V17.6_Full_Report.pdf"))
            shutil.copy2(pdf_path, r"C:\Users\strel\.gemini\antigravity\brain\6a2f6569-0174-4758-bbb9-311ec84e9bf3\RevOps_Enterprise_OS_V17.6_Full_Report.pdf")
        except Exception:
            pass

        # Фиксация успешной генерации в audit_log.csv
        duration = time.time() - start_time
        log_generation(
            status="success",
            duration_sec=duration,
            role="full",
            report_type="Role_Report",
            pdf_size_kb=size_kb
        )
    else:
        err_msg = res.stderr.strip() if res.stderr else f"Browser exited with code {res.returncode}"
        duration = time.time() - start_time
        log_generation(
            status="error",
            duration_sec=duration,
            role="full",
            report_type="Role_Report",
            error_message=err_msg
        )
        print("ERROR rendering Full Report PDF!")
        print("Stderr:", res.stderr)
        sys.exit(1)
except Exception as e:
    duration = time.time() - start_time
    log_generation(
        status="error",
        duration_sec=duration,
        role="full",
        report_type="Role_Report",
        error_message=str(e)
    )
    print(f"FATAL ERROR in generate_full_report_pdf: {e}", file=sys.stderr)
    raise
