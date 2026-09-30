import os
import sys
import json
import time
import shutil
import argparse
import subprocess
import jinja2
import pypdf

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# Include scratch root for audit_log and modules
PRESENTATION_DIR = os.path.dirname(os.path.abspath(__file__))
SCRATCH_ROOT = os.path.dirname(PRESENTATION_DIR)
if SCRATCH_ROOT not in sys.path:
    sys.path.insert(0, SCRATCH_ROOT)

from audit_log import log_generation
from presentation.roles_config import ROLE_CONFIGS, get_role_context

DATA_FILE = os.path.join(PRESENTATION_DIR, 'data.json')
TEMPLATE_FILE = os.path.join(PRESENTATION_DIR, 'template.html')

CURRENT_BRAIN_DIR = r'C:\Users\strel\.gemini\antigravity\brain\b958a22b-49ff-459c-a191-6c697ec14334'
LEGACY_BRAIN_DIR = r'C:\Users\strel\.gemini\antigravity\brain\6a2f6569-0174-4758-bbb9-311ec84e9bf3'

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


def render_html_for_role(role_key: str, data: dict, tmpl_content: str) -> tuple[str, str]:
    """Рендерит HTML для указанной роли и сохраняет во временный/целевой HTML файл."""
    role_ctx = get_role_context(role_key)
    full_ctx = {**data, **role_ctx}

    template = jinja2.Template(tmpl_content)
    rendered_html = template.render(**full_ctx)

    if role_key == 'all':
        out_html = os.path.join(PRESENTATION_DIR, 'presentation.html')
    else:
        out_html = os.path.join(PRESENTATION_DIR, f'presentation_{role_key}.html')

    with open(out_html, 'w', encoding='utf-8') as f:
        f.write(rendered_html)

    print(f"[{role_key.upper()}] HTML сохранен: {out_html} ({len(rendered_html)} chars)")
    return out_html, rendered_html


def build_pdf_for_role(role_key: str, data: dict, tmpl_content: str, browser_bin: str = None, no_browser: bool = False) -> str:
    """Генерирует HTML и PDF отчет для заданной роли."""
    start_time = time.time()
    cfg = ROLE_CONFIGS[role_key]
    out_pdf = os.path.join(PRESENTATION_DIR, cfg['filename'])

    try:
        out_html, _ = render_html_for_role(role_key, data, tmpl_content)

        if no_browser:
            print(f"[{role_key.upper()}] Флаг --no-browser активен. Генерация PDF пропущена.")
            return out_html

        if not browser_bin:
            browser_bin = find_browser()

        html_url = f"file:///{out_html.replace(os.sep, '/')}"
        cmd = [
            browser_bin,
            '--headless=new',
            '--disable-gpu',
            '--no-pdf-header-footer',
            '--allow-file-access-from-files',
            '--run-all-compositor-stages-before-draw',
            '--virtual-time-budget=5000',
            f'--print-to-pdf={out_pdf}',
            html_url
        ]

        print(f"[{role_key.upper()}] Рендер PDF через Chrome Headless...")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            print(f"Browser stderr: {result.stderr}")
            raise RuntimeError(f"Browser exited with code {result.returncode}")

        if not os.path.exists(out_pdf):
            raise FileNotFoundError(f"PDF was not generated at: {out_pdf}")

        pdf_size = os.path.getsize(out_pdf)
        print(f"[{role_key.upper()}] PDF сформирован: {out_pdf} ({pdf_size:,} bytes)")

        # Проверка страниц через PyPDF
        reader = pypdf.PdfReader(out_pdf)
        page_count = len(reader.pages)
        expected = cfg['expected_pages']
        print(f"[{role_key.upper()}] Страниц в PDF: {page_count} (ожидалось: {expected})")
        if page_count != expected:
            print(f"⚠️ ПРЕДУПРЕЖДЕНИЕ: Ожидалось {expected} стр., получено {page_count}!")
        else:
            print(f"✓ [{role_key.upper()}] ИДЕАЛЬНО: Ровно {expected} страниц.")

        # Копирование в корень проекта
        root_copy_path = os.path.join(SCRATCH_ROOT, cfg['root_copy'])
        shutil.copyfile(out_pdf, root_copy_path)
        print(f"  -> Скопировано в корень: {root_copy_path}")

        # Копирование в артефакты Brain (текущий и легаси)
        for b_dir in [CURRENT_BRAIN_DIR, LEGACY_BRAIN_DIR]:
            if os.path.exists(b_dir):
                target = os.path.join(b_dir, os.path.basename(root_copy_path))
                try:
                    shutil.copyfile(out_pdf, target)
                    print(f"  -> Скопировано в артефакты: {target}")
                except Exception as e:
                    print(f"  -> Не удалось скопировать в {target}: {e}")

        # Копирование на Рабочий стол
        desktop_dir = os.path.expanduser(r'~\Desktop')
        if os.path.exists(desktop_dir):
            desk_pdf = os.path.join(desktop_dir, cfg['root_copy'])
            try:
                shutil.copyfile(out_pdf, desk_pdf)
                print(f"  -> Скопировано на Desktop: {desk_pdf}")
            except Exception as e:
                print(f"  -> Не удалось скопировать на Desktop: {e}")

        # Логирование в audit_log
        duration = time.time() - start_time
        log_generation(
            status="success",
            duration_sec=duration,
            role=role_key,
            report_type=f"Report_{role_key.upper()}",
            pdf_size_kb=pdf_size / 1024.0
        )

        return out_pdf

    except Exception as e:
        duration = time.time() - start_time
        log_generation(
            status="error",
            duration_sec=duration,
            role=role_key,
            report_type=f"Report_{role_key.upper()}",
            error_message=str(e)
        )
        raise


def run_presteps():
    """Обновляет data.json из Excel и перегенерирует SVG графики."""
    print("=== Step 0: Updating data from Excel & re-rendering charts ===")
    try:
        subprocess.run([sys.executable, os.path.join(PRESENTATION_DIR, 'build_data.py')], cwd=SCRATCH_ROOT, check=True)
        subprocess.run([sys.executable, os.path.join(PRESENTATION_DIR, 'generate_charts.py')], cwd=SCRATCH_ROOT, check=True)
    except Exception as e:
        print(f"Warning: could not run auto-update step: {e}")


def main():
    parser = argparse.ArgumentParser(description="RevOps Enterprise OS V17.6 — Генератор ролевых PDF отчетов")
    parser.add_argument(
        '--role',
        choices=['all', 'ceo', 'rop', 'cfo', 'all-roles'],
        default='all',
        help="Роль для генерации: ceo (Собственник), rop (РОП), cfo (Финансы), all (Мастер-презентация), all-roles (Все 4 отчета сразу)"
    )
    parser.add_argument('--skip-presteps', action='store_true', help="Пропустить выгрузку данных из Excel и перерисовку SVG")
    parser.add_argument('--no-browser', action='store_true', help="Только скомпилировать HTML без вызова Chrome")

    args = parser.parse_args()

    if not args.skip_presteps:
        run_presteps()

    print("=== Step 1: Loading data.json & template.html ===")
    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        tmpl_content = f.read()

    browser_bin = None if args.no_browser else find_browser()

    if args.role == 'all-roles':
        print("\n" + "=" * 65)
        print("  ГЕНЕРАЦИЯ ВСЕХ 4 РОЛЕВЫХ ОТЧЕТОВ REVOPS ENTERPRISE")
        print("=" * 65)
        roles = ['all', 'ceo', 'rop', 'cfo']
        results = {}
        for r in roles:
            print(f"\n---> Сборка отчета для роли: {r.upper()} ({ROLE_CONFIGS[r]['name']})")
            pdf = build_pdf_for_role(r, data, tmpl_content, browser_bin=browser_bin, no_browser=args.no_browser)
            results[r] = pdf
        print("\n" + "=" * 65)
        print("🎉 ВСЕ 4 РОЛЕВЫХ ОТЧЕТА УСПЕШНО СФОРМИРОВАНЫ:")
        for r, path in results.items():
            print(f"   • [{r.upper()}] {ROLE_CONFIGS[r]['name']} -> {path}")
        print("=" * 65)
    else:
        print(f"\n---> Сборка отчета для роли: {args.role.upper()} ({ROLE_CONFIGS[args.role]['name']})")
        pdf = build_pdf_for_role(args.role, data, tmpl_content, browser_bin=browser_bin, no_browser=args.no_browser)
        print(f"\n✓ Успешно сформирован отчет: {pdf}")


if __name__ == '__main__':
    main()
