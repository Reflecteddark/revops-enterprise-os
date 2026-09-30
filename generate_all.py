import os
import sys
import subprocess

sys.stdout.reconfigure(encoding='utf-8')
SCRATCH_DIR = os.path.dirname(os.path.abspath(__file__))

def run_step(step_name, cmd, cwd=SCRATCH_DIR):
    print(f"\n---> {step_name}...")
    res = subprocess.run(cmd, cwd=cwd, shell=True)
    if res.returncode != 0:
        print(f"❌ Ошибка на шаге: {step_name} (Код: {res.returncode})")
        sys.exit(res.returncode)
    print(f"✓ {step_name} завершен успешно.")

def main():
    print("=" * 60)
    print("  RevOps Enterprise OS V17.6 — Генератор PDF Отчетов")
    print("=" * 60)

    # 0. Экспорт свежих данных из Excel книги
    run_step("0. Экспорт свежих данных из Excel (SSOT)", "python presentation/build_data.py")

    # 1. Генерация SVG графиков
    run_step("1. Построение 5 векторных SVG-графиков", "python presentation/generate_charts.py")

    # 2. Сборка 10-страничного Executive Summary через Chrome Headless
    run_step("2. Рендер Executive Summary PDF (10 слайдов C-Level)", "python presentation/build.py")

    # 3. Сборка 15-страничного Full Report
    run_step("3. Рендер Full Report PDF (15 аналитических витрин)", "python generate_full_report_pdf.py")

    print("\n" + "=" * 60)
    print("🎉 ВСЕ PDF УСПЕШНО СФОРМИРОВАНЫ:")
    print("   1. presentation/RevOps_Executive_Summary_2026-09-30.pdf")
    print("   2. RevOps_Enterprise_OS_V17.6_Full_Report.pdf")
    print("   3. Копии сохранены в артефактах brain для мгновенного просмотра.")
    print("=" * 60)

if __name__ == '__main__':
    main()
