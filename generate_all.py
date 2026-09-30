import os
import sys
import argparse
import subprocess

sys.stdout.reconfigure(encoding='utf-8', errors='replace')
SCRATCH_DIR = os.path.dirname(os.path.abspath(__file__))


def run_step(step_name, cmd, cwd=SCRATCH_DIR):
    print(f"\n---> {step_name}...")
    res = subprocess.run(cmd, cwd=cwd, shell=True)
    if res.returncode != 0:
        print(f"❌ Ошибка на шаге: {step_name} (Код: {res.returncode})")
        sys.exit(res.returncode)
    print(f"✓ {step_name} завершен успешно.")


def main():
    parser = argparse.ArgumentParser(description="RevOps Enterprise OS V17.6 — Комплексный генератор PDF отчетов")
    parser.add_argument(
        '--role',
        choices=['all', 'ceo', 'rop', 'cfo', 'all-roles'],
        default='all-roles',
        help="Роль отчета: ceo, rop, cfo, all или all-roles (по умолчанию генерируются все)"
    )
    args = parser.parse_args()

    print("=" * 65)
    print("  RevOps Enterprise OS V17.6 — Генератор Ролевых PDF Отчетов")
    print(f"  Выбранный режим: {args.role.upper()}")
    print("=" * 65)

    # 0. Экспорт свежих данных из Excel книги (SSOT)
    run_step("0. Экспорт свежих данных из Excel (SSOT)", f'"{sys.executable}" presentation/build_data.py')

    # 1. Генерация SVG графиков
    run_step("1. Построение 5 векторных SVG-графиков", f'"{sys.executable}" presentation/generate_charts.py')

    # 2. Сборка ролевых презентаций через Chrome Headless
    if args.role == 'all-roles':
        run_step("2. Рендер всех 4 ролевых отчетов (Master, CEO, РОП, CFO)", f'"{sys.executable}" presentation/build.py --role all-roles --skip-presteps')
    else:
        run_step(f"2. Рендер отчета для роли '{args.role}'", f'"{sys.executable}" presentation/build.py --role {args.role} --skip-presteps')

    # 3. Сборка 15-страничного Full Report (при all или all-roles)
    if args.role in ('all', 'all-roles'):
        run_step("3. Рендер Full Report PDF (15 аналитических витрин)", f'"{sys.executable}" generate_full_report_pdf.py')

    print("\n" + "=" * 65)
    print("🎉 ВСЕ ЦЕЛЕВЫЕ PDF УСПЕШНО СФОРМИРОВАНЫ:")
    if args.role == 'all-roles':
        print("   1. RevOps_Enterprise_OS_V17.6_Executive_Summary.pdf (10 стр. C-Level Master)")
        print("   2. RevOps_Report_CEO.pdf (5 стр. Стратегический срез CEO)")
        print("   3. RevOps_Report_ROP.pdf (5 стр. Пульт РОПа и аудит звонков)")
        print("   4. RevOps_Report_CFO.pdf (5 стр. Финансовый календарь и дебиторка)")
        print("   5. RevOps_Enterprise_OS_V17.6_Full_Report.pdf (15 стр. Полные витрины)")
    else:
        print(f"   • Отчет для роли '{args.role}' сохранен в корне проекта и на Desktop.")
    print("   • Файлы продублированы в артефактах brain для быстрого доступа.")
    print("=" * 65)


if __name__ == '__main__':
    main()
