"""Модуль валидации данных Excel книги RevOps Enterprise OS V17.6."""
import os, openpyxl
from datetime import datetime, date

def validate_data(workbook_path: str) -> list[str]:
    """Проверяет валидность данных в Excel-книге RevOps Platform. Возвращает список ошибок."""
    errors: list[str] = []
    if not os.path.exists(workbook_path): return [f"Файл Excel не найден: {workbook_path}"]
    try: wb = openpyxl.load_workbook(workbook_path, data_only=True)
    except Exception as e: return [f"Не удалось открыть Excel: {e}"]
    try:
        req_sheets = ["⚙️ Настройки", "⚡ Экспресс_Калькулятор_3_Цифры", "raw_deals", "📄 Executive_OnePager"]
        for s in req_sheets:
            if s not in wb.sheetnames: errors.append(f"Лист '{s}' не найден в книге")
        if errors:
            wb.close()
            return errors

        # Проверка формул на ошибки (#REF!, #VALUE!, #DIV/0!, #N/A, #NAME? и др.)
        check_cells = [
            ("⚙️ Настройки", "B3"), ("⚙️ Настройки", "B9"), ("⚙️ Настройки", "B6"), ("⚙️ Настройки", "B7"),
            ("⚡ Экспресс_Калькулятор_3_Цифры", "B5"), ("⚡ Экспресс_Калькулятор_3_Цифры", "B6"),
            ("⚡ Экспресс_Калькулятор_3_Цифры", "B7"), ("📄 Executive_OnePager", "A5")
        ]
        err_cells = set()
        try: wb_f = openpyxl.load_workbook(workbook_path, data_only=False, read_only=True)
        except Exception: wb_f = None

        for sname, addr in check_cells:
            vals = [wb_f[sname][addr].value] if wb_f else []
            vals.append(wb[sname][addr].value)
            for v in vals:
                if v is not None and isinstance(v, str):
                    t = v.strip().lstrip("=").strip()
                    if t.startswith("#"):
                        errors.append(f"Лист '{sname}', ячейка {addr}: формула содержит ошибку {t}")
                        err_cells.add((sname, addr))
                        break

        # A. Лист "⚙️ Настройки"
        ws_set = wb["⚙️ Настройки"]
        if ("⚙️ Настройки", "B3") not in err_cells:
            b3 = ws_set["B3"].value
            if b3 is None or not str(b3).strip(): errors.append("Лист '⚙️ Настройки', ячейка B3: наименование организации пустое или не указано")
        if ("⚙️ Настройки", "B9") not in err_cells:
            b9 = ws_set["B9"].value
            if b9 is None or isinstance(b9, bool): errors.append("Лист '⚙️ Настройки', ячейка B9: план выручки пустой или не число")
            else:
                try:
                    b9_n = float(b9)
                    if not (0 < b9_n < 1_000_000_000): errors.append(f"Лист '⚙️ Настройки', ячейка B9: план выручки ({b9_n:,.0f}) вне диапазона (0, 1 000 000 000)")
                except (ValueError, TypeError): errors.append(f"Лист '⚙️ Настройки', ячейка B9: план выручки '{b9}' не является числом")

        if ("⚙️ Настройки", "B6") not in err_cells and ("⚙️ Настройки", "B7") not in err_cells:
            b6, b7 = ws_set["B6"].value, ws_set["B7"].value
            if b6 is None: errors.append("Лист '⚙️ Настройки', ячейка B6: начальная дата периода не заполнена")
            if b7 is None: errors.append("Лист '⚙️ Настройки', ячейка B7: конечная дата периода не заполнена")
            if b6 is not None and b7 is not None:
                try:
                    d6 = b6.date() if isinstance(b6, datetime) else (b6 if isinstance(b6, date) else datetime.fromisoformat(str(b6)).date())
                    d7 = b7.date() if isinstance(b7, datetime) else (b7 if isinstance(b7, date) else datetime.fromisoformat(str(b7)).date())
                    if not (d6 < d7): errors.append(f"Лист '⚙️ Настройки', ячейки B6/B7: дата начала ({b6}) должна быть меньше даты окончания ({b7})")
                except Exception:
                    if not (b6 < d7 if isinstance(d7, date) else b6 < b7): errors.append(f"Лист '⚙️ Настройки', ячейки B6/B7: дата начала ({b6}) должна быть меньше даты окончания ({b7})")

        # B. Лист "⚡ Экспресс_Калькулятор_3_Цифры"
        ws_exp = wb["⚡ Экспресс_Калькулятор_3_Цифры"]
        if ("⚡ Экспресс_Калькулятор_3_Цифры", "B5") not in err_cells:
            b5 = ws_exp["B5"].value
            if b5 is None or isinstance(b5, bool): errors.append("Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B5: количество лидов пустое или не задано")
            else:
                try:
                    b5_f = float(b5)
                    if not b5_f.is_integer() or not (1 <= int(b5_f) <= 100_000): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B5: количество лидов ({b5}) должно быть целым числом от 1 до 100 000")
                except (ValueError, TypeError): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B5: значение '{b5}' не является числом")

        if ("⚡ Экспресс_Калькулятор_3_Цифры", "B6") not in err_cells:
            b6_e = ws_exp["B6"].value
            if b6_e is None or isinstance(b6_e, bool): errors.append("Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B6: средний чек пустой или не задан")
            else:
                try:
                    b6_n = float(b6_e)
                    if not (10_000 <= b6_n <= 10_000_000): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B6: средний чек ({b6_n:,.0f}) вне диапазона от 10 000 до 10 000 000")
                except (ValueError, TypeError): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B6: средний чек '{b6_e}' не является числом")

        if ("⚡ Экспресс_Калькулятор_3_Цифры", "B7") not in err_cells:
            b7_e = ws_exp["B7"].value
            if b7_e is None or isinstance(b7_e, bool): errors.append("Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B7: количество менеджеров пустое или не задано")
            else:
                try:
                    b7_f = float(b7_e)
                    if not b7_f.is_integer() or not (1 <= int(b7_f) <= 100): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B7: количество менеджеров ({b7_e}) должно быть целым числом от 1 до 100")
                except (ValueError, TypeError): errors.append(f"Лист '⚡ Экспресс_Калькулятор_3_Цифры', ячейка B7: значение '{b7_e}' не является числом")

        # C. Лист "raw_deals"
        ws_d = wb["raw_deals"]
        col_did, col_amt, col_stg = 1, 3, 4
        for c in range(1, ws_d.max_column + 1):
            h = str(ws_d.cell(1, c).value or "").strip().lower()
            if h == "deal_id": col_did = c
            elif h == "amount": col_amt = c
            elif h == "stage_id": col_stg = c

        s6_cnt, act_cnt, s6_sum = 0, 0, 0.0
        for r in range(2, ws_d.max_row + 1):
            row_v = [ws_d.cell(r, c).value for c in range(1, min(ws_d.max_column + 1, 15))]
            if not any(v is not None for v in row_v): continue
            did = ws_d.cell(r, col_did).value
            if did is None or not str(did).strip(): errors.append(f"Лист 'raw_deals', строка {r}: пустой deal_id")
            amt, amt_n = ws_d.cell(r, col_amt).value, None
            if amt is not None:
                try:
                    amt_n = float(amt)
                    if amt_n < 0: errors.append(f"Лист 'raw_deals', строка {r}: отрицательная сумма сделки {amt_n:,.0f} в колонке amount")
                except (ValueError, TypeError): errors.append(f"Лист 'raw_deals', строка {r}: некорректная сумма '{amt}' в колонке amount")
            stg = ws_d.cell(r, col_stg).value
            if stg is not None:
                try:
                    stg_val = int(float(stg))
                    if stg_val == 6:
                        s6_cnt += 1
                        if amt_n is not None: s6_sum += amt_n
                    elif 1 <= stg_val <= 5: act_cnt += 1
                except (ValueError, TypeError): pass

        if s6_cnt < 1: errors.append("Лист 'raw_deals': нет ни одной закрытой сделки на этапе 6 (Успешно реализовано)")
        if act_cnt < 1: errors.append("Лист 'raw_deals': нет ни одной сделки на этапах 1-5 (активный пайплайн)")

        # D. Лист "📄 Executive_OnePager"
        if ("📄 Executive_OnePager", "A5") not in err_cells:
            ws_op = wb["📄 Executive_OnePager"]
            a5 = ws_op["A5"].value
            if a5 is None and wb_f:
                try:
                    f_val = wb_f["📄 Executive_OnePager"]["A5"].value
                    if isinstance(f_val, str) and f_val.startswith("="): a5 = s6_sum
                except Exception: pass
            if a5 is None or isinstance(a5, bool): errors.append("Лист '📄 Executive_OnePager', ячейка A5: выручка Closed-Won пустая или не задана")
            else:
                try:
                    a5_n = float(a5)
                    if a5_n < 0: errors.append(f"Лист '📄 Executive_OnePager', ячейка A5: выручка Closed-Won ({a5_n:,.0f}) должна быть числом >= 0")
                except (ValueError, TypeError): errors.append(f"Лист '📄 Executive_OnePager', ячейка A5: значение '{a5}' не является числом")

        if wb_f:
            try: wb_f.close()
            except Exception: pass
        wb.close()
        return errors
    except Exception as e:
        try: wb.close()
        except Exception: pass
        return [f"Ошибка при валидации данных: {e}"]
