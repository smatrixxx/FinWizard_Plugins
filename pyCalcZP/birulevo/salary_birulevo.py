#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Калькулятор зарплат Бирулёво
Принимает папку с xlsx файлами и путь к выходному файлу
"""

import os
import sys
import glob
import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import json

# ============================================================
# Вспомогательные функции
# ============================================================

def sanitize_sheet_title(title: str, max_len=31) -> str:
    """Очищает название листа под требования Excel/openpyxl"""
    if not title:
        return "Без_фамилии"
    
    # Убираем запрещённые символы
    invalid = r'\/:*?"<>|'
    clean = ''.join(c for c in title if c not in invalid)
    clean = clean.strip()
    
    if not clean:
        clean = "Без_фамилии"
    
    # Excel ограничивает 31 символ
    if len(clean) > max_len:
        clean = clean[:max_len]
    
    return clean

def read_cell_double(ws, row, col):
    """Читает число из ячейки, обрабатывает строки вида '2,00', '50 000,00'"""
    val = ws.cell(row=row, column=col).value
    if val is None:
        return 0.0

    if isinstance(val, (int, float)):
        return float(val)

    if isinstance(val, str):
        s = val.strip()

        if s.startswith('='):
            print(f"  Предупреждение: формула без кэша [{row},{col}]: {s}")
            return 0.0

        s = s.replace(' ', '')
        s = s.replace('\u00A0', '')
        s = s.replace('\u202F', '')
        s = s.replace('\u2009', '')
        s = s.replace(',', '.')
        s = s.replace('₽', '')
        s = s.replace('$', '')
        s = s.strip()

        try:
            return float(s)
        except ValueError:
            return 0.0

    return 0.0


def read_cell_string(ws, row, col):
    """Читает строку из ячейки как есть"""
    val = ws.cell(row=row, column=col).value
    if val is None:
        return ""
    return str(val).strip()


def names_match(a, b):
    """Сравнение фамилий без учёта регистра и пробелов"""
    return a.strip().lower() == b.strip().lower()


def get_month_name(month_num):
    """Возвращает название месяца на русском"""
    months = {
        1: "Январь", 2: "Февраль", 3: "Март",
        4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь",
        10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(month_num, f"Месяц_{month_num}")


def format_time_to_minutes(time_str):
    """
    Округляет время до минут.
    '08:50:00' -> '08:50'
    datetime.time(8, 50) -> '08:50'
    """
    if not time_str:
        return ""

    if isinstance(time_str, datetime.time):
        return time_str.strftime("%H:%M")

    if isinstance(time_str, datetime.datetime):
        return time_str.strftime("%H:%M")

    s = str(time_str).strip()
    parts = s.split(':')
    if len(parts) >= 2:
        return f"{parts[0]}:{parts[1]}"

    return s


# ============================================================
# Основная функция
# ============================================================

def process_salary(input_folder, output_file):
    """
    Обрабатывает xlsx файлы из папки и создаёт сводный файл
    """

    xlsx_files = sorted(glob.glob(os.path.join(input_folder, "*.xlsx")))

    if not xlsx_files:
        print(f"Ошибка: в папке '{input_folder}' не найдены xlsx файлы")
        return False

    print(f"Найдено файлов: {len(xlsx_files)}")

    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    surnames = []
    surname_next_row = {}
    surname_to_sheetname = {}

    # ============ Стили ============

    header_font = Font(bold=True, size=14, color="000000")
    header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    col_header_font = Font(bold=True)
    col_header_alignment = Alignment(horizontal="center", vertical="center")
    data_alignment = Alignment(horizontal="center", vertical="center")

    sum_font = Font(bold=True, color="000000")
    sum_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    medium_border = Border(
        left=Side(style="medium"),
        right=Side(style="medium"),
        top=Side(style="medium"),
        bottom=Side(style="medium")
    )

    # ============ Обработка файлов ============

    for filepath in xlsx_files:
        print(f"\n=== Файл: {filepath} ===")

        try:
            wb = load_workbook(filepath, data_only=True)
        except Exception as e:
            print(f"  Не удалось открыть: {e}")
            continue

        ws = wb.active

        # === Дата из E1 ===
        e1_raw = read_cell_string(ws, 1, 5)
        date_val = read_cell_double(ws, 1, 5)
        month_num = round((date_val - int(date_val)) * 100)

        # Год из I1
        year = read_cell_string(ws, 1, 9)
        if year.endswith("г."):
            year = year[:-2].strip()

        month_name = get_month_name(month_num)
        full_date_label = f"{month_name} {year}"

        print(f"  E1 raw: {e1_raw} -> month: {month_num} {month_name} year: {year}")

        # ======== Читаем фамилии H4:H14 ========
        for i in range(4, 15):
            surname = read_cell_string(ws, i, 8)
            if not surname:
                continue

            print(f"\n  --- Строка {i} фамилия: {surname} ---")
            print(f"    G{i} (З/п) raw: {read_cell_string(ws, i, 7)} "
                  f"-> double: {read_cell_double(ws, i, 7)}")
            print(f"    I{i} (Время) raw: {ws.cell(row=i, column=9).value}")

            # Отладка УС
            for r in range(31, 36):
                n = read_cell_string(ws, r, 1)
                if n:
                    print(f"    УС: A{r}={n} D{r} raw:{read_cell_string(ws, r, 4)} "
                          f"-> double:{read_cell_double(ws, r, 4)}")

            # Отладка Аптека
            for r in range(40, 46):
                n = read_cell_string(ws, r, 1)
                if n:
                    print(f"    Аптека: A{r}={n} C{r} raw:{read_cell_string(ws, r, 3)} "
                          f"-> double:{read_cell_double(ws, r, 3)}")

            # === Если фамилия новая — создаём лист ===
            is_new = surname not in surnames
            if is_new:
                surnames.append(surname)
                surname_next_row[surname] = 3

                safe_title = sanitize_sheet_title(surname)
                
                # Если после очистки название стало неуникальным — добавляем номер
                original_title = safe_title
                counter = 1
                while safe_title in output_wb.sheetnames:
                    safe_title = f"{original_title}_{counter}"
                    counter += 1

                out_ws = output_wb.create_sheet(title=safe_title)
                surname_to_sheetname[surname] = safe_title

                # A1 — полная дата "Март 2025" (голубой фон)
                cell_a1 = out_ws.cell(row=1, column=1, value=full_date_label)
                cell_a1.fill = header_fill
                cell_a1.border = thin_border
                cell_a1.alignment = header_alignment
                cell_a1.font = Font(bold=True)

                # B1:F1 — фамилия
                out_ws.merge_cells("B1:F1")
                cell_b1 = out_ws.cell(row=1, column=2, value=surname)
                cell_b1.font = header_font
                cell_b1.fill = header_fill
                cell_b1.alignment = header_alignment
                cell_b1.border = thin_border

                # G1 голубой
                c = out_ws.cell(row=1, column=7)
                c.fill = header_fill
                c.border = thin_border

                # Шапка (строка 2)
                headers = ["Дата", "З/п Общая", "УС", "Аптека", "Вычет", "Время", "Итого"]
                for col_idx, header_text in enumerate(headers, start=1):
                    c = out_ws.cell(row=2, column=col_idx, value=header_text)
                    c.font = col_header_font
                    c.alignment = col_header_alignment
                    c.border = thin_border

                out_ws.column_dimensions['A'].width = 18
                for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
                    out_ws.column_dimensions[col_letter].width = 14

            # === Записываем данные ===
            out_ws = output_wb[surname_to_sheetname[surname]]
            row = surname_next_row[surname]

            # A: Дата — всегда сокращённая из E1 ("08.09")
            date_display = e1_raw if e1_raw else full_date_label
            out_ws.cell(row=row, column=1, value=date_display).alignment = data_alignment

            # B: З/п Общая — столбец G (7)
            zp_obshaya = read_cell_double(ws, i, 7)
            out_ws.cell(row=row, column=2, value=zp_obshaya).alignment = data_alignment

            # C: УС — ищем в A31:A35, берём D (4)
            us_val = 0.0
            for r in range(31, 36):
                us_name = read_cell_string(ws, r, 1)
                if us_name and names_match(us_name, surname):
                    us_val = read_cell_double(ws, r, 4)
                    print(f"    УС найден: строка {r} = {us_val}")
                    break
            out_ws.cell(row=row, column=3, value=us_val).alignment = data_alignment

            # D: Аптека — ищем в A40:A45, берём C (3)
            apteka_val = 0.0
            for r in range(40, 46):
                ap_name = read_cell_string(ws, r, 1)
                if ap_name and names_match(ap_name, surname):
                    apteka_val = read_cell_double(ws, r, 3)
                    print(f"    Аптека найдена: строка {r} = {apteka_val}")
                    break
            out_ws.cell(row=row, column=4, value=apteka_val).alignment = data_alignment

            # E: Вычет (пусто)
            out_ws.cell(row=row, column=5, value="").alignment = data_alignment

            # F: Время — как строку, округлённую до минут
            vremya_raw = ws.cell(row=i, column=9).value
            vremya = format_time_to_minutes(vremya_raw)
            out_ws.cell(row=row, column=6, value=vremya).alignment = data_alignment

            # G: Итого = З/п + УС + Аптека
            itogo = zp_obshaya + us_val
            out_ws.cell(row=row, column=7, value=itogo).alignment = data_alignment

            print(f"    ЗАПИСАНО: З/п={zp_obshaya} УС={us_val} "
                  f"Аптека={apteka_val} Время={vremya} Итого={itogo}")

            surname_next_row[surname] = row + 1

        wb.close()

    # ============ Суммы на каждом листе ============
    for surname in surnames:
        out_ws = output_wb[surname_to_sheetname[surname]]
        sum_row = surname_next_row[surname]

        c = out_ws.cell(row=sum_row, column=1, value="ИТОГО:")
        c.font = sum_font
        c.fill = sum_fill
        c.border = medium_border
        c.alignment = data_alignment

        total_zp = 0.0
        total_us = 0.0
        total_apteka = 0.0
        total_itogo = 0.0

        for r in range(3, sum_row):
            val_b = out_ws.cell(row=r, column=2).value
            if isinstance(val_b, (int, float)):
                total_zp += val_b

            val_c = out_ws.cell(row=r, column=3).value
            if isinstance(val_c, (int, float)):
                total_us += val_c

            val_d = out_ws.cell(row=r, column=4).value
            if isinstance(val_d, (int, float)):
                total_apteka += val_d

            val_g = out_ws.cell(row=r, column=7).value
            if isinstance(val_g, (int, float)):
                total_itogo += val_g

        sum_data = [
            (2, total_zp),
            (3, total_us),
            (4, total_apteka),
            (5, ""),
            (6, ""),
            (7, total_itogo),
        ]

        for col_idx, value in sum_data:
            c = out_ws.cell(row=sum_row, column=col_idx, value=value)
            c.font = sum_font
            c.fill = sum_fill
            c.border = medium_border
            c.alignment = data_alignment

    # ============ Сохранение ============
    try:
        output_wb.save(output_file)
        print(f"\n{'=' * 50}")
        print(f"Обработано файлов: {len(xlsx_files)}")
        print(f"Создано листов: {len(surnames)}")
        print(f"Сохранено в: {output_file}")
        return True
    except Exception as e:
        print(f"Ошибка сохранения: {e}")
        return False


# ============================================================
# Точка входа
# ============================================================

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Ошибка: Скрипт запущен без JSON-параметров FinWizard!")
        sys.exit(1)

    # Движок передает пути к json-файлам в качестве аргументов
    input_json_path = sys.argv[1]
    output_json_path = sys.argv[2]

    # Читаем параметры из input.json
    with open(input_json_path, 'r', encoding='utf-8') as f:
        params = json.load(f)

    # Вытаскиваем нужные данные (m_pluginManager передает список файлов, берем их директорию)
    xlsx_files_list = params.get("xlsxFiles", [])
    if not xlsx_files_list:
        print("Ошибка: Список xlsxFiles пуст!")
        sys.exit(1)
        
    # Так как скрипт ожидал папку, берем родительскую папку первого файла
    input_folder = os.path.dirname(xlsx_files_list[0])
    output_file = params.get("outputFile")

    # Запускаем твою функцию
    success = process_salary(input_folder, output_file)
    
    # Записываем статус для C++ движка, чтобы он понял, что всё прошло ок
    output_data = {"success": success}
    if success:
        output_data["message"] = "Расчет зарплат успешно завершен!"
        output_data["outputPath"] = output_file
    else:
        output_data["error"] = "Произошла ошибка при обработке Excel файлов."

    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=4)

    sys.exit(0 if success else 1)
