#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Калькулятор зарплат Марьино
Принимает папку с xlsx файлами и путь к выходному файлу
"""

import os
import re
import glob
import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

import finwizard_sdk as fw

# ============================================================
# Конфигурация ячеек (новая форма ведомости)
# ============================================================

COL_DATE = 6        # F1 — дата смены (24.10)
COL_SURNAME = 9     # I4:I8 — фамилии
COL_SALARY = 8      # H4:H8 — З/п общая (уже включает аптеку)
COL_TIME = 10       # J4:J8 — время
COL_APTEKA_NAME = 2  # B39:B43 — фамилия в блоке «Расчет заработной платы смены»
COL_APTEKA_SUM = 6   # F39:F43 — «Аптека %» (процент врача с аптеки)

SURNAME_ROWS = range(4, 9)    # строки 4-8
APTEKA_ROWS = range(39, 44)   # строки 39-43


# ============================================================
# Вспомогательные функции
# ============================================================

def sanitize_sheet_title(title: str, max_len=31) -> str:
    if not title:
        return "Без_фамилии"
    invalid_chars = r'\/:*?"<>|[]'
    clean = ''.join(c for c in title if c not in invalid_chars).strip()
    if not clean:
        clean = "Без_фамилии"
    return clean[:max_len]


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
        for ch in (' ', '\u00A0', '\u202F', '\u2009', '₽', '$'):
            s = s.replace(ch, '')
        s = s.replace(',', '.').strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    return 0.0


def read_cell_string(ws, row, col):
    val = ws.cell(row=row, column=col).value
    if val is None:
        return ""
    return str(val).strip()


def names_match(a, b):
    return a.strip().lower() == b.strip().lower()


def get_month_name(month_num):
    months = {
        1: "Январь", 2: "Февраль", 3: "Март",
        4: "Апрель", 5: "Май", 6: "Июнь",
        7: "Июль", 8: "Август", 9: "Сентябрь",
        10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
    }
    return months.get(month_num, f"Месяц_{month_num}")


def format_time_to_minutes(time_str):
    """'08:50:00' -> '08:50'"""
    if not time_str:
        return ""
    if isinstance(time_str, (datetime.time, datetime.datetime)):
        return time_str.strftime("%H:%M")
    s = str(time_str).strip()
    parts = s.split(':')
    if len(parts) >= 2:
        return f"{parts[0]}:{parts[1]}"
    return s


def read_month_year(ws, filepath):
    """
    Месяц и год из F1. Поддерживает:
    - дату (datetime) в F1,
    - число вида 16.09 (месяц после точки, как раньше),
    - текст вида 'Сентябрь 2026' / '16-30.09.2026'.
    Год, если его нет в F1, ищем в первой строке ('2026 г.'), затем в имени файла.
    """
    raw = ws.cell(row=1, column=COL_DATE).value
    month_num = 0
    year = ""

    if isinstance(raw, (datetime.datetime, datetime.date)):
        month_num = raw.month
        year = str(raw.year)
    elif isinstance(raw, (int, float)):
        month_num = round((float(raw) - int(raw)) * 100)
    elif isinstance(raw, str):
        s = raw.strip()
        low = s.lower()
        for num in range(1, 13):
            if get_month_name(num).lower()[:3] in low:
                month_num = num
                break
        if not month_num:
            m = re.search(r'\.(\d{1,2})(?:\.(\d{2,4}))?', s)
            if m:
                month_num = int(m.group(1))
                if m.group(2):
                    year = m.group(2) if len(m.group(2)) == 4 else "20" + m.group(2)
        ym = re.search(r'(20\d{2})', s)
        if ym and not year:
            year = ym.group(1)

    if not year:
        for c in range(1, 15):
            v = read_cell_string(ws, 1, c)
            ym = re.search(r'(20\d{2})', v)
            if ym:
                year = ym.group(1)
                break

    if not year:
        ym = re.search(r'(20\d{2})', os.path.basename(ws.parent.path)
                       if getattr(ws.parent, "path", None) else filepath)
        year = ym.group(1) if ym else str(datetime.date.today().year)

    return month_num, year, (str(raw) if raw is not None else "")


def format_shift_date(ws, month_num):
    """Дата смены из F1 в виде 'дд.мм' (например '24.10')."""
    raw = ws.cell(row=1, column=COL_DATE).value
    if isinstance(raw, (datetime.datetime, datetime.date)):
        return raw.strftime("%d.%m")
    if isinstance(raw, (int, float)):
        day = int(raw)
        return f"{day:02d}.{month_num:02d}" if month_num else str(raw)
    if isinstance(raw, str):
        m = re.search(r'(\d{1,2})\.(\d{1,2})', raw)
        if m:
            return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}"
        return raw.strip()
    return ""


# ============================================================
# Основная функция
# ============================================================

def process_salary(input_folder, output_file):
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
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    col_header_font = Font(bold=True)
    col_header_alignment = Alignment(horizontal="center", vertical="center")
    data_alignment = Alignment(horizontal="center", vertical="center")

    sum_font = Font(bold=True, color="000000")
    sum_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    medium_border = Border(
        left=Side(style="medium"), right=Side(style="medium"),
        top=Side(style="medium"), bottom=Side(style="medium")
    )

    # ============ Обработка файлов ============
    total_files = len(xlsx_files)
    for file_idx, filepath in enumerate(xlsx_files, start=1):
        fw.core.update_progress(int(file_idx / total_files * 100),
                                f"Обработка {os.path.basename(filepath)}")
        print(f"\n=== Файл: {filepath} ===")

        try:
            wb = load_workbook(filepath, data_only=True)
        except Exception as e:
            print(f"  Не удалось открыть: {e}")
            continue

        ws = wb.active

        month_num, year, f1_raw = read_month_year(ws, filepath)
        month_name = get_month_name(month_num)
        full_date_label = f"{month_name} {year}"
        shift_date = format_shift_date(ws, month_num)
        print(f"  F1 raw: {f1_raw} -> month: {month_num} {month_name} year: {year}, дата смены: {shift_date}")

        for i in SURNAME_ROWS:
            surname = read_cell_string(ws, i, COL_SURNAME)
            if not surname:
                continue

            # пропускаем мусор (0, числа)
            try:
                float(surname.replace(',', '.'))
                continue
            except ValueError:
                pass

            print(f"\n  --- Строка {i} фамилия: {surname} ---")

            # === Новая фамилия — создаём лист ===
            if surname not in surnames:
                surnames.append(surname)
                surname_next_row[surname] = 3

                safe_title = sanitize_sheet_title(surname)
                original_title = safe_title
                counter = 1
                while safe_title in output_wb.sheetnames:
                    safe_title = f"{original_title}_{counter}"
                    counter += 1

                out_ws = output_wb.create_sheet(title=safe_title)
                surname_to_sheetname[surname] = safe_title

                cell_a1 = out_ws.cell(row=1, column=1, value=full_date_label)
                cell_a1.fill = header_fill
                cell_a1.border = thin_border
                cell_a1.alignment = header_alignment
                cell_a1.font = Font(bold=True)

                out_ws.merge_cells("B1:F1")
                cell_b1 = out_ws.cell(row=1, column=2, value=surname)
                cell_b1.font = header_font
                cell_b1.fill = header_fill
                cell_b1.alignment = header_alignment
                cell_b1.border = thin_border

                c = out_ws.cell(row=1, column=7)
                c.fill = header_fill
                c.border = thin_border

                headers = ["Дата", "З/п Общая", "УС", "Аптека", "Вычет", "Время", "Итого"]
                for col_idx, header_text in enumerate(headers, start=1):
                    c = out_ws.cell(row=2, column=col_idx, value=header_text)
                    c.font = col_header_font
                    c.alignment = col_header_alignment
                    c.border = thin_border

                out_ws.column_dimensions['A'].width = 18
                for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
                    out_ws.column_dimensions[col_letter].width = 14

            out_ws = output_wb[surname_to_sheetname[surname]]
            row = surname_next_row[surname]

            # A: Дата
            date_display = shift_date or full_date_label
            out_ws.cell(row=row, column=1, value=date_display).alignment = data_alignment

            # B: З/п Общая (H4:H8)
            zp_obshaya = read_cell_double(ws, i, COL_SALARY)
            out_ws.cell(row=row, column=2, value=zp_obshaya).alignment = data_alignment

            # C: УС — в новой форме всегда 0
            us_val = 0.0
            out_ws.cell(row=row, column=3, value=us_val).alignment = data_alignment

            # D: Аптека (F39:F43), фамилия в A39:A43
            apteka_val = 0.0
            for r in APTEKA_ROWS:
                ap_name = read_cell_string(ws, r, COL_APTEKA_NAME)
                if ap_name and names_match(ap_name, surname):
                    apteka_val = read_cell_double(ws, r, COL_APTEKA_SUM)
                    print(f"    Аптека найдена: строка {r} = {apteka_val}")
                    break
            out_ws.cell(row=row, column=4, value=apteka_val).alignment = data_alignment

            # E: Вычет — пустое поле
            out_ws.cell(row=row, column=5, value="").alignment = data_alignment

            # F: Время (J4:J8)
            vremya = format_time_to_minutes(ws.cell(row=i, column=COL_TIME).value)
            out_ws.cell(row=row, column=6, value=vremya).alignment = data_alignment

            # G: Итого = З/п + УС (аптека уже входит в З/п общую: 6500 + 314.0375 = 6814.0375)
            itogo = zp_obshaya + us_val
            out_ws.cell(row=row, column=7, value=itogo).alignment = data_alignment

            print(f"    ЗАПИСАНО: З/п={zp_obshaya} Аптека={apteka_val} "
                  f"Время={vremya} Итого={itogo}")

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

        totals = {2: 0.0, 3: 0.0, 4: 0.0, 7: 0.0}
        for r in range(3, sum_row):
            for col in totals:
                v = out_ws.cell(row=r, column=col).value
                if isinstance(v, (int, float)):
                    totals[col] += v

        sum_data = [
            (2, totals[2]), (3, totals[3]), (4, totals[4]),
            (5, ""), (6, ""), (7, totals[7]),
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

@fw.entrypoint
def run(params):
    xlsx_files_list = params.get("xlsxFiles", [])
    if not xlsx_files_list:
        raise fw.PluginError("Список xlsxFiles пуст!")

    input_folder = os.path.dirname(xlsx_files_list[0])
    output_file = params.get("outputFile")

    success = process_salary(input_folder, output_file)

    if not success:
        raise fw.PluginError("Произошла ошибка при обработке Excel файлов.")

    return {
        "message": "Расчет зарплат успешно завершен!",
        "outputPath": output_file,
    }


if __name__ == "__main__":
    run()
