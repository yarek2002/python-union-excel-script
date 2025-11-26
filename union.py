import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from collections import Counter, defaultdict

def is_numeric(s):
    if pd.isna(s):
        return False
    s = str(s).strip()
    try:
        float(s)
        return True
    except:
        return False

def make_unique_columns(headers):
    count = Counter(headers)
    version = defaultdict(int)
    unique_cols = []
    for h in headers:
        if count[h] > 1:
            version[h] += 1
            unique_cols.append(f"{h}-{version[h]}")
        else:
            unique_cols.append(h)
    return unique_cols

def find_header_info(file_path):
    wb = load_workbook(file_path, read_only=True)
    ws = wb.active
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value and str(cell_value).startswith('№'):
                headers = []
                start_col = col
                while col <= ws.max_column:
                    cell = ws.cell(row=row, column=col)
                    if cell.value is None:
                        break
                    headers.append(str(cell.value))
                    col += 1
                return row - 1, start_col - 1, headers  # start_col 0-based
    return 0, 0, []

def get_max_headers(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    union = []
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        _, _, headers = find_header_info(file_path)
        if headers:
            for h in headers:
                if h not in union:
                    union.append(h)
    return ['Файл'] + union  # ← БЕЗ уникализации здесь!


def merge_excel_files(folder_path, output_file, max_headers):
    all_dfs = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        raw = pd.read_excel(file_path, header=None, engine='openpyxl', dtype=str)

        header_start, start_col, headers = find_header_info(file_path)
        header_row = header_start

        if not headers:
            continue

        #  Уникализируем шапку один раз, целиком (все секции будут совпадать между файлами)
        unique_headers = make_unique_columns(headers)

        #  Вырезаем до конца документа, но ограничиваем первую секцию как раньше по №
        body = raw.iloc[header_row + 1:, start_col: start_col + len(unique_headers)].copy()
        body.columns = unique_headers
        body = body.dropna(how='all')

        # обрезка первой секции по numeric №
        if "№-1" in body.columns or "№" in body.columns:
            first_col = body.columns[0]
            stop_idx = None
            for i in range(len(body)):
                val = body.iloc[i, 0]
                if pd.isna(val) or not is_numeric(val):
                    stop_idx = i
                    break
            if stop_idx is not None:
                body = body.iloc[:stop_idx]

        #  Обрезаем последнюю секцию по пустой строке
        stop_idx = None
        for i in range(len(body)):
            if body.iloc[i].isna().all():
                stop_idx = i
                break
        if stop_idx is not None:
            body = body.iloc[:stop_idx]

        #  Вставляем столбец Файл
        body.insert(0, "Файл", file_name)

        #  Делаем reindex по сырым заголовкам + unique-суффиксы уже совпадают
        body = body.reindex(columns=max_headers, fill_value=pd.NA)

        all_dfs.append(body)

    if not all_dfs:
        all_dfs = [pd.DataFrame(columns=max_headers)]

    merged_df = pd.concat(all_dfs, ignore_index=True)

    #  Убираем время из всех Дата-X
    date_columns = [c for c in merged_df.columns if c.startswith("Дата")]
    for col in date_columns:
        merged_df[col] = pd.to_datetime(merged_df[col], errors='coerce').dt.strftime("%d-%m-%Y")

    merged_df.to_excel(output_file, index=False)