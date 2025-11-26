import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from collections import Counter, defaultdict

def is_numeric(s):
    if pd.isna(s):
        return False
    try:
        float(str(s).replace(',', '.'))
        return True
    except:
        return False

def unique_within_file(headers):
    """Уникализация повторов только внутри одного файла"""
    count = Counter(headers)
    version = defaultdict(int)
    result = []
    for h in headers:
        if count[h] > 1:
            version[h] += 1
            result.append(f"{h}-{version[h]}")
        else:
            result.append(h)
    return result

def find_header_info(file_path):
    wb = load_workbook(file_path, read_only=True)
    ws = wb.active
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and str(v).startswith("№"):
                headers = []
                start_col = c
                while c <= ws.max_column:
                    cell = ws.cell(r, c)
                    if cell.value is None:
                        break
                    headers.append(str(cell.value))
                    c += 1
                return r - 1, start_col - 1, headers
    return 0, 0, []

def merge_excel_files(folder_path, output_file):
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    merged = []

    for name in files:
        if name.startswith("объединенный файл"):
            continue

        path = os.path.join(folder_path, name)
        raw = pd.read_excel(path, header=None, engine="openpyxl", dtype=str)

        header_start, start_col, headers = find_header_info(path)
        if not headers:
            continue

        # 1. Делаем уникальные заголовки внутри файла
        headers = unique_within_file(headers)

        # 2. Находим позиции колонок "Дата-1", "Дата-2", "Дата-3" ...
        date_positions = [i for i, h in enumerate(headers) if h.startswith("Дата")]

        # 3. Нарезаем секции по этим позициям (как раньше)
        idx = 0
        for pos in date_positions:
            cols = headers[idx:pos + 1]
            cs = start_col + idx
            ce = cs + len(cols)
            sec = raw.iloc[header_start + 1:, cs:ce].dropna(how="all")
            sec.columns = cols

            # Обрезка первой секции по numeric № (как было раньше)
            if idx == 0:
                stop = None
                for i in range(len(sec)):
                    if not is_numeric(sec.iloc[i, 0]):
                        stop = i
                        break
                if stop is not None:
                    sec = sec.iloc[:stop]

            # Обрезка секции по первой полностью пустой строке
            stop = None
            for i in range(len(sec)):
                if sec.iloc[i].isna().all():
                    stop = i
                    break
            if stop is not None:
                sec = sec.iloc[:stop]

            merged.append(sec)
            idx = pos + 1

        # Последняя секция после последней "Дата-X"
        if idx < len(headers):
            cols = headers[idx:]
            sec = raw.iloc[header_start + 1:, start_col + idx:].dropna(how="all")
            sec.columns = cols

            stop = None
            for i in range(len(sec)):
                if sec.iloc[i].isna().all():
                    stop = i
                    break
            if stop is not None:
                sec = sec.iloc[:stop]

            merged.append(sec)

    # Горизонтальное объединение секций
    final_df = pd.concat(merged, axis=1)
    final_df.insert(0, "Файл", name)

    # Убираем время из всех колонок Дата-X (как раньше в коде)
    for c in final_df.columns:
        if c.startswith("Дата"):
            final_df[c] = pd.to_datetime(final_df[c], errors='coerce').dt.strftime("%d-%m-%Y")

    final_df.to_excel(output_file, index=False)
    print("готово")

if __name__ == "__main__":
    folder = os.getcwd()
    date = datetime.now().strftime("%Y-%m-%d")
    out = f"объединенный файл {date}.xlsx"
    merge_excel_files(folder, out)
    os.system("pause")
