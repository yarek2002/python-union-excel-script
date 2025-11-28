import os
import traceback
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
    count = Counter(headers)
    version = defaultdict(int)
    result = []
    for h in headers:
        if count[h] > 1:
            version[h] += 1
            result.append(f"{h}-{version[h]}")
        else:
            result.append(h)
    return [h.replace("Дата", f"Дата-{i+1}") if h.startswith("Дата") else h for i, h in enumerate(result)]

def find_header_info(file_path):
    wb = load_workbook(file_path, read_only=True)
    ws = wb.active
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v and str(v).startswith("№"):
                headers = []
                start_col = c
                cc = c
                while cc <= ws.max_column:
                    cell = ws.cell(r, cc)
                    if cell.value is None:
                        break
                    headers.append(str(cell.value))
                    cc += 1
                return r - 1, start_col - 1, headers
    return 0, 0, []

def get_union_headers(folder_path):
    union = []
    for name in os.listdir(folder_path):
        if name.endswith(".xlsx") and not name.startswith("объединенный файл"):
            path = os.path.join(folder_path, name)
            _, _, headers = find_header_info(path)
            if headers:
                for h in headers:
                    if h not in union:
                        union.append(h)
    return ['Файл', '№', 'Запрос от', 'Комментарий от', 'Документ', 'Раздел', 'Лист', 'Дата-1', 'Дата-2',
            'Комментарий заказчика', 'Ответ проектной организации']

def extract_file_data(file_path):
    raw = pd.read_excel(file_path, header=None, engine="openpyxl", dtype=str)
    header_start, start_col, headers = find_header_info(file_path)

    if not headers:
        return None

    headers_unique = unique_within_file(headers)
    body = raw.iloc[header_start + 1:, start_col: start_col + len(headers_unique)].copy()
    body.columns = headers_unique
    body = body.dropna(how="all")

    record = defaultdict(lambda: pd.NA)
    record["Файл"] = os.path.basename(file_path)

    stop = None
    for i in range(len(body)):
        val = body.iloc[i, 0]

    # если NaN, пусто или не число → конец 1 секции
        if pd.isna(val) or str(val).strip() == "" or not is_numeric(val):
            stop = i
            break

    if stop is not None:
        body = body.iloc[:stop]


    # Документ (номер или название)
    if "№ документа-1" in body.columns:
        record["Документ"] = body["№ документа-1"].dropna().iloc[0]
    elif "Название документа-1" in body.columns:
        record["Документ"] = body["Название документа-1"].dropna().iloc[0]
    elif "№ документа" in body.columns:
        record["Документ"] = body["№ документа"].dropna().iloc[0]
    elif "Название документа" in body.columns:
        record["Документ"] = body["Название документа"].dropna().iloc[0]

    if "Запрос от-1" in body.columns:
        record["Запрос от"] = body["Запрос от-1"].dropna().iloc[0]
    elif "Запрос от" in body.columns:
        record["Запрос от"] = body["Запрос от"].dropna().iloc[0]

    if "Комментарий от-1" in body.columns:
        record["Комментарий от"] = body["Комментарий от-1"].dropna().iloc[0]
    elif "Комментарий от" in body.columns:
        record["Комментарий от"] = body["Комментарий от"].dropna().iloc[0]

    if "Раздел-1" in body.columns:
        record["Раздел"] = body["Раздел-1"].dropna().iloc[0]
    elif "Раздел" in body.columns:
        record["Раздел"] = body["Раздел"].dropna().iloc[0]

    if "Лист-1" in body.columns:
        record["Лист"] = body["Лист-1"].dropna().iloc[0]
    elif "Лист" in body.columns:
        record["Лист"] = body["Лист"].dropna().iloc[0]

    # Дата-1 первая и Дата-2 последняя, без времени
    date_cols = [c for c in body.columns if c.startswith("Дата")]
    if date_cols:
        dates = []
        for col in date_cols:
            for v in body[col].dropna():
                dv = pd.to_datetime(v, errors="coerce")
                if not pd.isna(dv):
                    dates.append(dv)
        if dates:
            record["Дата-1"] = dates[0].strftime("%d-%m-%Y")
            record["Дата-2"] = dates[-1].strftime("%d-%m-%Y")

    # Комментарий заказчика (собираем все X в один список)
    cust_cols = [c for c in body.columns if c.startswith("Комментарий заказчика")]
    all_cust = []
    for col in cust_cols:
        for v in body[col].dropna():
            t = v.strip()
            if t:
                all_cust.append(t)
    if all_cust:
        record["Комментарий заказчика"] = "; ".join(f"{i+1}. {t}" for i, t in enumerate(all_cust))

    # Ответ проектной организации (аналогично)
    ans_cols = [c for c in body.columns if c.startswith("Ответ проектной")]
    all_ans = []
    for col in ans_cols:
        for v in body[col].dropna():
            t = v.strip()
            if t:
                all_ans.append(t)
    if all_ans:
        record["Ответ проектной организации"] = "; ".join(f"{i+1}. {t}" for i, t in enumerate(all_ans))

    record_renamed = {
        "Файл": record["Файл"],
        "№": record.get("№", pd.NA),
        "Запрос от": record.get("Запрос от", pd.NA),
        "Комментарий от": record.get("Комментарий от", pd.NA),
        "Документ": record.get("Документ", pd.NA),
        "Раздел": record.get("Раздел", pd.NA),
        "Лист": record.get("Лист", pd.NA),
        "Дата-1": record["Дата-1"],
        "Дата-2": record["Дата-2"],
        "Комментарий заказчика": record["Комментарий заказчика"],
        "Ответ проектной организации": record["Ответ проектной организации"],
    }

    return record_renamed

def process_folder(folder_path, output_file):
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("объединенный файл")]
    records = []
    for name in files:
        path = os.path.join(folder_path, name)
        r = extract_file_data(path)
        if r:
            records.append(r)

    merged = pd.DataFrame(records, columns=get_union_headers(folder_path))
    merged.to_excel(output_file, index=False)
    return output_file

if __name__ == "__main__":
    try:
        folder = os.getcwd()
        out = f"объединенный файл {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        res = process_folder(folder, out)
        print("Готово:", res)
    except Exception as e:
        tb = traceback.format_exc()
        print("Ошибка:\n", tb)
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(tb)
    finally:
        os.system("pause")
