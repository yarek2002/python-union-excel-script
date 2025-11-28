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
    
    # Вырезаем тело таблицы
    body = raw.iloc[header_start + 1:, start_col: start_col + len(headers_unique)].copy()
    if body.empty:
        print(f"Нет строк под шапкой: {os.path.basename(file_path)}")
        return None
        
    body.columns = headers_unique
    body = body.dropna(how="all").reset_index(drop=True)
    
    if body.empty:
        print(f"Все строки пустые: {os.path.basename(file_path)}")
        return None

    record = defaultdict(lambda: pd.NA)
    record["Файл"] = os.path.basename(file_path)

    # --- Безопасно обрезаем по первому нечисловому в первом столбце ---
    first_col = body.iloc[:, 0]
    stop_idx = None
    for idx, val in enumerate(first_col):
        if pd.isna(val) or str(val).strip() == "" or not is_numeric(val):
            stop_idx = idx
            break
    
    if stop_idx is not None:
        body = body.iloc[:stop_idx].reset_index(drop=True)
    
    # Если после обрезки ничего не осталось — выходим
    if body.empty:
        print(f"После обрезки по номеру строки — пусто: {os.path.basename(file_path)}")
        return None

    # --- Теперь безопасно берём данные ---
    def safe_first(col_name_candidates):
        for col in col_name_candidates:
            if col in body.columns:
                cleaned = body[col].dropna()
                if not cleaned.empty:
                    return cleaned.iloc[0]
        return pd.NA

    # Документ
    record["Документ"] = safe_first([
        "№ документа-1", "Название документа-1",
        "№ документа", "Название документа"
    ])

    # Запрос от / Комментарий от
    record["Запрос от"] = safe_first(["Запрос от-1", "Запрос от"])
    record["Комментарий от"] = safe_first(["Комментарий от-1", "Комментарий от"])

    # Раздел / Лист
    record["Раздел"] = safe_first(["Раздел-1", "Раздел"])
    record["Лист"] = safe_first(["Лист-1", "Лист"])

    # Даты
    date_cols = [c for c in body.columns if c.startswith("Дата")]
    dates = []
    for col in date_cols:
        for v in body[col]:
            if pd.isna(v):
                continue
            dv = pd.to_datetime(v, errors="coerce")
            if not pd.isna(dv):
                dates.append(dv.date())
    if dates:
        record["Дата-1"] = dates[0].strftime("%d-%m-%Y")
        record["Дата-2"] = dates[-1].strftime("%d-%m-%Y")

    # Комментарии заказчика
    cust_cols = [c for c in body.columns if "Комментарий заказчика" in c]
    comments = []
    for col in cust_cols:
        for v in body[col].dropna():
            t = str(v).strip()
            if t and t.lower() not in ["nan", "none"]:
                comments.append(t)
    if comments:
        record["Комментарий заказчика"] = "; ".join(f"{i+1}. {t}" for i, t in enumerate(comments))

    # Ответ проектной организации
    ans_cols = [c for c in body.columns if "Ответ проектной" in c]
    answers = []
    for col in ans_cols:
        for v in body[col].dropna():
            t = str(v).strip()
            if t and t.lower() not in ["nan", "none"]:
                answers.append(t)
    if answers:
        record["Ответ проектной организации"] = "; ".join(f"{i+1}. {t}" for i, t in enumerate(answers))

    # Формируем итоговую запись
    result = {
        "Файл": record["Файл"],
        "№": pd.NA,
        "Запрос от": record.get("Запрос от", pd.NA),
        "Комментарий от": record.get("Комментарий от", pd.NA),
        "Документ": record.get("Документ", pd.NA),
        "Раздел": record.get("Раздел", pd.NA),
        "Лист": record.get("Лист", pd.NA),
        "Дата-1": record.get("Дата-1", pd.NA),
        "Дата-2": record.get("Дата-2", pd.NA),
        "Комментарий заказчика": record.get("Комментарий заказчика", pd.NA),
        "Ответ проектной организации": record.get("Ответ проектной организации", pd.NA),
    }
    return result

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
