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
    # Фиксированная шапка, как вам нужна
    return ['Файл', '№', 'Запрос от', 'Комментарий от', 'Документ', 'Раздел', 'Лист',
            'Дата-1', 'Дата-2', 'Комментарий Заказчика', 'Ответ Проектной Организации']

def extract_file_data(file_path):
    raw = pd.read_excel(file_path, header=None, engine="openpyxl", dtype=str)
    header_start, start_col, headers = find_header_info(file_path)
    if not headers:
        return None

    headers_unique = unique_within_file(headers)

    body = raw.iloc[header_start + 1:, start_col: start_col + len(headers_unique)].copy()

    # Защита от пустого body — НЕ ломает логику
    if body.empty:
        print(f"Нет строк под шапкой: {os.path.basename(file_path)}")
        return None

    body.columns = headers_unique
    body = body.dropna(how="all").reset_index(drop=True)

    if body.empty:
        print(f"Все строки пустые после dropna: {os.path.basename(file_path)}")
        return None

    records = []
    for i in range(len(body)):
        val = body.iloc[i, 0]

        # Если в колонке № есть число → делаем запись
        if is_numeric(val) and str(val).strip() != "":
            row = body.iloc[i]

            record = {
                "Файл": os.path.basename(file_path),
                "№": str(val).strip()
            }

            record["Запрос от"] = row.get("Запрос от-1", row.get("Запрос от", pd.NA))
            record["Комментарий от"] = row.get("Комментарий от-1", row.get("Комментарий от", pd.NA))
            record["Документ"] = row.get("№ документа-1", row.get("Название документа-1",
                                row.get("№ документа", row.get("Название документа", pd.NA)))
            )
            record["Раздел"] = row.get("Раздел-1", row.get("Раздел", pd.NA))
            record["Лист"] = row.get("Лист-1", row.get("Лист", pd.NA))

            # Обработка дат
            date_vals = [row[c] for c in body.columns if c.startswith("Дата") and not pd.isna(row[c])]
            if len(date_vals) == 1:
                d = pd.to_datetime(date_vals[0], errors='coerce')
                if not pd.isna(d):
                    record["Дата-1"] = d.strftime("%d-%m-%Y")
            elif len(date_vals) > 1:
                d1 = pd.to_datetime(date_vals[0], errors='coerce')
                d2 = pd.to_datetime(date_vals[-1], errors='coerce')
                if not pd.isna(d1): record["Дата-1"] = d1.strftime("%d-%m-%Y")
                if not pd.isna(d2): record["Дата-2"] = d2.strftime("%d-%m-%Y")

            #  Сбор и нумерация комментариев заказчика
            cust_vals = [row[c] for c in body.columns if "Комментарий Заказчика" in c and not pd.isna(row[c])]
            if cust_vals:
                cust_list = []
                for v in cust_vals:
                    t = str(v).strip()
                    if t and t.lower() not in ["nan", "none"]:
                        cust_list.append(t)
                if cust_list:
                    record["Комментарий Заказчика"] = " ".join(f"{idx+1}) {c}" for idx, c in enumerate(cust_list))

            #  Сбор и нумерация ответов проектной организации
            ans_vals = [row[c] for c in body.columns if "Ответ Проектной Организации" in c and not pd.isna(row[c])]
            if ans_vals:
                ans_list = []
                for v in ans_vals:
                    t = str(v).strip()
                    if t and t.lower() not in ["nan", "none"]:
                        ans_list.append(t)
                if ans_list:
                    record["Ответ Проектной Организации"] = " ".join(f"{idx+1}) {a}" for idx, a in enumerate(ans_list))

            records.append(record)

    if not records:
        print(f"В файле {os.path.basename(file_path)} не найдено строк с числовым №.")
        return None

    # Если возвращаем, то именно список строк
    return records

def process_folder(folder_path, output_file):
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("объединенный файл")]
    records = []
    for name in files:
        path = os.path.join(folder_path, name)
        r = extract_file_data(path)
        if r:
            records.extend(r)  # ← добавляем ВСЕ строки

    merged = pd.DataFrame(records, columns=get_union_headers(folder_path))
    
    # ⚠ Если только 1 дата, убираем её из Дата-2
    if not merged.empty:
        for i, row in merged.iterrows():
            if not pd.isna(row["Дата-1"]) and pd.isna(row["Дата-2"]):
                merged.at[i, "Дата-2"] = pd.NA

    merged.to_excel(output_file, index=False)
    return output_file

if __name__ == "__main__":
    try:
        folder = os.getcwd()
        out = f"объединенный файл {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        res = process_folder(folder, out)
        print("Готово:", res)
    except Exception:
        tb = traceback.format_exc()
        print("Ошибка:\n", tb)
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(tb)
    finally:
        os.system("pause")
