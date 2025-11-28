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

    records = []  # ← сюда будем складывать записи по строкам
    num_col = body.columns[0]  # Первая колонка — это № или №-1 и т.д.

    for i in range(len(body)):
        val = body.iloc[i, 0]

        # ЕСЛИ в колонке № есть число → делаем запись
        if is_numeric(val) and str(val).strip() != "":
            row = body.iloc[i]  # ← берём всю строку

            record = {
                "Файл": os.path.basename(file_path),
                "№": str(val).strip()
            }

            # Остальные значения берем из этой же строки
            record["Запрос от"] = row.get("Запрос от-1", row.get("Запрос от", pd.NA))
            record["Комментарий от"] = row.get("Комментарий от-1", row.get("Комментарий от", pd.NA))
            record["Документ"] = row.get("№ документа-1", row.get("Название документа-1",
                                    row.get("№ документа", row.get("Название документа", pd.NA))))
            record["Раздел"] = row.get("Раздел-1", row.get("Раздел", pd.NA))
            record["Лист"] = row.get("Лист-1", row.get("Лист", pd.NA))

            # Обработка дат для строки
            date_vals = [row[c] for c in body.columns if c.startswith("Дата") and not pd.isna(row[c])]
            if len(date_vals) == 1:  # ← только одна дата
                d = pd.to_datetime(date_vals[0], errors='coerce')
                if not pd.isna(d):
                    record["Дата-1"] = d.strftime("%d-%m-%Y")
                record["Дата-2"] = pd.NA  # можно вообще не писать, но лучше занулить, чтобы колонка осталась
            elif len(date_vals) > 1:  # ← даты две и более
                d1 = pd.to_datetime(date_vals[0], errors='coerce')
                d2 = pd.to_datetime(date_vals[-1], errors='coerce')
                if not pd.isna(d1): record["Дата-1"] = d1.strftime("%d-%m-%Y")
                if not pd.isna(d2): record["Дата-2"] = d2.strftime("%d-%m-%Y")


            # Комментарий заказчика (если несколько колонок в строке — все собираем)
            cust_vals = [row[c] for c in body.columns if "Комментарий заказчика" in c and not pd.isna(row[c])]
            if cust_vals:
                record["Комментарий заказчика"] = " | ".join(cust_vals)

            # Ответ проектной (если несколько колонок в строке — все собираем)
            ans_vals = [row[c] for c in body.columns if "Ответ проектной" in c and not pd.isna(row[c])]
            if ans_vals:
                record["Ответ проектной организации"] = " | ".join(ans_vals)

            records.append(record)

    # Если ни одна строка не подошла — файл игнорируем
    if not records:
        print(f"В файле {os.path.basename(file_path)} не найдено строк с числовым №.")
        return None

    return records  # ← возвращаем список записей, а не одну!


def process_folder(folder_path, output_file):
    files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("объединенный файл")]
    records = []
    for name in files:
        path = os.path.join(folder_path, name)
        r = extract_file_data(path)
        if r:
            records.extend(r)  # ← было append (добавляло 1 объект), теперь extend (добавляет все строки)


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
