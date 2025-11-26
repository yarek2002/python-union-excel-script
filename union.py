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
                cc = c
                while cc <= ws.max_column:
                    cell = ws.cell(r, cc)
                    if cell.value is None:
                        break
                    headers.append(str(cell.value))
                    cc += 1
                return r - 1, start_col - 1, headers
    return 0, 0, []

def get_all_files(folder_path):
    return [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

def merge_excel_files(folder_path, output_file):
    all_dfs = []
    files = get_all_files(folder_path)

    for name in files:
        # не обрабатывать выходной файл, если он в той же папке
        if name.startswith("объединенный файл"):
            continue

        path = os.path.join(folder_path, name)
        try:
            raw = pd.read_excel(path, header=None, engine="openpyxl", dtype=str)
        except Exception as e:
            print(f"Не удалось прочитать {name}: {e}")
            continue

        header_start, start_col, headers = find_header_info(path)
        if not headers:
            print(f"В файле {name} не найдена шапка (№...), пропускаю.")
            continue

        # уникализируем заголовки внутри файла: Дата -> Дата-1, Дата-2 ...
        headers_unique = unique_within_file(headers)

        # позиции колонок с "Дата" (теперь с суффиксом, но проверяем по началу)
        date_positions = [i for i, h in enumerate(headers_unique) if h.startswith("Дата")]

        sections = []
        idx = 0
        # нарезаем секции до каждой Даты (включительно)
        for pos in date_positions:
            cols = headers_unique[idx:pos + 1]
            cs = start_col + idx
            ce = cs + len(cols)
            sec = raw.iloc[header_start + 1:, cs:ce].copy()
            sec.columns = cols
            sec = sec.dropna(how="all")

            # для первой секции — обрезаем по первому нечисловому в колонке №
            if idx == 0:
                stop = None
                for i in range(len(sec)):
                    if not is_numeric(sec.iloc[i, 0]):
                        stop = i
                        break
                if stop is not None:
                    sec = sec.iloc[:stop]

            # обрезаем секцию по первой полностью пустой строке
            stop = None
            for i in range(len(sec)):
                if sec.iloc[i].isna().all():
                    stop = i
                    break
            if stop is not None:
                sec = sec.iloc[:stop]

            # если секция пустая — создаём пустой DF с нужными столбцами, чтобы concat не падал
            if sec.shape[0] == 0:
                sec = pd.DataFrame(columns=cols)

            sections.append(sec)
            idx = pos + 1

        # последняя секция после последней "Дата"
        if idx < len(headers_unique):
            cols = headers_unique[idx:]
            sec = raw.iloc[header_start + 1:, start_col + idx:].copy()
            sec.columns = cols
            sec = sec.dropna(how="all")
            stop = None
            for i in range(len(sec)):
                if sec.iloc[i].isna().all():
                    stop = i
                    break
            if stop is not None:
                sec = sec.iloc[:stop]
            if sec.shape[0] == 0:
                sec = pd.DataFrame(columns=cols)
            sections.append(sec)

        # если не найдено ни одной секции — пропускаем файл
        if not sections:
            print(f"В файле {name} не найдено секций после нарезки — пропускаю.")
            continue

        # теперь горизонтально склеиваем секции внутри этого файла
        try:
            file_df = pd.concat(sections, axis=1, ignore_index=False)
        except Exception as e:
            # для отладки: сохраним размеры секций
            sizes = [s.shape for s in sections]
            raise RuntimeError(f"Ошибка при горизонтальном concat в файле {name}. sizes={sizes}. error={e}")

        # вставляем имя файла и выравниваем индексы
        file_df.insert(0, "Файл", name)

        # удаляем полностью пустые колонки (если они появились)
        file_df = file_df.dropna(axis=1, how="all")

        all_dfs.append(file_df)

    if not all_dfs:
        merged_df = pd.DataFrame()
    else:
        # вертикально соединяем все файлы (каждый file_df должен иметь одинаковые кол-ва колонок, но если нет — pandas проставит NaN)
        merged_df = pd.concat(all_dfs, ignore_index=True, sort=False)

    # убираем время из колонок Дата-*
    if not merged_df.empty:
        date_columns = [c for c in merged_df.columns if str(c).startswith("Дата")]
        for col in date_columns:
            merged_df[col] = pd.to_datetime(merged_df[col], errors='coerce').dt.strftime("%d-%m-%Y")

    # сохраняем результат
    merged_df.to_excel(output_file, index=False)
    return output_file

if __name__ == "__main__":
    try:
        folder = os.getcwd()
        outname = f"объединенный файл {datetime.now().strftime('%Y-%m-%d')}.xlsx"
        print("Папка:", folder)
        print("Файлы для обработки:", get_all_files(folder))
        res = merge_excel_files(folder, outname)
        print("Успешно сохранено:", res)
    except Exception as exc:
        tb = traceback.format_exc()
        print("Скрипт упал с ошибкой:\n", tb)
        # записать лог для удобного анализа
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(tb)
        print("Трассировка записана в error_log.txt")
    finally:
        os.system("pause")
