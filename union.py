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
    max_headers = []
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        _, _, headers = find_header_info(file_path)
        if len(headers) > len(max_headers):
            max_headers = headers
    return ['Файл'] + make_unique_columns(max_headers)

def merge_excel_files(folder_path, output_file, max_headers):
    all_dfs = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path, header=None, engine='openpyxl', dtype=str)
        header_start, start_col, headers = find_header_info(file_path)
        header_row = header_start
        sections = []
        if headers:
            positions = [i for i, h in enumerate(headers) if h == 'Дата']
            start_idx = 0
            for end_idx in positions:
                section_cols = headers[start_idx:end_idx + 1]
                col_start = start_col + start_idx
                col_end = col_start + len(section_cols)
                section_df = df.iloc[header_row + 1:, col_start:col_end].copy()
                section_df = section_df.dropna(how='all')  # drop empty rows
                section_df.columns = make_unique_columns(section_cols)
                if start_idx == 0:  # only for first section
                    # filter rows where first column is not numeric
                    stop_idx = None
                    for i in range(len(section_df)):
                        val = section_df.iloc[i, 0]
                        if pd.isna(val) or not is_numeric(val):
                            stop_idx = i
                            break
                    if stop_idx is not None:
                        section_df = section_df.iloc[:stop_idx]
                sections.append(section_df)
                start_idx = end_idx + 1 
            # last section
            if start_idx < len(headers):
                section_cols = headers[start_idx:]
                col_start = start_col + start_idx
                section_df = df.iloc[header_row + 1:, col_start:].copy()
                section_df = section_df.dropna(how='all')  # drop empty rows
                section_df.columns = make_unique_columns(section_cols)
                # filter rows where first column is not numeric
                stop_idx = None
                for i in range(len(section_df)):
                    val = section_df.iloc[i, 0]
                    if pd.isna(val) or not is_numeric(val):
                        stop_idx = i
                        break
                if stop_idx is not None:
                    section_df = section_df.iloc[:stop_idx]
                sections.append(section_df)
        if sections:
            # file_df = horizontal concat of sections
            file_df = pd.concat(sections, axis=1, ignore_index=False)
            file_df.columns = make_unique_columns(list(file_df.columns))
            file_df.insert(0, 'Файл', file_name)
            file_df_reindexed = file_df.reindex(columns=max_headers, fill_value=pd.NA)
            all_dfs.append(file_df_reindexed)

    if not all_dfs:
        all_dfs = [pd.DataFrame(columns=max_headers)]
    merged_df = pd.concat(all_dfs, ignore_index=True)
    merged_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    folder_path =	os.getcwd()  # current directory
    current_date = datetime.now().strftime("%Y-%m-%d")
    output_file = f"объединенный файл {current_date}.xlsx"
    max_headers = get_max_headers(folder_path)
    merge_excel_files(folder_path, output_file, max_headers)
    print(f"Объединенный файл сохранен как {output_file}")
