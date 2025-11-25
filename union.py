import os
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from collections import Counter, defaultdict

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
        cell_value = ws[f'A{row}'].value
        if cell_value and str(cell_value).startswith('№'):
            headers = []
            col = 1
            while True:
                cell = ws.cell(row=row, column=col)
                if cell.value is None:
                    break
                headers.append(str(cell.value))
                col += 1
            return row - 1, headers
    return 0, []

def get_max_headers(folder_path):
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    max_headers = []
    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        _, headers = find_header_info(file_path)
        if len(headers) > len(max_headers):
            max_headers = headers
    return make_unique_columns(max_headers)

def merge_excel_files(folder_path, output_file, max_headers):
    all_dfs = []
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    for file_name in excel_files:
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
        sections = []
        i = 0
        while i < len(df):
            if df.iloc[i, 0] == "№":
                start = i
                # find end: next "Дата"
                end = len(df)
                for j in range(i+1, len(df)):
                    if df.iloc[j, 0] == "Дата":
                        end = j
                        break
                headers = df.iloc[start]
                data = df.iloc[start+1:end]
                columns = make_unique_columns(list(headers.values))
                part_df = pd.DataFrame(data.values, columns=columns)
                sections.append(part_df)
                i = end
            else:
                i += 1
        for section in sections:
            section_reindexed = section.reindex(columns=max_headers, fill_value=pd.NA)
            all_dfs.append(section_reindexed)

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
