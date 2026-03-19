# -*- coding: utf-8 -*-
import re
import ast
from openpyxl.utils import get_column_letter

def sanitize_excel_text(text):
    """Excelに書き込む際にエラーとなる制御文字（IllegalCharacterErrorの原因）を除去する"""
    if text is None: return ""
    # 0x00-0x08, 0x0B-0x0C, 0x0E-0x1F を除去。タブ(\t)、改行(\n)、復帰(\r)は保持
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(text))

def auto_adjust_excel_column_width(ws):
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    for line in str(cell.value).split('\n'):
                        length = sum(2 if ord(c) > 255 else 1 for c in line)
                        if length > max_length: max_length = length
            except: pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

def analyze_column_profile(col_data):
    if not col_data: return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
    pure_num_cnt, fraction_cnt, total_num_len, total_cells = 0, 0, 0, 0
    for val in col_data:
        s = str(val).strip()
        if not s or s == "None": continue
        total_cells += 1
        if re.match(r'^\d+/\d+$', s):
            fraction_cnt += 1; total_num_len += len(s)
        else:
            s_clean = s.replace(",", "").replace(".", "", 1).replace("-", "", 1)
            if s_clean.isdigit(): pure_num_cnt += 1; total_num_len += len(s_clean)
    if total_cells == 0: return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
    return {
        "pure_num_ratio": pure_num_cnt / total_cells,
        "fraction_ratio": fraction_cnt / total_cells,
        "avg_num_len": total_num_len / (pure_num_cnt + fraction_cnt) if (pure_num_cnt + fraction_cnt) > 0 else 0.0,
        "is_text": ((pure_num_cnt / total_cells) < 0.2 and (fraction_cnt / total_cells) < 0.2)
    }

def get_profile_similarity(p1, p2):
    diff_pure = abs(p1["pure_num_ratio"] - p2["pure_num_ratio"])
    diff_frac = abs(p1["fraction_ratio"] - p2["fraction_ratio"])
    max_len = max(p1["avg_num_len"], p2["avg_num_len"])
    diff_len = abs(p1["avg_num_len"] - p2["avg_num_len"]) / max_len if max_len > 0 else 0.0
    return max(0.0, 1.0 - (diff_pure * 0.4 + diff_frac * 0.4 + diff_len * 0.2))

def parse_row_data(row_data):
    if isinstance(row_data, list) and len(row_data) == 1: row_data = row_data[0]
    if isinstance(row_data, str):
        row_data = row_data.strip()
        if (row_data.startswith('(') and row_data.endswith(')')) or (row_data.startswith('[') and row_data.endswith(']')):
            try:
                parsed = ast.literal_eval(row_data)
                if isinstance(parsed, (list, tuple)): return [str(x) if x is not None else "" for x in parsed]
            except:
                return [x.strip().strip("'\"") for x in row_data.strip("()[]").split(",")]
        return [row_data]
    if isinstance(row_data, tuple): return [str(x) if x is not None else "" for x in row_data]
    if not isinstance(row_data, list): return [str(row_data)]
    return [str(x) if x is not None else "" for x in row_data]

def apply_text_inheritance(final_aggregated_data):
    if len(final_aggregated_data) <= 1: return
    def is_text_to_inherit(text):
        s = str(text).strip()
        if not s or s in ["〃", "”", "\"", "''", "””", "''", "同上"]: return False
        return bool(re.search(r'[a-zA-Zａ-ｚＡ-Ｚぁ-んァ-ン一-龥]', s))
    header = final_aggregated_data[0]
    skip_cols = {idx for idx, h in enumerate(header) if "備考" in str(h)}
    for col_idx in range(1, len(header)):
        if col_idx in skip_cols: continue
        last_text = ""
        for row_idx in range(1, len(final_aggregated_data)):
            cell_val = str(final_aggregated_data[row_idx][col_idx]).strip()
            if cell_val in ["", "None", "〃", "”", "\"", "''", "””", "''", "同上"]:
                if last_text: final_aggregated_data[row_idx][col_idx] = last_text
            else:
                last_text = cell_val if is_text_to_inherit(cell_val) else ""

def merge_2d_arrays_horizontally(arrays_list):
    if not arrays_list: return []
    max_rows = max((len(arr) for arr in arrays_list), default=0)
    merged = []
    region_max_cols = [max((len(row) for row in arr), default=0) if arr else 0 for arr in arrays_list]
    for r in range(max_rows):
        merged_row = []
        for i, arr in enumerate(arrays_list):
            max_c = region_max_cols[i]
            if arr and r < len(arr):
                row_data = list(arr[r])
                row_data += [""] * (max_c - len(row_data))
                merged_row.extend(row_data)
            else: merged_row.extend([""] * max_c)
        merged.append(merged_row)
    return merged