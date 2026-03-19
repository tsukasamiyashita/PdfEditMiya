# -*- coding: utf-8 -*-
import os, sys, cv2, csv, time, json, random, gc, re
import numpy as np
import fitz
import pytesseract
import openpyxl
from openpyxl import Workbook
from PIL import Image
import google.generativeai as genai

from utils import (
    auto_adjust_excel_column_width,
    analyze_column_profile,
    get_profile_similarity,
    parse_row_data,
    apply_text_inheritance,
    merge_2d_arrays_horizontally,
    sanitize_excel_text
)

# ==============================
# AI抽出・データ集約タスク
# ==============================
def check_tesseract_installation():
    if sys.platform.startswith("win"):
        tess_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(tess_path): pytesseract.pytesseract.tesseract_cmd = tess_path
    try: pytesseract.get_tesseract_version()
    except Exception: raise Exception("Tesseract OCRがインストールされていないか、PATHが通っていません。")

def extract_tesseract_task(files, save_dir, options, ui):
    check_tesseract_installation()
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    out_format = options.get("out_format", "xlsx")
    crop_regions = options.get("crop_regions", [])
    
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(f))[0]
        doc = fitz.open(f)
        digits = max(2, len(str(len(doc))))
        total_pages = len(doc)
        for page_num in range(total_pages):
            ui.set_indeterminate(f"Tesseractで解析中... ( {page_num+1} / {total_pages} ページ )")
            pix = doc[page_num].get_pixmap(dpi=300)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4: img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
            elif pix.n == 1: img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
            
            cropped_images = []
            if crop_regions:
                h, w = img_array.shape[:2]
                for (rx1, ry1, rx2, ry2) in crop_regions:
                    cropped_images.append(img_array[int(min(ry1, ry2)*h):int(max(ry1, ry2)*h), int(min(rx1, rx2)*w):int(max(rx1, rx2)*w)])
            else: cropped_images.append(img_array)
                
            all_regions_data = []
            for crop_img in cropped_images:
                try:
                    text = pytesseract.image_to_string(Image.fromarray(crop_img), lang="jpn+eng")
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    if lines:
                        if (crop_regions and crop_img.shape[0] > crop_img.shape[1] * 1.5) or (len(lines) > 1 and all(len(l) <= 2 for l in lines)):
                            all_regions_data.append([["".join(lines)]])
                        else: all_regions_data.append([[l] for l in lines])
                    else: all_regions_data.append([[""]])
                except Exception as e: raise Exception(f"Tesseract OCRエラー: {e}")
                    
            merged_data = merge_2d_arrays_horizontally(all_regions_data)
            final_data = []
            page_info = f"{page_num+1}/{len(doc)}"
            if crop_regions:
                final_data.append(["ページ番号"] + [f"抽出範囲{idx+1}" for idx in range(len(cropped_images))])
            else:
                final_data.append(["ページ番号", "抽出テキスト"])
            for row in merged_data: final_data.append([page_info] + row)
            
            save_path = os.path.join(save_dir, f"{base}_Page_{str(page_num+1).zfill(digits)}_Tesseract抽出")
            if out_format == "xlsx":
                wb = Workbook(); ws = wb.active; ws.title = f"Page_{str(page_num+1).zfill(digits)}"
                for r_idx, row_data in enumerate(final_data, 1):
                    for c_idx, val in enumerate(row_data, 1): 
                        ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
                auto_adjust_excel_column_width(ws); wb.save(f"{save_path}.xlsx")
            elif out_format == "csv":
                with open(f"{save_path}.csv", "w", encoding="utf-8-sig", newline="") as f_out: csv.writer(f_out).writerows(final_data)
            elif out_format == "txt":
                with open(f"{save_path}.txt", "w", encoding="utf-8") as f_out:
                    for row_data in final_data: f_out.write("\t".join(row_data) + "\n")
        doc.close(); gc.collect()
        ui.set_determinate(total_pages, total_pages, "完了")

def extract_gemini_task(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    genai.configure(api_key=options.get("api_key", ""))
    
    base_models_to_try = options.get("models_to_try", ['gemini-1.5-pro', 'gemini-1.5-flash'])
    valid_models_to_try = []
    try:
        available_models = [m.name.replace('models/', '') for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        for bm in base_models_to_try:
            if bm in available_models: valid_models_to_try.append(bm)
            elif f"{bm}-latest" in available_models: valid_models_to_try.append(f"{bm}-latest")
        if not valid_models_to_try:
            safe_fallbacks = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-flash-latest']
            for fb in safe_fallbacks:
                if fb in available_models:
                    valid_models_to_try.append(fb)
                    break
        if not valid_models_to_try:
            flash_models = [m for m in available_models if 'flash' in m and 'image' not in m and 'embedding' not in m]
            if flash_models: valid_models_to_try.append(flash_models[0])
            elif available_models: valid_models_to_try.append(available_models[0])
    except Exception:
        valid_models_to_try = ['gemini-1.5-flash', 'gemini-1.5-pro']
        
    if not valid_models_to_try: valid_models_to_try = ['gemini-1.5-flash']

    out_format = options.get("out_format", "xlsx")
    crop_regions = options.get("crop_regions", [])

    if out_format in ["csv", "xlsx"]:
        if crop_regions:
            prompt = """あなたは優秀なデータ入力オペレーターです。添付画像のテキストを読み取りJSONを作成してください。
            【特別ルール】出力は「1つの列」にまとめ、セル分けしないでください。縦書きテキストは繋げて横書きに変換してください。
            【出力形式】 {"rows": ["1行目のテキスト...", "2行目のテキスト..."]}"""
        else:
            prompt = """あなたは優秀なデータ入力オペレーターです。表画像を読み取りJSONを作成してください。
            【特別ルール】見た目より「物理的な列の区切り(罫線・空白)」を最優先し分割してください。空白セルは `""` を挿入してください。縦書きテキストは繋げて横書きにしてください。
            【出力形式】 {"header": ["列1", ...], "rows": [["データ1", ""], ...]}"""
        generation_config = {"response_mime_type": "application/json"}
    else:
        prompt = "この画像の文字を可能な限り正確に読み取りプレーンテキストとして出力してください。縦書きの文章は横書きに変換してください。"
        generation_config = None

    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(f))[0]
        doc = fitz.open(f)
        digits = max(2, len(str(len(doc))))
        total_pages = len(doc)
        for page_num in range(total_pages):
            pix = doc[page_num].get_pixmap(dpi=300)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4: img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
            elif pix.n == 1: img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)

            cropped_images = []
            if crop_regions:
                h, w = img_array.shape[:2]
                for (rx1, ry1, rx2, ry2) in crop_regions:
                    cropped_images.append(img_array[int(min(ry1, ry2)*h):int(max(ry1, ry2)*h), int(min(rx1, rx2)*w):int(max(rx1, rx2)*w)])
            else: cropped_images.append(img_array)
                
            all_regions_data = []
            for region_idx, crop_img_array in enumerate(cropped_images):
                gray = cv2.cvtColor(crop_img_array, cv2.COLOR_RGB2GRAY)
                blurred = cv2.GaussianBlur(gray, (3, 3), 0)
                clean_bg = np.where(cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 5) == 255, 255, gray)
                enhanced = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(clean_bg.astype(np.uint8))
                img = Image.fromarray(cv2.cvtColor(cv2.addWeighted(enhanced, 1.5, cv2.GaussianBlur(enhanced, (0, 0), 2), -0.5, 0), cv2.COLOR_GRAY2RGB))
                
                extracted_text, success, last_error = "", False, ""
                for attempt in range(4):
                    for model_name in valid_models_to_try:
                        ui.set_indeterminate(f"AI解析中... ( {page_num+1}/{total_pages}頁 | 領域 {region_idx+1}/{len(cropped_images)} )")
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content([prompt, img], generation_config=generation_config) if generation_config else model.generate_content([prompt, img])
                            if not response.parts: raise Exception("安全フィルタブロック")
                            extracted_text = response.text.strip(); success = True; break
                        except Exception as api_err: 
                            err_str = str(api_err)
                            if "429" in err_str or "Quota" in err_str: last_error = f"API制限(429): {model_name}の利用枠上限に達しました。"
                            elif "404" in err_str: last_error = f"モデル未発見(404): {model_name}が利用できません。"
                            else: last_error = err_str
                            continue
                    if success: break
                    time.sleep((2 ** attempt) + random.uniform(0.5, 1.5))

                if success:
                    time.sleep(4.0) 
                    if out_format in ["xlsx", "csv"]:
                        try:
                            data = json.loads(extracted_text)
                            if crop_regions:
                                rows = data.get("rows", []) if isinstance(data, dict) else data
                                page_data_to_write = []
                                h_c, w_c = crop_img_array.shape[:2]
                                clean_rows = ["".join([l.strip() for l in (" ".join(str(x) for x in r) if isinstance(r, list) else str(r)).split('\n') if l.strip()]) if all(len(l)<=2 for l in (" ".join(str(x) for x in r) if isinstance(r, list) else str(r)).split('\n')) else (" ".join(str(x) for x in r) if isinstance(r, list) else str(r)) for r in rows]
                                if h_c > w_c * 1.5 and all(len(x.strip()) <= 2 for x in clean_rows if x.strip()): page_data_to_write.append(["".join(clean_rows)])
                                else:
                                    for val in clean_rows: page_data_to_write.append([val])
                                all_regions_data.append(page_data_to_write)
                            else:
                                header = data.get("header", [])
                                rows = data.get("rows", [])
                                if not header and not rows and isinstance(data, list) and data:
                                    header, rows = (data[0] if isinstance(data[0], list) else [str(data[0])]), data[1:]
                                safe_header = [str(x).strip() for x in header] if isinstance(header, list) else []
                                page_col_count = max(len(safe_header), max([len(r) for r in rows if isinstance(r, list)] + [0]))
                                if not safe_header: safe_header = [f"列{i+1}" for i in range(page_col_count)]
                                page_data_to_write = [(safe_header + [""] * page_col_count)[:page_col_count]]
                                for row_data in rows:
                                    safe_row = [("".join([l.strip() for l in str(v).split('\n')]) if '\n' in str(v) and len([l for l in str(v).split('\n') if l.strip()]) > 1 and all(len(l) <= 2 for l in str(v).split('\n') if l.strip()) else str(v)) for v in (parse_row_data(row_data) + [""] * page_col_count)[:page_col_count]]
                                    if any(v != "" for v in safe_row): page_data_to_write.append(safe_row)
                                all_regions_data.append(page_data_to_write)
                        except Exception as e: all_regions_data.append([[f"JSONパースエラー: {e}"]])
                    else: all_regions_data.append([[line] for line in extracted_text.split('\n')])
                else: all_regions_data.append([[f"抽出失敗: {last_error}"]])
            
            merged_data = merge_2d_arrays_horizontally(all_regions_data)
            final_data = []
            page_info = f"{page_num+1}/{len(doc)}"
            if crop_regions:
                final_data.append(["ページ番号"] + [f"抽出範囲{idx+1}" for idx in range(len(cropped_images))])
                for row in merged_data: final_data.append([page_info] + row)
            else:
                for r_idx, row in enumerate(merged_data):
                    final_data.append(["ページ番号" if r_idx == 0 and out_format in ["xlsx", "csv"] else page_info] + row)
            
            save_path = os.path.join(save_dir, f"{base}_Page_{str(page_num+1).zfill(digits)}_AI抽出")
            if out_format == "xlsx":
                wb = Workbook(); ws = wb.active; ws.title = f"Page_{str(page_num+1).zfill(digits)}"
                for r_idx, row_data in enumerate(final_data, 1):
                    for c_idx, val in enumerate(row_data, 1): 
                        ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
                auto_adjust_excel_column_width(ws); wb.save(f"{save_path}.xlsx")
            elif out_format == "csv":
                with open(f"{save_path}.csv", "w", encoding="utf-8-sig", newline="") as f_out: csv.writer(f_out).writerows(final_data)
            elif out_format == "txt":
                with open(f"{save_path}.txt", "w", encoding="utf-8") as f_out:
                    for row_data in final_data: f_out.write("\t".join(row_data) + "\n")
        doc.close(); gc.collect()
        ui.set_determinate(total_pages, total_pages, "完了")

def aggregate_only_task(files, save_dir, options, ui):
    out_format = options.get("out_format", "xlsx")
    search_ext = "xlsx" if out_format in ["jpg", "png", "dxf"] else out_format
    
    if not files: 
        raise Exception("処理対象のファイルまたはフォルダが選択されていません。")

    target_files_set = set()
    
    # 1. UIから渡されたファイル・フォルダリストを素直に展開し、対象拡張子のファイルだけを確実に追加
    for f in files:
        if os.path.isdir(f):
            try:
                for fn in os.listdir(f):
                    if fn.lower().endswith(f".{search_ext}") and "データ集約" not in fn and not fn.startswith("~$"):
                        target_files_set.add(os.path.abspath(os.path.join(f, fn)))
            except Exception:
                pass
        elif os.path.isfile(f):
            if f.lower().endswith(f".{search_ext}") and "データ集約" not in os.path.basename(f) and not os.path.basename(f).startswith("~$"):
                target_files_set.add(os.path.abspath(f))

    # ファイル名の昇順（アルファベット・数字順）でソートし、上から下へ順番に結合する
    target_files = sorted(list(target_files_set))
    
    if not target_files: 
        raise Exception(f"選択したファイルやフォルダ内に集約可能な (.{search_ext}) データが見つかりません。")

    agg_header, agg_rows, agg_texts = ["元ファイル名"], [], []
    
    # 2. レイアウト崩れを完全に防ぐ、シンプルかつ確実なマッピングロジック
    def map_to_master(fname, curr_header, curr_rows):
        if not curr_header:
            curr_header = [f"列{i+1}" for i in range(max(1, len(curr_rows[0]) if curr_rows else 1))]
            
        safe_header = [str(h).strip() if h is not None else "" for h in parse_row_data(curr_header)]
        for i in range(len(safe_header)):
            if not safe_header[i]: safe_header[i] = f"列{i+1}"
            
        col_mapping = {}
        for i, h in enumerate(safe_header):
            match_idx = -1
            
            # 【絶対ルール1】ヘッダー名が完全に一致する列があれば、そこへ結合
            for m_idx, m_h in enumerate(agg_header):
                if m_idx == 0: continue
                if h == str(m_h).strip() and h != "":
                    match_idx = m_idx
                    break
            
            # 【絶対ルール2】一致する名前が無ければ、同じ「列の位置（インデックス）」へ強制結合
            if match_idx == -1:
                target_idx = i + 1
                if target_idx < len(agg_header):
                    match_idx = target_idx
                    # マスター側のヘッダーが仮の名前「列X」だった場合は、ちゃんとした名前で上書き
                    if str(agg_header[target_idx]).startswith("列") and not h.startswith("列"):
                        agg_header[target_idx] = h
                else:
                    # マスターの列数が足りない場合のみ、右側に新しい列を追加
                    agg_header.append(h)
                    match_idx = len(agg_header) - 1
            
            col_mapping[i] = match_idx
                
        # データ行の追加
        for r in curr_rows:
            r_list = parse_row_data(r)
            row = [""] * len(agg_header)
            row[0] = fname  # 最初の列はファイル名
            for i, val in enumerate(r_list):
                if i in col_mapping:
                    m_idx = col_mapping[i]
                    if m_idx >= len(row): 
                        row.extend([""] * (m_idx - len(row) + 1))
                    row[m_idx] = str(val).strip() if val is not None and str(val).strip() != "None" else ""
            
            # 全てが空欄の行でなければマスターデータに追加
            if any(v != "" for v in row[1:]): 
                agg_rows.append(row)

    # 3. 各ファイルの読み込みと集約実行
    for i, f in enumerate(target_files, 1):
        ui.update_overall(i, len(target_files), f"データを集約中... ( {i} / {len(target_files)} ファイル )")
        ui.set_determinate(50, 100, f"読み込み中: {os.path.basename(f)}")
        
        fname = os.path.basename(f)
        # 本アプリのAI抽出などでついた不要なサフィックスを消して見やすくする（外部ファイルはそのまま）
        fname = re.sub(r'(_Page_\d+)?(_AI抽出|_Tesseract抽出|_Excel|_CSV|_Text)\.' + search_ext + '$', '.pdf', fname)
        
        try:
            if search_ext == "xlsx":
                wb = openpyxl.load_workbook(f, data_only=True)
                for sheet in wb.sheetnames:
                    rows = list(wb[sheet].iter_rows(values_only=True))
                    if rows and len(rows) > 0:
                        if len(rows) > 1: map_to_master(fname, rows[0], rows[1:])
                        else: map_to_master(fname, rows[0], [])
                wb.close()
            elif search_ext == "csv":
                with open(f, "r", encoding="utf-8-sig") as f_in:
                    rows = list(csv.reader(f_in))
                    if rows and len(rows) > 0:
                        if len(rows) > 1: map_to_master(fname, rows[0], rows[1:])
                        else: map_to_master(fname, rows[0], [])
            elif search_ext == "txt":
                with open(f, "r", encoding="utf-8") as f_in: 
                    agg_texts.append(f"[{fname}]\n{f_in.read()}")
        except Exception as e: 
            print(f"Read Error in {f}: {e}")
            
        ui.set_determinate(100, 100, "完了"); time.sleep(0.05)
        
    # 4. 集約データの保存
    if len(target_files) > 0 and save_dir:
        ui.set_indeterminate("集約データを保存中...")
        if search_ext in ["xlsx", "csv"]:
            # 行の長さをヘッダーに合わせる
            final_data = [agg_header] + [(r + [""] * len(agg_header))[:len(agg_header)] for r in agg_rows]
            apply_text_inheritance(final_data)
            
            if search_ext == "xlsx" and len(final_data) > 1:
                wb = Workbook(); ws = wb.active; ws.title = "集約データ"
                for r_idx, r_data in enumerate(final_data, 1):
                    for c_idx, val in enumerate(r_data, 1): 
                        # エラーとなる特殊文字を除外して書き込み
                        ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
                auto_adjust_excel_column_width(ws); wb.save(os.path.join(save_dir, "データ集約.xlsx"))
            elif search_ext == "csv" and len(final_data) > 1:
                with open(os.path.join(save_dir, "データ集約.csv"), "w", encoding="utf-8-sig", newline="") as f_out: 
                    csv.writer(f_out).writerows(final_data)
        elif search_ext == "txt" and agg_texts:
            with open(os.path.join(save_dir, "データ集約.txt"), "w", encoding="utf-8") as f_out: 
                f_out.write("\n\n".join(agg_texts))