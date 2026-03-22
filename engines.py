# -*- coding: utf-8 -*-
import os, sys, cv2, csv, json, time, gc, re
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
import fitz  # PyMuPDF
import pytesseract
import ezdxf
import numpy as np
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from PIL import Image

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import xlrd
except ImportError:
    xlrd = None

from common import (
    auto_adjust_excel_column_width,
    sanitize_excel_text,
    merge_2d_arrays_horizontally,
    parse_row_data,
    apply_text_inheritance
)

# ==============================
# コア処理関数群 (PDF編集・標準抽出)
# ==============================
def merge_pdfs(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    writer = PdfWriter()
    ui.update_overall(0, len(files), f"全体の進捗 ( 0 / {len(files)} ファイル )")
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        reader = PdfReader(f)
        for j, p in enumerate(reader.pages, 1):
            ui.set_determinate(j, len(reader.pages), f"ファイルを結合中... ( {j} / {len(reader.pages)} ページ )")
            writer.add_page(p)
    ui.set_determinate(1, 1, "PDFを保存中...")
    if save_dir:
        with open(os.path.join(save_dir, f"{options.get('folder_name', 'Merged')}_Merge.pdf"), "wb") as out: writer.write(out)

def split_pdfs(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        reader = PdfReader(f)
        base = os.path.splitext(os.path.basename(f))[0]
        digits = max(2, len(str(len(reader.pages))))
        for n, p in enumerate(reader.pages, 1):
            ui.set_determinate(n, len(reader.pages), f"ファイルを分割中... ( {n} / {len(reader.pages)} ページ )")
            writer = PdfWriter(); writer.add_page(p)
            with open(os.path.join(save_dir, f"{base}_Split_{str(n).zfill(digits)}.pdf"), "wb") as out: writer.write(out)

def rotate_pdfs(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        reader = PdfReader(f); writer = PdfWriter()
        for j, p in enumerate(reader.pages, 1):
            ui.set_determinate(j, len(reader.pages), f"ページを回転中... ( {j} / {len(reader.pages)} ページ )")
            p.rotate(options.get("rotate_deg", 270)); writer.add_page(p)
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Rotate.pdf"), "wb") as out: writer.write(out)

def extract_text_internal(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(f))[0]; text_list = []
        with pdfplumber.open(f) as pdf:
            for j, p in enumerate(pdf.pages, 1):
                ui.set_determinate(j, len(pdf.pages), f"テキストを抽出中... ( {j} / {len(pdf.pages)} ページ )")
                if crop_regions:
                    page_texts = []
                    for (rx1, ry1, rx2, ry2) in crop_regions:
                        x0, top, x1, bottom = min(rx1, rx2) * p.width, min(ry1, ry2) * p.height, max(rx1, rx2) * p.width, max(ry1, ry2) * p.height
                        txt = p.crop((x0, top, x1, bottom)).extract_text()
                        if txt: page_texts.append(txt)
                    if page_texts: text_list.append("\n".join(page_texts))
                else:
                    txt = p.extract_text()
                    if txt: text_list.append(txt)
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out: out.write("\n".join(text_list))

def convert_to_excel_internal(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    border_style = Side(border_style="thin", color="000000")
    crop_regions = options.get("crop_regions", [])
    for i, pdf_path in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        wb = Workbook(); wb.remove(wb.active)
        with pdfplumber.open(pdf_path) as pdf:
            digits = max(2, len(str(len(pdf.pages))))
            for page_idx, page in enumerate(pdf.pages, 1):
                ui.set_determinate(page_idx, len(pdf.pages), f"表データをExcelへ変換中... ( {page_idx} / {len(pdf.pages)} ページ )")
                tables = []
                if crop_regions:
                    for (rx1, ry1, rx2, ry2) in crop_regions:
                        x0, top, x1, bottom = min(rx1, rx2) * page.width, min(ry1, ry2) * page.height, max(rx1, rx2) * page.width, max(ry1, ry2) * page.height
                        tbls = page.crop((x0, top, x1, bottom)).extract_tables()
                        if tbls: tables.extend(tbls)
                else: tables = page.extract_tables()
                if not tables: continue
                ws = wb.create_sheet(f"Page_{str(page_idx).zfill(digits)}"); current_row = 1
                for table in tables:
                    for row_data in table:
                        for col_idx, cell_value in enumerate(row_data, 1):
                            clean_value = sanitize_excel_text(cell_value)
                            cell = ws.cell(row=current_row, column=col_idx, value=clean_value if clean_value else "")
                            cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                        current_row += 1
                    current_row += 2
                auto_adjust_excel_column_width(ws)
        if len(wb.sheetnames) > 0: wb.save(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_Excel.xlsx"))

def convert_to_csv_internal(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    for i, pdf_path in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        with pdfplumber.open(pdf_path) as pdf:
            digits = max(2, len(str(len(pdf.pages))))
            for page_idx, page in enumerate(pdf.pages, 1):
                ui.set_determinate(page_idx, len(pdf.pages), f"表データをCSVへ変換中... ( {page_idx} / {len(pdf.pages)} ページ )")
                tables = []
                if crop_regions:
                    for (rx1, ry1, rx2, ry2) in crop_regions:
                        x0, top, x1, bottom = min(rx1, rx2) * page.width, min(ry1, ry2) * page.height, max(rx1, rx2) * page.width, max(ry1, ry2) * page.height
                        tbls = page.crop((x0, top, x1, bottom)).extract_tables()
                        if tbls: tables.extend(tbls)
                else: tables = page.extract_tables()
                if not tables: continue
                with open(os.path.join(save_dir, f"{base}_Page_{str(page_idx).zfill(digits)}_CSV.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                    writer = csv.writer(f_out)
                    for table in tables:
                        for row_data in table: writer.writerow([str(cell).strip() if cell else "" for cell in row_data])
                        writer.writerow([]) 

def convert_to_image_jpg(files, save_dir, options, ui): _convert_image(files, save_dir, options, ui, "jpg")
def convert_to_image_png(files, save_dir, options, ui): _convert_image(files, save_dir, options, ui, "png")
def convert_to_image_tiff(files, save_dir, options, ui): _convert_image(files, save_dir, options, ui, "tiff")
def convert_to_image_bmp(files, save_dir, options, ui): _convert_image(files, save_dir, options, ui, "bmp")

def _convert_image(files, save_dir, options, ui, ext):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        doc = fitz.open(f)
        base = os.path.splitext(os.path.basename(f))[0]
        digits = max(2, len(str(len(doc))))
        for n, page in enumerate(doc, 1):
            ui.set_determinate(n, len(doc), f"画像へ変換中... ( {n} / {len(doc)} ページ )")
            n_str = str(n).zfill(digits)
            pix = page.get_pixmap(dpi=200)
            if crop_regions:
                img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                if pix.n == 4: img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
                elif pix.n == 1: img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
                h, w = img_array.shape[:2]
                for idx, (rx1, ry1, rx2, ry2) in enumerate(crop_regions):
                    x1, y1 = int(min(rx1, rx2) * w), int(min(ry1, ry2) * h)
                    x2, y2 = int(max(rx1, rx2) * w), int(max(ry1, ry2) * h)
                    Image.fromarray(img_array[y1:y2, x1:x2]).save(os.path.join(save_dir, f"{base}_{n_str}_crop{idx+1}.{ext}"))
            else: 
                # PNG, JPG 以外は PIL を経由して保存
                if ext in ["tiff", "bmp"]:
                    Image.frombytes("RGB" if pix.n >= 3 else "L", [pix.width, pix.height], pix.samples).save(os.path.join(save_dir, f"{base}_{n_str}.{ext}"))
                else:
                    pix.save(os.path.join(save_dir, f"{base}_{n_str}.{ext}"))

def convert_to_svg(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        doc = fitz.open(f)
        base = os.path.splitext(os.path.basename(f))[0]
        digits = max(2, len(str(len(doc))))
        for n, page in enumerate(doc, 1):
            ui.set_determinate(n, len(doc), f"SVGへ変換中... ( {n} / {len(doc)} ページ )")
            n_str = str(n).zfill(digits)
            if crop_regions:
                h, w = page.rect.height, page.rect.width
                for idx, (rx1, ry1, rx2, ry2) in enumerate(crop_regions):
                    rect = fitz.Rect(min(rx1, rx2)*w, min(ry1, ry2)*h, max(rx1, rx2)*w, max(ry1, ry2)*h)
                    page.set_cropbox(rect)
                    svg_xml = page.get_svg_image()
                    with open(os.path.join(save_dir, f"{base}_{n_str}_crop{idx+1}.svg"), "w", encoding="utf-8") as svg_file:
                        svg_file.write(svg_xml)
                    page.set_cropbox(page.mediabox) # 元に戻す
            else:
                svg_xml = page.get_svg_image()
                with open(os.path.join(save_dir, f"{base}_{n_str}.svg"), "w", encoding="utf-8") as svg_file:
                    svg_file.write(svg_xml)

def convert_to_dxf(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    def in_crop(x, y, w, h):
        if not crop_regions: return True
        for (rx1, ry1, rx2, ry2) in crop_regions:
            if min(rx1, rx2)*w <= x <= max(rx1, rx2)*w and min(ry1, ry2)*h <= y <= max(ry1, ry2)*h: return True
        return False

    for i, f in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        try:
            doc = fitz.open(f); dwg = ezdxf.new('R2010'); msp = dwg.modelspace()
            for page_num, page in enumerate(doc, 1):
                ui.set_determinate(page_num, len(doc), f"DXFへ変換中... ( {page_num} / {len(doc)} ページ )")
                h, w = page.rect.height, page.rect.width
                paths = page.get_drawings()
                is_vector_rich = len(paths) > 0
                if is_vector_rich:
                    for path in paths:
                        for item in path["items"]:
                            if item[0] == "l":
                                p1, p2 = item[1], item[2]
                                if in_crop(p1.x, p1.y, w, h) or in_crop(p2.x, p2.y, w, h): msp.add_line((p1.x, h - p1.y), (p2.x, h - p2.y))
                            elif item[0] == "re":
                                rect = item[1]
                                if in_crop(rect.x0, rect.y0, w, h) or in_crop(rect.x1, rect.y1, w, h):
                                    msp.add_lwpolyline([(rect.x0, h - rect.y0), (rect.x1, h - rect.y0), (rect.x1, h - rect.y1), (rect.x0, h - rect.y1)], close=True)
                            elif item[0] == "c":
                                p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                                if any(in_crop(pt.x, pt.y, w, h) for pt in [p1, p2, p3, p4]):
                                    msp.add_spline([(p1.x, h - p1.y), (p2.x, h - p2.y), (p3.x, h - p3.y), (p4.x, h - p4.y)])
                if not is_vector_rich or len(paths) < 5:
                    pix = page.get_pixmap(dpi=300)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY) if pix.n >= 3 else img
                    if crop_regions:
                        mask = np.zeros_like(gray)
                        for (rx1, ry1, rx2, ry2) in crop_regions:
                            mask[int(min(ry1, ry2)*gray.shape[0]):int(max(ry1, ry2)*gray.shape[0]), int(min(rx1, rx2)*gray.shape[1]):int(max(rx1, rx2)*gray.shape[1])] = 255
                        gray = cv2.bitwise_and(gray, mask)
                    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
                    binary = cv2.morphologyEx(cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2), cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
                    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                    scale_x, scale_y = w / pix.w, h / pix.h
                    for cnt in contours:
                        if cv2.contourArea(cnt) < 15: continue 
                        pts = [(p[0][0] * scale_x, h - p[0][1] * scale_y) for p in cv2.approxPolyDP(cnt, 0.003 * cv2.arcLength(cnt, True), True)]
                        if len(pts) > 1: msp.add_lwpolyline(pts, close=True)
            ui.set_determinate(len(doc), len(doc), "DXFファイルを保存中...")
            dwg.saveas(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(f))[0]}_CAD.dxf"))
        except Exception as e: print(f"DXF Conversion Error: {e}")

# ==============================
# ローカルOCR抽出タスク (Tesseract)
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
        if ui.is_cancelled(): return
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(f))[0]
        doc = fitz.open(f)
        digits = max(2, len(str(len(doc))))
        total_pages = len(doc)
        for page_num in range(total_pages):
            if ui.is_cancelled(): return
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
            elif out_format == "json":
                with open(f"{save_path}.json", "w", encoding="utf-8") as f_out: json.dump(final_data, f_out, ensure_ascii=False, indent=2)
            elif out_format == "md":
                with open(f"{save_path}.md", "w", encoding="utf-8") as f_out:
                    if final_data:
                        f_out.write("| " + " | ".join(map(str, final_data[0])) + " |\n")
                        f_out.write("|" + "|".join(["---"] * len(final_data[0])) + "|\n")
                        for row in final_data[1:]: f_out.write("| " + " | ".join(map(lambda x: str(x).replace('\n', '<br>'), row)) + " |\n")
            elif out_format == "docx":
                if Document is None: raise Exception("python-docx ライブラリがインストールされていません。")
                doc_out = Document()
                if final_data:
                    table = doc_out.add_table(rows=len(final_data), cols=max(len(r) for r in final_data))
                    table.style = 'Table Grid'
                    for r_idx, row_data in enumerate(final_data):
                        row_cells = table.rows[r_idx].cells
                        for c_idx, val in enumerate(row_data):
                            if c_idx < len(row_cells):
                                row_cells[c_idx].text = str(val)
                doc_out.save(f"{save_path}.docx")
        doc.close(); gc.collect()
        ui.set_determinate(total_pages, total_pages, "完了")

# ==============================
# データ集約タスク (ローカル)
# ==============================
def aggregate_local_task(files, save_dir, options, ui):
    out_format = options.get("out_format", "xlsx")
    search_ext = "xlsx" if out_format in ["jpg", "png", "dxf", "svg", "tiff", "bmp"] else out_format
    
    # 実行時のタイムスタンプを取得
    run_timestamp = time.strftime("%Y%m%d_%H%M%S")
    
    if not files: raise Exception("処理対象のファイルまたはフォルダが選択されていません。")

    search_exts = ["xlsx", "xlsm", "xls"] if search_ext == "xlsx" else [search_ext]

    target_files_set = set()
    for f in files:
        if os.path.isdir(f):
            try:
                for fn in os.listdir(f):
                    if any(fn.lower().endswith(f".{ext}") for ext in search_exts) and "データ集約" not in fn and not fn.startswith("~$"):
                        target_files_set.add(os.path.abspath(os.path.join(f, fn)))
            except Exception: pass
        elif os.path.isfile(f):
            if any(f.lower().endswith(f".{ext}") for ext in search_exts) and "データ集約" not in os.path.basename(f) and not os.path.basename(f).startswith("~$"):
                target_files_set.add(os.path.abspath(f))

    target_files = sorted(list(target_files_set))
    if not target_files: raise Exception(f"選択したファイルやフォルダ内に集約可能な ({', '.join(['.' + ext for ext in search_exts])}) データが見つかりません。")

    agg_header, agg_rows, agg_texts = ["元ファイル名"], [], []
    
    def map_to_master(fname, curr_header, curr_rows):
        if not curr_header: curr_header = [f"列{i+1}" for i in range(max(1, len(curr_rows[0]) if curr_rows else 1))]
        safe_header = [str(h).strip() if h is not None else "" for h in parse_row_data(curr_header)]
        for i in range(len(safe_header)):
            if not safe_header[i]: safe_header[i] = f"列{i+1}"
            
        col_mapping = {}
        for i, h in enumerate(safe_header):
            match_idx = -1
            for m_idx, m_h in enumerate(agg_header):
                if m_idx == 0: continue
                if h == str(m_h).strip() and h != "": match_idx = m_idx; break
            
            if match_idx == -1:
                # 既存の列名と一致しない場合は、必ず新しい列として末尾に追加する
                agg_header.append(h)
                match_idx = len(agg_header) - 1
                
            col_mapping[i] = match_idx
                
        for r in curr_rows:
            r_list = parse_row_data(r)
            row = [""] * len(agg_header)
            row[0] = fname 
            for i, val in enumerate(r_list):
                if i in col_mapping:
                    m_idx = col_mapping[i]
                    if m_idx >= len(row): row.extend([""] * (m_idx - len(row) + 1))
                    row[m_idx] = str(val).strip() if val is not None and str(val).strip() != "None" else ""
            if any(v != "" for v in row[1:]): agg_rows.append(row)

    for i, f in enumerate(target_files, 1):
        if ui.is_cancelled(): return
        ui.update_overall(i, len(target_files), f"データを集約中... ( {i} / {len(target_files)} ファイル )")
        ui.set_determinate(50, 100, f"読み込み中: {os.path.basename(f)}")
        
        ext_lower = os.path.splitext(f)[1].lower().strip('.')
        fname = os.path.basename(f)
        fname = re.sub(r'(?i)(_Page_\d+)?(_AI抽出|_Tesseract抽出|_Excel|_CSV|_Text)\.' + ext_lower + '$', '.pdf', fname)
        
        try:
            if ext_lower in ["xlsx", "xlsm"]:
                wb = openpyxl.load_workbook(f, data_only=True)
                for sheet in wb.sheetnames:
                    all_rows = list(wb[sheet].iter_rows(values_only=True))
                    valid_rows = []
                    for r in all_rows:
                        if r and any(c is not None and str(c).strip() != "" for c in r):
                            valid_rows.append(r)
                    if valid_rows:
                        if len(valid_rows) > 1: map_to_master(fname, valid_rows[0], valid_rows[1:])
                        else: map_to_master(fname, valid_rows[0], [])
                wb.close()
            elif ext_lower == "xls":
                if xlrd is None: raise Exception("xlsファイルの読み込みには 'xlrd' が必要です。")
                wb = xlrd.open_workbook(f)
                for sheet_idx in range(wb.nsheets):
                    sheet = wb.sheet_by_index(sheet_idx)
                    all_rows = [sheet.row_values(r_idx) for r_idx in range(sheet.nrows)]
                    valid_rows = []
                    for r in all_rows:
                        if r and any(c is not None and str(c).strip() != "" for c in r):
                            valid_rows.append(r)
                    if valid_rows:
                        if len(valid_rows) > 1: map_to_master(fname, valid_rows[0], valid_rows[1:])
                        else: map_to_master(fname, valid_rows[0], [])
            elif ext_lower == "csv":
                rows = []
                try:
                    with open(f, "r", encoding="utf-8-sig") as f_in: rows = list(csv.reader(f_in))
                except UnicodeDecodeError:
                    try:
                        with open(f, "r", encoding="cp932") as f_in: rows = list(csv.reader(f_in))
                    except Exception:
                        with open(f, "r", encoding="utf-8", errors="ignore") as f_in: rows = list(csv.reader(f_in))
                
                valid_rows = []
                for r in rows:
                    if r and any(c.strip() != "" for c in r):
                        valid_rows.append(r)
                if valid_rows:
                    if len(valid_rows) > 1: map_to_master(fname, valid_rows[0], valid_rows[1:])
                    else: map_to_master(fname, valid_rows[0], [])
            elif search_ext == "json":
                with open(f, "r", encoding="utf-8") as f_in:
                    try:
                        data = json.load(f_in)
                        rows = data.get("rows", []) if isinstance(data, dict) else data
                        if rows and isinstance(rows, list) and len(rows) > 0:
                            if len(rows) > 1: map_to_master(fname, rows[0], rows[1:])
                            else: map_to_master(fname, rows[0], [])
                    except Exception: pass
            elif search_ext == "md":
                with open(f, "r", encoding="utf-8") as f_in:
                    rows = []
                    for line in f_in:
                        line = line.strip()
                        if line.startswith('|') and line.endswith('|'):
                            cols = [c.strip().replace('<br>', '\n') for c in line[1:-1].split('|')]
                            if all(c.strip() == '-' * len(c.strip()) or c.strip() == '' or ':' in c for c in cols): continue
                            rows.append(cols)
                    if rows and len(rows) > 0:
                        if len(rows) > 1: map_to_master(fname, rows[0], rows[1:])
                        else: map_to_master(fname, rows[0], [])
            elif search_ext == "docx":
                if Document is None: raise Exception("python-docx ライブラリがインストールされていません。")
                doc_in = Document(f)
                for table in doc_in.tables:
                    rows = []
                    for row in table.rows:
                        rows.append([cell.text for cell in row.cells])
                    if rows and len(rows) > 0:
                        if len(rows) > 1: map_to_master(fname, rows[0], rows[1:])
                        else: map_to_master(fname, rows[0], [])
            elif search_ext == "txt":
                text_content = ""
                try:
                    with open(f, "r", encoding="utf-8-sig") as f_in: text_content = f_in.read()
                except UnicodeDecodeError:
                    try:
                        with open(f, "r", encoding="cp932") as f_in: text_content = f_in.read()
                    except Exception:
                        with open(f, "r", encoding="utf-8", errors="ignore") as f_in: text_content = f_in.read()
                if text_content.strip():
                    agg_texts.append(f"[{fname}]\n{text_content}")
        except Exception as e: 
            print(f"Read Error in {f}: {e}")
            
        ui.set_determinate(100, 100, "完了"); time.sleep(0.05)
        
    if len(target_files) > 0 and save_dir:
        if ui.is_cancelled(): return
        ui.set_indeterminate("集約データを保存中...")
        if search_ext in ["xlsx", "csv", "json", "md", "docx"]:
            final_data = [agg_header] + [(r + [""] * len(agg_header))[:len(agg_header)] for r in agg_rows]
            
            # --- 修正: 空白セルの自動補完を廃止し、点々や同上記号の場合のみ前の行の値を引き継ぐ ---
            for r_idx in range(2, len(final_data)): # 0はヘッダー、1は最初のデータ行なので2行目から
                for c_idx in range(len(final_data[r_idx])):
                    val = str(final_data[r_idx][c_idx]).strip()
                    # 点々や同上を表す記号（〃, …, .., ・, 同上 など）の判定
                    if re.match(r'^([〃…・]+|\.+|,+|同上|”+|"+)$', val):
                        final_data[r_idx][c_idx] = final_data[r_idx-1][c_idx]
            # ---------------------------------------------------------------------------------
            
            if search_ext == "xlsx" and len(final_data) > 1:
                wb = Workbook(); ws = wb.active; ws.title = "集約データ"
                for r_idx, r_data in enumerate(final_data, 1):
                    for c_idx, val in enumerate(r_data, 1): 
                        ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
                auto_adjust_excel_column_width(ws); wb.save(os.path.join(save_dir, f"データ集約_{run_timestamp}.xlsx"))
            elif search_ext == "csv" and len(final_data) > 1:
                with open(os.path.join(save_dir, f"データ集約_{run_timestamp}.csv"), "w", encoding="utf-8-sig", newline="") as f_out: 
                    csv.writer(f_out).writerows(final_data)
            elif search_ext == "json" and len(final_data) > 1:
                with open(os.path.join(save_dir, f"データ集約_{run_timestamp}.json"), "w", encoding="utf-8") as f_out:
                    json.dump(final_data, f_out, ensure_ascii=False, indent=2)
            elif search_ext == "md" and len(final_data) > 1:
                with open(os.path.join(save_dir, f"データ集約_{run_timestamp}.md"), "w", encoding="utf-8") as f_out:
                    f_out.write("| " + " | ".join(map(str, final_data[0])) + " |\n")
                    f_out.write("|" + "|".join(["---"] * len(final_data[0])) + "|\n")
                    for row in final_data[1:]: f_out.write("| " + " | ".join(map(lambda x: str(x).replace('\n', '<br>'), row)) + " |\n")
            elif search_ext == "docx" and len(final_data) > 1:
                if Document is None: raise Exception("python-docx ライブラリがインストールされていません。")
                doc_out = Document()
                table = doc_out.add_table(rows=len(final_data), cols=max(len(r) for r in final_data))
                table.style = 'Table Grid'
                for r_idx, row_data in enumerate(final_data):
                    row_cells = table.rows[r_idx].cells
                    for c_idx, val in enumerate(row_data):
                        if c_idx < len(row_cells):
                            row_cells[c_idx].text = str(val)
                doc_out.save(os.path.join(save_dir, f"データ集約_{run_timestamp}.docx"))
        elif search_ext == "txt" and agg_texts:
            with open(os.path.join(save_dir, f"データ集約_{run_timestamp}.txt"), "w", encoding="utf-8") as f_out: 
                f_out.write("\n\n".join(agg_texts))