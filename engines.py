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
# 画像前処理タスク (OCR精度向上・クロップ拡張)
# ==============================
def expand_crop_rect_for_intersecting_objects(img_array, rx1, ry1, rx2, ry2):
    """
    指定された相対座標(rx1, ry1, rx2, ry2)のクロップ枠に対し、
    枠の境界に接している文字や図形（輪郭）が途切れないよう、
    輪郭が完全に含まれるようにピクセル座標レベルで枠を自動拡張して返す。
    """
    h, w = img_array.shape[:2]
    x1, y1 = int(min(rx1, rx2) * w), int(min(ry1, ry2) * h)
    x2, y2 = int(max(rx1, rx2) * w), int(max(ry1, ry2) * h)
    
    # 枠が画像全体に近い場合はそのまま返す
    if x1 == 0 and y1 == 0 and x2 == w and y2 == h:
        return x1, y1, x2, y2

    # グレースケール化
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:
        gray = img_array.copy()

    # コントラスト強調・ノイズ除去
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    gray = cv2.medianBlur(gray, 3)

    # 適応的二値化 (文字を白、背景を黒に)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    
    # 輪郭抽出
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    new_x1, new_y1, new_x2, new_y2 = x1, y1, x2, y2
    
    for cnt in contours:
        cx, cy, cw, ch = cv2.boundingRect(cnt)
        cx1, cy1, cx2, cy2 = cx, cy, cx + cw, cy + ch
        
        # 大きすぎる輪郭（ページ枠線や背景）や小さすぎる輪郭（ノイズ）は無視
        if cw > w * 0.5 or ch > h * 0.5: continue
        if cw < 3 or ch < 3: continue
            
        # 輪郭のバウンディングボックスが元のクロップ枠と交差（接触）しているか判定
        if (cx1 <= x2 and cx2 >= x1 and cy1 <= y2 and cy2 >= y1):
            new_x1 = min(new_x1, cx1)
            new_y1 = min(new_y1, cy1)
            new_x2 = max(new_x2, cx2)
            new_y2 = max(new_y2, cy2)
            
    # 画像サイズをはみ出さないように制限し、マージンを持たせる
    margin = 4
    new_x1 = max(0, new_x1 - margin)
    new_y1 = max(0, new_y1 - margin)
    new_x2 = min(w, new_x2 + margin)
    new_y2 = min(h, new_y2 + margin)
    
    return new_x1, new_y1, new_x2, new_y2

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
        is_scanned_pdf = False
        with pdfplumber.open(f) as pdf:
            for j, p in enumerate(pdf.pages, 1):
                ui.set_determinate(j, len(pdf.pages), f"テキストを抽出中... ( {j} / {len(pdf.pages)} ページ )")
                if not p.chars: is_scanned_pdf = True
                
                if crop_regions:
                    page_texts = []
                    for idx, (rx1, ry1, rx2, ry2) in enumerate(crop_regions, 1):
                        m = 0.005
                        x0 = max(0, min(rx1, rx2) - m) * p.width
                        top = max(0, min(ry1, ry2) - m) * p.height
                        x1 = min(1, max(rx1, rx2) + m) * p.width
                        bottom = min(1, max(ry1, ry2) + m) * p.height
                        
                        cropped_page = p.crop((x0, top, x1, bottom), strict=False)
                        txt = cropped_page.extract_text(layout=True)
                        if txt: page_texts.append(f"--- 範囲{idx} ---\n{txt}")
                    if page_texts: text_list.append(f"【Page {j}】\n" + "\n".join(page_texts))
                else:
                    txt = p.extract_text()
                    if txt: text_list.append(f"【Page {j}】\n{txt}")
        
        if text_list:
            output_content = "\n\n".join(text_list)
        elif is_scanned_pdf:
            output_content = "このPDFは「画像（スキャンされたPDF）」として保存されており、標準ライブラリでは文字を読み取れません。\n「Gemini API」または「Tesseract」エンジンを使用して再試行してください。"
        else:
            output_content = "指定された範囲内にテキストデータが見つかりませんでした。枠を少し広げて選択するか、AIエンジンの使用を検討してください。"
            
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out: out.write(output_content)

def convert_to_excel_internal(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    border_style = Side(border_style="thin", color="000000")
    crop_regions = options.get("crop_regions", [])
    for i, pdf_path in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        wb = Workbook(); wb.remove(wb.active)
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        
        has_any_table = False
        is_scanned_pdf = False
        with pdfplumber.open(pdf_path) as pdf:
            digits = max(2, len(str(len(pdf.pages))))
            for page_idx, page in enumerate(pdf.pages, 1):
                ui.set_determinate(page_idx, len(pdf.pages), f"表データを抽出中... ( {page_idx} / {len(pdf.pages)} ページ )")
                if not page.chars: is_scanned_pdf = True
                
                tables = []
                if crop_regions:
                    all_regions_data = []
                    for (rx1, ry1, rx2, ry2) in crop_regions:
                        m = 0.005
                        x0 = max(0, min(rx1, rx2) - m) * page.width
                        top = max(0, min(ry1, ry2) - m) * page.height
                        x1 = min(1, max(rx1, rx2) + m) * page.width
                        bottom = min(1, max(ry1, ry2) + m) * page.height
                        
                        cropped_page = page.crop((x0, top, x1, bottom), strict=False)
                        tbls = cropped_page.extract_tables()
                        
                        if not tbls:
                            tbl_settings = {"vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 3}
                            tbls = cropped_page.extract_tables(table_settings=tbl_settings)
                        
                        if not tbls:
                            txt = cropped_page.extract_text()
                            if txt and txt.strip():
                                dummy_table = [[line] for line in txt.strip().split('\n')]
                                tbls = [dummy_table]
                                
                        region_table = []
                        if tbls:
                            for tbl in tbls: region_table.extend(tbl)
                        
                        if not region_table: region_table = [[""]]
                        all_regions_data.append(region_table)
                        
                    merged_table = merge_2d_arrays_horizontally(all_regions_data)
                    is_empty = True
                    for row in merged_table:
                        for cell in row:
                            if cell and str(cell).strip():
                                is_empty = False; break
                        if not is_empty: break
                            
                    if not is_empty: tables = [merged_table]
                else: 
                    tables = page.extract_tables()
                    if not tables:
                        txt = page.extract_text()
                        if txt and txt.strip():
                            dummy_table = [[line] for line in txt.strip().split('\n')]
                            tables = [dummy_table]
                
                if not tables: continue
                has_any_table = True
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
        
        if has_any_table:
            wb.save(os.path.join(save_dir, f"{base_name}_Excel.xlsx"))
        else:
            empty_wb = Workbook(); ws = empty_wb.active; ws.title = "No_Data"
            if is_scanned_pdf: msg = "このPDFは「スキャンされた画像」です。Gemini APIまたはTesseractを使用してください。"
            else: msg = "指定範囲内に表構造やテキストデータが見つかりませんでした。"
            ws.cell(row=1, column=1, value=msg)
            empty_wb.save(os.path.join(save_dir, f"{base_name}_Excel_NoData.xlsx"))

def convert_to_csv_internal(files, save_dir, options, ui):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    crop_regions = options.get("crop_regions", [])
    for i, pdf_path in enumerate(files, 1):
        ui.update_overall(i, len(files), f"全体の進捗 ( {i} / {len(files)} ファイル )")
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        
        has_any_table = False
        with pdfplumber.open(pdf_path) as pdf:
            digits = max(2, len(str(len(pdf.pages))))
            for page_idx, page in enumerate(pdf.pages, 1):
                ui.set_determinate(page_idx, len(pdf.pages), f"CSV変換中... ( {page_idx} / {len(pdf.pages)} ページ )")
                tables = []
                if crop_regions:
                    all_regions_data = []
                    for (rx1, ry1, rx2, ry2) in crop_regions:
                        m = 0.005
                        x0, top, x1, bottom = (max(0, min(rx1, rx2)-m) * page.width, max(0, min(ry1, ry2)-m) * page.height, 
                                              min(1, max(rx1, rx2)+m) * page.width, min(1, max(ry1, ry2)+m) * page.height)
                        cropped_page = page.crop((x0, top, x1, bottom), strict=False)
                        
                        tbls = cropped_page.extract_tables()
                        if not tbls:
                            tbls = cropped_page.extract_tables(table_settings={"vertical_strategy": "text", "horizontal_strategy": "text"})
                        
                        if not tbls:
                            txt = cropped_page.extract_text()
                            if txt and txt.strip():
                                tbls = [[[line] for line in txt.strip().split('\n')]]
                                
                        region_table = []
                        if tbls:
                            for tbl in tbls: region_table.extend(tbl)
                                
                        if not region_table: region_table = [[""]]
                        all_regions_data.append(region_table)
                        
                    merged_table = merge_2d_arrays_horizontally(all_regions_data)
                    is_empty = True
                    for row in merged_table:
                        for cell in row:
                            if cell and str(cell).strip():
                                is_empty = False; break
                        if not is_empty: break
                            
                    if not is_empty: tables = [merged_table]
                else: 
                    tables = page.extract_tables()
                    if not tables:
                        txt = page.extract_text()
                        if txt and txt.strip(): tables = [[[line] for line in txt.strip().split('\n')]]
                
                if not tables: continue
                has_any_table = True
                with open(os.path.join(save_dir, f"{base}_Page_{str(page_idx).zfill(digits)}_CSV.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                    writer = csv.writer(f_out)
                    for table in tables:
                        for row_data in table: writer.writerow([str(cell).strip() if cell else "" for cell in row_data])
                        writer.writerow([]) 
        
        if not has_any_table:
            with open(os.path.join(save_dir, f"{base}_NoData_CSV.txt"), "w", encoding="utf-8") as f_out:
                f_out.write("データが見つかりませんでした。PDFがスキャン形式であるか、範囲が狭すぎる可能性があります。")

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
                    page.set_cropbox(page.mediabox) 
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
            ui.set_determinate(len(doc), len(doc), "DXFファイルを保存中...")
            dwg.saveas(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(f))[0]}_CAD.dxf"))
        except Exception as e: print(f"DXF Conversion Error: {e}")

# ==============================
# 画像前処理タスク (OCR精度向上)
# ==============================
def preprocess_image_for_ocr(img_array, is_digital=False):
    """
    デジタルPDFとスキャンPDFで適切な前処理を自動で使い分ける
    """
    if len(img_array.shape) == 3: gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else: gray = img_array
    
    if is_digital:
        # デジタル生成PDF（Excel等からの変換）の場合
        # 元々クリアな文字なので、過剰な補正(適応的二値化や平滑化)を行うと逆に文字が劣化します。
        # 大津の二値化のみを適用し、くっきりとした白黒画像にします。
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        return binary
    else:
        # スキャンされたPDFの場合
        # 影や照明ムラを取り除くため、コントラストを強調してから適応的二値化を行います。
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        gray = clahe.apply(gray)
        # 解像度が高い(dpi=400)場合、文字の線が太くなるため、ブロックサイズを大きく(51)して太い文字の中抜けを防止します。
        binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 51, 15)
        return binary

# ==============================
# ローカルOCR抽出タスク (Tesseract)
# ==============================
def check_tesseract_installation():
    if sys.platform.startswith("win"):
        tess_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if os.path.exists(tess_path): pytesseract.pytesseract.tesseract_cmd = tess_path
    try: pytesseract.get_tesseract_version()
    except Exception: raise Exception("Tesseract OCRが見つかりません。")

def extract_tesseract_task(files, save_dir, options, ui):
    check_tesseract_installation()
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFが含まれていません。")
    out_format = options.get("out_format", "xlsx")
    crop_regions = options.get("crop_regions", [])
    
    for i, f in enumerate(files, 1):
        if ui.is_cancelled(): return
        ui.update_overall(i, len(files), f"全体進捗 ( {i} / {len(files)} )")
        base = os.path.splitext(os.path.basename(f))[0]
        doc = fitz.open(f)
        total_pages = len(doc)
        for page_num in range(total_pages):
            if ui.is_cancelled(): return
            ui.set_indeterminate(f"OCR解析中... ({page_num+1}/{total_pages}ページ)")
            
            # 【改善】PDFページ自体にテキストデータが埋め込まれているか(デジタルPDFか)を自動判定
            page_obj = doc[page_num]
            is_digital = bool(page_obj.get_text().strip())
            
            # 【改善】DPIを300から400に引き上げ、Tesseractの認識精度を向上（文字が小さい表などに効果絶大）
            pix = page_obj.get_pixmap(dpi=400)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4: img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
            elif pix.n == 1: img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)
            
            cropped_images = []
            if crop_regions:
                h, w = img_array.shape[:2]
                for (rx1, ry1, rx2, ry2) in crop_regions:
                    if out_format not in ["xlsx", "csv"]: x1, y1, x2, y2 = expand_crop_rect_for_intersecting_objects(img_array, rx1, ry1, rx2, ry2)
                    else: x1, y1, x2, y2 = int(min(rx1, rx2)*w), int(min(ry1, ry2)*h), int(max(rx1, rx2)*w), int(max(ry1, ry2)*h)
                    cropped_images.append(img_array[y1:y2, x1:x2])
            else: cropped_images.append(img_array)
                
            all_regions_data = []
            for crop_img in cropped_images:
                try:
                    # デジタルかスキャンかを渡して、最適な前処理を適用する
                    processed_img = preprocess_image_for_ocr(crop_img, is_digital)
                    
                    # 【改善】範囲指定抽出(セル単位等)の場合は PSM 6 (均一なテキストブロック)、ページ全体の場合は PSM 3 を使用
                    psm_val = 6 if crop_regions else 3
                    custom_config = f'--oem 3 --psm {psm_val}'
                    
                    text = pytesseract.image_to_string(Image.fromarray(processed_img), lang="jpn+jpn_vert+eng", config=custom_config)
                    
                    if crop_regions:
                        text = text.strip()
                        if text: all_regions_data.append([[text]])
                        else: all_regions_data.append([[""]])
                    else:
                        lines = [l.strip() for l in text.split('\n') if l.strip()]
                        if lines: all_regions_data.append([[l] for l in lines])
                        else: all_regions_data.append([[""]])
                except Exception: all_regions_data.append([["Error"]])
                    
            merged_data = merge_2d_arrays_horizontally(all_regions_data)
            final_data = [["ページ番号"] + ([f"範囲{idx+1}" for idx in range(len(cropped_images))] if crop_regions else ["抽出テキスト"])]
            for row in merged_data: final_data.append([f"{page_num+1}/{total_pages}"] + row)
            
            save_path = os.path.join(save_dir, f"{base}_P{page_num+1}_OCR")
            if out_format == "xlsx":
                wb = Workbook(); ws = wb.active; ws.title = "OCR"
                for r_idx, r_data in enumerate(final_data, 1):
                    for c_idx, val in enumerate(r_data, 1): ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
                auto_adjust_excel_column_width(ws); wb.save(f"{save_path}.xlsx")
            elif out_format == "csv":
                with open(f"{save_path}.csv", "w", encoding="utf-8-sig", newline="") as f_out: csv.writer(f_out).writerows(final_data)
            elif out_format == "txt":
                with open(f"{save_path}.txt", "w", encoding="utf-8") as f_out:
                    for r in final_data: f_out.write("\t".join(r) + "\n")
        doc.close(); gc.collect()

# ==============================
# データ集約タスク (ローカル)
# ==============================
def aggregate_local_task(files, save_dir, options, ui):
    out_format = options.get("out_format", "xlsx")
    search_ext = "xlsx" if out_format in ["jpg", "png", "dxf", "svg", "tiff", "bmp"] else out_format
    run_timestamp = time.strftime("%Y%m%d_%H%M%S")
    
    if not files: raise Exception("対象が選択されていません。")
    search_exts = ["xlsx", "xlsm", "xls"] if search_ext == "xlsx" else [search_ext]
    target_files_set = set()
    for f in files:
        if os.path.isdir(f):
            for fn in os.listdir(f):
                if any(fn.lower().endswith(f".{ext}") for ext in search_exts) and "集約" not in fn:
                    target_files_set.add(os.path.abspath(os.path.join(f, fn)))
        elif os.path.isfile(f):
            if any(f.lower().endswith(f".{ext}") for ext in search_exts): target_files_set.add(os.path.abspath(f))

    target_files = sorted(list(target_files_set))
    if not target_files: raise Exception("集約対象が見つかりません。")

    agg_header, agg_rows = ["元ファイル名"], []
    
    def map_to_master(fname, curr_header, curr_rows):
        safe_header = [str(h).strip() if h else f"列{i+1}" for i, h in enumerate(curr_header)]
        col_mapping = {}
        for i, h in enumerate(safe_header):
            if h not in agg_header: agg_header.append(h)
            col_mapping[i] = agg_header.index(h)
        for r in curr_rows:
            row = [""] * len(agg_header); row[0] = fname
            for i, val in enumerate(r):
                if i in col_mapping:
                    idx = col_mapping[i]
                    if idx >= len(row): row.extend([""] * (idx - len(row) + 1))
                    row[idx] = str(val).strip()
            agg_rows.append(row)

    for i, f in enumerate(target_files, 1):
        if ui.is_cancelled(): return
        ui.update_overall(i, len(target_files), f"集約中... ({i}/{len(target_files)})")
        ext = os.path.splitext(f)[1].lower().strip('.')
        fname = os.path.basename(f)
        try:
            if ext in ["xlsx", "xlsm"]:
                wb = openpyxl.load_workbook(f, data_only=True)
                for sheet in wb.sheetnames:
                    rows = [r for r in wb[sheet].iter_rows(values_only=True) if any(c is not None for c in r)]
                    if rows: map_to_master(fname, rows[0], rows[1:])
                wb.close()
            elif ext == "csv":
                with open(f, "r", encoding="utf-8-sig") as f_in: rows = list(csv.reader(f_in))
                if rows: map_to_master(fname, rows[0], rows[1:])
        except Exception: pass
        
    if agg_rows and save_dir:
        final_data = [agg_header] + [r + [""]*(len(agg_header)-len(r)) for r in agg_rows]
        if search_ext == "xlsx":
            wb = Workbook(); ws = wb.active; ws.title = "集約"
            for r_idx, r_data in enumerate(final_data, 1):
                for c_idx, val in enumerate(r_data, 1): ws.cell(row=r_idx, column=c_idx, value=sanitize_excel_text(val))
            auto_adjust_excel_column_width(ws); wb.save(os.path.join(save_dir, f"データ集約_{run_timestamp}.xlsx"))
        elif search_ext == "csv":
            with open(os.path.join(save_dir, f"データ集約_{run_timestamp}.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                csv.writer(f_out).writerows(final_data)