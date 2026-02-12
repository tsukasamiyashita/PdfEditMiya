# -*- coding: utf-8 -*-
"""
PdfEditMiya - スキャン画像対応・DXF変換・完全版
"""

import os
import threading
import cv2
import numpy as np
from tkinter import *
from tkinter import ttk, filedialog
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
import fitz  # PyMuPDF
import ezdxf

# ==============================
# 基本設定
# ==============================

APP_TITLE = "PdfEditMiya"
WINDOW_WIDTH = 560
WINDOW_HEIGHT = 600  # 元のサイズを維持

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
SUCCESS = "#2E7D32"
ERROR = "#C62828"
INACTIVE = "#90A4AE"

# ==============================
# グローバル変数
# ==============================

selected_files = []
selected_folder = ""
current_mode = None
preset_save_dir = ""
processing_popup = None
progress_bar = None
cancelled = False

# ==============================
# 共通ロジック
# ==============================

def run_task(func):
    global cancelled
    cancelled = False
    try:
        files = get_target_files()
        if not files: return

        test_dir = get_save_dir(files[0])
        if not test_dir or cancelled: return

        show_processing(len(files))
        func(files)
        close_processing()

        if not cancelled:
            show_message("✅ 完了", SUCCESS)
    except Exception as e:
        print(f"Error: {e}")
        close_processing()
        show_message("❌ エラー発生", ERROR)

def safe_run(func):
    threading.Thread(target=run_task, args=(func,), daemon=True).start()

# ==============================
# UI補助
# ==============================

def show_message(msg, color=PRIMARY):
    win = Toplevel(root)
    win.geometry("220x90")
    win.configure(bg=LIGHT)
    win.attributes("-topmost", True)
    Label(win, text=msg, bg=LIGHT, fg=color, font=("Segoe UI", 10, "bold")).pack(expand=True)
    win.after(3000, win.destroy)

def show_processing(total_steps=1):
    global processing_popup, progress_bar
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("300x120")
    processing_popup.configure(bg=LIGHT)
    processing_popup.grab_set()

    Label(processing_popup, text="処理中...", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold")).pack(pady=10)
    progress_bar = ttk.Progressbar(processing_popup, mode="determinate", maximum=total_steps, length=240)
    progress_bar.pack(pady=10)

def close_processing():
    global processing_popup
    if processing_popup:
        processing_popup.destroy()
        processing_popup = None

def update_progress(step):
    if progress_bar:
        progress_bar["value"] = step
        progress_bar.update()

# ==============================
# 保存・選択ロジック
# ==============================

def get_save_dir(original_path):
    global preset_save_dir, cancelled
    if save_option.get() == 1: return os.path.dirname(original_path)
    if preset_save_dir: return preset_save_dir
    folder = filedialog.askdirectory(title="保存先フォルダを選択")
    if folder:
        preset_save_dir = folder
        save_label.config(text=preset_save_dir)
        return folder
    cancelled = True
    return None

def select_save_dir():
    global preset_save_dir
    folder = filedialog.askdirectory(title="保存先フォルダを選択")
    if folder:
        preset_save_dir = folder
        save_label.config(text=preset_save_dir)
        save_option.set(2)

def on_save_mode_change():
    global preset_save_dir
    preset_save_dir = ""
    save_label.config(text="同じフォルダ" if save_option.get() == 1 else "未選択")

def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    if files:
        selected_files, selected_folder, current_mode = list(files), "", "file"
        update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory(title="PDFフォルダを選択")
    if folder:
        selected_folder, selected_files, current_mode = folder, [], "folder"
        update_ui()

def get_target_files():
    if current_mode == "file": return selected_files
    if current_mode == "folder" and selected_folder:
        return [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith(".pdf")]
    return []

# ==============================
# PDF操作コア機能
# ==============================

def merge_pdfs(files):
    writer = PdfWriter()
    for i, f in enumerate(files, 1):
        reader = PdfReader(f)
        for p in reader.pages: writer.add_page(p)
        update_progress(i)
    save_dir = get_save_dir(files[0])
    if save_dir:
        name = os.path.basename(selected_folder) if selected_folder else "Merged"
        with open(os.path.join(save_dir, f"{name}_Merge.pdf"), "wb") as out:
            writer.write(out)

def split_pdfs(files):
    for i, f in enumerate(files, 1):
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        for n, p in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(p)
            with open(os.path.join(save_dir, f"{base}_Split_{n+1}.pdf"), "wb") as out:
                writer.write(out)
        update_progress(i)

def rotate_pdfs(files):
    deg = rotate_option.get()
    for i, f in enumerate(files, 1):
        reader = PdfReader(f)
        writer = PdfWriter()
        for p in reader.pages:
            p.rotate(deg)
            writer.add_page(p)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)
        update_progress(i)

def extract_text(files):
    for i, f in enumerate(files, 1):
        reader = PdfReader(f)
        text = "".join([p.extract_text() or "" for p in reader.pages])
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out:
            out.write(text)
        update_progress(i)

def convert_to_excel(files):
    border_style = Side(border_style="thin", color="000000")
    for i, pdf_path in enumerate(files, 1):
        wb = Workbook()
        wb.remove(wb.active)
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if not tables: continue
                ws = wb.create_sheet(f"Page_{page_idx+1}")
                current_row = 1
                for table in tables:
                    for row_data in table:
                        for col_idx, cell_value in enumerate(row_data, 1):
                            val = str(cell_value).strip() if cell_value else ""
                            cell = ws.cell(row=current_row, column=col_idx, value=val)
                            cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                        current_row += 1
                    current_row += 2
        save_dir = get_save_dir(pdf_path)
        if not save_dir: return
        wb.save(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_Excel.xlsx"))
        update_progress(i)

def convert_to_image(files, ext):
    for i, f in enumerate(files, 1):
        doc = fitz.open(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        for n, page in enumerate(doc):
            page.get_pixmap(dpi=200).save(os.path.join(save_dir, f"{base}_{n+1}.{ext}"))
        update_progress(i)

def convert_to_dxf(files):
    """ベクタ要素とスキャン画像の両方に対応したDXF変換ロジック"""
    for i, f in enumerate(files, 1):
        doc = fitz.open(f)
        dwg = ezdxf.new('R2010')
        msp = dwg.modelspace()
        save_dir = get_save_dir(f)
        if not save_dir: return

        for page in doc:
            h = page.rect.height
            paths = page.get_drawings()
            
            if len(paths) > 20: # ベクタデータが豊富にある場合
                for path in paths:
                    for item in path["items"]:
                        if item[0] == "l":
                            msp.add_line((item[1].x, h-item[1].y), (item[2].x, h-item[2].y))
            else: # スキャン画像(ラスタ)として処理
                pix = page.get_pixmap(dpi=300)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
                # エッジ抽出
                binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                
                scale = page.rect.width / pix.w
                for cnt in contours:
                    if cv2.contourArea(cnt) < 10: continue # 小さなノイズを無視
                    pts = [(p[0][0]*scale, h - p[0][1]*scale) for p in cnt]
                    if len(pts) > 1: msp.add_lwpolyline(pts, close=True)

        dwg.saveas(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(f))[0]}_CAD.dxf"))
        update_progress(i)

# ==============================
# UI構築
# ==============================

def update_ui():
    path_text = "\n".join(selected_files) if current_mode == "file" else (f"フォルダ: {selected_folder}" if selected_folder else "未選択")
    path_label.config(text=path_text)
    is_active = current_mode is not None
    btns = [btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf]
    for b in btns: b.config(state=NORMAL if is_active else DISABLED, bg="#1E88E5" if is_active else LIGHT, fg="white" if is_active else INACTIVE)
    btn_merge.config(state=NORMAL if current_mode=="folder" else DISABLED, bg="#1E88E5" if current_mode=="folder" else LIGHT, fg="white" if current_mode=="folder" else INACTIVE)

root = Tk()
root.title(APP_TITLE)
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg=LIGHT)
root.resizable(False, False)

rotate_option, save_option = IntVar(value=270), IntVar(value=1)

Label(root, text=APP_TITLE, bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 15, "bold")).pack(pady=8)

file_frame = Frame(root, bg=LIGHT)
file_frame.pack(pady=5)
Button(file_frame, text="📄 ファイル選択", command=select_files, width=22).grid(row=0, column=0, padx=5)
Button(file_frame, text="📁 フォルダ選択", command=select_folder, width=22).grid(row=0, column=1, padx=5)

Label(root, text="選択パス", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold")).pack(pady=5)
path_label = Label(root, text="未選択", bg=LIGHT, wraplength=520, justify="left")
path_label.pack(pady=2)

save_frame = LabelFrame(root, text="保存先設定", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"), padx=5, pady=5)
save_frame.pack(pady=5, fill="x", padx=10)
Radiobutton(save_frame, text="同じフォルダ", variable=save_option, value=1, bg=LIGHT, command=on_save_mode_change).pack(anchor="w")
Radiobutton(save_frame, text="任意フォルダ", variable=save_option, value=2, bg=LIGHT, command=on_save_mode_change).pack(anchor="w")
Button(save_frame, text="📂 保存先を選択", command=select_save_dir, width=22).pack(pady=3)
save_label = Label(save_frame, text="同じフォルダ", bg=LIGHT)
save_label.pack()

rotate_frame = LabelFrame(root, text="回転設定", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"), padx=5, pady=5)
rotate_frame.pack(pady=5, fill="x", padx=10)
for t, v in [("左（270°）", 270), ("上下（180°）", 180), ("右（90°）", 90)]:
    Radiobutton(rotate_frame, text=t, variable=rotate_option, value=v, bg=LIGHT).pack(anchor="w")

op_frame = LabelFrame(root, text="操作", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"), padx=5, pady=5)
op_frame.pack(pady=10)
btn_merge = Button(op_frame, text="結合", width=12, command=lambda: safe_run(merge_pdfs))
btn_split = Button(op_frame, text="分割", width=12, command=lambda: safe_run(split_pdfs))
btn_rotate = Button(op_frame, text="回転", width=12, command=lambda: safe_run(rotate_pdfs))
btn_text = Button(op_frame, text="Text抽出", width=12, command=lambda: safe_run(extract_text))
btn_excel = Button(op_frame, text="Excel変換", width=12, command=lambda: safe_run(convert_to_excel))
btn_jpeg = Button(op_frame, text="JPEG変換", width=12, command=lambda: safe_run(lambda fs: convert_to_image(fs, "jpg")))
btn_png = Button(op_frame, text="PNG変換", width=12, command=lambda: safe_run(lambda fs: convert_to_image(fs, "png")))
btn_dxf = Button(op_frame, text="DXF変換", width=12, command=lambda: safe_run(convert_to_dxf))

# レイアウト配置
op_list = [btn_merge, btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf]
for i, b in enumerate(op_list):
    b.grid(row=i//4, column=i%4, padx=5, pady=3)

update_ui()
root.mainloop()