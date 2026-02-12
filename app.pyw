# -*- coding: utf-8 -*-
"""
PdfEditMiya - 高機能・保存先制御強化版
"""

import os
import threading
from tkinter import *
from tkinter import ttk, filedialog
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter
import fitz  # PyMuPDF

# ==============================
# 基本設定
# ==============================

APP_TITLE = "PdfEditMiya"
WINDOW_WIDTH = 560
WINDOW_HEIGHT = 600

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

        # 実行直前の保存先チェック
        test_dir = get_save_dir(files[0])
        if not test_dir or cancelled:
            return

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
    
    # 同じフォルダに保存
    if save_option.get() == 1:
        return os.path.dirname(original_path)
    
    # 任意フォルダが既に選択済み
    if preset_save_dir:
        return preset_save_dir

    # 任意フォルダが未選択の場合、選択画面を出す
    folder = filedialog.askdirectory(title="保存先フォルダを選択してください")
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
    if save_option.get() == 1:
        preset_save_dir = ""
        save_label.config(text="同じフォルダ")
    else:
        preset_save_dir = ""
        save_label.config(text="未選択")

def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    if files:
        selected_files = list(files)
        selected_folder = ""
        current_mode = "file"
        update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory(title="PDFフォルダを選択")
    if folder:
        selected_folder = folder
        selected_files = []
        current_mode = "folder"
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
    if not save_dir: return
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
    table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines", "snap_tolerance": 3, "join_tolerance": 3}

    for i, pdf_path in enumerate(files, 1):
        wb = Workbook()
        wb.remove(wb.active)
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables(table_settings)
                if not tables: tables = page.extract_tables()
                if not tables: continue
                
                ws = wb.create_sheet(f"Page_{page_idx+1}")
                current_row = 1
                for table in tables:
                    for row_data in table:
                        for col_idx, cell_value in enumerate(row_data, 1):
                            val = cell_value.strip() if cell_value else ""
                            clean_val = val.replace(',', '').replace('¥', '')
                            try:
                                val = float(clean_val) if '.' in clean_val else int(clean_val)
                            except ValueError: pass
                            
                            cell = ws.cell(row=current_row, column=col_idx, value=val)
                            cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        current_row += 1
                    current_row += 2
                
                for col in ws.columns:
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = 15

        save_dir = get_save_dir(pdf_path)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        wb.save(os.path.join(save_dir, f"{base}_Excel.xlsx"))
        update_progress(i)

def convert_to_image(files, ext):
    for i, f in enumerate(files, 1):
        doc = fitz.open(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        for n, page in enumerate(doc):
            pix = page.get_pixmap(dpi=200)
            pix.save(os.path.join(save_dir, f"{base}_{n+1}.{ext}"))
        update_progress(i)

# ==============================
# UI構築
# ==============================

def update_ui():
    if current_mode == "file": path_text = "\n".join(selected_files)
    elif current_mode == "folder": path_text = f"フォルダ: {selected_folder}"
    else: path_text = "未選択"
    path_label.config(text=path_text)
    
    is_active = current_mode is not None
    for btn in [btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png]:
        btn.config(state=NORMAL if is_active else DISABLED, bg="#1E88E5" if is_active else LIGHT, fg="white" if is_active else INACTIVE)
    btn_merge.config(state=NORMAL if current_mode=="folder" else DISABLED, bg="#1E88E5" if current_mode=="folder" else LIGHT, fg="white" if current_mode=="folder" else INACTIVE)

root = Tk()
root.title(APP_TITLE)
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg=LIGHT)
root.resizable(False, False)

rotate_option = IntVar(value=270)
save_option = IntVar(value=1)

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
for txt, val in [("左（270°）", 270), ("上下（180°）", 180), ("右（90°）", 90)]:
    Radiobutton(rotate_frame, text=txt, variable=rotate_option, value=val, bg=LIGHT).pack(anchor="w")

op_frame = LabelFrame(root, text="操作", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"), padx=5, pady=5)
op_frame.pack(pady=10)
btn_merge = Button(op_frame, text="結合", width=12, command=lambda: safe_run(merge_pdfs))
btn_split = Button(op_frame, text="分割", width=12, command=lambda: safe_run(split_pdfs))
btn_rotate = Button(op_frame, text="回転", width=12, command=lambda: safe_run(rotate_pdfs))
btn_text = Button(op_frame, text="Text抽出", width=12, command=lambda: safe_run(extract_text))
btn_excel = Button(op_frame, text="Excel変換", width=12, command=lambda: safe_run(convert_to_excel))
btn_jpeg = Button(op_frame, text="JPEG変換", width=12, command=lambda: safe_run(lambda fs: convert_to_image(fs, "jpg")))
btn_png = Button(op_frame, text="PNG変換", width=12, command=lambda: safe_run(lambda fs: convert_to_image(fs, "png")))

btn_merge.grid(row=0, column=0, padx=5, pady=3)
btn_split.grid(row=0, column=1, padx=5, pady=3)
btn_rotate.grid(row=0, column=2, padx=5, pady=3)
btn_text.grid(row=0, column=3, padx=5, pady=3)
btn_excel.grid(row=1, column=0, padx=5, pady=3)
btn_jpeg.grid(row=1, column=1, padx=5, pady=3)
btn_png.grid(row=1, column=2, padx=5, pady=3)

update_ui()
root.mainloop()