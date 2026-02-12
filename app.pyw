# -*- coding: utf-8 -*-
"""
PdfEditMiya - 安定版 + Excel変換強化

追加機能:
・PDF表をExcel変換
・罫線再現
・列幅自動調整
・中央揃え
"""

import os
import threading
from tkinter import *
from tkinter import ttk, filedialog
from PyPDF2 import PdfReader, PdfWriter
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
import pdfplumber

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
cancelled = False

# ==============================
# メインウィンドウ
# ==============================

root = Tk()
root.title(APP_TITLE)
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg=LIGHT)
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")
style.configure("TProgressbar", thickness=12)

# ==============================
# ユーティリティ
# ==============================

def safe_run(func):
    threading.Thread(target=run_task, args=(func,), daemon=True).start()

def show_message(msg, color=PRIMARY):
    win = Toplevel(root)
    win.geometry("220x90")
    win.configure(bg=LIGHT)
    win.resizable(False, False)
    Label(win, text=msg, bg=LIGHT, fg=color,
          font=("Segoe UI", 10, "bold")).pack(expand=True)
    win.after(3000, win.destroy)

def show_processing(total_steps=1):
    global processing_popup, progress_bar
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("300x120")
    processing_popup.configure(bg=LIGHT)
    processing_popup.resizable(False, False)
    processing_popup.grab_set()

    Label(processing_popup, text="処理中...",
          bg=LIGHT, fg=PRIMARY,
          font=("Segoe UI", 10, "bold")).pack(pady=10)

    progress_bar = ttk.Progressbar(processing_popup,
                                   mode="determinate",
                                   maximum=total_steps,
                                   length=240)
    progress_bar.pack(pady=10)

def close_processing():
    global processing_popup
    if processing_popup:
        processing_popup.destroy()

def update_progress(step):
    progress_bar["value"] = step
    progress_bar.update()

# ==============================
# 保存先処理
# ==============================

def get_save_dir(original_path):
    global preset_save_dir, cancelled
    if save_option.get() == 1:
        return os.path.dirname(original_path)
    if preset_save_dir:
        return preset_save_dir
    folder = filedialog.askdirectory(title="保存先フォルダを選択")
    if folder:
        preset_save_dir = folder
        save_label.config(text=preset_save_dir)
        save_option.set(2)
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

def on_save_change():
    global preset_save_dir
    if save_option.get() == 1:
        preset_save_dir = ""
        save_label.config(text="同じフォルダ")
    else:
        preset_save_dir = ""
        save_label.config(text="未選択")

# ==============================
# 選択処理
# ==============================

def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    if files:
        selected_files = list(files)
        selected_folder = ""
        current_mode = "file"
        update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode, preset_save_dir
    folder = filedialog.askdirectory(title="PDFフォルダを選択")
    if folder:
        selected_folder = folder
        selected_files = []
        current_mode = "folder"
        if save_option.get() == 2:
            preset_save_dir = ""
            save_label.config(text="未選択")
        update_ui()

# ==============================
# UI更新
# ==============================

def set_button_state(btn, enabled):
    if enabled:
        btn.config(state=NORMAL, bg="#1E88E5", fg="white")
    else:
        btn.config(state=DISABLED, bg=LIGHT, fg=INACTIVE)

def update_ui():
    if current_mode == "file":
        path_text = "\n".join(selected_files) if selected_files else "未選択"
    elif current_mode == "folder":
        path_text = selected_folder if selected_folder else "未選択"
    else:
        path_text = "未選択"

    path_label.config(text=path_text)

    set_button_state(btn_merge, current_mode == "folder")
    set_button_state(btn_split, current_mode == "file")
    set_button_state(btn_rotate, current_mode == "file")
    set_button_state(btn_text, current_mode == "file")
    set_button_state(btn_excel, current_mode == "file")

# ==============================
# 共通処理実行
# ==============================

def run_task(func):
    global cancelled
    cancelled = False
    try:
        files = get_target_files()
        if not files:
            raise Exception()
        total = len(files)
        show_processing(total)
        func()
        close_processing()
        if cancelled:
            return
        show_message("✅ 完了", SUCCESS)
    except Exception:
        close_processing()
        show_message("❌ エラー", ERROR)

def get_target_files():
    if selected_files:
        return selected_files
    if selected_folder:
        return [os.path.join(selected_folder, f)
                for f in os.listdir(selected_folder)
                if f.lower().endswith(".pdf")]
    return []

# ==============================
# PDF操作（元機能維持）
# ==============================

def merge_pdfs():
    files = get_target_files()
    writer = PdfWriter()
    for i, f in enumerate(files, 1):
        reader = PdfReader(f)
        for p in reader.pages:
            writer.add_page(p)
        update_progress(i)

    save_dir = get_save_dir(files[0])
    if not save_dir:
        return

    name = os.path.basename(selected_folder)
    with open(os.path.join(save_dir, name + "_Merge.pdf"), "wb") as out:
        writer.write(out)

def split_pdfs():
    for i, f in enumerate(selected_files, 1):
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        for n, p in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(p)
            with open(os.path.join(save_dir, f"{base}_Split_{n+1}.pdf"), "wb") as out:
                writer.write(out)
        update_progress(i)

def rotate_pdfs():
    deg = rotate_option.get()
    for i, f in enumerate(selected_files, 1):
        reader = PdfReader(f)
        writer = PdfWriter()
        for p in reader.pages:
            p.rotate(deg)
            writer.add_page(p)
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)
        update_progress(i)

def extract_text():
    for i, f in enumerate(selected_files, 1):
        reader = PdfReader(f)
        text = ""
        for p in reader.pages:
            t = p.extract_text()
            text += t if t else ""
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out:
            out.write(text)
        update_progress(i)

# ==============================
# ★ 追加機能：表→Excel変換
# ==============================

def convert_table_to_excel():

    thin = Side(border_style="thin", color="000000")

    for i, pdf_path in enumerate(selected_files, 1):

        wb = Workbook()
        wb.remove(wb.active)

        with pdfplumber.open(pdf_path) as pdf:
            for page_i, page in enumerate(pdf.pages):

                tables = page.extract_tables()
                if not tables:
                    continue

                ws = wb.create_sheet(f"Page_{page_i+1}")

                row_cursor = 1

                for table in tables:
                    for r, row in enumerate(table):
                        for c, cell in enumerate(row):
                            value = cell.strip() if cell else ""
                            excel_cell = ws.cell(
                                row=row_cursor + r,
                                column=c + 1,
                                value=value
                            )
                            excel_cell.alignment = Alignment(
                                horizontal="center",
                                vertical="center"
                            )
                            excel_cell.border = Border(
                                left=thin, right=thin,
                                top=thin, bottom=thin
                            )
                    row_cursor += len(table) + 2

                # 列幅自動調整
                for col in ws.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_length = max(
                                max_length,
                                len(str(cell.value))
                            )
                    ws.column_dimensions[col_letter].width = max_length + 4

        save_dir = get_save_dir(pdf_path)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(pdf_path))[0]

        wb.save(os.path.join(
            save_dir,
            f"{base}_Table.xlsx"
        ))

        update_progress(i)

# ==============================
# UI構築（レイアウト維持）
# ==============================

Label(root, text=APP_TITLE,
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 15, "bold")).pack(pady=8)

file_frame = Frame(root, bg=LIGHT)
file_frame.pack(pady=5)

Button(file_frame, text="📄 ファイル選択", command=select_files, width=22).grid(row=0, column=0, padx=5)
Button(file_frame, text="📁 フォルダ選択", command=select_folder, width=22).grid(row=0, column=1, padx=5)

Label(root, text="選択パス", bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=5)

path_label = Label(root, text="未選択", bg=LIGHT,
                   wraplength=520, justify="left")
path_label.pack(pady=2)

save_frame = LabelFrame(root, text="保存先設定",
                        bg=LIGHT, fg=PRIMARY,
                        font=("Segoe UI", 10, "bold"),
                        padx=5, pady=5)
save_frame.pack(pady=5, fill="x", padx=10)

save_option = IntVar(value=1)

Radiobutton(save_frame, text="同じフォルダ（初期）",
            variable=save_option, value=1,
            command=on_save_change, bg=LIGHT).pack(anchor="w")

Radiobutton(save_frame, text="任意フォルダ",
            variable=save_option, value=2,
            command=on_save_change, bg=LIGHT).pack(anchor="w")

Button(save_frame, text="📂 保存先を選択",
       command=select_save_dir, width=22).pack(pady=3)

save_label = Label(save_frame, text="同じフォルダ", bg=LIGHT)
save_label.pack()

rotate_frame = LabelFrame(root, text="回転設定",
                          bg=LIGHT, fg=PRIMARY,
                          font=("Segoe UI", 10, "bold"),
                          padx=5, pady=5)
rotate_frame.pack(pady=5, fill="x", padx=10)

rotate_option = IntVar(value=270)

Radiobutton(rotate_frame, text="左回転（270°）",
            variable=rotate_option, value=270,
            bg=LIGHT).pack(anchor="w")

Radiobutton(rotate_frame, text="上下回転（180°）",
            variable=rotate_option, value=180,
            bg=LIGHT).pack(anchor="w")

Radiobutton(rotate_frame, text="右回転（90°）",
            variable=rotate_option, value=90,
            bg=LIGHT).pack(anchor="w")

op_frame = LabelFrame(root, text="操作",
                      bg=LIGHT, fg=PRIMARY,
                      font=("Segoe UI", 10, "bold"),
                      padx=5, pady=5)
op_frame.pack(pady=10)

btn_merge = Button(op_frame, text="結合", width=12,
                   command=lambda: safe_run(merge_pdfs))
btn_split = Button(op_frame, text="分割", width=12,
                   command=lambda: safe_run(split_pdfs))
btn_rotate = Button(op_frame, text="回転", width=12,
                    command=lambda: safe_run(rotate_pdfs))
btn_text = Button(op_frame, text="Text抽出", width=12,
                  command=lambda: safe_run(extract_text))
btn_excel = Button(op_frame, text="表→Excel", width=12,
                   command=lambda: safe_run(convert_table_to_excel))

btn_merge.grid(row=0, column=0, padx=5, pady=3)
btn_split.grid(row=0, column=1, padx=5, pady=3)
btn_rotate.grid(row=0, column=2, padx=5, pady=3)
btn_text.grid(row=0, column=3, padx=5, pady=3)
btn_excel.grid(row=1, column=1, columnspan=2, pady=5)

update_ui()
root.mainloop()
