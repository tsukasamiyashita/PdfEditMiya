# -*- coding: utf-8 -*-
"""
PdfEditMiya

■ 機能
・PDF結合（フォルダ選択時のみ有効）
・PDF分割（ファイル選択時のみ有効）
・PDF回転（左回転/上下回転/右回転 ラジオボタン）
・テキスト抽出（OCRエンジン単一選択）
・保存先 初期＝同じフォルダ
・任意フォルダ選択時は保存先未選択表示
・同じフォルダを選択し直すと表示も戻る
・保存先キャンセル時は完了画面を表示しない
・処理中ポップアップ表示
・完了画面は3秒後自動クローズ
・青ベースUI
"""

import os
import threading
from tkinter import *
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter

# ===== OCR =====
try:
    import pytesseract
    from pdf2image import convert_from_path
    TESS_AVAILABLE = True
except Exception:
    TESS_AVAILABLE = False

# ==========================
# グローバル
# ==========================

selected_files = []
selected_folder = ""
current_mode = None
preset_save_dir = ""
processing_popup = None
cancelled = False

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
WHITE = "#FFFFFF"

# ==========================
# メイン画面
# ==========================

root = Tk()
root.title("PdfEditMiya")
root.geometry("620x800")
root.minsize(620, 800)
root.configure(bg=LIGHT)

# ==========================
# ポップアップ
# ==========================

def show_processing(msg="処理実行中..."):
    global processing_popup
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("260x100")
    processing_popup.configure(bg=LIGHT)
    Label(processing_popup, text=msg,
          bg=LIGHT, fg=PRIMARY,
          font=("Segoe UI", 10, "bold")).pack(expand=True)
    processing_popup.grab_set()
    processing_popup.update()

def close_processing():
    global processing_popup
    if processing_popup:
        processing_popup.destroy()
        processing_popup = None

def auto_close_message(title, msg, error=False):
    win = Toplevel(root)
    win.title(title)
    win.geometry("260x100")
    bg = "#FFEBEE" if error else LIGHT
    fg = "#C62828" if error else PRIMARY
    win.configure(bg=bg)
    Label(win, text=msg, bg=bg, fg=fg,
          font=("Segoe UI", 10, "bold")).pack(expand=True)
    win.after(3000, win.destroy)

# ==========================
# 保存先制御
# ==========================

def on_save_option_change():
    global preset_save_dir
    if save_option.get() == 2:
        preset_save_dir = ""
        save_dir_label.config(text="保存先: 未選択")
    else:
        preset_save_dir = ""
        save_dir_label.config(text="保存先: 同じフォルダ")

def choose_preset_folder():
    global preset_save_dir
    folder = filedialog.askdirectory()
    if folder:
        preset_save_dir = folder
        save_dir_label.config(text=f"保存先: {preset_save_dir}")

def get_save_dir(original_path):
    global preset_save_dir, cancelled

    if save_option.get() == 1:
        return os.path.dirname(original_path)

    if preset_save_dir:
        return preset_save_dir

    folder = filedialog.askdirectory()
    if folder:
        preset_save_dir = folder
        save_dir_label.config(text=f"保存先: {preset_save_dir}")
        return folder

    cancelled = True
    return None

# ==========================
# 選択処理
# ==========================

def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    if files:
        selected_files = list(files)
        selected_folder = ""
        current_mode = "file"
        path_label.config(text=f"ファイル選択: {len(files)}件")
        update_buttons()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        selected_files = []
        current_mode = "folder"
        path_label.config(text=f"フォルダ選択: {folder}")
        update_buttons()

def update_buttons():
    if current_mode == "file":
        split_btn.config(state=NORMAL)
        rotate_btn.config(state=NORMAL)
        text_btn.config(state=NORMAL)
        merge_btn.config(state=DISABLED)
    elif current_mode == "folder":
        merge_btn.config(state=NORMAL)
        split_btn.config(state=DISABLED)
        rotate_btn.config(state=DISABLED)
        text_btn.config(state=DISABLED)

# ==========================
# 実行制御
# ==========================

def run_task(func):
    def task():
        global cancelled
        cancelled = False
        try:
            show_processing()
            func()
            close_processing()
            if cancelled:
                return
            auto_close_message("完了", "処理が完了しました")
        except Exception:
            close_processing()
            auto_close_message("エラー", "処理失敗", True)
    threading.Thread(target=task).start()

# ==========================
# PDF処理
# ==========================

def merge_pdfs():
    files = [os.path.join(selected_folder, f)
             for f in os.listdir(selected_folder)
             if f.lower().endswith(".pdf")]

    writer = PdfWriter()
    for f in files:
        reader = PdfReader(f)
        for p in reader.pages:
            writer.add_page(p)

    save_dir = get_save_dir(files[0])
    if not save_dir:
        return

    output = os.path.join(save_dir, "Merged_Merge.pdf")
    with open(output, "wb") as f:
        writer.write(f)

def split_pdfs():
    for f in selected_files:
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(f))[0]
        for i, page in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(page)
            with open(os.path.join(save_dir,
                     f"{base}_Split_{i+1}.pdf"), "wb") as out:
                writer.write(out)

def rotate_pdfs():
    angle = rotation.get()
    for f in selected_files:
        reader = PdfReader(f)
        writer = PdfWriter()
        for p in reader.pages:
            p.rotate(angle)
            writer.add_page(p)

        save_dir = get_save_dir(f)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir,
                 f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)

def extract_text():
    engine = ocr_engine.get()
    if engine == 0:
        raise Exception()

    for f in selected_files:
        text = ""

        if engine == 1:  # PyPDF2
            reader = PdfReader(f)
            for p in reader.pages:
                t = p.extract_text()
                text += t if t else ""

        elif engine == 2 and TESS_AVAILABLE:  # Tesseract
            images = convert_from_path(f)
            for img in images:
                text += pytesseract.image_to_string(img, lang="jpn+eng")

        save_dir = get_save_dir(f)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir,
                 f"{base}_Text.txt"),
                 "w", encoding="utf-8") as out:
            out.write(text)

# ==========================
# UI構築
# ==========================

Label(root, text="PdfEditMiya",
      font=("Segoe UI", 18, "bold"),
      bg=LIGHT, fg=PRIMARY).pack(pady=10)

Button(root, text="📄 ファイル選択",
       command=select_files,
       bg=PRIMARY, fg=WHITE,
       width=25).pack(pady=5)

Button(root, text="📂 フォルダ選択",
       command=select_folder,
       bg=PRIMARY, fg=WHITE,
       width=25).pack(pady=5)

path_label = Label(root, text="未選択",
                   bg=LIGHT, fg=PRIMARY)
path_label.pack(pady=5)

# 保存先
Label(root, text="保存先設定",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 11, "bold")).pack(pady=10)

save_option = IntVar(value=1)

Radiobutton(root, text="同じフォルダ（初期）",
            variable=save_option, value=1,
            command=on_save_option_change,
            bg=LIGHT).pack()

Radiobutton(root, text="任意のフォルダ",
            variable=save_option, value=2,
            command=on_save_option_change,
            bg=LIGHT).pack()

Button(root, text="📂 任意保存先を事前選択",
       command=choose_preset_folder,
       bg=PRIMARY, fg=WHITE,
       width=25).pack(pady=5)

save_dir_label = Label(root,
                       text="保存先: 同じフォルダ",
                       bg=LIGHT)
save_dir_label.pack()

# 回転
Label(root, text="回転方向",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 11, "bold")).pack(pady=10)

rotation = IntVar(value=270)

Radiobutton(root, text="左回転",
            variable=rotation, value=270,
            bg=LIGHT).pack()

Radiobutton(root, text="上下回転",
            variable=rotation, value=180,
            bg=LIGHT).pack()

Radiobutton(root, text="右回転",
            variable=rotation, value=90,
            bg=LIGHT).pack()

# OCR（単一選択）
Label(root, text="テキスト抽出エンジン（単一選択）",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 11, "bold")).pack(pady=10)

ocr_engine = IntVar(value=1)

Radiobutton(root,
            text="PyPDF2（高速・埋め込みテキスト向け）",
            variable=ocr_engine,
            value=1,
            bg=LIGHT).pack(anchor="w", padx=40)

Radiobutton(root,
            text="Tesseract OCR（画像PDF対応・要インストール）",
            variable=ocr_engine,
            value=2,
            bg=LIGHT,
            state=NORMAL if TESS_AVAILABLE else DISABLED).pack(anchor="w", padx=40)

# 実行ボタン
merge_btn = Button(root, text="📎 結合",
                   command=lambda: run_task(merge_pdfs),
                   bg=PRIMARY, fg=WHITE,
                   width=25, state=DISABLED)
merge_btn.pack(pady=5)

split_btn = Button(root, text="✂ 分割",
                   command=lambda: run_task(split_pdfs),
                   bg=PRIMARY, fg=WHITE,
                   width=25, state=DISABLED)
split_btn.pack(pady=5)

rotate_btn = Button(root, text="🔄 回転",
                    command=lambda: run_task(rotate_pdfs),
                    bg=PRIMARY, fg=WHITE,
                    width=25, state=DISABLED)
rotate_btn.pack(pady=5)

text_btn = Button(root, text="📝 テキスト抽出",
                  command=lambda: run_task(extract_text),
                  bg=PRIMARY, fg=WHITE,
                  width=25, state=DISABLED)
text_btn.pack(pady=10)

root.mainloop()
