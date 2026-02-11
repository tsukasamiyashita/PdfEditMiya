# -*- coding: utf-8 -*-
"""
PdfEditMiya
青ベースデザイン
・処理実行中は「実行中画面」を表示
・完了画面は3秒後に自動で閉じる
機能：分割 / 結合 / 回転 / テキスト抽出
"""

import os
import threading
from tkinter import *
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter

# ==========================
# 共通変数
# ==========================

selected_files = []
selected_folder = ""
processing_popup = None

# ==========================
# ポップアップ関連
# ==========================

def show_processing(message="処理実行中..."):
    global processing_popup
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("300x120")
    processing_popup.resizable(False, False)
    processing_popup.configure(bg="#E3F2FD")

    Label(processing_popup,
          text=message,
          bg="#E3F2FD",
          fg="#1565C0",
          font=("Segoe UI", 11, "bold")).pack(expand=True)

    processing_popup.grab_set()
    processing_popup.update()

def close_processing():
    global processing_popup
    if processing_popup:
        processing_popup.destroy()
        processing_popup = None

def show_auto_close_message(title, message, is_error=False):
    popup = Toplevel(root)
    popup.title(title)
    popup.geometry("300x120")
    popup.resizable(False, False)

    bg_color = "#E3F2FD" if not is_error else "#FFEBEE"
    fg_color = "#1565C0" if not is_error else "#C62828"

    popup.configure(bg=bg_color)

    Label(popup, text=message,
          bg=bg_color,
          fg=fg_color,
          font=("Segoe UI", 10, "bold")).pack(expand=True)

    popup.after(3000, popup.destroy)

# ==========================
# 選択処理
# ==========================

def select_files():
    global selected_files, selected_folder
    files = filedialog.askopenfilenames(filetypes=[("PDFファイル", "*.pdf")])
    if files:
        selected_files = list(files)
        selected_folder = ""
        update_path_display()
        update_button_state("file")

def select_folder():
    global selected_folder, selected_files
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        selected_files = []
        update_path_display()
        update_button_state("folder")

def update_path_display():
    text_paths.config(state=NORMAL)
    text_paths.delete(1.0, END)
    if selected_files:
        text_paths.insert(END, "\n".join(selected_files))
    elif selected_folder:
        text_paths.insert(END, selected_folder)
    text_paths.config(state=DISABLED)

def update_button_state(mode=None):
    btn_merge.config(state=DISABLED)
    btn_split.config(state=DISABLED)
    btn_rotate.config(state=DISABLED)
    btn_text.config(state=DISABLED)

    if mode == "file":
        btn_split.config(state=NORMAL)
        btn_rotate.config(state=NORMAL)
        btn_text.config(state=NORMAL)
    elif mode == "folder":
        btn_merge.config(state=NORMAL)

# ==========================
# 共通ユーティリティ
# ==========================

def get_target_files():
    if selected_files:
        return selected_files
    elif selected_folder:
        return [
            os.path.join(selected_folder, f)
            for f in os.listdir(selected_folder)
            if f.lower().endswith(".pdf")
        ]
    return []

def get_save_dir(original_path):
    if save_option.get() == 1:
        return os.path.dirname(original_path)
    else:
        return filedialog.askdirectory()

# ==========================
# 実行ラッパー（スレッド対応）
# ==========================

def run_with_loading(task_func):
    def task():
        try:
            show_processing()
            task_func()
            close_processing()
            show_auto_close_message("完了", "処理が完了しました")
        except Exception:
            close_processing()
            show_auto_close_message("エラー", "処理失敗（0扱い）", True)

    threading.Thread(target=task).start()

# ==========================
# PDF操作
# ==========================

def merge_pdfs():
    files = get_target_files()
    if not files:
        raise Exception()

    writer = PdfWriter()
    for file in files:
        reader = PdfReader(file)
        for page in reader.pages:
            writer.add_page(page)

    save_dir = get_save_dir(files[0])
    if not save_dir:
        return

    base = os.path.basename(selected_folder)
    output_path = os.path.join(save_dir, base + "_Merge.pdf")

    with open(output_path, "wb") as f:
        writer.write(f)

def split_pdfs():
    for file in selected_files:
        reader = PdfReader(file)
        save_dir = get_save_dir(file)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(file))[0]

        for i, page in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(page)
            output_path = os.path.join(save_dir, f"{base}_Split_{i+1}.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)

def rotate_pdfs():
    degree = rotate_option.get()
    if degree == 0:
        raise Exception()

    for file in selected_files:
        reader = PdfReader(file)
        writer = PdfWriter()

        for page in reader.pages:
            page.rotate(degree)
            writer.add_page(page)

        save_dir = get_save_dir(file)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(file))[0]
        output_path = os.path.join(save_dir, f"{base}_Rotate.pdf")

        with open(output_path, "wb") as f:
            writer.write(f)

def extract_text():
    for file in selected_files:
        reader = PdfReader(file)
        text = ""

        for page in reader.pages:
            t = page.extract_text()
            text += t if t else ""

        save_dir = get_save_dir(file)
        if not save_dir:
            return

        base = os.path.splitext(os.path.basename(file))[0]
        output_path = os.path.join(save_dir, f"{base}_Text.txt")

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text)

# ==========================
# UIデザイン
# ==========================

PRIMARY = "#1565C0"
ACCENT = "#1E88E5"
LIGHT = "#E3F2FD"
WHITE = "#FFFFFF"

root = Tk()
root.title("PdfEditMiya")
root.geometry("620x700")
root.configure(bg=LIGHT)

Label(root,
      text="PdfEditMiya",
      font=("Segoe UI", 20, "bold"),
      bg=LIGHT,
      fg=PRIMARY).pack(pady=15)

btn_style = {
    "font": ("Segoe UI", 10, "bold"),
    "bg": PRIMARY,
    "fg": WHITE,
    "activebackground": ACCENT,
    "activeforeground": WHITE,
    "bd": 0,
    "width": 22,
    "height": 1
}

Button(root, text="📄 ファイル選択", command=select_files, **btn_style).pack(pady=5)
Button(root, text="📁 フォルダ選択", command=select_folder, **btn_style).pack(pady=5)

Label(root, text="選択パス", bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=8)

text_paths = Text(root, height=6, width=70,
                  bg=WHITE, fg="#333333",
                  font=("Consolas", 9),
                  bd=0)
text_paths.pack(pady=5)
text_paths.config(state=DISABLED)

Label(root, text="保存先", bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=10)

save_option = IntVar(value=1)
Radiobutton(root, text="同じフォルダ",
            variable=save_option, value=1,
            bg=LIGHT, selectcolor=WHITE).pack()

Radiobutton(root, text="任意のフォルダ",
            variable=save_option, value=2,
            bg=LIGHT, selectcolor=WHITE).pack()

Label(root, text="回転方法", bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=12)

rotate_option = IntVar(value=0)
frame_rotate = Frame(root, bg=LIGHT)
frame_rotate.pack()

toggle_style = {
    "indicatoron": False,
    "width": 9,
    "font": ("Segoe UI", 9, "bold"),
    "bg": PRIMARY,
    "fg": WHITE,
    "selectcolor": ACCENT,
    "bd": 0
}

Radiobutton(frame_rotate, text="左回転",
            variable=rotate_option, value=270,
            **toggle_style).grid(row=0, column=0, padx=5)

Radiobutton(frame_rotate, text="上下回転",
            variable=rotate_option, value=180,
            **toggle_style).grid(row=0, column=1, padx=5)

Radiobutton(frame_rotate, text="右回転",
            variable=rotate_option, value=90,
            **toggle_style).grid(row=0, column=2, padx=5)

Label(root, text="操作", bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=15)

btn_merge = Button(root, text="🔗 結合",
                   command=lambda: run_with_loading(merge_pdfs),
                   state=DISABLED, **btn_style)

btn_split = Button(root, text="✂ 分割",
                   command=lambda: run_with_loading(split_pdfs),
                   state=DISABLED, **btn_style)

btn_rotate = Button(root, text="🔄 回転実行",
                    command=lambda: run_with_loading(rotate_pdfs),
                    state=DISABLED, **btn_style)

btn_text = Button(root, text="📝 テキスト抽出",
                  command=lambda: run_with_loading(extract_text),
                  state=DISABLED, **btn_style)

btn_merge.pack(pady=5)
btn_split.pack(pady=5)
btn_rotate.pack(pady=5)
btn_text.pack(pady=5)

root.mainloop()
