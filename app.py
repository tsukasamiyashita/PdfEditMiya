# -*- coding: utf-8 -*-
"""
PdfEditMiya（保存先 初期＝同じフォルダ）
・青ベースUI
・保存先は初期状態で「同じフォルダ」
・任意保存先は事前選択可能（未選択なら実行時に選択）
・回転はラジオボタン（初期：左回転）
・処理中ポップアップ表示
・完了は3秒後に自動クローズ
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
current_mode = None
processing_popup = None
preset_save_dir = ""   # 任意保存先

# ==========================
# ポップアップ
# ==========================

def show_processing(msg="処理実行中..."):
    global processing_popup
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("260x100")
    processing_popup.configure(bg="#E3F2FD")
    processing_popup.resizable(False, False)

    Label(processing_popup, text=msg,
          bg="#E3F2FD", fg="#1565C0",
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
    win.resizable(False, False)

    bg = "#FFEBEE" if error else "#E3F2FD"
    fg = "#C62828" if error else "#1565C0"

    win.configure(bg=bg)
    Label(win, text=msg, bg=bg, fg=fg,
          font=("Segoe UI", 10, "bold")).pack(expand=True)

    win.after(3000, win.destroy)

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
        update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        selected_files = []
        current_mode = "folder"
        update_ui()

def select_save_dir():
    global preset_save_dir
    folder = filedialog.askdirectory()
    if folder:
        preset_save_dir = folder
        save_dir_label.config(text=f"保存先: {preset_save_dir}")

def update_ui():
    if current_mode == "file":
        mode_label.config(text="📄 ファイル選択中", fg="#1565C0")
    elif current_mode == "folder":
        mode_label.config(text="📁 フォルダ選択中", fg="#2E7D32")
    else:
        mode_label.config(text="未選択", fg="#666666")

    text_paths.config(state=NORMAL)
    text_paths.delete(1.0, END)
    if selected_files:
        text_paths.insert(END, "\n".join(selected_files))
    elif selected_folder:
        text_paths.insert(END, selected_folder)
    text_paths.config(state=DISABLED)

    btn_merge.config(state=DISABLED)
    btn_split.config(state=DISABLED)
    btn_rotate.config(state=DISABLED)
    btn_text.config(state=DISABLED)

    if current_mode == "file":
        btn_split.config(state=NORMAL)
        btn_rotate.config(state=NORMAL)
        btn_text.config(state=NORMAL)
    elif current_mode == "folder":
        btn_merge.config(state=NORMAL)

# ==========================
# 共通処理
# ==========================

def get_target_files():
    if selected_files:
        return selected_files
    if selected_folder:
        return [os.path.join(selected_folder, f)
                for f in os.listdir(selected_folder)
                if f.lower().endswith(".pdf")]
    return []

def get_save_dir(original_path):
    # ★ 初期は同じフォルダ
    if save_option.get() == 1:
        return os.path.dirname(original_path)

    # ★ 任意フォルダ選択
    global preset_save_dir
    if preset_save_dir:
        return preset_save_dir

    folder = filedialog.askdirectory()
    if folder:
        preset_save_dir = folder
        save_dir_label.config(text=f"保存先: {preset_save_dir}")
        return folder

    return None

def run_task(func):
    def task():
        try:
            show_processing()
            func()
            close_processing()
            auto_close_message("完了", "処理が完了しました")
        except Exception:
            close_processing()
            auto_close_message("エラー", "処理失敗（0扱い）", True)

    threading.Thread(target=task).start()

# ==========================
# PDF処理
# ==========================

def merge_pdfs():
    files = get_target_files()
    if not files:
        raise Exception()

    writer = PdfWriter()
    for f in files:
        reader = PdfReader(f)
        for p in reader.pages:
            writer.add_page(p)

    save_dir = get_save_dir(files[0])
    if not save_dir:
        return

    name = os.path.basename(selected_folder)
    with open(os.path.join(save_dir, name + "_Merge.pdf"), "wb") as out:
        writer.write(out)

def split_pdfs():
    for f in selected_files:
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        for i, p in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(p)
            with open(os.path.join(save_dir,
                     f"{base}_Split_{i+1}.pdf"), "wb") as out:
                writer.write(out)

def rotate_pdfs():
    deg = rotate_option.get()
    for f in selected_files:
        reader = PdfReader(f)
        writer = PdfWriter()
        for p in reader.pages:
            p.rotate(deg)
            writer.add_page(p)
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir,
                 f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)

def extract_text():
    for f in selected_files:
        reader = PdfReader(f)
        text = ""
        for p in reader.pages:
            t = p.extract_text()
            text += t if t else ""
        save_dir = get_save_dir(f)
        if not save_dir:
            return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir,
                 f"{base}_Text.txt"), "w",
                 encoding="utf-8") as out:
            out.write(text)

# ==========================
# UI
# ==========================

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
WHITE = "#FFFFFF"

root = Tk()
root.title("PdfEditMiya")
root.geometry("600x780")
root.minsize(600, 780)
root.configure(bg=LIGHT)

Label(root, text="PdfEditMiya",
      font=("Segoe UI", 18, "bold"),
      bg=LIGHT, fg=PRIMARY).pack(pady=10)

mode_label = Label(root, text="未選択",
                   bg=LIGHT, font=("Segoe UI", 11, "bold"))
mode_label.pack(pady=5)

btn_style = {
    "font": ("Segoe UI", 9, "bold"),
    "bg": PRIMARY,
    "fg": WHITE,
    "activebackground": "#1E88E5",
    "bd": 0,
    "width": 20,
    "height": 1
}

Button(root, text="📄 ファイル選択",
       command=select_files, **btn_style).pack(pady=4)

Button(root, text="📁 フォルダ選択",
       command=select_folder, **btn_style).pack(pady=4)

Label(root, text="選択パス",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=6)

text_paths = Text(root, height=5, width=70,
                  font=("Consolas", 9), bd=0)
text_paths.pack()
text_paths.config(state=DISABLED)

# ==========================
# 保存先設定
# ==========================

Label(root, text="保存先設定",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=8)

save_option = IntVar(value=1)  # ★ 初期＝同じフォルダ

Radiobutton(root, text="同じフォルダ（初期）",
            variable=save_option, value=1,
            bg=LIGHT).pack()

Radiobutton(root, text="任意のフォルダ",
            variable=save_option, value=2,
            bg=LIGHT).pack()

Button(root, text="📂 任意保存先を事前選択",
       command=select_save_dir, **btn_style).pack(pady=3)

save_dir_label = Label(root,
                       text="保存先: 同じフォルダ",
                       bg=LIGHT, font=("Segoe UI", 9))
save_dir_label.pack(pady=3)

# ==========================
# 回転
# ==========================

Label(root, text="回転方法",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=8)

rotate_option = IntVar(value=270)

Radiobutton(root, text="左回転（270°）",
            variable=rotate_option, value=270,
            bg=LIGHT).pack()

Radiobutton(root, text="上下回転（180°）",
            variable=rotate_option, value=180,
            bg=LIGHT).pack()

Radiobutton(root, text="右回転（90°）",
            variable=rotate_option, value=90,
            bg=LIGHT).pack()

# ==========================
# 操作
# ==========================

Label(root, text="操作",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 11, "bold")).pack(pady=10)

btn_merge = Button(root, text="🔗 結合",
                   command=lambda: run_task(merge_pdfs),
                   state=DISABLED, **btn_style)
btn_split = Button(root, text="✂ 分割",
                   command=lambda: run_task(split_pdfs),
                   state=DISABLED, **btn_style)
btn_rotate = Button(root, text="🔄 回転",
                    command=lambda: run_task(rotate_pdfs),
                    state=DISABLED, **btn_style)
btn_text = Button(root, text="📝 テキスト抽出",
                  command=lambda: run_task(extract_text),
                  state=DISABLED, **btn_style)

btn_merge.pack(pady=3)
btn_split.pack(pady=3)
btn_rotate.pack(pady=3)
btn_text.pack(pady=3)

root.mainloop()
