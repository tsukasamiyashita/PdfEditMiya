# -*- coding: utf-8 -*-
"""
PdfEditMiya
・青ベースUI
・操作ボタンは実行可能時のみ強調色
・実行不可時は背景を通常（透明風＝画面色と同じ）に
・保存先 初期＝同じフォルダ
・処理中ポップアップ表示
・完了は3秒後自動クローズ
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
preset_save_dir = ""
cancelled = False

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
WHITE = "#FFFFFF"

RUN_COLOR = "#43A047"
RUN_ACTIVE = "#2E7D32"

# ==========================
# ポップアップ
# ==========================

def show_processing(msg="処理実行中..."):
    global processing_popup
    processing_popup = Toplevel(root)
    processing_popup.title("実行中")
    processing_popup.geometry("220x90")
    processing_popup.configure(bg=LIGHT)
    processing_popup.resizable(False, False)

    Label(processing_popup, text="⏳ " + msg,
          bg=LIGHT, fg=PRIMARY,
          font=("Segoe UI", 9, "bold")).pack(expand=True)

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
    win.geometry("220x90")
    win.resizable(False, False)

    bg = "#FFEBEE" if error else LIGHT
    fg = "#C62828" if error else PRIMARY

    win.configure(bg=bg)
    Label(win, text=msg, bg=bg, fg=fg,
          font=("Segoe UI", 9, "bold")).pack(expand=True)

    win.after(3000, win.destroy)

# ==========================
# ボタン有効/無効デザイン制御
# ==========================

def set_button_state(btn, enabled):
    if enabled:
        btn.config(
            state=NORMAL,
            bg=RUN_COLOR,
            activebackground=RUN_ACTIVE,
            fg=WHITE,
            cursor="hand2"
        )
    else:
        btn.config(
            state=DISABLED,
            bg=LIGHT,              # 透明風（画面と同色）
            activebackground=LIGHT,
            fg="#90A4AE",
            cursor="arrow"
        )

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

def on_save_option_change():
    global preset_save_dir
    if save_option.get() == 2:
        preset_save_dir = ""
        save_dir_label.config(text="保存先: 未選択")
    else:
        preset_save_dir = ""
        save_dir_label.config(text="保存先: 同じフォルダ")

# ==========================
# UI更新
# ==========================

def update_ui():
    if current_mode == "file":
        mode_label.config(text="📄 ファイル選択中", fg=PRIMARY)
        if len(selected_files) == 1:
            path_text = selected_files[0]
        else:
            path_text = f"{len(selected_files)}件のPDFを選択中"
    elif current_mode == "folder":
        mode_label.config(text="📁 フォルダ選択中", fg="#2E7D32")
        path_text = selected_folder
    else:
        mode_label.config(text="未選択", fg="#666666")
        path_text = "未選択"

    path_label.config(text=f"選択パス:\n{path_text}")

    # いったん全無効
    set_button_state(btn_merge, False)
    set_button_state(btn_split, False)
    set_button_state(btn_rotate, False)
    set_button_state(btn_text, False)

    if current_mode == "file":
        set_button_state(btn_split, True)
        set_button_state(btn_rotate, True)
        set_button_state(btn_text, True)

    elif current_mode == "folder":
        set_button_state(btn_merge, True)

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

            auto_close_message("完了", "✅ 処理が完了しました")

        except Exception:
            close_processing()
            auto_close_message("エラー", "❌ 処理失敗", True)

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
# UI構築
# ==========================

root = Tk()
root.title("PdfEditMiya")
root.geometry("360x650")
root.minsize(360, 650)
root.configure(bg=LIGHT)

Label(root, text="PdfEditMiya",
      font=("Segoe UI", 14, "bold"),
      bg=LIGHT, fg=PRIMARY).pack(pady=6)

mode_label = Label(root, text="未選択",
                   bg=LIGHT, font=("Segoe UI", 9, "bold"))
mode_label.pack(pady=3)

Button(root, text="📄 ファイル選択",
       command=select_files,
       bg=PRIMARY, fg=WHITE,
       activebackground="#1E88E5",
       width=22, height=1,
       bd=0).pack(pady=2)

Button(root, text="📁 フォルダ選択",
       command=select_folder,
       bg=PRIMARY, fg=WHITE,
       activebackground="#1E88E5",
       width=22, height=1,
       bd=0).pack(pady=2)

Label(root, text="選択パス",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 9, "bold")).pack(pady=4)

path_label = Label(root,
                   text="選択パス:\n未選択",
                   bg=LIGHT,
                   wraplength=320,
                   justify=LEFT,
                   font=("Segoe UI", 8))
path_label.pack(pady=2)

Label(root, text="保存先設定",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 9, "bold")).pack(pady=4)

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
       command=select_save_dir,
       bg=PRIMARY, fg=WHITE,
       activebackground="#1E88E5",
       width=22, height=1,
       bd=0).pack(pady=2)

save_dir_label = Label(root,
                       text="保存先: 同じフォルダ",
                       bg=LIGHT,
                       font=("Segoe UI", 8))
save_dir_label.pack(pady=2)

Label(root, text="回転方法",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 9, "bold")).pack(pady=4)

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

Label(root, text="操作",
      bg=LIGHT, fg=PRIMARY,
      font=("Segoe UI", 10, "bold")).pack(pady=6)

btn_merge = Button(root, text="▶ 結合を実行", width=22, height=2, bd=0)
btn_split = Button(root, text="▶ 分割を実行", width=22, height=2, bd=0)
btn_rotate = Button(root, text="▶ 回転を実行", width=22, height=2, bd=0)
btn_text = Button(root, text="▶ テキスト抽出を実行", width=22, height=2, bd=0)

btn_merge.config(command=lambda: run_task(merge_pdfs))
btn_split.config(command=lambda: run_task(split_pdfs))
btn_rotate.config(command=lambda: run_task(rotate_pdfs))
btn_text.config(command=lambda: run_task(extract_text))

btn_merge.pack(pady=3)
btn_split.pack(pady=3)
btn_rotate.pack(pady=3)
btn_text.pack(pady=3)

update_ui()

root.mainloop()
