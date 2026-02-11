# -*- coding: utf-8 -*-
"""
シンプルPDF編集デスクトップアプリ
機能：分割 / 結合 / 回転 / テキスト抽出
回転は 90 / 180 / 270 度のトグル選択式
選択したファイル／フォルダのパスを表示
1ファイル完結版
"""

import os
from tkinter import *
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

# ==========================
# 共通処理
# ==========================

selected_files = []
selected_folder = ""

def select_files():
    global selected_files, selected_folder
    files = filedialog.askopenfilenames(
        filetypes=[("PDFファイル", "*.pdf")]
    )
    if files:
        selected_files = list(files)
        selected_folder = ""
        update_path_display()

def select_folder():
    global selected_folder, selected_files
    folder = filedialog.askdirectory()
    if folder:
        selected_folder = folder
        selected_files = []
        update_path_display()

def update_path_display():
    if selected_files:
        paths = "\n".join(selected_files)
        text_paths.delete(1.0, END)
        text_paths.insert(END, paths)
    elif selected_folder:
        text_paths.delete(1.0, END)
        text_paths.insert(END, selected_folder)
    else:
        text_paths.delete(1.0, END)

def get_target_files():
    if selected_files:
        return selected_files
    elif selected_folder:
        pdfs = []
        for f in os.listdir(selected_folder):
            if f.lower().endswith(".pdf"):
                pdfs.append(os.path.join(selected_folder, f))
        return pdfs
    else:
        return []

def get_save_dir(original_path):
    if save_option.get() == 1:
        return os.path.dirname(original_path)
    else:
        return filedialog.askdirectory()

# ==========================
# PDF操作
# ==========================

def merge_pdfs():
    try:
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

        base = os.path.splitext(os.path.basename(files[0]))[0]
        output_path = os.path.join(save_dir, base + "_Merge.pdf")

        with open(output_path, "wb") as f:
            writer.write(f)

        messagebox.showinfo("完了", "結合完了")
    except Exception:
        messagebox.showerror("エラー", "結合失敗（0扱い）")

def split_pdfs():
    try:
        files = get_target_files()
        if not files:
            raise Exception()

        for file in files:
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

        messagebox.showinfo("完了", "分割完了")
    except Exception:
        messagebox.showerror("エラー", "分割失敗（0扱い）")

def rotate_pdfs():
    try:
        files = get_target_files()
        if not files:
            raise Exception()

        degree = rotate_option.get()
        if degree == 0:
            raise Exception()

        for file in files:
            reader = PdfReader(file)
            writer = PdfWriter()

            for page in reader.pages:
                page.rotate(degree)
                writer.add_page(page)

            save_dir = get_save_dir(file)
            if not save_dir:
                return

            base = os.path.splitext(os.path.basename(file))[0]
            output_path = os.path.join(save_dir, f"{base}_Rotate_{degree}.pdf")

            with open(output_path, "wb") as f:
                writer.write(f)

        messagebox.showinfo("完了", "回転完了")
    except Exception:
        messagebox.showerror("エラー", "回転失敗（角度未選択は0扱い）")

def extract_text():
    try:
        files = get_target_files()
        if not files:
            raise Exception()

        for file in files:
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

        messagebox.showinfo("完了", "テキスト抽出完了")
    except Exception:
        messagebox.showerror("エラー", "テキスト抽出失敗（0扱い）")

# ==========================
# UI構築
# ==========================

root = Tk()
root.title("PDF編集アプリ")
root.geometry("600x700")
root.minsize(600, 700)

Label(root, text="PDF編集ツール", font=("Arial", 16)).pack(pady=10)

Button(root, text="ファイル選択", command=select_files, width=25).pack(pady=5)
Button(root, text="フォルダ選択", command=select_folder, width=25).pack(pady=5)

Label(root, text="選択パス").pack(pady=5)

text_paths = Text(root, height=10, width=70)
text_paths.pack(pady=5)

Label(root, text="保存先").pack()
save_option = IntVar(value=1)
Radiobutton(root, text="同じフォルダ", variable=save_option, value=1).pack()
Radiobutton(root, text="任意のフォルダ", variable=save_option, value=2).pack()

Label(root, text="回転角度").pack(pady=5)

rotate_option = IntVar(value=0)

frame_rotate = Frame(root)
frame_rotate.pack()

Radiobutton(frame_rotate, text="90°", variable=rotate_option, value=90, indicatoron=False, width=8).grid(row=0, column=0, padx=5)
Radiobutton(frame_rotate, text="180°", variable=rotate_option, value=180, indicatoron=False, width=8).grid(row=0, column=1, padx=5)
Radiobutton(frame_rotate, text="270°", variable=rotate_option, value=270, indicatoron=False, width=8).grid(row=0, column=2, padx=5)

Label(root, text="操作").pack(pady=10)

Button(root, text="結合", command=merge_pdfs, width=25).pack(pady=5)
Button(root, text="分割", command=split_pdfs, width=25).pack(pady=5)
Button(root, text="回転", command=rotate_pdfs, width=25).pack(pady=5)
Button(root, text="テキスト抽出", command=extract_text, width=25).pack(pady=5)

root.mainloop()
