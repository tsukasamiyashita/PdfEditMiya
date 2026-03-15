# -*- coding: utf-8 -*-
"""
PdfEditMiya v1.4.0
------------------
更新情報:
- v1.4.0: メニューバーの設定を廃止し、UIのAPIキー枠に一本化（テストボタン追加）。ヘルプにインストール手順等を追加。
- v1.4.0_fix3: 図面名称列の省略記号（点々、「〃」など）がある場合、上の行の図面名称を引き継ぐようプロンプトを修正。
- v1.3.0: 手書き文字抽出に生成AI（Gemini API）を導入、Tesseractとの切り替え対応。
"""

import os
import sys
import threading
import io
import gc
import cv2
import csv
import time
import numpy as np
from tkinter import *
from tkinter import ttk, filedialog, messagebox, Menu
from tkinter.simpledialog import askstring
import tkinter.scrolledtext as st
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side
import fitz  # PyMuPDF
import ezdxf
from PIL import Image

import pytesseract
import google.generativeai as genai

# ==============================
# リソースパス取得関数
# ==============================
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==============================
# 基本設定
# ==============================
APP_TITLE = "PdfEditMiya"
VERSION = "v1.4.0"
WINDOW_WIDTH = 680
WINDOW_HEIGHT = 860

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
SUCCESS = "#2E7D32"
ERROR = "#C62828"
INACTIVE = "#90A4AE"
INFO_TEXT = "#455A64"

USER_HOME = os.path.expanduser("~")
API_KEY_FILE = os.path.join(USER_HOME, ".pdfeditmiya_api_key.txt")

VERSION_HISTORY = """
[ v1.4.0 ]
- AI抽出のプロンプト精度を極限まで最適化（行ズレの絶対防止、連番推論、専門用語補正、省略記号の引継ぎ）
- メニューバーの「設定」を廃止し、メイン画面のAPIキー入力枠に一本化
- APIキーの「テスト」ボタンを入力枠の横に配置
- メニューバーの「ヘルプ」に、TesseractやGemini APIの導入手順(使い方)を追加

[ v1.3.0 ]
- 手書き文字抽出に生成AI（Gemini API）を搭載
- AI抽出設定UIを追加し、Tesseractとの切り替え、出力形式(txt/csv/xlsx)の選択に対応
- 処理中の進捗画面を「全体」と「個別(ページ)」の2段階表示に進化
- APIキーの保存先を変更し、GitHubへの流出を防止するセキュリティ対策を実施
- 利用可能なAIモデルの自動検索システムを搭載
"""

AI_HELP_TEXT = """
【 AI抽出機能を使うための準備 】

■ Gemini API を使う場合（推奨・超高精度）
1. 以下のURLにアクセスし、Googleアカウントでログインします。
   https://aistudio.google.com/app/apikey
2. 「Create API key」ボタンを押し、新しいプロジェクトでAPIキーを作成します。
3. 発行された文字列（AIza...から始まるもの）をコピーします。
4. 本アプリの「AI抽出設定」の「APIキー」枠に貼り付け、「テスト」ボタンを押して認証成功と出れば準備完了です。
※ APIキーはPC内に安全に隠しファイルとして自動保存されるため、GitHub等へ流出することはありません。

■ Tesseract を使う場合（オフライン・簡易抽出）
Windows環境にTesseract OCR本体がインストールされている必要があります。
1. 以下のサイト等からWindows用のインストーラーをダウンロードし、インストールします。
   https://github.com/UB-Mannheim/tesseract/wiki
2. インストール時、「Additional language data (download)」を展開し、「Japanese」に必ずチェックを入れてください。
3. インストール先のパス（例: C:\\Program Files\\Tesseract-OCR）に環境変数(Path)を通すか、PCを再起動してください。
"""

# ==============================
# グローバル変数
# ==============================
selected_files = []
selected_folder = ""
current_mode = None
preset_save_dir = ""

processing_popup = None
overall_label = None
overall_progress = None
file_label = None
file_progress = None
cancelled = False

# ==============================
# APIキー管理 ＆ モデル自動検索
# ==============================
def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return None

def get_available_models(api_key):
    genai.configure(api_key=api_key)
    models_to_try = []
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.replace('models/', '')
                if name not in models_to_try:
                    models_to_try.append(name)
    except Exception:
        pass
        
    fallbacks = [
        'gemini-2.5-flash', 'gemini-2.0-flash', 
        'gemini-1.5-flash', 'gemini-1.5-pro',
        'gemini-1.0-pro-vision-latest', 'gemini-pro-vision'
    ]
    for f in fallbacks:
        if f not in models_to_try:
            models_to_try.append(f)
    return models_to_try

def test_api_key_ui():
    key = api_key_var.get().strip()
    if not key:
        messagebox.showwarning("警告", "APIキーが入力されていません。")
        return
        
    models_to_try = get_available_models(key)
    success = False
    last_err = ""
    for model_name in models_to_try:
        try:
            model = genai.GenerativeModel(model_name)
            model.generate_content("Test")
            success = True
            break
        except Exception as e:
            last_err = str(e)
            continue
            
    if success:
        with open(API_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(key)
        messagebox.showinfo("認証成功！", "APIキーは正しく認識されました。\n安全に保存しました。")
    else:
        err_msg = last_err.lower()
        if "404" in err_msg or "not found" in err_msg:
            messagebox.showerror("権限エラー", "APIキーは認識されましたが、AIモデルを利用する権限がありません（404エラー）。\nGoogle AI Studioで新しくキーを作成し直してください。")
        else:
            messagebox.showerror("通信エラー", f"APIキー確認中にエラーが発生しました。\n{last_err}")

# ==============================
# 共通ロジック
# ==============================
def get_target_files():
    if current_mode == "file": return selected_files
    if current_mode == "folder" and selected_folder:
        return [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith(".pdf")]
    return []

def get_save_dir(original_path):
    global preset_save_dir
    if save_option.get() == 1: 
        return os.path.dirname(original_path)
    if preset_save_dir: 
        return preset_save_dir
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

def run_task(func):
    global cancelled
    cancelled = False
    try:
        files = get_target_files()
        if not files: return
        func(files)
        close_processing()
        if not cancelled:
            show_message("✅ 完了", SUCCESS)
    except Exception as e:
        print(f"Error: {e}")
        close_processing()
        show_message("❌ エラー発生", ERROR)

def safe_run(func):
    files = get_target_files()
    if not files: return
    global preset_save_dir
    if save_option.get() == 2 and not preset_save_dir:
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        if not folder: return
        preset_save_dir = folder
        save_label.config(text=preset_save_dir)

    show_processing(len(files))
    threading.Thread(target=run_task, args=(func,), daemon=True).start()

# ==============================
# UI補助機能
# ==============================
def show_message(msg, color=PRIMARY):
    def _task():
        win = Toplevel(root)
        win.geometry("220x90")
        win.configure(bg=LIGHT)
        win.attributes("-topmost", True)
        x = root.winfo_x() + (WINDOW_WIDTH // 2) - 110
        y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 45
        win.geometry(f"+{x}+{y}")
        Label(win, text=msg, bg=LIGHT, fg=color, font=("Segoe UI", 12, "bold")).pack(expand=True)
        win.after(2500, win.destroy)
    root.after(0, _task)

def show_processing(total_files=1):
    global processing_popup, overall_label, overall_progress, file_label, file_progress
    processing_popup = Toplevel(root)
    processing_popup.title("処理中")
    processing_popup.geometry("400x180")
    processing_popup.configure(bg=LIGHT)
    processing_popup.grab_set()
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 200
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 90
    processing_popup.geometry(f"+{x}+{y}")
    
    overall_label = Label(processing_popup, text=f"全体の進捗 ( 0 / {total_files} ファイル )", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"))
    overall_label.pack(pady=(15, 2))
    overall_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=total_files, length=340)
    overall_progress.pack(pady=(0, 10))
    
    file_label = Label(processing_popup, text="現在のファイルを準備中...", bg=LIGHT, fg=INFO_TEXT, font=("Segoe UI", 9))
    file_label.pack(pady=(5, 2))
    file_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=1, length=340)
    file_progress.pack(pady=(0, 10))

def close_processing():
    def _task():
        global processing_popup
        if processing_popup:
            processing_popup.destroy()
            processing_popup = None
    root.after(0, _task)

def update_overall_progress(step, max_val=None, text=None):
    def _task():
        if processing_popup and processing_popup.winfo_exists():
            if max_val is not None: overall_progress["maximum"] = max_val
            overall_progress["value"] = step
            if text: overall_label.config(text=text)
            overall_progress.update()
    root.after(0, _task)

def update_file_progress(step, max_val=None, text=None):
    def _task():
        if processing_popup and processing_popup.winfo_exists():
            if max_val is not None: file_progress["maximum"] = max_val
            file_progress["value"] = step
            if text: file_label.config(text=text)
            file_progress.update()
    root.after(0, _task)

# ==============================
# メニューバー機能
# ==============================
def show_text_window(title, content):
    win = Toplevel(root)
    win.title(title)
    win.geometry("580x450")
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 290
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 225
    win.geometry(f"+{x}+{y}")
    text_area = st.ScrolledText(win, wrap=WORD, font=("Meiryo UI", 10))
    text_area.pack(expand=True, fill=BOTH, padx=10, pady=10)
    text_area.insert(END, content)
    text_area.configure(state=DISABLED)

def show_ai_help():
    show_text_window("AI抽出の準備 (使い方)", AI_HELP_TEXT.strip())

def show_version_info():
    msg = f"{APP_TITLE}\nバージョン: {VERSION}\n\nPython & Tkinter製 PDF編集ツール"
    messagebox.showinfo("バージョン情報", msg)

def show_history():
    show_text_window("バージョン履歴", VERSION_HISTORY.strip())

def show_readme():
    readme_path = resource_path("readme.md")
    content = ""
    if os.path.exists(readme_path):
        with open(readme_path, "r", encoding="utf-8") as f: content = f.read()
    else:
        content = "readme.md ファイルが見つかりませんでした。"
    show_text_window("Readme", content)

# ==============================
# PDF操作コア機能
# ==============================
def merge_pdfs(files):
    total_files = len(files)
    writer = PdfWriter()
    update_overall_progress(0, total_files, f"全体の進捗 ( 0 / {total_files} ファイル )")
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        total_pages = len(reader.pages)
        for j, p in enumerate(reader.pages, 1):
            update_file_progress(j, total_pages, f"ファイルを結合中... ( {j} / {total_pages} ページ )")
            writer.add_page(p)
    update_file_progress(1, 1, "PDFを保存中...")
    save_dir = get_save_dir(files[0])
    if save_dir:
        name = os.path.basename(selected_folder) if selected_folder else "Merged"
        with open(os.path.join(save_dir, f"{name}_Merge.pdf"), "wb") as out:
            writer.write(out)

def split_pdfs(files):
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        total_pages = len(reader.pages)
        for n, p in enumerate(reader.pages, 1):
            update_file_progress(n, total_pages, f"ファイルを分割中... ( {n} / {total_pages} ページ )")
            writer = PdfWriter()
            writer.add_page(p)
            with open(os.path.join(save_dir, f"{base}_Split_{n}.pdf"), "wb") as out:
                writer.write(out)

def rotate_pdfs(files):
    deg = rotate_option.get()
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        for j, p in enumerate(reader.pages, 1):
            update_file_progress(j, total_pages, f"ページを回転中... ( {j} / {total_pages} ページ )")
            p.rotate(deg)
            writer.add_page(p)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)

def extract_text(files):
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        total_pages = len(reader.pages)
        text_list = []
        for j, p in enumerate(reader.pages, 1):
            update_file_progress(j, total_pages, f"テキストを抽出中... ( {j} / {total_pages} ページ )")
            text_list.append(p.extract_text() or "")
        text = "".join(text_list)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out:
            out.write(text)

# ==============================
# AI抽出ロジック
# ==============================
def run_ai_extraction():
    engine = ai_engine_var.get()
    if engine == "Gemini":
        key = api_key_var.get().strip()
        if not key:
            messagebox.showerror("エラー", "Gemini APIキーを入力してください。")
            return
        with open(API_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(key)
        safe_run(extract_gemini_task)
    else:
        safe_run(extract_tesseract_task)

def extract_tesseract_task(files):
    out_format = output_format_var.get()
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            save_dir = get_save_dir(f)
            if not save_dir: return
            base = os.path.splitext(os.path.basename(f))[0]
            doc = fitz.open(f)
            total_pages = len(doc)
            
            wb = Workbook() if out_format == "xlsx" else None
            if wb: wb.remove(wb.active)
            all_text_list = []
            
            for page_num in range(total_pages):
                update_file_progress(page_num + 1, total_pages, f"Tesseractで解析中... ( {page_num+1} / {total_pages} ページ )")
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                mode = "RGB" if pix.n == 3 else "L"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")
                
                try:
                    text = pytesseract.image_to_string(img, lang="jpn+eng")
                    if out_format == "xlsx":
                        ws = wb.create_sheet(f"Page_{page_num+1}")
                        for row_idx, line in enumerate(text.split('\n'), 1):
                            ws.cell(row=row_idx, column=1, value=line.strip())
                    else:
                        all_text_list.append(text)
                except Exception as e:
                    err_msg = f"[ --- ページ {page_num+1} 解析失敗 --- ]\nTesseract OCRエラー\n詳細: {e}"
                    if out_format == "xlsx":
                        ws = wb.create_sheet(f"Page_{page_num+1}")
                        ws.cell(row=1, column=1, value=err_msg)
                    else:
                        all_text_list.append(err_msg)
                gc.collect()
            
            doc.close()
            update_file_progress(total_pages, total_pages, "ファイルを保存中...")
            
            if out_format == "xlsx":
                wb.save(os.path.join(save_dir, f"{base}_Tesseract抽出.xlsx"))
            elif out_format == "csv":
                with open(os.path.join(save_dir, f"{base}_Tesseract抽出.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                    f_out.write("\n\n".join(all_text_list))
            elif out_format == "txt":
                with open(os.path.join(save_dir, f"{base}_Tesseract抽出.txt"), "w", encoding="utf-8") as f_out:
                    f_out.write("\n\n".join(all_text_list))

        except Exception as e:
            print(f"Tesseract Error: {e}")
            raise e

def extract_gemini_task(files):
    key = api_key_var.get().strip()
    genai.configure(api_key=key)
    models_to_try = get_available_models(key)
    out_format = output_format_var.get()

    if out_format in ["csv", "xlsx"]:
        prompt = """
        この画像は手書きの図面台帳（表）です。行のズレは致命的なエラーとなるため、画像の表の「物理的な行数」とぴったり一致するように、1行ずつ絶対にスキップせずに抽出してください。

        【出力フォーマット（厳守）】
        ・1行目は画像右上のタイトルをE列に配置（例: `,,,,佐渡C振`）
        ・2行目はヘッダー（固定）: `機械名,図面名称,,整理番号,備考` （C列のヘッダーは必ず空欄にしてください）
        ・3行目以降は画像内の表データ。各行必ず5列になるようにカンマを4つ含めてください。

        【各列の抽出・補完ルール（重要）】
        1列目 (機械名): 「No1送り改造」や「No.2 修理組立説明図」など。下の行が空白（または省略記号）になっている場合、必ず【上の行の文字列をそのままコピー】してすべての行の1列目を埋めてください。
        2列目 (図面名称): 読み取りが難しい手書き文字を以下のルールに従って補正してください。
           - 「細立図」や「立図」 → 「組立図」
           - 「新品図」や「部品図」のかすれ → 「部品図」
           - 「写真事」など → 「ケッチ台車」
           - 「チェーン送り送り」 → 「チェーン送り」
           - 「チェーンコンベア駆動部新品図」等 → 「チェーンコンベア駆動部改造」
           - 「修理相互説明図」 → 「修理組立説明図」
           - ★重要1★ 点々（「〃」など）の同じを表す省略記号が書かれている場合は、必ず【直上の行の図面名称をそのままコピーして引き継いで】ください。（機械名と同じ処理）
           - ★重要2★「No.2 修理組立説明図」のブロックにおいて、3列目が「4/4」の行のみ図面名称を「部品図」とし、「1/4」「2/4」「3/4」の行は省略記号ではなく完全な空白であるため、図面名称は絶対に空欄のままにしてください。
        3列目 (ページ・分数): 絶対に日付変換しないでください。「No.1 送り改造」のブロックの分母は「12」です。AIがかすれによって「/2」と誤読しやすいですが、正しくは「1/12, 2/12, 3/12, 4/12, 5/12, 6/12, 7/12, 8/12, 9/12, 10/12, 11/12, 12/12」という連番になります。また「No.2 修理組立説明図」のブロックは「1/4, 2/4, 3/4, 4/4」となります。この完全な連番規則に従って出力してください。
        4列目 (整理番号): 「10194」から始まる5桁の数字です。上から下へ「10194, 10195, 10196...」と【1行につき1ずつ増える完全な連番】です。欠損していても、必ず前の行に+1した数値を出力し、行のズレを防ぐためのアンカー（基準）としてください。
        5列目 (備考): 空白で構いません。

        【最も重要な禁止事項】
        ・行を勝手に結合したり、読み飛ばしたりしないでください。表の行数と出力する行数が完全に一致しなければなりません。
        ・余計な挨拶や ```csv などの記号は一切出力せず、純粋なテキストのみ出力してください。
        """
    else:
        prompt = "この画像に記載されている手書きの文字や文章を可能な限り正確に読み取り、プレーンテキストとして出力してください。余計な挨拶や説明文は一切含めないでください。"

    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            save_dir = get_save_dir(f)
            if not save_dir: return
            
            base = os.path.splitext(os.path.basename(f))[0]
            doc = fitz.open(f)
            total_pages = len(doc)
            
            wb = Workbook() if out_format == "xlsx" else None
            if wb: wb.remove(wb.active)
            all_text_list = []
            
            for page_num in range(total_pages):
                update_file_progress(page_num + 1, total_pages, f"AIが解析中... ( {page_num+1} / {total_pages} ページ )")
                ws = wb.create_sheet(f"Page_{page_num+1}") if wb else None
                
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                mode = "RGB" if pix.n == 3 else "L"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")
                
                max_retries = 3
                extracted_text, success, last_error, used_model = "", False, "", ""

                for attempt in range(max_retries):
                    for model_name in models_to_try:
                        update_file_progress(page_num + 1, text=f"AIが解析中... ( {page_num+1} / {total_pages} ページ ) [使用: {model_name}]")
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content([prompt, img])
                            if not response.parts: raise Exception("安全フィルタ等によりブロックされました。")
                            extracted_text = response.text.strip()
                            success, used_model = True, model_name
                            break
                        except Exception as api_err:
                            last_error = str(api_err)
                            continue
                    if success: break
                    time.sleep(2 ** attempt)

                if success:
                    update_file_progress(page_num + 1, text=f"解析成功！ [モデル: {used_model}]")
                    markdown_marker = "`" * 3
                    if extracted_text.startswith(markdown_marker):
                        lines = extracted_text.split('\n')
                        if len(lines) > 0 and lines[0].startswith(markdown_marker): lines = lines[1:]
                        if len(lines) > 0 and lines[-1].startswith(markdown_marker): lines = lines[:-1]
                        extracted_text = '\n'.join(lines)
                    
                    if out_format == "xlsx":
                        reader = csv.reader(io.StringIO(extracted_text))
                        for row_idx, row in enumerate(reader, 1):
                            for col_idx, val in enumerate(row, 1):
                                ws.cell(row=row_idx, column=col_idx, value=val.strip())
                    else:
                        all_text_list.append(extracted_text)
                else:
                    err_msg = f"[ --- ページ {page_num+1} の解析に失敗しました --- ]\nエラー詳細: {last_error}"
                    if out_format == "xlsx":
                        ws.cell(row=1, column=1, value=err_msg)
                    else:
                        all_text_list.append(err_msg)
                gc.collect()
            
            doc.close()
            update_file_progress(total_pages, total_pages, "ファイルを保存中...")
            
            if out_format == "xlsx":
                wb.save(os.path.join(save_dir, f"{base}_AI抽出.xlsx"))
            elif out_format == "csv":
                with open(os.path.join(save_dir, f"{base}_AI抽出.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                    f_out.write("\n\n".join(all_text_list))
            elif out_format == "txt":
                with open(os.path.join(save_dir, f"{base}_AI抽出.txt"), "w", encoding="utf-8") as f_out:
                    f_out.write("\n\n".join(all_text_list))
                
        except Exception as e:
            print(f"AI Task Error: {e}")
            raise e

def convert_to_excel(files):
    border_style = Side(border_style="thin", color="000000")
    total_files = len(files)
    for i, pdf_path in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        wb = Workbook()
        wb.remove(wb.active)
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                for page_idx, page in enumerate(pdf.pages, 1):
                    update_file_progress(page_idx, total_pages, f"表データをExcelへ変換中... ( {page_idx} / {total_pages} ページ )")
                    tables = page.extract_tables()
                    if not tables: continue
                    ws = wb.create_sheet(f"Page_{page_idx}")
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
            if save_dir:
                wb.save(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_Excel.xlsx"))
        except Exception as e:
            print(f"Excel Error: {e}")

def convert_to_image(files, ext):
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        doc = fitz.open(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        total_pages = len(doc)
        for n, page in enumerate(doc, 1):
            update_file_progress(n, total_pages, f"画像へ変換中... ( {n} / {total_pages} ページ )")
            page.get_pixmap(dpi=200).save(os.path.join(save_dir, f"{base}_{n}.{ext}"))

def convert_to_dxf(files):
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            doc = fitz.open(f)
            dwg = ezdxf.new('R2010')
            msp = dwg.modelspace()
            save_dir = get_save_dir(f)
            if not save_dir: return

            total_pages = len(doc)
            for page_num, page in enumerate(doc, 1):
                update_file_progress(page_num, total_pages, f"DXFへ変換中... ( {page_num} / {total_pages} ページ )")
                h = page.rect.height
                paths = page.get_drawings()
                is_vector_rich = len(paths) > 0

                if is_vector_rich:
                    for path in paths:
                        for item in path["items"]:
                            if item[0] == "l":
                                msp.add_line((item[1].x, h - item[1].y), (item[2].x, h - item[2].y))
                            elif item[0] == "re":
                                rect = item[1]
                                pts = [(rect.x0, h - rect.y0), (rect.x1, h - rect.y0),
                                       (rect.x1, h - rect.y1), (rect.x0, h - rect.y1)]
                                msp.add_lwpolyline(pts, close=True)
                            elif item[0] == "c":
                                p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                                msp.add_spline([(p1.x, h - p1.y), (p2.x, h - p2.y),
                                                (p3.x, h - p3.y), (p4.x, h - p4.y)])
                
                if not is_vector_rich or len(paths) < 5:
                    pix = page.get_pixmap(dpi=300)
                    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY) if pix.n >= 3 else img
                    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
                    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
                    binary = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
                    contours, _ = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
                    scale_x = page.rect.width / pix.w
                    scale_y = page.rect.height / pix.h

                    for cnt in contours:
                        if cv2.contourArea(cnt) < 15: continue 
                        epsilon = 0.003 * cv2.arcLength(cnt, True)
                        pts = [(p[0][0] * scale_x, h - p[0][1] * scale_y) for p in cv2.approxPolyDP(cnt, epsilon, True)]
                        if len(pts) > 1:
                            msp.add_lwpolyline(pts, close=True)

            update_file_progress(total_pages, total_pages, "DXFファイルを保存中...")
            dwg.saveas(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(f))[0]}_CAD.dxf"))
        except Exception as e:
            print(f"DXF Conversion Error: {e}")

# ==============================
# UI構築
# ==============================
def update_ui():
    path_text = "\n".join(selected_files) if current_mode == "file" else (f"フォルダ: {selected_folder}" if selected_folder else "未選択")
    path_label.config(text=path_text)
    is_active = current_mode is not None
    
    btns = [btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_ai_extract]
    for b in btns: b.config(state=NORMAL if is_active else DISABLED, bg="#1E88E5" if is_active else LIGHT, fg="white" if is_active else INACTIVE)
    btn_merge.config(state=NORMAL if current_mode=="folder" else DISABLED, bg="#1E88E5" if current_mode=="folder" else LIGHT, fg="white" if current_mode=="folder" else INACTIVE)

def toggle_api_key_entry(*args):
    if ai_engine_var.get() == "Gemini":
        api_key_entry.config(state=NORMAL)
        btn_api_test.config(state=NORMAL)
    else:
        api_key_entry.config(state=DISABLED)
        btn_api_test.config(state=DISABLED)

root = Tk()
root.title(f"{APP_TITLE} {VERSION}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg=LIGHT)
root.resizable(False, False)

# 変数の初期化
rotate_option = IntVar(value=270)
save_option = IntVar(value=1)
ai_engine_var = StringVar(value="Gemini")
api_key_var = StringVar(value=get_api_key() or "")
output_format_var = StringVar(value="xlsx")

menubar = Menu(root)
help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="AI抽出の準備 (使い方)", command=show_ai_help)
help_menu.add_separator()
help_menu.add_command(label="Readmeを表示", command=show_readme)
help_menu.add_command(label="バージョン履歴", command=show_history)
help_menu.add_separator()
help_menu.add_command(label="バージョン情報", command=show_version_info)
menubar.add_cascade(label="ヘルプ", menu=help_menu)
root.config(menu=menubar)

title_frame = Frame(root, bg=LIGHT)
title_frame.pack(pady=(10, 2))
Label(title_frame, text=APP_TITLE, bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 16, "bold")).pack(side=LEFT)
Label(title_frame, text=f" {VERSION}", bg=LIGHT, fg=INACTIVE, font=("Segoe UI", 11)).pack(side=LEFT, pady=(5, 0))

info_text = "✨ Update: メニューバーをスッキリさせ、APIキー設定を画面内に統合しました！"
Label(root, text=info_text, bg=LIGHT, fg=INFO_TEXT, font=("Meiryo UI", 9)).pack(pady=(0, 2))

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

ai_frame = LabelFrame(root, text="AI抽出設定", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"), padx=5, pady=5)
ai_frame.pack(pady=5, fill="x", padx=10)

engine_frame = Frame(ai_frame, bg=LIGHT)
engine_frame.pack(anchor="w", fill="x", pady=2)
Label(engine_frame, text="エンジン:", bg=LIGHT, width=10, anchor="e").pack(side=LEFT, padx=5)
Radiobutton(engine_frame, text="Gemini API (推奨)", variable=ai_engine_var, value="Gemini", bg=LIGHT).pack(side=LEFT)
Radiobutton(engine_frame, text="Tesseract", variable=ai_engine_var, value="Tesseract", bg=LIGHT).pack(side=LEFT)

api_key_frame = Frame(ai_frame, bg=LIGHT)
api_key_frame.pack(anchor="w", fill="x", pady=2)
Label(api_key_frame, text="APIキー:", bg=LIGHT, width=10, anchor="e").pack(side=LEFT, padx=5)
api_key_entry = Entry(api_key_frame, textvariable=api_key_var, width=42, show="*")
api_key_entry.pack(side=LEFT)
btn_api_test = Button(api_key_frame, text="テスト", command=test_api_key_ui, width=6, bg="#E0E0E0")
btn_api_test.pack(side=LEFT, padx=5)

format_frame = Frame(ai_frame, bg=LIGHT)
format_frame.pack(anchor="w", fill="x", pady=2)
Label(format_frame, text="出力形式:", bg=LIGHT, width=10, anchor="e").pack(side=LEFT, padx=5)
Radiobutton(format_frame, text=".xlsx", variable=output_format_var, value="xlsx", bg=LIGHT).pack(side=LEFT)
Radiobutton(format_frame, text=".csv", variable=output_format_var, value="csv", bg=LIGHT).pack(side=LEFT)
Radiobutton(format_frame, text=".txt", variable=output_format_var, value="txt", bg=LIGHT).pack(side=LEFT)

ai_engine_var.trace("w", toggle_api_key_entry)

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
btn_ai_extract = Button(op_frame, text="AI抽出", width=12, command=run_ai_extraction)

op_list = [btn_merge, btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_ai_extract]
for i, b in enumerate(op_list):
    b.grid(row=i//4, column=i%4, padx=5, pady=3)

update_ui()
toggle_api_key_entry()
root.mainloop()