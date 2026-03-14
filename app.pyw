# -*- coding: utf-8 -*-
"""
PdfEditMiya v1.3.0
------------------
更新情報:
- v1.3.0: 手書き文字抽出に生成AI（Gemini API）を導入
- v1.3.0_fix2: 処理中画面を「全体進捗」と「個別進捗」の2段プログレスバーに改良
- v1.3.0_fix1: APIキーの保存先をユーザーディレクトリに変更し、GitHubへの流出を完全防止
- v1.2.1: メニューバー追加（バージョン履歴、Readme表示機能）
- v1.2.0: DXF変換精度の向上(曲線・スキャン対応)
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

import google.generativeai as genai

# ==============================
# リソースパス取得関数 (exe化対応)
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
VERSION = "v1.3.0"

WINDOW_WIDTH = 680
WINDOW_HEIGHT = 700

PRIMARY = "#1565C0"
LIGHT = "#E3F2FD"
SUCCESS = "#2E7D32"
ERROR = "#C62828"
INACTIVE = "#90A4AE"
INFO_TEXT = "#455A64"

# セキュリティ対策: APIキーの保存先をホームディレクトリの隠しファイルに指定
USER_HOME = os.path.expanduser("~")
API_KEY_FILE = os.path.join(USER_HOME, ".pdfeditmiya_api_key.txt")

VERSION_HISTORY = """
[ v1.3.0 ]
- 手書き文字抽出に生成AI（Gemini API）を搭載
- 処理中の進捗画面を「全体」と「個別(ページ)」の2段階表示に進化
- APIキーの保存先を変更し、GitHubへの流出を防止するセキュリティ対策を実施
- 404エラーを回避するため「利用可能なAIモデルの自動検索システム」を搭載
- 抽出結果を理想的なフォーマットの Excel (.xlsx) として出力

[ v1.2.1 ]
- メニューバーの実装、Readme表示機能追加

[ v1.2.0 ]
- DXF変換機能の強化
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
# APIキー管理 ＆ モデル自動検索機能
# ==============================

def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    return None

def get_available_models(api_key):
    """現在のAPIキーで利用可能なAIモデルのリストを動的に取得する"""
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

def set_api_key():
    key = askstring("Gemini APIキー設定", "Gemini APIキーを入力してください:\n（https://aistudio.google.com/app/apikey で取得したもの）", show='*')
    if key:
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
            messagebox.showinfo("認証成功！", "APIキーは正しく認識されました。\n設定を保存しました！")
            return key
        else:
            err_msg = last_err.lower()
            if "404" in err_msg or "not found" in err_msg:
                msg = ("APIキーは認識されましたが、AIモデルを利用する権限がありません（404エラー）。\n\n"
                       "【原因の可能性】\n"
                       "・Google AI Studio 以外(GCP等)で取得し、Gemini APIが有効になっていない。\n"
                       "・別のサービス用のAPIキーを入力している。\n\n"
                       "※Google AI Studio で新しくキーを作成し直すことをお勧めします。")
                messagebox.showerror("権限エラー", msg)
            else:
                messagebox.showerror("通信エラー", f"APIキーの確認中にエラーが発生しました。\n\n【詳細】\n{last_err}")
            return None
    return None

# ==============================
# 共通ロジック (UI安全化対応)
# ==============================

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
        if not folder:
            return
        preset_save_dir = folder
        save_label.config(text=preset_save_dir)

    show_processing(len(files))
    threading.Thread(target=run_task, args=(func,), daemon=True).start()

# ==============================
# UI補助機能 (ダブルプログレスバー対応)
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
    
    # 上段：全体の進捗
    overall_label = Label(processing_popup, text=f"全体の進捗 ( 0 / {total_files} ファイル )", bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 10, "bold"))
    overall_label.pack(pady=(15, 2))
    overall_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=total_files, length=340)
    overall_progress.pack(pady=(0, 10))
    
    # 下段：個別ファイルの進捗
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
            if max_val is not None:
                overall_progress["maximum"] = max_val
            overall_progress["value"] = step
            if text:
                overall_label.config(text=text)
            overall_progress.update()
    root.after(0, _task)

def update_file_progress(step, max_val=None, text=None):
    def _task():
        if processing_popup and processing_popup.winfo_exists():
            if max_val is not None:
                file_progress["maximum"] = max_val
            file_progress["value"] = step
            if text:
                file_label.config(text=text)
            file_progress.update()
    root.after(0, _task)

# ==============================
# メニューバー機能
# ==============================

def show_text_window(title, content):
    win = Toplevel(root)
    win.title(title)
    win.geometry("500x400")
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 250
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 200
    win.geometry(f"+{x}+{y}")
    text_area = st.ScrolledText(win, wrap=WORD, font=("Consolas", 10))
    text_area.pack(expand=True, fill=BOTH, padx=5, pady=5)
    text_area.insert(END, content)
    text_area.configure(state=DISABLED)

def show_version_info():
    msg = f"{APP_TITLE}\nバージョン: {VERSION}\n\nPython & Tkinter製 PDF編集ツール"
    messagebox.showinfo("バージョン情報", msg)

def show_history():
    show_text_window("バージョン履歴", VERSION_HISTORY.strip())

def show_readme():
    readme_path = resource_path("readme.md")
    content = ""
    if os.path.exists(readme_path):
        try:
            with open(readme_path, "r", encoding="utf-8") as f:
                content = f.read()
        except Exception as e:
            content = f"ファイルの読み込みに失敗しました:\n{e}"
    else:
        content = f"readme.md ファイルが見つかりませんでした。\n検索パス: {readme_path}"
    show_text_window("Readme", content)

# ==============================
# 保存・選択ロジック
# ==============================

def get_save_dir(original_path):
    global preset_save_dir, cancelled
    if save_option.get() == 1: return os.path.dirname(original_path)
    if preset_save_dir: return preset_save_dir
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

def run_extract_handwriting():
    api_key = get_api_key()
    if not api_key:
        api_key = set_api_key()
        if not api_key:
            return
    safe_run(extract_handwriting_task)

def extract_handwriting_task(files):
    api_key = get_api_key()
    genai.configure(api_key=api_key)
    models_to_try = get_available_models(api_key)

    prompt = """
    この画像は手書きの表や日報です。以下のルールに厳密に従って、カンマ区切りのCSV形式でデータ化してください。
    1. 表の項目（機械名, 図面名称, 空欄(日付), 整理番号, 備考など）を正確に読み取り、セルごとに分けてください。
    2. 表の最上部にあるタイトル（例：「佐渡 C振」など）は、CSVの1行目の右側の列（E列など）に配置してください。
    3. 左側の列（例：「No1送り改造」「No2修理部組立説明図」など）の下が空欄になっている場合、文脈を判断し、上の行の値を自動的に引き継いで空白を埋めてください。
    4. 日付（3/12 など）は '2026-03-12' のように2026年の形式（YYYY-MM-DD）に自動変換して、日付用の列に出力してください。
    5. 記号（「No.」を「No」にするなど）は統一してください。
    6. 余計な挨拶や説明文、またマークダウンのコードブロック記号(```csv等)は一切出力せず、純粋なCSVのテキストデータのみを出力してください。
    """

    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            save_dir = get_save_dir(f)
            if not save_dir: return
            
            base = os.path.splitext(os.path.basename(f))[0]
            wb = Workbook()
            wb.remove(wb.active)
            
            doc = fitz.open(f)
            total_pages = len(doc)
            
            for page_num in range(total_pages):
                update_file_progress(page_num + 1, total_pages, f"AIが手書き文字を解析中... ( {page_num+1} / {total_pages} ページ ) [モデル検索中]")
                
                ws = wb.create_sheet(f"Page_{page_num+1}")
                page = doc[page_num]
                pix = page.get_pixmap(dpi=150)
                mode = "RGB" if pix.n == 3 else "L"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")
                
                max_retries = 3
                csv_text = ""
                success = False
                last_error = ""
                used_model = ""

                for attempt in range(max_retries):
                    for model_name in models_to_try:
                        update_file_progress(page_num + 1, text=f"AIが手書き文字を解析中... ( {page_num+1} / {total_pages} ページ ) [使用: {model_name}]")
                        try:
                            model = genai.GenerativeModel(model_name)
                            response = model.generate_content([prompt, img])
                            if not response.parts:
                                raise Exception("AIが安全フィルタ等により回答をブロックしました。")
                            csv_text = response.text.strip()
                            success = True
                            used_model = model_name
                            break
                        except Exception as api_err:
                            last_error = str(api_err)
                            continue
                            
                    if success:
                        break
                    time.sleep(2 ** attempt)

                if success:
                    update_file_progress(page_num + 1, text=f"解析成功！ ( {page_num+1} / {total_pages} ページ ) [モデル: {used_model}]")
                    markdown_marker = "`" * 3
                    if csv_text.startswith(markdown_marker):
                        lines = csv_text.split('\n')
                        if len(lines) > 0 and lines[0].startswith(markdown_marker): 
                            lines = lines[1:]
                        if len(lines) > 0 and lines[-1].startswith(markdown_marker): 
                            lines = lines[:-1]
                        csv_text = '\n'.join(lines)
                    
                    reader = csv.reader(io.StringIO(csv_text))
                    for row_idx, row in enumerate(reader, 1):
                        for col_idx, val in enumerate(row, 1):
                            ws.cell(row=row_idx, column=col_idx, value=val.strip())
                else:
                    update_file_progress(page_num + 1, text=f"解析失敗 ( {page_num+1} / {total_pages} ページ )")
                    print(f"Page {page_num+1} AI Error: {last_error}")
                    error_msg = (f"[ --- ページ {page_num+1} の解析に失敗しました --- ]\n"
                                 f"エラー詳細: {last_error}\n"
                                 "【原因と対策】\n"
                                 "1. Google AI StudioでAPIキーを取得しましたか？ (GCPの場合はGeminiの有効化が必要です)\n"
                                 "2. APIキーが間違っていないか、別のサービスのキーでないか確認してください。\n"
                                 "3. もう一度「Gemini APIキー設定」から正しいキーを登録し直してください。")
                    ws.cell(row=1, column=1, value=error_msg)
                
                gc.collect()
            
            doc.close()
            update_file_progress(total_pages, total_pages, "Excelファイルを保存中...")
            out_path = os.path.join(save_dir, f"{base}_抽出後.xlsx")
            wb.save(out_path)
                
        except Exception as e:
            print(f"Handwriting Task Error: {e}")
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
    
    btns = [btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_handwriting]
    for b in btns: b.config(state=NORMAL if is_active else DISABLED, bg="#1E88E5" if is_active else LIGHT, fg="white" if is_active else INACTIVE)
    btn_merge.config(state=NORMAL if current_mode=="folder" else DISABLED, bg="#1E88E5" if current_mode=="folder" else LIGHT, fg="white" if current_mode=="folder" else INACTIVE)

root = Tk()
root.title(f"{APP_TITLE} {VERSION}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
root.configure(bg=LIGHT)
root.resizable(False, False)

menubar = Menu(root)
setting_menu = Menu(menubar, tearoff=0)
setting_menu.add_command(label="Gemini APIキー設定", command=set_api_key)
menubar.add_cascade(label="設定", menu=setting_menu)

help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="Readmeを表示", command=show_readme)
help_menu.add_command(label="バージョン履歴", command=show_history)
help_menu.add_separator()
help_menu.add_command(label="バージョン情報", command=show_version_info)
menubar.add_cascade(label="ヘルプ", menu=help_menu)

root.config(menu=menubar)

rotate_option, save_option = IntVar(value=270), IntVar(value=1)

title_frame = Frame(root, bg=LIGHT)
title_frame.pack(pady=(10, 2))
Label(title_frame, text=APP_TITLE, bg=LIGHT, fg=PRIMARY, font=("Segoe UI", 16, "bold")).pack(side=LEFT)
Label(title_frame, text=f" {VERSION}", bg=LIGHT, fg=INACTIVE, font=("Segoe UI", 11)).pack(side=LEFT, pady=(5, 0))

info_text = "✨ Update: 進捗画面を「全体」と「個別ページ」の2段表示に改良しました"
Label(root, text=info_text, bg=LIGHT, fg=INFO_TEXT, font=("Meiryo UI", 9)).pack(pady=(0, 2))

note_text = (
    "【手書き抽出の使い方】\n"
    "事前にメニューの「設定」>「Gemini APIキー設定」からAPIキーを登録してください。\n"
    "AIが表の結合（空欄補完）や日付変換を自動で処理し、理想的なExcelを出力します。"
)
Label(root, text=note_text, bg=LIGHT, fg=ERROR, font=("Meiryo UI", 8), justify="left").pack(pady=(0, 8))

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
btn_handwriting = Button(op_frame, text="手書き抽出", width=12, command=run_extract_handwriting)

op_list = [btn_merge, btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_handwriting]
for i, b in enumerate(op_list):
    b.grid(row=i//4, column=i%4, padx=5, pady=3)

update_ui()
root.mainloop()