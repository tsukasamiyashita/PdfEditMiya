# -*- coding: utf-8 -*-
"""
PdfEditMiya v1.10.0
------------------
更新情報:
- v1.10.0: ウィンドウの初期表示位置を画面左上に固定。その他、集約ロジックやAI抽出機能は最新状態を保持。
"""

import os
import sys
import threading
import io
import gc
import cv2
import csv
import time
import json
import ast
import difflib
import re
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter.scrolledtext as st
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
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
# 基本設定 & 洗練されたカラーパレット
# ==============================
APP_TITLE = "PdfEditMiya"
VERSION = "v1.10.0"
WINDOW_WIDTH = 700
WINDOW_HEIGHT = 890

BG_COLOR = "#F0F4F8"          
CARD_BG = "#FFFFFF"           
PRIMARY = "#0D6EFD"           
PRIMARY_HOVER = "#0B5ED7"     
TEXT_COLOR = "#212529"        
MUTED_TEXT = "#6C757D"        
BORDER_COLOR = "#DEE2E6"      
SUCCESS = "#198754"           
ERROR = "#DC3545"             

USER_HOME = os.path.expanduser("~")
API_KEY_FILE = os.path.join(USER_HOME, ".pdfeditmiya_api_key.txt")

VERSION_HISTORY = """
[ v1.10.0 ]
- 【レイアウト固定】アプリ起動時のウィンドウ表示位置を画面左上に固定しました。
- 【引継ぎ処理の強化】「〃」などの繰り返し記号がある場合も、上の行のテキストを引き継ぐように改善。
- 【引継ぎの例外設定】「備考」の列については自動引継ぎを行わないようルールを追加。
- 【列ズレの根本的解決】1つのセル内にタプル文字列としてデータが合体してしまうバグを修正。astモジュールで強制的に複数列へ展開。
- 【スマートカラムマッピング】テキストでの判定を廃止し、「数値・分数の割合と桁数の整合性」のみに基づいて列を自動判別・整列させる仕様。
- 【列幅の自動調整】作成されたExcelデータの列幅が、セル内の文字数や桁数に合わせて自動的に広がるようになりました。
"""

AI_HELP_TEXT = """
【 AI抽出機能を使うための準備 】

■ Gemini API を使う場合（推奨・超高精度）
1. 以下のURLにアクセスし、Googleアカウントでログインします。
   https://aistudio.google.com/app/apikey
2. 「Create API key」ボタンを押し、新しいプロジェクトでAPIキーを作成します。
3. 発行された文字列をコピーし、本アプリの「APIキー」枠に貼り付け「テスト」ボタンを押してください。

■ Tesseract を使う場合（オフライン・簡易抽出）
Windows環境にTesseract OCR本体がインストールされている必要があります。
1. https://github.com/UB-Mannheim/tesseract/wiki からダウンロード・インストール。
2. インストール時、「Additional language data (download)」から「Japanese」を選択してください。
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
# Excel列幅の自動調整機能
# ==============================
def auto_adjust_excel_column_width(ws):
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    lines = str(cell.value).split('\n')
                    for line in lines:
                        length = sum(2 if ord(c) > 255 else 1 for c in line)
                        if length > max_length:
                            max_length = length
            except:
                pass
        adjusted_width = (max_length + 2)
        if adjusted_width > 60:
            adjusted_width = 60
        ws.column_dimensions[column_letter].width = adjusted_width

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
        'gemini-1.5-pro', 'gemini-1.5-pro-latest', 
        'gemini-2.5-pro', 'gemini-2.0-pro',
        'gemini-1.5-flash', 'gemini-2.5-flash'
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
            messagebox.showerror("権限エラー", "APIキーは認識されましたが、AIモデルを利用権限がありません。")
        else:
            messagebox.showerror("通信エラー", f"APIキー確認中にエラーが発生しました。\n{last_err}")

# ==============================
# 共通ロジック
# ==============================
def get_target_files():
    if current_mode == "file": return selected_files
    if current_mode == "folder" and selected_folder:
        return [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".csv", ".txt"))]
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
    files = filedialog.askopenfilenames(filetypes=[("すべての対応ファイル", "*.pdf;*.xlsx;*.csv;*.txt"), ("PDF", "*.pdf"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Text", "*.txt")])
    if files:
        selected_files, selected_folder, current_mode = list(files), "", "file"
        update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory(title="フォルダを選択")
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
            show_message("✅ 処理が完了しました", SUCCESS)
    except Exception as e:
        print(f"Error: {e}")
        close_processing()
        show_message(f"❌ エラーが発生しました\n{str(e)[:30]}...", ERROR)

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
# UI補助・プログレス機能
# ==============================
def show_message(msg, color=PRIMARY):
    def _task():
        win = tk.Toplevel(root)
        win.geometry("260x90")
        win.configure(bg=CARD_BG)
        win.attributes("-topmost", True)
        x = root.winfo_x() + (WINDOW_WIDTH // 2) - 130
        y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 45
        win.geometry(f"+{x}+{y}")
        win.overrideredirect(True)
        
        frame = tk.Frame(win, bg=CARD_BG, highlightbackground=color, highlightthickness=2)
        frame.pack(expand=True, fill=tk.BOTH)
        ttk.Label(frame, text=msg, foreground=color, font=("Segoe UI", 10, "bold"), background=CARD_BG, wraplength=240).pack(expand=True)
        win.after(2500, win.destroy)
    root.after(0, _task)

def show_processing(total_files=1):
    global processing_popup, overall_label, overall_progress, file_label, file_progress
    processing_popup = tk.Toplevel(root)
    processing_popup.title("処理を実行中...")
    processing_popup.geometry("440x210")
    processing_popup.configure(bg=CARD_BG)
    processing_popup.grab_set()
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 220
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 105
    processing_popup.geometry(f"+{x}+{y}")
    
    overall_label = ttk.Label(processing_popup, text=f"全体の進捗 ( 0 / {total_files} ファイル )", font=("Segoe UI", 10, "bold"), background=CARD_BG, foreground=PRIMARY)
    overall_label.pack(pady=(25, 5))
    overall_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=total_files, length=380)
    overall_progress.pack(pady=(0, 20))
    
    file_label = ttk.Label(processing_popup, text="現在のファイルを準備中...", font=("Segoe UI", 9), background=CARD_BG, foreground=MUTED_TEXT)
    file_label.pack(pady=(5, 5))
    file_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=1, length=380)
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
    root.after(0, _task)

def set_file_progress_indeterminate(text=None):
    def _task():
        if processing_popup and processing_popup.winfo_exists():
            file_progress.config(mode="indeterminate")
            file_progress.start(15)
            if text: file_label.config(text=text)
    root.after(0, _task)

def set_file_progress_determinate(step, max_val=None, text=None):
    def _task():
        if processing_popup and processing_popup.winfo_exists():
            file_progress.stop()
            file_progress.config(mode="determinate")
            if max_val is not None: file_progress["maximum"] = max_val
            file_progress["value"] = step
            if text: file_label.config(text=text)
    root.after(0, _task)

# ==============================
# メニューバー機能
# ==============================
def show_text_window(title, content):
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("620x500")
    win.configure(bg=BG_COLOR)
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 310
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 250
    win.geometry(f"+{x}+{y}")
    text_area = st.ScrolledText(win, wrap=tk.WORD, font=("Meiryo UI", 10), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.FLAT, padx=15, pady=15)
    text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
    text_area.insert(tk.END, content)
    text_area.configure(state=tk.DISABLED)

def show_ai_help(): show_text_window("AI抽出の準備 (使い方)", AI_HELP_TEXT.strip())
def show_version_info(): messagebox.showinfo("バージョン情報", f"{APP_TITLE}\nバージョン: {VERSION}\n\nPython & Tkinter製 PDF編集ツール")
def show_history(): show_text_window("バージョン履歴", VERSION_HISTORY.strip())
def show_readme():
    p = resource_path("README.md")
    content = open(p, "r", encoding="utf-8").read() if os.path.exists(p) else "READMEが見つかりません。"
    show_text_window("Readme", content)

# ==============================
# PDF操作コア機能
# ==============================
def merge_pdfs(files):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    total_files = len(files)
    writer = PdfWriter()
    update_overall_progress(0, total_files, f"全体の進捗 ( 0 / {total_files} ファイル )")
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        total_pages = len(reader.pages)
        for j, p in enumerate(reader.pages, 1):
            set_file_progress_determinate(j, total_pages, f"ファイルを結合中... ( {j} / {total_pages} ページ )")
            writer.add_page(p)
    set_file_progress_determinate(1, 1, "PDFを保存中...")
    save_dir = get_save_dir(files[0])
    if save_dir:
        name = os.path.basename(selected_folder) if selected_folder else "Merged"
        with open(os.path.join(save_dir, f"{name}_Merge.pdf"), "wb") as out:
            writer.write(out)

def split_pdfs(files):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        total_pages = len(reader.pages)
        digits = max(2, len(str(total_pages)))
        for n, p in enumerate(reader.pages, 1):
            set_file_progress_determinate(n, total_pages, f"ファイルを分割中... ( {n} / {total_pages} ページ )")
            writer = PdfWriter()
            writer.add_page(p)
            n_str = str(n).zfill(digits)
            with open(os.path.join(save_dir, f"{base}_Split_{n_str}.pdf"), "wb") as out:
                writer.write(out)

def rotate_pdfs(files):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    deg = rotate_option.get()
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        for j, p in enumerate(reader.pages, 1):
            set_file_progress_determinate(j, total_pages, f"ページを回転中... ( {j} / {total_pages} ページ )")
            p.rotate(deg)
            writer.add_page(p)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Rotate.pdf"), "wb") as out:
            writer.write(out)

def extract_text(files):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        reader = PdfReader(f)
        total_pages = len(reader.pages)
        text_list = []
        for j, p in enumerate(reader.pages, 1):
            set_file_progress_determinate(j, total_pages, f"テキストを抽出中... ( {j} / {total_pages} ページ )")
            text_list.append(p.extract_text() or "")
        text = "".join(text_list)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        with open(os.path.join(save_dir, f"{base}_Text.txt"), "w", encoding="utf-8") as out:
            out.write(text)

# ==============================
# プロファイリング・スマートマッピング機能 (数値整合性特化)
# ==============================
def analyze_column_profile(col_data):
    if not col_data: 
        return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
    
    pure_num_cnt = 0
    fraction_cnt = 0
    total_num_len = 0
    num_cell_cnt = 0
    total_cells = 0
    
    for val in col_data:
        s = str(val).strip()
        if not s or s == "None": continue
        total_cells += 1
        
        if re.match(r'^\d+/\d+$', s):
            fraction_cnt += 1
            num_cell_cnt += 1
            total_num_len += len(s)
        else:
            s_clean = s.replace(",", "").replace(".", "", 1).replace("-", "", 1)
            if s_clean.isdigit():
                pure_num_cnt += 1
                num_cell_cnt += 1
                total_num_len += len(s_clean)
                
    if total_cells == 0:
        return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
        
    pure_num_ratio = pure_num_cnt / total_cells
    fraction_ratio = fraction_cnt / total_cells
    is_text = (pure_num_ratio < 0.2 and fraction_ratio < 0.2)
    
    return {
        "pure_num_ratio": pure_num_ratio,
        "fraction_ratio": fraction_ratio,
        "avg_num_len": total_num_len / num_cell_cnt if num_cell_cnt > 0 else 0.0,
        "is_text": is_text
    }

def get_profile_similarity(p1, p2):
    diff_pure = abs(p1["pure_num_ratio"] - p2["pure_num_ratio"])
    diff_frac = abs(p1["fraction_ratio"] - p2["fraction_ratio"])
    
    max_len = max(p1["avg_num_len"], p2["avg_num_len"])
    diff_len = abs(p1["avg_num_len"] - p2["avg_num_len"]) / max_len if max_len > 0 else 0.0
    
    sim = 1.0 - (diff_pure * 0.4 + diff_frac * 0.4 + diff_len * 0.2)
    return max(0.0, sim)

def parse_row_data(row_data):
    """
    タプル化された文字列などを強制的に配列に展開する処理
    "(None, '図番', ...)" などのエラー形式を救済する。
    """
    if isinstance(row_data, list) and len(row_data) == 1:
        row_data = row_data[0]

    if isinstance(row_data, str):
        row_data = row_data.strip()
        if (row_data.startswith('(') and row_data.endswith(')')) or (row_data.startswith('[') and row_data.endswith(']')):
            try:
                parsed = ast.literal_eval(row_data)
                if isinstance(parsed, (list, tuple)):
                    return [str(x) if x is not None else "" for x in parsed]
            except:
                return [x.strip().strip("'\"") for x in row_data.strip("()[]").split(",")]
        return [row_data]
        
    if isinstance(row_data, tuple):
        return [str(x) if x is not None else "" for x in row_data]
        
    if not isinstance(row_data, list):
        return [str(row_data)]
    
    return [str(x) if x is not None else "" for x in row_data]

# ==============================
# 空白セルの自動引継ぎ処理
# ==============================
def apply_text_inheritance(final_aggregated_data):
    if len(final_aggregated_data) <= 1:
        return
        
    def is_text_to_inherit(text):
        if not text: return False
        s = str(text).strip()
        if s in ["〃", "”", "\"", "''", "””", "''", "同上"]:
            return False
        if re.search(r'[a-zA-Zａ-ｚＡ-Ｚぁ-んァ-ン一-龥]', s):
            return True
        return False

    header = final_aggregated_data[0]
    
    skip_cols = set()
    for idx, h in enumerate(header):
        if "備考" in str(h):
            skip_cols.add(idx)

    for col_idx in range(1, len(header)):
        if col_idx in skip_cols:
            continue
            
        last_text = ""
        for row_idx in range(1, len(final_aggregated_data)):
            cell_val = str(final_aggregated_data[row_idx][col_idx]).strip()
            
            if cell_val == "" or cell_val == "None" or cell_val in ["〃", "”", "\"", "''", "””", "''", "同上"]:
                if last_text:
                    final_aggregated_data[row_idx][col_idx] = last_text
            else:
                if is_text_to_inherit(cell_val):
                    last_text = cell_val
                else:
                    last_text = ""

# ==============================
# データ集約のみタスク
# ==============================
def aggregate_only_task(files):
    out_format = output_format_var.get()
    target_files = []
    
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        if ext == ".pdf":
            base = os.path.splitext(f)[0]
            cand_ai = f"{base}_AI抽出.{out_format}"
            cand_tess = f"{base}_Tesseract抽出.{out_format}"
            if os.path.exists(cand_ai): target_files.append(cand_ai)
            elif os.path.exists(cand_tess): target_files.append(cand_tess)
        elif ext == f".{out_format}":
            target_files.append(f)
            
    if not target_files:
        raise Exception(f"指定された出力形式 (.{out_format}) のデータが見つかりません。\n事前に抽出処理を行ってください。")

    total_files = len(target_files)
    aggregated_master_header = ["元ファイル名"]
    aggregated_master_rows = []
    master_profiles = {}
    aggregated_all_texts = []
    
    def map_to_master(fname, curr_header, curr_rows):
        if isinstance(curr_header, list) and len(curr_header) == 1:
            curr_header = parse_row_data(curr_header)
        
        safe_header = []
        for i, h in enumerate(curr_header):
            h_str = str(h).strip() if h is not None else ""
            if not h_str: h_str = f"列{i+1}"
            safe_header.append(h_str)
            
        col_count = len(safe_header)
        col_data_list = [[] for _ in range(col_count)]
        parsed_rows = []
        
        for r in curr_rows:
            r_list = parse_row_data(r)
            parsed_rows.append(r_list)
            for i, val in enumerate(r_list):
                if i < col_count: col_data_list[i].append(val)
                
        curr_profiles = [analyze_column_profile(col_data_list[i]) for i in range(col_count)]
            
        col_mapping = {}
        mapped_master_indices = set()
        
        for i, h in enumerate(safe_header):
            best_match_idx = -1
            best_score = -1
            
            for m_idx, m_h in enumerate(aggregated_master_header):
                if m_idx == 0: continue
                if m_idx in mapped_master_indices: continue
                
                if m_idx in master_profiles:
                    p_curr = curr_profiles[i]
                    p_master = master_profiles[m_idx]
                    
                    if p_curr["is_text"] and p_master["is_text"]:
                        # テキスト同士は元の列のインデックス位置の近さを最優先
                        if i == (m_idx - 1):
                            total_score = 1.0
                        else:
                            total_score = 0.5 - abs(i - (m_idx - 1)) * 0.1
                    else:
                        # 数値・分数データは、プロファイル（整合性・桁数）を絶対優先
                        p_score = get_profile_similarity(p_curr, p_master)
                        total_score = p_score - abs(i - (m_idx - 1)) * 0.05
                    
                    if total_score > best_score and total_score > 0.4:
                        best_score = total_score
                        best_match_idx = m_idx
            
            if best_match_idx != -1:
                col_mapping[i] = best_match_idx
                mapped_master_indices.add(best_match_idx)
                
                if best_match_idx in master_profiles:
                    old_p = master_profiles[best_match_idx]
                    new_p = curr_profiles[i]
                    master_profiles[best_match_idx] = {
                        "pure_num_ratio": (old_p["pure_num_ratio"] + new_p["pure_num_ratio"]) / 2,
                        "fraction_ratio": (old_p["fraction_ratio"] + new_p["fraction_ratio"]) / 2,
                        "avg_num_len": (old_p["avg_num_len"] + new_p["avg_num_len"]) / 2,
                        "is_text": old_p["is_text"] and new_p["is_text"]
                    }
            else:
                aggregated_master_header.append(h)
                new_idx = len(aggregated_master_header) - 1
                col_mapping[i] = new_idx
                mapped_master_indices.add(new_idx)
                master_profiles[new_idx] = curr_profiles[i]
                    
        for r in parsed_rows:
            aligned_row = [""] * len(aggregated_master_header)
            aligned_row[0] = fname
            for i, val in enumerate(r):
                if i in col_mapping:
                    m_idx = col_mapping[i]
                    if m_idx >= len(aligned_row):
                        aligned_row.extend([""] * (m_idx - len(aligned_row) + 1))
                    val_str = str(val).strip() if val is not None else ""
                    if val_str == "None": val_str = ""
                    aligned_row[m_idx] = val_str
            
            # 元ファイル名以外のセルがすべて空文字なら、集約データに追加しない
            if any(v != "" for v in aligned_row[1:]):
                aggregated_master_rows.append(aligned_row)

    for i, f in enumerate(target_files, 1):
        update_overall_progress(i, total_files, f"データを集約中... ( {i} / {total_files} ファイル )")
        set_file_progress_determinate(50, 100, f"読み込み中: {os.path.basename(f)}")
        
        filename = os.path.basename(f)
        for suffix in [f"_AI抽出.{out_format}", f"_Tesseract抽出.{out_format}"]:
            if filename.endswith(suffix):
                filename = filename.replace(suffix, ".pdf")
                break
                
        try:
            if out_format == "xlsx":
                import openpyxl
                wb = openpyxl.load_workbook(f, data_only=True)
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows: continue
                    map_to_master(filename, rows[0], rows[1:])
                wb.close()
                
            elif out_format == "csv":
                with open(f, "r", encoding="utf-8-sig") as f_in:
                    reader = csv.reader(f_in)
                    rows = list(reader)
                    if not rows: continue
                    map_to_master(filename, rows[0], rows[1:])
                        
            elif out_format == "txt":
                with open(f, "r", encoding="utf-8") as f_in:
                    text = f_in.read()
                    aggregated_all_texts.append(f"[{filename}]\n{text}")
                    
        except Exception as e:
            print(f"Error reading {f}: {e}")
            
        set_file_progress_determinate(100, 100, "完了")
        time.sleep(0.05)
        
    if total_files > 0:
        save_dir = get_save_dir(target_files[0])
        if save_dir:
            set_file_progress_indeterminate("集約データを保存中...")
            
            if out_format in ["xlsx", "csv"]:
                final_aggregated_data = [aggregated_master_header]
                for row_data in aggregated_master_rows:
                    padded_row = (row_data + [""] * len(aggregated_master_header))[:len(aggregated_master_header)]
                    final_aggregated_data.append(padded_row)

                apply_text_inheritance(final_aggregated_data)

                if out_format == "xlsx" and len(final_aggregated_data) > 1:
                    wb_agg = Workbook()
                    ws_agg = wb_agg.active
                    ws_agg.title = "集約データ"
                    for r_idx, row_data in enumerate(final_aggregated_data, 1):
                        for c_idx, val in enumerate(row_data, 1):
                            ws_agg.cell(row=r_idx, column=c_idx, value=str(val).strip())
                    auto_adjust_excel_column_width(ws_agg)
                    wb_agg.save(os.path.join(save_dir, "データ集約.xlsx"))
                    
                elif out_format == "csv" and len(final_aggregated_data) > 1:
                    with open(os.path.join(save_dir, "データ集約.csv"), "w", encoding="utf-8-sig", newline="") as f_out:
                        writer = csv.writer(f_out)
                        writer.writerows(final_aggregated_data)
                        
            elif out_format == "txt" and aggregated_all_texts:
                with open(os.path.join(save_dir, "データ集約.txt"), "w", encoding="utf-8") as f_out:
                    f_out.write("\n\n".join(aggregated_all_texts))

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
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    out_format = output_format_var.get()
    total_files = len(files)
    aggregated_all_rows = []
    aggregated_all_texts = []
    
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            save_dir = get_save_dir(f)
            if not save_dir: return
            base = os.path.splitext(os.path.basename(f))[0]
            filename = os.path.basename(f)
            doc = fitz.open(f)
            total_pages = len(doc)
            digits = max(2, len(str(total_pages)))
            
            wb = Workbook() if out_format == "xlsx" else None
            if wb: wb.remove(wb.active)
            all_text_list = []
            
            for page_num in range(total_pages):
                set_file_progress_indeterminate(f"Tesseractで解析中... ( {page_num+1} / {total_pages} ページ )")
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                mode = "RGB" if pix.n == 3 else "L"
                img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")
                
                try:
                    text = pytesseract.image_to_string(img, lang="jpn+eng")
                    if out_format == "xlsx":
                        page_str = str(page_num+1).zfill(digits)
                        ws = wb.create_sheet(f"Page_{page_str}")
                        for row_idx, line in enumerate(text.split('\n'), 1):
                            ws.cell(row=row_idx, column=1, value=line.strip())
                            aggregated_all_rows.append([filename, line.strip()])
                        auto_adjust_excel_column_width(ws)
                    else:
                        all_text_list.append(text)
                        aggregated_all_texts.append(f"[{filename}] {text}")
                except Exception as e:
                    err_msg = f"[ --- ページ {page_num+1} 解析失敗 --- ]\nTesseract OCRエラー\n詳細: {e}"
                    if out_format == "xlsx":
                        page_str = str(page_num+1).zfill(digits)
                        ws = wb.create_sheet(f"Page_{page_str}")
                        ws.cell(row=1, column=1, value=err_msg)
                        auto_adjust_excel_column_width(ws)
                    else:
                        all_text_list.append(err_msg)
            
            doc.close()
            gc.collect()
            set_file_progress_determinate(total_pages, total_pages, "ファイルを保存中...")
            
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
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    key = api_key_var.get().strip()
    genai.configure(api_key=key)
    models_to_try = get_available_models(key)
    out_format = output_format_var.get()

    if out_format in ["csv", "xlsx"]:
        prompt = """
        あなたは優秀なデータ入力オペレーターです。添付された図面管理台帳などの表画像を読み取り、正確なJSONデータを作成してください。
        
        【データの分離精度を最優先する特別ルール（超重要）】
        - データの見た目や文脈の意味よりも、「表の縦の罫線」や「文字列間の大きな空白」などの【物理的な列の区切り】を絶対的な基準として最優先し、物理的に分かれているデータは必ず別の要素として分割してください。
        - 意味的に繋がっているように見えても、罫線や空白で区切られていれば絶対に1つの要素に結合しないでください。
        - データが存在しない「空白セル」の場合は無視せず、必ず `""` (空文字) を該当位置に挿入し、列の位置を厳密に揃えてください。
        - 1行のデータを丸ごと1つの文字列に繋げて出力することは絶対に禁止します。必ず各セルごとにリスト内で分割してください。

        【出力形式（絶対厳守）】
        シンプルな配列（リスト）で出力してください。
        {
          "header": ["列1名", "列2名", "列3名", ...],
          "rows": [
            ["データ1", "データ2", "", "データ4", ...],
            ["データ1", "データ2", "データ3", "", ...]
          ]
        }
        """
        generation_config = {"response_mime_type": "application/json"}
    else:
        prompt = "この画像に記載されている手書きの文字や文章を可能な限り正確に読み取り、プレーンテキストとして出力してください。"
        generation_config = None

    total_files = len(files)
    
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        try:
            save_dir = get_save_dir(f)
            if not save_dir: return
            
            base = os.path.splitext(os.path.basename(f))[0]
            doc = fitz.open(f)
            total_pages = len(doc)
            digits = max(2, len(str(total_pages)))
            
            wb = Workbook() if out_format == "xlsx" else None
            if wb: wb.remove(wb.active)
            all_text_list = []
            
            for page_num in range(total_pages):
                set_file_progress_determinate(0, 100, f"AIが画像を補正中... ( {page_num+1} / {total_pages} ページ )")
                ws = wb.create_sheet(f"Page_{str(page_num+1).zfill(digits)}") if wb else None
                
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                
                if pix.n == 4: img_array = cv2.cvtColor(img_array, cv2.COLOR_RGBA2RGB)
                elif pix.n == 1: img_array = cv2.cvtColor(img_array, cv2.COLOR_GRAY2RGB)

                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
                blurred = cv2.GaussianBlur(gray, (3, 3), 0)
                binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 5)
                clean_bg = np.where(binary == 255, 255, gray)
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                enhanced = clahe.apply(clean_bg.astype(np.uint8))
                blur_for_sharp = cv2.GaussianBlur(enhanced, (0, 0), 2)
                sharp = cv2.addWeighted(enhanced, 1.5, blur_for_sharp, -0.5, 0)
                img = Image.fromarray(cv2.cvtColor(sharp, cv2.COLOR_GRAY2RGB))
                
                max_retries = 3
                extracted_text, success, last_error, used_model = "", False, "", ""

                for attempt in range(max_retries):
                    for model_name in models_to_try:
                        set_file_progress_indeterminate(f"AIが物理的な列構造を解析中... ( {page_num+1} / {total_pages} ページ )")
                        try:
                            model = genai.GenerativeModel(model_name)
                            if generation_config:
                                response = model.generate_content([prompt, img], generation_config=generation_config)
                            else:
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
                    set_file_progress_determinate(100, 100, f"解析成功！ [モデル: {used_model}]")
                    
                    if out_format in ["xlsx", "csv"]:
                        try:
                            data = json.loads(extracted_text)
                            
                            header = data.get("header", [])
                            rows = data.get("rows", [])
                            
                            if not header and not rows and isinstance(data, list):
                                if data:
                                    header = data[0] if isinstance(data[0], list) else [str(data[0])]
                                    rows = data[1:]
                            
                            safe_header = [str(x).strip() for x in header] if isinstance(header, list) else []
                            
                            page_col_count = len(safe_header)
                            for r in rows:
                                if isinstance(r, list) and len(r) > page_col_count:
                                    page_col_count = len(r)
                                    
                            if not safe_header:
                                safe_header = [f"列{idx+1}" for idx in range(page_col_count)]

                            page_data_to_write = []
                            padded_header = (safe_header + [""] * page_col_count)[:page_col_count]
                            page_data_to_write.append(padded_header)
                            
                            for row_data in rows:
                                parsed_r = parse_row_data(row_data)
                                safe_row_local = (parsed_r + [""] * page_col_count)[:page_col_count]
                                if any(v != "" for v in safe_row_local):
                                    page_data_to_write.append(safe_row_local)

                            if out_format == "xlsx":
                                for row_idx, r_data in enumerate(page_data_to_write, 1):
                                    for col_idx, val in enumerate(r_data, 1):
                                        val_str = val if val != "None" else ""
                                        ws.cell(row=row_idx, column=col_idx, value=val_str)
                                auto_adjust_excel_column_width(ws)
                            else: # csv
                                csv_lines = []
                                for r_data in page_data_to_write:
                                    clean_row = [v if v != "None" else "" for v in r_data]
                                    # Python 3.11以下の構文エラーを防ぐため .format() を使用して修正
                                    csv_lines.append(",".join('"{}"'.format(val.replace('"', '""')) if ',' in val or '"' in val else val for val in clean_row))
                                all_text_list.append("\n".join(csv_lines))

                        except json.JSONDecodeError as e:
                            err_msg = f"[ --- ページ {page_num+1} の解析結果が不正です --- ]\n詳細: {e}\n出力:\n{extracted_text}"
                            if out_format == "xlsx":
                                ws.cell(row=1, column=1, value=err_msg)
                                auto_adjust_excel_column_width(ws)
                            else:
                                all_text_list.append(err_msg)
                    else: # txt
                        all_text_list.append(extracted_text)
                else:
                    err_msg = f"[ --- ページ {page_num+1} の解析に失敗しました --- ]\nエラー詳細: {last_error}"
                    if out_format == "xlsx":
                        ws.cell(row=1, column=1, value=err_msg)
                        auto_adjust_excel_column_width(ws)
                    else:
                        all_text_list.append(err_msg)
            
            doc.close()
            gc.collect()
            set_file_progress_determinate(total_pages, total_pages, "ファイルを保存中...")
            
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
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    border_style = Side(border_style="thin", color="000000")
    total_files = len(files)
    for i, pdf_path in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        wb = Workbook()
        wb.remove(wb.active)
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                digits = max(2, len(str(total_pages)))
                for page_idx, page in enumerate(pdf.pages, 1):
                    set_file_progress_determinate(page_idx, total_pages, f"表データをExcelへ変換中... ( {page_idx} / {total_pages} ページ )")
                    tables = page.extract_tables()
                    if not tables: continue
                    ws = wb.create_sheet(f"Page_{str(page_idx).zfill(digits)}")
                    current_row = 1
                    for table in tables:
                        for row_data in table:
                            for col_idx, cell_value in enumerate(row_data, 1):
                                val = str(cell_value).strip() if cell_value else ""
                                cell = ws.cell(row=current_row, column=col_idx, value=val)
                                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                            current_row += 1
                        current_row += 2
                    auto_adjust_excel_column_width(ws)
            save_dir = get_save_dir(pdf_path)
            if save_dir:
                wb.save(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_Excel.xlsx"))
        except Exception as e:
            print(f"Excel Error: {e}")

def convert_to_image(files, ext):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
    total_files = len(files)
    for i, f in enumerate(files, 1):
        update_overall_progress(i, total_files, f"全体の進捗 ( {i} / {total_files} ファイル )")
        doc = fitz.open(f)
        save_dir = get_save_dir(f)
        if not save_dir: return
        base = os.path.splitext(os.path.basename(f))[0]
        total_pages = len(doc)
        digits = max(2, len(str(total_pages)))
        for n, page in enumerate(doc, 1):
            set_file_progress_determinate(n, total_pages, f"画像へ変換中... ( {n} / {total_pages} ページ )")
            n_str = str(n).zfill(digits)
            page.get_pixmap(dpi=200).save(os.path.join(save_dir, f"{base}_{n_str}.{ext}"))

def convert_to_dxf(files):
    files = [f for f in files if f.lower().endswith(".pdf")]
    if not files: raise Exception("PDFファイルが含まれていません。")
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
                set_file_progress_determinate(page_num, total_pages, f"DXFへ変換中... ( {page_num} / {total_pages} ページ )")
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

            set_file_progress_determinate(total_pages, total_pages, "DXFファイルを保存中...")
            dwg.saveas(os.path.join(save_dir, f"{os.path.splitext(os.path.basename(f))[0]}_CAD.dxf"))
        except Exception as e:
            print(f"DXF Conversion Error: {e}")

# ==============================
# UI構築 (モダン＆洗練化)
# ==============================
def update_ui():
    path_text = "\n".join(selected_files) if current_mode == "file" else (f"フォルダ: {selected_folder}" if selected_folder else "未選択")
    path_label.config(text=path_text)
    is_active = current_mode is not None
    
    state_val = tk.NORMAL if is_active else tk.DISABLED
    for b in [btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_ai_extract, btn_aggregate]:
        b.config(state=state_val)
    btn_merge.config(state=tk.NORMAL if current_mode=="folder" else tk.DISABLED)

def toggle_api_key_entry(*args):
    if ai_engine_var.get() == "Gemini":
        api_key_entry.config(state=tk.NORMAL)
        btn_api_test.config(state=tk.NORMAL)
    else:
        api_key_entry.config(state=tk.DISABLED)
        btn_api_test.config(state=tk.DISABLED)

# --- ウィンドウ初期位置を左上（+0+0）に設定 ---
root = tk.Tk()
root.title(f"{APP_TITLE} {VERSION}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+0+0")
root.configure(bg=BG_COLOR)

style = ttk.Style()
if "clam" in style.theme_names():
    style.theme_use("clam")
style.configure(".", background=BG_COLOR, font=("Segoe UI", 10))
style.configure("Card.TFrame", background=CARD_BG)
style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1, bordercolor=BORDER_COLOR)
style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=PRIMARY, font=("Segoe UI", 11, "bold"))

style.configure("TButton", padding=6, font=("Segoe UI", 10), background="#E9ECEF", foreground=TEXT_COLOR, borderwidth=1)
style.map("TButton", background=[("active", "#DEE2E6")])
style.configure("Primary.TButton", background=PRIMARY, foreground="white", borderwidth=0)
style.map("Primary.TButton", background=[("active", PRIMARY_HOVER)])

# 機能ごとのカラー設定
COLOR_SUCCESS = "#198754"
COLOR_SUCCESS_HOVER = "#157347"
COLOR_INFO = "#0DCAF0"
COLOR_INFO_HOVER = "#0BACCE"
COLOR_WARNING = "#FFC107"
COLOR_WARNING_HOVER = "#E0A800"
COLOR_DANGER = "#DC3545"
COLOR_DANGER_HOVER = "#B02A37"
COLOR_PURPLE = "#6F42C1"
COLOR_PURPLE_HOVER = "#59339D"

style.configure("Success.TButton", background=COLOR_SUCCESS, foreground="white", borderwidth=0)
style.map("Success.TButton", background=[("active", COLOR_SUCCESS_HOVER)])
style.configure("Info.TButton", background=COLOR_INFO, foreground="white", borderwidth=0)
style.map("Info.TButton", background=[("active", COLOR_INFO_HOVER)])
style.configure("Warning.TButton", background=COLOR_WARNING, foreground="black", borderwidth=0)
style.map("Warning.TButton", background=[("active", COLOR_WARNING_HOVER)])
style.configure("Danger.TButton", background=COLOR_DANGER, foreground="white", borderwidth=0)
style.map("Danger.TButton", background=[("active", COLOR_DANGER_HOVER)])
style.configure("Purple.TButton", background=COLOR_PURPLE, foreground="white", borderwidth=0)
style.map("Purple.TButton", background=[("active", COLOR_PURPLE_HOVER)])

style.configure("TRadiobutton", background=CARD_BG, font=("Segoe UI", 10), foreground=TEXT_COLOR)

# メニューバー
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

# 変数の初期化
rotate_option = tk.IntVar(value=270)
save_option = tk.IntVar(value=1)
ai_engine_var = tk.StringVar(value="Gemini")
api_key_var = tk.StringVar(value=get_api_key() or "")
output_format_var = tk.StringVar(value="xlsx")

# メインコンテナ
main_container = ttk.Frame(root, padding=25)
main_container.pack(fill=tk.BOTH, expand=True)

# タイトルエリア
title_frame = ttk.Frame(main_container)
title_frame.pack(fill=tk.X, pady=(0, 15))
ttk.Label(title_frame, text=APP_TITLE, font=("Segoe UI", 20, "bold"), foreground=PRIMARY).pack(side=tk.LEFT)
ttk.Label(title_frame, text=f" {VERSION}", font=("Segoe UI", 12), foreground=MUTED_TEXT).pack(side=tk.LEFT, pady=(8, 0))
ttk.Label(main_container, text="✨ Update: タプル文字列化バグを修正し、数値整合性のみで列を自動マッピングします。", font=("Meiryo UI", 9), foreground=MUTED_TEXT).pack(anchor="w", pady=(0, 20))

# ファイル選択カード
file_card = ttk.Frame(main_container, style="Card.TFrame", padding=15)
file_card.pack(fill=tk.X, pady=5)
btn_frame = ttk.Frame(file_card, style="Card.TFrame")
btn_frame.pack()
ttk.Button(btn_frame, text="📄 ファイルを選択", command=select_files, width=22, style="Primary.TButton").grid(row=0, column=0, padx=8)
ttk.Button(btn_frame, text="📁 フォルダを選択", command=select_folder, width=22, style="Primary.TButton").grid(row=0, column=1, padx=8)

path_label = ttk.Label(file_card, text="未選択", background=CARD_BG, foreground=TEXT_COLOR, wraplength=580, justify="center")
path_label.pack(pady=(12, 0))

# 設定グリッド (保存先 & 回転)
settings_grid = ttk.Frame(main_container)
settings_grid.pack(fill=tk.X, pady=15)
settings_grid.columnconfigure(0, weight=1)
settings_grid.columnconfigure(1, weight=1)

save_frame = ttk.LabelFrame(settings_grid, text=" 保存先設定 ", style="Card.TLabelframe", padding=12)
save_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
ttk.Radiobutton(save_frame, text="元のファイルと同じフォルダ", variable=save_option, value=1, command=on_save_mode_change).pack(anchor="w", pady=4)
ttk.Radiobutton(save_frame, text="任意のフォルダを指定", variable=save_option, value=2, command=on_save_mode_change).pack(anchor="w", pady=4)
ttk.Button(save_frame, text="📂 フォルダ参照", command=select_save_dir).pack(pady=(10, 5))
save_label = ttk.Label(save_frame, text="同じフォルダ", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 9))
save_label.pack()

rotate_frame = ttk.LabelFrame(settings_grid, text=" 回転設定 ", style="Card.TLabelframe", padding=12)
rotate_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
for t, v in [("左（270°）", 270), ("上下（180°）", 180), ("右（90°）", 90)]:
    ttk.Radiobutton(rotate_frame, text=t, variable=rotate_option, value=v).pack(anchor="w", pady=5)

# AI設定カード
ai_frame = ttk.LabelFrame(main_container, text=" AI抽出設定 ", style="Card.TLabelframe", padding=12)
ai_frame.pack(fill=tk.X, pady=10)

engine_frame = ttk.Frame(ai_frame, style="Card.TFrame")
engine_frame.pack(fill=tk.X, pady=5)
ttk.Label(engine_frame, text="エンジン:", width=10, background=CARD_BG, font=("Segoe UI", 10, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
ttk.Radiobutton(engine_frame, text="Gemini API (推奨)", variable=ai_engine_var, value="Gemini").pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(engine_frame, text="Tesseract", variable=ai_engine_var, value="Tesseract").pack(side=tk.LEFT, padx=5)

api_key_frame = ttk.Frame(ai_frame, style="Card.TFrame")
api_key_frame.pack(fill=tk.X, pady=5)
ttk.Label(api_key_frame, text="APIキー:", width=10, background=CARD_BG, font=("Segoe UI", 10, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
api_key_entry = ttk.Entry(api_key_frame, textvariable=api_key_var, width=50, show="*")
api_key_entry.pack(side=tk.LEFT, padx=(0, 8))
btn_api_test = ttk.Button(api_key_frame, text="テスト", command=test_api_key_ui, width=8)
btn_api_test.pack(side=tk.LEFT)

format_frame = ttk.Frame(ai_frame, style="Card.TFrame")
format_frame.pack(fill=tk.X, pady=5)
ttk.Label(format_frame, text="出力形式:", width=10, background=CARD_BG, font=("Segoe UI", 10, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
for fmt in ["xlsx", "csv", "txt"]:
    ttk.Radiobutton(format_frame, text=f".{fmt}", variable=output_format_var, value=fmt).pack(side=tk.LEFT, padx=5)

ai_engine_var.trace("w", toggle_api_key_entry)

# 操作ボタン群カード
op_frame = ttk.LabelFrame(main_container, text=" 実行アクション ", style="Card.TLabelframe", padding=15)
op_frame.pack(fill=tk.X, pady=10)
op_inner = ttk.Frame(op_frame, style="Card.TFrame")
op_inner.pack()

btn_merge = ttk.Button(op_inner, text="結合", width=14, command=lambda: safe_run(merge_pdfs), style="Primary.TButton")
btn_split = ttk.Button(op_inner, text="分割", width=14, command=lambda: safe_run(split_pdfs), style="Primary.TButton")
btn_rotate = ttk.Button(op_inner, text="回転", width=14, command=lambda: safe_run(rotate_pdfs), style="Primary.TButton")
btn_text = ttk.Button(op_inner, text="Text抽出", width=14, command=lambda: safe_run(extract_text), style="Success.TButton")

btn_excel = ttk.Button(op_inner, text="Excel変換", width=14, command=lambda: safe_run(convert_to_excel), style="Success.TButton")
btn_jpeg = ttk.Button(op_inner, text="JPEG変換", width=14, command=lambda: safe_run(lambda fs: convert_to_image(fs, "jpg")), style="Info.TButton")
btn_png = ttk.Button(op_inner, text="PNG変換", width=14, command=lambda: safe_run(lambda fs: convert_to_image(fs, "png")), style="Info.TButton")
btn_dxf = ttk.Button(op_inner, text="DXF変換", width=14, command=lambda: safe_run(convert_to_dxf), style="Warning.TButton")

btn_ai_extract = ttk.Button(op_inner, text="AIデータ抽出", width=14, command=run_ai_extraction, style="Danger.TButton")
btn_aggregate = ttk.Button(op_inner, text="データ集約", width=14, command=lambda: safe_run(aggregate_only_task), style="Purple.TButton")

op_list = [btn_merge, btn_split, btn_rotate, btn_text, btn_excel, btn_jpeg, btn_png, btn_dxf, btn_ai_extract, btn_aggregate]
for i, b in enumerate(op_list):
    b.grid(row=i//4, column=i%4, padx=10, pady=10)

update_ui()
toggle_api_key_entry()
root.mainloop()