# -*- coding: utf-8 -*-
import os, sys, threading, re, json, warnings
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter.scrolledtext as st
import fitz
from PIL import Image, ImageTk

# 今後使えなくなる予定の google.generativeai に関する警告(FutureWarning)をコンソールに表示しないよう安全に抑制
warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai

from pdf_engine import (
    merge_pdfs, split_pdfs, rotate_pdfs,
    extract_text_internal, convert_to_excel_internal, convert_to_csv_internal,
    convert_to_image_jpg, convert_to_image_png, convert_to_dxf,
    convert_to_image_tiff, convert_to_image_bmp, convert_to_svg
)
from ai_engine import (
    extract_tesseract_task, extract_gemini_task, aggregate_local_task, aggregate_gemini_task
)

# ==============================
# 基本設定 & カラーパレット
# ==============================
APP_TITLE, VERSION = "PdfEditMiya", "v2.1.0"
WINDOW_WIDTH, WINDOW_HEIGHT = 900, 750

BG_COLOR, CARD_BG = "#F0F4F8", "#FFFFFF"
PRIMARY, PRIMARY_HOVER = "#0D6EFD", "#0B5ED7"
TEXT_COLOR, MUTED_TEXT, BORDER_COLOR = "#212529", "#6C757D", "#DEE2E6"

SUCCESS, ERROR = "#198754", "#DC3545"
COLOR_SUCCESS, COLOR_SUCCESS_HOVER = "#198754", "#157347"
COLOR_INFO, COLOR_INFO_HOVER = "#0DCAF0", "#0BACCE"
COLOR_WARNING, COLOR_WARNING_HOVER = "#FFC107", "#E0A800"
COLOR_DANGER, COLOR_DANGER_HOVER = "#DC3545", "#B02A37"
COLOR_PURPLE, COLOR_PURPLE_HOVER = "#6F42C1", "#59339D"

USER_HOME = os.path.expanduser("~")
API_KEY_FILE = os.path.join(USER_HOME, ".pdfeditmiya_api_key.txt")
SETTINGS_FILE = os.path.join(USER_HOME, ".pdfeditmiya_settings.json") 

# ==============================
# ヘルプ・履歴テキスト
# ==============================
VERSION_HISTORY = """
[ v2.1.0 ]
- 【UI改善】API詳細設定画面の「制限と仕様を確認」ボタンを強化し、プラン比較や各モデルの特徴に加え「モデルごとのRPM・スレッド数の制限と推奨値」を詳細なエクセル風テーブルで一目で確認できるよう改善しました。
- 【UI改善】API詳細設定画面のレイアウトを大幅に見直し、項目を左右に並べることで縦幅を削減。画面の低いノートPC等でも見切れずに全体が収まるように改善しました。
- 【機能改善】AI抽出のカスタムプロンプトを「左右分割のチェックボックス型UI」へと大幅に刷新しました。指示を1行ずつ追加でき、お気に入りの保存・復元もチェックボックスで直感的に行えるようになりました。
- 【機能追加】API詳細設定画面に、抽出精度を高める「Temperature」、エラーを回避する「安全フィルタ無効化」、出力上限を増やす「最大トークン数」、処理スピードを調整する「同時処理スレッド数（直列/並列）」などの高度な設定項目をすべて追加しました。
- 【機能追加】API詳細設定画面を実装し、Gemini APIの設定を「無料枠」と「課金枠」それぞれ独立して設定・保持できるようになりました。
- 【機能改善】「実行プランの選択」と「設定タブの表示」を分離し、選択状態を維持したまま各プランの設定を自由に確認・編集できるようになりました。
- 【機能改善】「フォルダを選択」からのデータ集約時、異なる文字コード（Shift_JISなど）のCSV・テキストファイルを自動判定して読み込むように改善しました。
- 【高速化】Gemini APIでのクロップ（範囲抽出）時、複数領域を1回のリクエストで同時解析するように変更し、大幅な高速化を実現しました。
"""

AI_HELP_TEXT = """
【 AI抽出機能の使い方と準備 】

PDF内の表データや手書き文字を解析し、Excel(xlsx)・CSV・テキスト・Word・JSON・Markdownデータとして抽出する機能です。
用途に合わせて2つのAIエンジンを切り替えて使用できます。

───────────────────────────
■ Gemini API を使う場合（推奨・超高精度）
───────────────────────────
最新のAIモデルを利用し、かすれた文字や複雑な表の罫線を高精度に認識します。
インターネット接続と、無料の「APIキー」が必要です。

[APIキーの取得手順]
1. ブラウザで以下のURLにアクセスします。
   https://aistudio.google.com/app/apikey
2. お持ちのGoogleアカウントでログインします。
3. 画面左上の「Create API key」ボタンを押します。
4. 「Create API key in new project」を選択します。
5. 発行された長い英数字の文字列（APIキー）をコピーします。
6. 本アプリの「詳細設定」ボタンを押し、開いた画面の「APIキー」欄で右クリックして貼り付け、「テスト」ボタンを押してください。

───────────────────────────
■ Tesseract を使う場合（オフライン・簡易抽出）
───────────────────────────
インターネットに繋がっていない環境でも使用できる、無料のOCRソフトです。

[インストール手順]
1. ブラウザで以下のURLにアクセスします。
   https://github.com/UB-Mannheim/tesseract/wiki
2. 「tesseract-ocr-w64-setup...exe」など、最新の64bit版インストーラーをダウンロードして実行します。
3. インストール中の「Choose Components」画面で、「Additional language data (download)」の中にある「Japanese」および「Japanese (vertical)」に必ずチェックを入れてください。
4. インストール先は初期設定のまま（C:\\Program Files\\Tesseract-OCR）進めてください。
"""

# ==============================
# グローバル状態
# ==============================
selected_files, selected_folder, current_mode = [], "", None
preset_save_dir, selected_crop_regions = [], []
processing_popup, overall_label, overall_progress = None, None, None
file_label, file_progress, cancelled = None, None, False
saved_custom_prompts = []  

# ==============================
# UI共通コンポーネント (スクロール可能なチェックボックスリスト)
# ==============================
class ScrollableCheckboxList(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg=CARD_BG, highlightthickness=1, highlightbackground=BORDER_COLOR, height=120)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas, style="Card.TFrame")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        self.items = []

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def set_items(self, item_texts):
        for item in self.items:
            item["cb"].destroy()
        self.items.clear()
        for text in item_texts:
            self.add_item(text)

    def add_item(self, text):
        var = tk.BooleanVar(value=False)
        cb = ttk.Checkbutton(self.scrollable_frame, text=text, variable=var, style="TCheckbutton")
        cb.pack(anchor="w", padx=5, pady=2, fill="x")
        self.items.append({"text": text, "var": var, "cb": cb})

    def get_all_items(self):
        return [item["text"] for item in self.items]

    def get_selected_items(self):
        return [item["text"] for item in self.items if item["var"].get()]

    def remove_selected(self):
        new_items = []
        for item in self.items:
            if item["var"].get():
                item["cb"].destroy()
            else:
                new_items.append(item)
        self.items = new_items

class UIController:
    def update_overall(self, step, max_val=None, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                if max_val is not None: overall_progress["maximum"] = max_val
                overall_progress["value"] = step
                if text: overall_label.config(text=text)
        root.after(0, _task)
    def set_indeterminate(self, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                file_progress.config(mode="indeterminate"); file_progress.start(15)
                if text: file_label.config(text=text)
        root.after(0, _task)
    def set_determinate(self, step, max_val=None, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                file_progress.stop(); file_progress.config(mode="determinate")
                if max_val is not None: file_progress["maximum"] = max_val
                file_progress["value"] = step
                if text: file_label.config(text=text)
        root.after(0, _task)
    def is_cancelled(self):
        global cancelled
        return cancelled

# ==============================
# 設定の保存・読み込み機能
# ==============================
def load_settings():
    global preset_save_dir
    global saved_custom_prompts
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
            
            if "rotate_option" in settings: rotate_option.set(settings["rotate_option"])
            if "engine_var" in settings: engine_var.set(settings["engine_var"])
            if "output_format_var" in settings: output_format_var.set(settings["output_format_var"])
            if "api_plan_var" in settings: api_plan_var.set(settings["api_plan_var"])
            
            if "api_key_free_var" in settings: api_key_free_var.set(settings["api_key_free_var"])
            elif "api_key_var" in settings: api_key_free_var.set(settings["api_key_var"])
            
            if "api_key_paid_var" in settings: api_key_paid_var.set(settings["api_key_paid_var"])
            elif "api_key_var" in settings: api_key_paid_var.set(settings["api_key_var"])

            if "gemini_model_free_var" in settings: gemini_model_free_var.set(settings["gemini_model_free_var"])
            elif "gemini_model_var" in settings: gemini_model_free_var.set(settings["gemini_model_var"])

            if "gemini_model_paid_var" in settings: gemini_model_paid_var.set(settings["gemini_model_paid_var"])
            elif "gemini_model_var" in settings: gemini_model_paid_var.set(settings["gemini_model_var"])

            if "api_rpm_free_var" in settings: api_rpm_free_var.set(settings["api_rpm_free_var"])
            if "api_rpm_paid_var" in settings: api_rpm_paid_var.set(settings["api_rpm_paid_var"])

            if "temperature_free_var" in settings: temperature_free_var.set(settings["temperature_free_var"])
            if "temperature_paid_var" in settings: temperature_paid_var.set(settings["temperature_paid_var"])
            if "safety_free_var" in settings: safety_free_var.set(settings["safety_free_var"])
            if "safety_paid_var" in settings: safety_paid_var.set(settings["safety_paid_var"])
            if "max_tokens_free_var" in settings: max_tokens_free_var.set(settings["max_tokens_free_var"])
            if "max_tokens_paid_var" in settings: max_tokens_paid_var.set(settings["max_tokens_paid_var"])
            if "custom_prompt_free_var" in settings: custom_prompt_free_var.set(settings["custom_prompt_free_var"])
            if "custom_prompt_paid_var" in settings: custom_prompt_paid_var.set(settings["custom_prompt_paid_var"])
            if "threads_free_var" in settings: threads_free_var.set(settings["threads_free_var"])
            if "threads_paid_var" in settings: threads_paid_var.set(settings["threads_paid_var"])
            
            if "saved_custom_prompts" in settings:
                saved_custom_prompts = settings["saved_custom_prompts"]

            if "save_option" in settings: 
                save_option.set(settings["save_option"])
                if settings["save_option"] == 1:
                    save_label.config(text="同じフォルダ")
                elif settings["save_option"] == 2:
                    if "preset_save_dir" in settings and settings["preset_save_dir"]:
                        preset_save_dir = settings["preset_save_dir"]
                        save_label.config(text=preset_save_dir)
                    else:
                        save_label.config(text="未選択")
            
            if "window_width" in settings and "window_height" in settings:
                w = max(settings["window_width"], 760)
                h = max(settings["window_height"], 650)
                root.geometry(f"{w}x{h}")
                
        except Exception as e:
            print(f"Failed to load settings: {e}")

def save_settings():
    root.update_idletasks()
    settings = {
        "rotate_option": rotate_option.get(),
        "save_option": save_option.get(),
        "preset_save_dir": preset_save_dir,
        "engine_var": engine_var.get(),
        "output_format_var": output_format_var.get(),
        "api_plan_var": api_plan_var.get(),
        "api_key_free_var": api_key_free_var.get().strip(),
        "api_key_paid_var": api_key_paid_var.get().strip(),
        "gemini_model_free_var": gemini_model_free_var.get(),
        "gemini_model_paid_var": gemini_model_paid_var.get(),
        "api_rpm_free_var": api_rpm_free_var.get(),
        "api_rpm_paid_var": api_rpm_paid_var.get(),
        "temperature_free_var": temperature_free_var.get(),
        "temperature_paid_var": temperature_paid_var.get(),
        "safety_free_var": safety_free_var.get(),
        "safety_paid_var": safety_paid_var.get(),
        "max_tokens_free_var": max_tokens_free_var.get(),
        "max_tokens_paid_var": max_tokens_paid_var.get(),
        "custom_prompt_free_var": custom_prompt_free_var.get(),
        "custom_prompt_paid_var": custom_prompt_paid_var.get(),
        "threads_free_var": threads_free_var.get(),
        "threads_paid_var": threads_paid_var.get(),
        "saved_custom_prompts": saved_custom_prompts,
        "window_width": root.winfo_width(),
        "window_height": root.winfo_height()
    }
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
            
        plan = api_plan_var.get()
        current_key = api_key_free_var.get().strip() if plan == "free" else api_key_paid_var.get().strip()
        if current_key:
            with open(API_KEY_FILE, "w", encoding="utf-8") as f:
                f.write(current_key)
                
        messagebox.showinfo("保存完了", "現在の選択項目を保存しました。\n次回起動時もこの設定が適用されます。")
    except Exception as e:
        messagebox.showerror("エラー", f"設定の保存に失敗しました。\n{e}")

# ==============================
# ユーティリティ・ヘルパー関数
# ==============================
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    return None

def show_context_menu(event, target_widget=None):
    widget = target_widget if target_widget else event.widget
    menu = Menu(root, tearoff=0)
    menu.add_command(label="貼り付け", command=lambda: paste_to_entry(widget))
    menu.post(event.x_root, event.y_root)

def paste_to_entry(widget):
    try:
        text = root.clipboard_get()
        try: widget.delete("sel.first", "sel.last")
        except tk.TclError: pass
        widget.insert(tk.INSERT, text)
    except tk.TclError: pass

def show_text_context_menu(event, text_widget):
    menu = Menu(root, tearoff=0)
    menu.add_command(label="コピー", command=lambda: text_widget.event_generate("<<Copy>>"))
    menu.add_command(label="切り取り", command=lambda: text_widget.event_generate("<<Cut>>"))
    menu.add_command(label="貼り付け", command=lambda: text_widget.event_generate("<<Paste>>"))
    menu.post(event.x_root, event.y_root)

def show_message(msg, color=PRIMARY):
    def _task():
        win = tk.Toplevel(root)
        win.geometry("260x90"); win.configure(bg=CARD_BG); win.attributes("-topmost", True)
        x = root.winfo_x() + (WINDOW_WIDTH // 2) - 130; y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 45
        win.geometry(f"+{x}+{y}"); win.overrideredirect(True)
        frame = tk.Frame(win, bg=CARD_BG, highlightbackground=color, highlightthickness=2); frame.pack(expand=True, fill=tk.BOTH)
        ttk.Label(frame, text=msg, foreground=color, font=("Segoe UI", 10, "bold"), background=CARD_BG, wraplength=240).pack(expand=True)
        win.after(2500, win.destroy)
    root.after(0, _task)

# ==============================
# 処理実行・制御系関数
# ==============================
def show_processing(total_files=1):
    global processing_popup, overall_label, overall_progress, file_label, file_progress, btn_cancel
    processing_popup = tk.Toplevel(root)
    processing_popup.title("処理を実行中...")
    processing_popup.geometry("440x280")
    processing_popup.configure(bg=CARD_BG)
    processing_popup.grab_set()
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 220
    y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 140
    processing_popup.geometry(f"+{x}+{y}")
    
    engine_name = engine_var.get()
    if engine_name == "Gemini":
        plan = api_plan_var.get()
        model_name = gemini_model_free_var.get() if plan == "free" else gemini_model_paid_var.get()
        engine_text = f"使用エンジン: Gemini API ( {model_name} / {plan.capitalize()} )"
    elif engine_name == "Tesseract":
        engine_text = "使用エンジン: Tesseract (ローカルOCR)"
    else:
        engine_text = "使用エンジン: Python標準ライブラリ"
        
    engine_label = ttk.Label(processing_popup, text=engine_text, font=("Segoe UI", 9, "bold"), background=CARD_BG, foreground=COLOR_PURPLE)
    engine_label.pack(pady=(20, 0))

    overall_label = ttk.Label(processing_popup, text=f"全体の進捗 ( 0 / {total_files} ファイル )", font=("Segoe UI", 10, "bold"), background=CARD_BG, foreground=PRIMARY)
    overall_label.pack(pady=(15, 5))
    overall_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=total_files, length=380); overall_progress.pack(pady=(0, 20))
    file_label = ttk.Label(processing_popup, text="現在のファイルを準備中...", font=("Segoe UI", 9), background=CARD_BG, foreground=MUTED_TEXT); file_label.pack(pady=(5, 5))
    file_progress = ttk.Progressbar(processing_popup, mode="determinate", maximum=1, length=380); file_progress.pack(pady=(0, 10))
    
    def cancel_processing():
        global cancelled
        if messagebox.askyesno("確認", "処理を中止しますか？", parent=processing_popup):
            cancelled = True
            btn_cancel.config(text="中止処理中...", state=tk.DISABLED)
            
    btn_cancel = ttk.Button(processing_popup, text="処理を中止", command=cancel_processing, style="Warning.TButton")
    btn_cancel.pack(pady=(5, 10))

def close_processing():
    def _task():
        global processing_popup
        if processing_popup: processing_popup.destroy(); processing_popup = None
    root.after(0, _task)

def run_task(func, task_name):
    global cancelled; cancelled = False
    try:
        def _start_status(): status_label.config(text=f"ステータス: {task_name} を実行中...", foreground=PRIMARY)
        root.after(0, _start_status)
        
        files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
        if not files: return
        save_dir = os.path.dirname(files[0]) if save_option.get() == 1 else preset_save_dir
        
        plan = api_plan_var.get()
        is_free = (plan == "free")
        api_key = api_key_free_var.get().strip() if is_free else api_key_paid_var.get().strip()
        model = gemini_model_free_var.get() if is_free else gemini_model_paid_var.get()
        rpm = api_rpm_free_var.get() if is_free else api_rpm_paid_var.get()
        
        options = {
            "rotate_deg": rotate_option.get(), "crop_regions": selected_crop_regions, "out_format": output_format_var.get(),
            "folder_name": os.path.basename(selected_folder) if selected_folder else "Merged",
            "api_key": api_key,
            "models_to_try": [model] if engine_var.get() == "Gemini" else [],
            "api_plan": plan,
            "api_rpm": rpm,
            "temperature": temperature_free_var.get() if is_free else temperature_paid_var.get(),
            "disable_safety": safety_free_var.get() if is_free else safety_paid_var.get(),
            "max_tokens": max_tokens_free_var.get() if is_free else max_tokens_paid_var.get(),
            "custom_prompt": custom_prompt_free_var.get() if is_free else custom_prompt_paid_var.get(),
            "threads": threads_free_var.get() if is_free else threads_paid_var.get()
        }
        func(files, save_dir, options, UIController())
        close_processing()
        
        def _end_status():
            if cancelled:
                show_message("⚠️ 処理が中止されました", COLOR_WARNING)
                status_label.config(text=f"ステータス: {task_name} は中止されました", foreground=COLOR_WARNING)
            else:
                show_message("✅ 処理が完了しました", SUCCESS)
                status_label.config(text=f"ステータス: {task_name} が完了しました", foreground=SUCCESS)
        root.after(0, _end_status)
        
    except Exception as e:
        print(f"Error: {e}"); close_processing()
        def _err_status():
            show_message(f"❌ エラーが発生しました\n{str(e)[:40]}...", ERROR)
            status_label.config(text=f"ステータス: {task_name} でエラーが発生しました", foreground=ERROR)
        root.after(0, _err_status)

def safe_run(func, task_name="処理"):
    files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
    if not files: return
    global preset_save_dir
    if save_option.get() == 2 and not preset_save_dir:
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        if not folder: return
        preset_save_dir = folder; save_label.config(text=preset_save_dir)
    show_processing(len(files))
    threading.Thread(target=run_task, args=(func, task_name), daemon=True).start()

def run_selected_extraction():
    engine = engine_var.get(); fmt = output_format_var.get()
    
    engine_ja = {"Internal": "標準ライブラリ", "Tesseract": "Tesseract OCR", "Gemini": "Gemini API"}.get(engine, engine)
    task_name = f"抽出・変換 ({engine_ja} -> {fmt.upper()})"
    
    if fmt == "jpg": safe_run(convert_to_image_jpg, task_name)
    elif fmt == "png": safe_run(convert_to_image_png, task_name)
    elif fmt == "tiff": safe_run(convert_to_image_tiff, task_name)
    elif fmt == "bmp": safe_run(convert_to_image_bmp, task_name)
    elif fmt == "svg": safe_run(convert_to_svg, task_name)
    elif fmt == "dxf": safe_run(convert_to_dxf, task_name)
    elif engine == "Internal":
        if fmt == "txt": safe_run(extract_text_internal, task_name)
        elif fmt == "xlsx": safe_run(convert_to_excel_internal, task_name)
        elif fmt == "csv": safe_run(convert_to_csv_internal, task_name)
        elif fmt in ["json", "md", "docx"]: 
            messagebox.showinfo("情報", "標準ライブラリ（Internal）はJSON / Markdown / Word出力に未対応です。\nExcelやCSV、またはAIエンジンを使用してください。")
            return
    elif engine == "Tesseract": safe_run(extract_tesseract_task, task_name)
    elif engine == "Gemini":
        plan = api_plan_var.get()
        api_key = api_key_free_var.get().strip() if plan == "free" else api_key_paid_var.get().strip()
        if not api_key: 
            return messagebox.showerror("エラー", f"Gemini APIキー({plan.capitalize()}枠用)を入力してください。\n「⚙️ 詳細設定」ボタンから設定できます。")
        safe_run(extract_gemini_task, task_name)

# ==============================
# API詳細設定ダイアログ
# ==============================
def open_api_settings_dialog():
    dialog = tk.Toplevel(root)
    dialog.title("⚙️ AI詳細設定 (Gemini API)")
    
    screen_h = root.winfo_screenheight()
    dialog_h = min(780, screen_h - 80) 
    dialog.geometry(f"1050x{dialog_h}") 
    dialog.configure(bg=BG_COLOR)
    dialog.grab_set()
    
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 525
    y = max(10, root.winfo_y() + (WINDOW_HEIGHT // 2) - (dialog_h // 2))
    dialog.geometry(f"+{x}+{y}")

    fav_lists = []

    def update_all_fav_lists():
        for f_list in fav_lists:
            f_list.set_items(saved_custom_prompts)

    original_values = {
        "plan": api_plan_var.get(),
        "key_free": api_key_free_var.get(), "key_paid": api_key_paid_var.get(),
        "model_free": gemini_model_free_var.get(), "model_paid": gemini_model_paid_var.get(),
        "rpm_free": api_rpm_free_var.get(), "rpm_paid": api_rpm_paid_var.get(),
        "temp_free": temperature_free_var.get(), "temp_paid": temperature_paid_var.get(),
        "safety_free": safety_free_var.get(), "safety_paid": safety_paid_var.get(),
        "tokens_free": max_tokens_free_var.get(), "tokens_paid": max_tokens_paid_var.get(),
        "prompt_free": custom_prompt_free_var.get(), "prompt_paid": custom_prompt_paid_var.get(),
        "threads_free": threads_free_var.get(), "threads_paid": threads_paid_var.get(),
        "saved_prompts": list(saved_custom_prompts)
    }

    def apply_and_close():
        dialog.destroy()

    def has_changes():
        if api_plan_var.get() != original_values["plan"]: return True
        if api_key_free_var.get() != original_values["key_free"]: return True
        if api_key_paid_var.get() != original_values["key_paid"]: return True
        if gemini_model_free_var.get() != original_values["model_free"]: return True
        if gemini_model_paid_var.get() != original_values["model_paid"]: return True
        if api_rpm_free_var.get() != original_values["rpm_free"]: return True
        if api_rpm_paid_var.get() != original_values["rpm_paid"]: return True
        if temperature_free_var.get() != original_values["temp_free"]: return True
        if temperature_paid_var.get() != original_values["temp_paid"]: return True
        if safety_free_var.get() != original_values["safety_free"]: return True
        if safety_paid_var.get() != original_values["safety_paid"]: return True
        if max_tokens_free_var.get() != original_values["tokens_free"]: return True
        if max_tokens_paid_var.get() != original_values["tokens_paid"]: return True
        if custom_prompt_free_var.get() != original_values["prompt_free"]: return True
        if custom_prompt_paid_var.get() != original_values["prompt_paid"]: return True
        if threads_free_var.get() != original_values["threads_free"]: return True
        if threads_paid_var.get() != original_values["threads_paid"]: return True
        if saved_custom_prompts != original_values["saved_prompts"]: return True
        return False

    def cancel_and_close():
        if has_changes():
            if not messagebox.askyesno("確認", "変更が適用されていません。\n破棄して設定画面を閉じますか？", parent=dialog):
                return 
                
        api_plan_var.set(original_values["plan"])
        api_key_free_var.set(original_values["key_free"])
        api_key_paid_var.set(original_values["key_paid"])
        gemini_model_free_var.set(original_values["model_free"])
        gemini_model_paid_var.set(original_values["model_paid"])
        api_rpm_free_var.set(original_values["rpm_free"])
        api_rpm_paid_var.set(original_values["rpm_paid"])
        temperature_free_var.set(original_values["temp_free"])
        temperature_paid_var.set(original_values["temp_paid"])
        safety_free_var.set(original_values["safety_free"])
        safety_paid_var.set(original_values["safety_paid"])
        max_tokens_free_var.set(original_values["tokens_free"])
        max_tokens_paid_var.set(original_values["tokens_paid"])
        custom_prompt_free_var.set(original_values["prompt_free"])
        custom_prompt_paid_var.set(original_values["prompt_paid"])
        threads_free_var.set(original_values["threads_free"])
        threads_paid_var.set(original_values["threads_paid"])
        saved_custom_prompts.clear()
        saved_custom_prompts.extend(original_values["saved_prompts"])
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", cancel_and_close)

    lbl_title = ttk.Label(dialog, text="Gemini API 詳細設定", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
    lbl_title.pack(pady=(10, 5))

    # --- 実行プランの選択 ---
    plan_frame = ttk.LabelFrame(dialog, text=" 実行プランの選択 ", style="Card.TLabelframe", padding=8)
    plan_frame.pack(fill=tk.X, padx=15, pady=(0, 5))
    
    plan_inner = ttk.Frame(plan_frame, style="Card.TFrame")
    plan_inner.pack(anchor="w", padx=5, pady=2)

    ttk.Label(plan_inner, text="実際に抽出で使用するプランを選んでください（下のタブとは連動しません）:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 15))
    rb_free = ttk.Radiobutton(plan_inner, text="無料枠 (Free Tier)", variable=api_plan_var, value="free")
    rb_free.pack(side=tk.LEFT, padx=(0, 15))
    rb_paid = ttk.Radiobutton(plan_inner, text="課金枠 (Paid Tier)", variable=api_plan_var, value="paid")
    rb_paid.pack(side=tk.LEFT)

    # --- 個別設定タブの構築 ---
    notebook = ttk.Notebook(dialog)
    notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
    
    tab_free = ttk.Frame(notebook, style="Main.TFrame")
    tab_paid = ttk.Frame(notebook, style="Main.TFrame")
    notebook.add(tab_free, text=" 🟢 無料枠 (Free Tier) の設定 ")
    notebook.add(tab_paid, text=" 🔵 課金枠 (Paid Tier) の設定 ")

    models = [
        ("Gemini 2.5 Flash (高速・万能 / 推奨)", "gemini-2.5-flash"),
        ("Gemini 2.5 Pro (主力・高精度)", "gemini-2.5-pro"),
        ("Gemini 2.5 Flash-Lite (最軽量・低コスト)", "gemini-2.5-flash-8b"),
        ("Gemini 3.1 Pro Preview (次世代プレビュー)", "gemini-3.1-pro-preview"),
        ("Gemini 3.0 Flash Preview (次世代プレビュー)", "gemini-3.0-flash-preview")
    ]

    def build_tab(parent_tab, plan_type):
        is_free = (plan_type == "free")
        key_var = api_key_free_var if is_free else api_key_paid_var
        model_var = gemini_model_free_var if is_free else gemini_model_paid_var
        rpm_var = api_rpm_free_var if is_free else api_rpm_paid_var
        temp_var = temperature_free_var if is_free else temperature_paid_var
        safety_var = safety_free_var if is_free else safety_paid_var
        tokens_var = max_tokens_free_var if is_free else max_tokens_paid_var
        prompt_var = custom_prompt_free_var if is_free else custom_prompt_paid_var
        threads_var = threads_free_var if is_free else threads_paid_var
        
        # ① APIキー
        key_frame = ttk.LabelFrame(parent_tab, text=" ① APIキー ", style="Card.TLabelframe", padding=8)
        key_frame.pack(fill=tk.X, padx=10, pady=5)
        
        key_inner = ttk.Frame(key_frame, style="Card.TFrame")
        key_inner.pack(fill=tk.X)
        
        ttk.Label(key_inner, text=f"{plan_type.capitalize()} 用のAPIキー:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        
        entry_key = ttk.Entry(key_inner, textvariable=key_var, width=55, show="*")
        entry_key.pack(side=tk.LEFT, padx=(0, 5))
        entry_key.bind("<Button-3>", lambda e, widget=entry_key: show_context_menu(e, widget))
        
        btn_toggle = ttk.Button(key_inner, text="確認", width=6)
        btn_toggle.pack(side=tk.LEFT, padx=(0, 5))
        
        def toggle_key(e=entry_key, b=btn_toggle):
            if e.cget('show') == '*':
                e.configure(show='')
                b.configure(text="隠す")
            else:
                e.configure(show='*')
                b.configure(text="確認")
        btn_toggle.config(command=toggle_key)

        def test_key(k_var=key_var, m_var=model_var):
            key = k_var.get().strip()
            if not key: return messagebox.showwarning("警告", "APIキーが入力されていません。", parent=dialog)
            genai.configure(api_key=key)
            model_name = m_var.get()
            try:
                model = genai.GenerativeModel(model_name)
                model.generate_content("Test")
                messagebox.showinfo("テスト成功", f"APIキーは正しく認識されました。\nAIモデル「{model_name}」による通信は正常です！", parent=dialog)
            except Exception as e:
                err_str = str(e).lower()
                if "404" in err_str or "not found" in err_str:
                    messagebox.showerror("モデル利用不可", f"エラー: モデル「{model_name}」が存在しないか、利用する権限がありません。\n詳細:\n{e}", parent=dialog)
                elif "429" in err_str or "quota" in err_str:
                    msg = f"エラー: APIの利用枠（クォータ）を超過しています。\n\n"
                    m = re.search(r'retry in ([\d\.]+)s', err_str)
                    if not m: m = re.search(r'seconds:\s*(\d+)', err_str, re.IGNORECASE | re.DOTALL)
                    if m:
                        wait_sec = int(float(m.group(1)))
                        msg += f"⚠️ Googleのバースト制限です。約 {wait_sec} 秒後に利用枠が回復します。\n"
                    else:
                        if "perday" in err_str.lower() or "limit: 20" in err_str.lower(): msg += "⚠️ 【1日の利用上限】に達した可能性があります。\n"
                        else: msg += "⚠️ APIの制限に達しました。\n"
                    messagebox.showerror("利用枠超過", msg, parent=dialog)
                else:
                    messagebox.showerror("通信エラー", f"APIキーまたは通信に問題が発生しました。\n詳細:\n{e}", parent=dialog)
        
        btn_test = ttk.Button(key_inner, text="テスト", command=test_key, width=6)
        btn_test.pack(side=tk.LEFT)

        # 縦幅圧縮のため、②と③を横並びにするレイアウト枠
        middle_frame = ttk.Frame(parent_tab, style="Main.TFrame")
        middle_frame.pack(fill=tk.X, padx=10, pady=5)
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.columnconfigure(1, weight=1)

        # ② モデル・パフォーマンス設定 (左側)
        perf_frame = ttk.LabelFrame(middle_frame, text=" ② モデル・パフォーマンス設定 ", style="Card.TLabelframe", padding=8)
        perf_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        model_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        model_inner.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(model_inner, text="使用モデル:", width=10, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        model_combo = ttk.Combobox(model_inner, values=[m[0] for m in models], state="readonly", width=42)
        current_val = model_var.get()
        for m in models:
            if m[1] == current_val:
                model_combo.set(m[0]); break
        if not model_combo.get(): model_combo.set(models[0][0])
                
        def on_model_select(event, cb=model_combo, m_var=model_var):
            selected_display = cb.get()
            for m in models:
                if m[0] == selected_display:
                    m_var.set(m[1]); break
        model_combo.bind("<<ComboboxSelected>>", on_model_select)
        model_combo.pack(side=tk.LEFT)

        speed_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        speed_inner.pack(fill=tk.X, pady=2)
        
        ttk.Label(speed_inner, text="RPM:", width=5, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        spin_rpm = ttk.Spinbox(speed_inner, from_=1, to=2000, textvariable=rpm_var, width=5)
        spin_rpm.pack(side=tk.LEFT, padx=(0, 2))
        
        ttk.Label(speed_inner, text="スレッド:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        spin_threads = ttk.Spinbox(speed_inner, from_=1, to=20, textvariable=threads_var, width=4)
        spin_threads.pack(side=tk.LEFT, padx=(0, 5))
        
        perf_action_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        perf_action_inner.pack(fill=tk.X, pady=(10, 0))

        def show_limit_info(m_var=model_var, is_f=is_free):
            info_win = tk.Toplevel(dialog)
            info_win.title("Gemini API 制限と仕様一覧")
            info_win.geometry("950x700") 
            info_win.configure(bg=BG_COLOR)
            info_win.grab_set()
            
            x = dialog.winfo_x() + 30
            y = dialog.winfo_y() + 30
            info_win.geometry(f"+{x}+{y}")
            
            canvas = tk.Canvas(info_win, bg=BG_COLOR, highlightthickness=0)
            scrollbar = ttk.Scrollbar(info_win, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas, style="Main.TFrame")
            
            scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            
            def _on_canvas_config(e):
                canvas.itemconfig(canvas_window, width=e.width)
            canvas.bind("<Configure>", _on_canvas_config)
            
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            lbl_title = ttk.Label(scrollable_frame, text="Gemini API 仕様・制限一覧", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
            lbl_title.pack(pady=(15, 5))
            
            def create_table(parent, headers, data, col_widths):
                table_frame = tk.Frame(parent, bg=BORDER_COLOR) 
                table_frame.pack(fill=tk.X, expand=True, padx=20, pady=5)
                
                for col_idx, w in enumerate(col_widths):
                    table_frame.columnconfigure(col_idx, weight=1, minsize=w)
                
                for col_idx, header_text in enumerate(headers):
                    lbl = tk.Label(table_frame, text=header_text, font=("Segoe UI", 9, "bold"), bg="#E9ECEF", fg=TEXT_COLOR, padx=8, pady=8, wraplength=col_widths[col_idx]-16)
                    lbl.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                
                for row_idx, row_data in enumerate(data, 1):
                    for col_idx, cell_text in enumerate(row_data):
                        lbl = tk.Label(table_frame, text=cell_text, font=("Segoe UI", 9), bg="white", fg=TEXT_COLOR, padx=8, pady=8, justify="left", anchor="nw", wraplength=col_widths[col_idx]-16)
                        lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)
            
            ttk.Label(scrollable_frame, text="▼ プラン比較", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(10, 0))
            
            headers_plan = ["比較項目", "無料枠 (Free Tier)", "課金枠 (Paid Tier)"]
            data_plan = [
                ["利用料金", "完全無料（クレジットカード登録不要）", "従量課金（トークンと呼ばれるデータ量に応じて支払い）"],
                ["利用できるモデル", "2.5 Pro, 2.5 Flash, 3 Flash など", "すべてのモデル（3.1 Pro Previewなども利用可）"],
                ["データの\nプライバシー", "入力データがGoogleのAI学習に利用される可能性がある", "入力データはAI学習に利用されない"]
            ]
            col_widths_plan = [150, 360, 360]
            create_table(scrollable_frame, headers_plan, data_plan, col_widths_plan)

            ttk.Label(scrollable_frame, text="▼ 各モデルの制限目安 (RPM と 推奨スレッド数)", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(15, 0))
            
            headers_limit = ["モデル名", "無料枠の制限目安\n(RPM / スレッド数)", "課金枠の制限目安\n(RPM / スレッド数)"]
            data_limit = [
                ["Gemini 2.5 Pro", "2 RPM, 50 RPD\n[推奨: 2 RPM / 直列(1)]", "360 RPM\n[推奨: 150 RPM / 並列(5〜)]"],
                ["Gemini 2.5 Flash", "15 RPM, 1500 RPD\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                ["Gemini 2.5 Flash-Lite", "15 RPM, 1500 RPD\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                ["Gemini 3.1 Pro Preview\n/ 3.0 Flash Preview", "非常に厳しい (2 RPM未満など)\n[推奨: 1 RPM / 直列(1)]", "時期・モデルにより変動\n[推奨: 150 RPM / 並列(5)]"]
            ]
            col_widths_limit = [200, 335, 335]
            create_table(scrollable_frame, headers_limit, data_limit, col_widths_limit)

            ttk.Label(scrollable_frame, text="▼ 各モデルの特徴と適した用途", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(15, 0))
            
            headers_model = ["モデル名", "特徴", "得意なこと", "適した用途"]
            data_model = [
                ["Gemini 2.5 Pro", "主力・高精度モデル", "複雑な論理的推論、高度なプログラミング、非常に長い文章の文脈理解", "複雑な問題を解かせるAIアシスタント、コード生成・レビュー、大量の資料の要約・分析"],
                ["Gemini 2.5 Flash", "高速・万能モデル", "スピードと性能のバランスが良く、画像・動画・音声の認識（マルチモーダル）にも強い", "一般的なチャットボット、リアルタイム応答、画像内容解析（日常的なAI開発向け）"],
                ["Gemini 2.5 Flash-Lite", "最軽量・低コストモデル", "応答スピードが非常に速く、APIの利用コストが最も安い", "単純なテキスト分類、短い文章の翻訳、大量データを安価に高速処理したい場合"],
                ["Gemini 3.1 Pro Preview\n/ 3.0 Flash Preview", "次世代プレビュー版", "新しいアーキテクチャや最先端の推論能力の提供", "最新鋭のモデルをいち早く試したい開発者向け"]
            ]
            col_widths_model = [160, 140, 270, 290]
            create_table(scrollable_frame, headers_model, data_model, col_widths_model)

            current_plan = "無料枠 (Free Tier)" if is_f else "課金枠 (Paid Tier)"
            current_model = m_var.get()
            status_text = f"【現在、このタブで選択中の設定】\nプラン: {current_plan}　／　モデル: {current_model}"
            
            ttk.Label(scrollable_frame, text=status_text, font=("Segoe UI", 10, "bold"), background=CARD_BG, foreground=PRIMARY, relief="solid", borderwidth=1, padding=10).pack(fill=tk.X, padx=20, pady=15)

            btn_close = ttk.Button(scrollable_frame, text="閉じる", command=info_win.destroy, width=15)
            btn_close.pack(pady=(0, 20))
            
        btn_show_limit = ttk.Button(perf_action_inner, text="ℹ️ 制限と仕様を確認", command=lambda m=model_var, f=is_free: show_limit_info(m, f))
        btn_show_limit.pack(side=tk.LEFT)

        def reset_perf(m_var=model_var, r_var=rpm_var, t_var=threads_var, is_f=is_free):
            model = m_var.get()
            if is_f:
                if "pro" in model: r_var.set(2)
                else: r_var.set(12)
                t_var.set(1) 
            else:
                if "pro" in model: r_var.set(150)
                else: r_var.set(300)
                t_var.set(5) 
                    
        btn_reset_perf = ttk.Button(perf_action_inner, text="🔄 推奨値", command=lambda m=model_var, r=rpm_var, t=threads_var, f=is_free: reset_perf(m, r, t, f))
        btn_reset_perf.pack(side=tk.RIGHT)

        # ③ 抽出パラメータ設定 (右側)
        param_frame = ttk.LabelFrame(middle_frame, text=" ③ AI抽出パラメータ設定 ", style="Card.TLabelframe", padding=8)
        param_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        param_row1 = ttk.Frame(param_frame, style="Card.TFrame")
        param_row1.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(param_row1, text="Temp:", width=6, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        spin_temp = ttk.Spinbox(param_row1, from_=0.0, to=2.0, increment=0.1, textvariable=temp_var, width=4)
        spin_temp.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(param_row1, text="最大トークン:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        spin_tokens = ttk.Spinbox(param_row1, from_=1024, to=2097152, increment=1024, textvariable=tokens_var, width=8)
        spin_tokens.pack(side=tk.LEFT, padx=(0, 5))
        
        param_row2 = ttk.Frame(param_frame, style="Card.TFrame")
        param_row2.pack(fill=tk.X, pady=2)
        chk_safety = ttk.Checkbutton(param_row2, text="安全フィルタ無効化 (エラー回避)", variable=safety_var, style="TCheckbutton")
        chk_safety.pack(side=tk.LEFT)

        def reset_param(t_var=temp_var, tok_var=tokens_var, s_var=safety_var, is_f=is_free):
            t_var.set(0.0)
            tok_var.set(8192)
            s_var.set(True)
            
        btn_reset_param = ttk.Button(param_row2, text="🔄 推奨値", command=lambda t=temp_var, tok=tokens_var, s=safety_var, f=is_free: reset_param(t, tok, s, f))
        btn_reset_param.pack(side=tk.RIGHT, pady=(10, 0))

        # ④ 独自の追加指示 (カスタムプロンプト) - 左右分割リスト型UI
        prompt_frame = ttk.LabelFrame(parent_tab, text=" ④ 独自の追加指示 (カスタムプロンプト) - 任意 ", style="Card.TLabelframe", padding=8)
        prompt_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        input_inner = ttk.Frame(prompt_frame, style="Card.TFrame")
        input_inner.pack(fill=tk.X, pady=(0, 8))
        
        entry_new_prompt = ttk.Entry(input_inner)
        entry_new_prompt.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        entry_new_prompt.bind("<Button-3>", lambda e, widget=entry_new_prompt: show_context_menu(e, widget))
        
        def add_current_prompt(e=None):
            text = entry_new_prompt.get().strip()
            if text:
                current_list.add_item(text)
                entry_new_prompt.delete(0, tk.END)
                sync_current_to_var()

        entry_new_prompt.bind("<Return>", add_current_prompt)
        
        btn_add_prompt = ttk.Button(input_inner, text="＋ 指示を追加", command=add_current_prompt, style="Primary.TButton")
        btn_add_prompt.pack(side=tk.LEFT)

        lists_frame = ttk.Frame(prompt_frame, style="Card.TFrame")
        lists_frame.pack(fill=tk.BOTH, expand=True)
        lists_frame.columnconfigure(0, weight=1)
        lists_frame.columnconfigure(1, weight=1)
        lists_frame.rowconfigure(0, weight=1)

        left_frame = ttk.Frame(lists_frame, style="Card.TFrame")
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        
        ttk.Label(left_frame, text="▼ 現在の抽出に使用する指示", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        current_list = ScrollableCheckboxList(left_frame)
        current_list.pack(fill=tk.BOTH, expand=True, pady=5)
        
        left_actions = ttk.Frame(left_frame, style="Card.TFrame")
        left_actions.pack(fill=tk.X)
        
        def sync_current_to_var():
            prompt_var.set('\n'.join(current_list.get_all_items()))

        def delete_current_selected():
            current_list.remove_selected()
            sync_current_to_var()

        def save_selected_to_fav():
            sel = current_list.get_selected_items()
            if not sel: return
            added = 0
            for text in sel:
                if text not in saved_custom_prompts:
                    saved_custom_prompts.append(text)
                    added += 1
            if added > 0:
                update_all_fav_lists()
                messagebox.showinfo("保存", f"{added}件の指示をお気に入りに保存しました。", parent=dialog)
            else:
                messagebox.showinfo("情報", "選択された指示は既にお気に入りに保存されています。", parent=dialog)

        ttk.Button(left_actions, text="🗑 選択を削除", command=delete_current_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(left_actions, text="⭐ 選択をお気に入りに保存", command=save_selected_to_fav).pack(side=tk.LEFT)

        right_frame = ttk.Frame(lists_frame, style="Card.TFrame")
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        
        ttk.Label(right_frame, text="⭐ お気に入り (よく使う指示)", background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=COLOR_WARNING).pack(anchor="w")
        fav_list = ScrollableCheckboxList(right_frame)
        fav_list.pack(fill=tk.BOTH, expand=True, pady=5)
        fav_lists.append(fav_list)
        
        right_actions = ttk.Frame(right_frame, style="Card.TFrame")
        right_actions.pack(fill=tk.X)

        def add_fav_to_current():
            sel = fav_list.get_selected_items()
            if not sel: return
            for text in sel:
                current_list.add_item(text)
            sync_current_to_var()
            
        def delete_fav_selected():
            sel = fav_list.get_selected_items()
            if not sel: return
            if messagebox.askyesno("確認", "選択したお気に入りを削除しますか？", parent=dialog):
                for text in sel:
                    if text in saved_custom_prompts:
                        saved_custom_prompts.remove(text)
                update_all_fav_lists()

        ttk.Button(right_actions, text="◀ 選択を左に追加", command=add_fav_to_current, style="Primary.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(right_actions, text="🗑 選択を削除", command=delete_fav_selected).pack(side=tk.LEFT)

        initial_prompts = [p for p in prompt_var.get().split('\n') if p.strip()]
        current_list.set_items(initial_prompts)

    build_tab(tab_free, "free")
    build_tab(tab_paid, "paid")
    
    update_all_fav_lists() 
    
    if api_plan_var.get() == "free":
        notebook.select(tab_free)
    else:
        notebook.select(tab_paid)

    btn_action_frame = ttk.Frame(dialog, style="Main.TFrame")
    btn_action_frame.pack(pady=(10, 15))
    
    btn_cancel = ttk.Button(btn_action_frame, text="キャンセル", command=cancel_and_close, width=15)
    btn_cancel.pack(side=tk.LEFT, padx=10)
    
    btn_apply = ttk.Button(btn_action_frame, text="設定を適用して閉じる", command=apply_and_close, style="Primary.TButton", width=25)
    btn_apply.pack(side=tk.LEFT, padx=10)

# ==============================
# クロップ(範囲指定)関連
# ==============================
class CropSelector:
    def __init__(self, master, pdf_path):
        self.top = tk.Toplevel(master)
        self.top.title("抽出範囲の選択 (複数選択可)")
        self.top.configure(bg=BG_COLOR)
        self.top.grab_set()

        self.pdf_path = pdf_path
        self.zoom = 1.0
        
        btn_frame = ttk.Frame(self.top, padding=10)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="クリア", command=self.clear_rects, style="Warning.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="設定して閉じる", command=self.save_and_close, style="Primary.TButton").pack(side=tk.RIGHT, padx=5)
        
        zoom_frame = ttk.Frame(btn_frame)
        zoom_frame.pack(side=tk.RIGHT, padx=20)
        ttk.Button(zoom_frame, text="拡大 (+)", command=self.zoom_in, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="縮小 (-)", command=self.zoom_out, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="フィット", command=self.zoom_fit, width=8).pack(side=tk.LEFT, padx=2)

        ttk.Label(btn_frame, text="【使い方】ドラッグで範囲を選択。Ctrl+ホイール: 拡縮 / Shift+ホイール: 横スクロール", foreground=PRIMARY, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=10)

        canvas_frame = ttk.Frame(self.top)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.vbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        self.hbar.grid(row=1, column=0, sticky="ew")

        self.canvas = tk.Canvas(canvas_frame, cursor="cross", bg="white", xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.vbar.config(command=self.canvas.yview)
        self.hbar.config(command=self.canvas.xview)

        self.start_x, self.start_y, self.current_rect, self.rectangles = None, None, None, []
        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        if sys.platform.startswith("win"): 
            self.canvas.bind("<MouseWheel>", self.on_mousewheel_y)
            self.canvas.bind("<Shift-MouseWheel>", self.on_mousewheel_x)
            self.canvas.bind("<Control-MouseWheel>", self.on_mousewheel_zoom)

        try:
            self.doc = fitz.open(pdf_path)
            self.page = self.doc[0]
            self.zoom_fit()
        except Exception as e:
            self.top.destroy()
            raise Exception(f"プレビュー生成失敗: {e}")

        self.top.update_idletasks()
        try: self.top.state('zoomed')
        except Exception:
            w, h = master.winfo_screenwidth(), master.winfo_screenheight()
            self.top.geometry(f"{w}x{h}+0+0")

    def draw_image(self):
        mat = fitz.Matrix(self.zoom, self.zoom); pix = self.page.get_pixmap(matrix=mat)
        self.tk_image = ImageTk.PhotoImage(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
        self.canvas.delete("all"); self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image); self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.img_w, self.img_h = pix.width, pix.height
        for r in self.rectangles: r['id'] = self.canvas.create_rectangle(r['rx1']*self.img_w, r['ry1']*self.img_h, r['rx2']*self.img_w, r['ry2']*self.img_h, outline="red", width=2)

    def zoom_in(self): self.zoom = min(5.0, self.zoom * 1.2); self.draw_image()
    def zoom_out(self): self.zoom = max(0.2, self.zoom / 1.2); self.draw_image()
    def zoom_fit(self): self.zoom = min(2.0, (self.top.winfo_screenheight() * 0.7) / self.page.rect.height); self.draw_image()
    
    def on_mousewheel_y(self, event): self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    def on_mousewheel_x(self, event): self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
    def on_mousewheel_zoom(self, event):
        if event.delta > 0: self.zoom_in()
        else: self.zoom_out()
        
    def on_press(self, event):
        self.start_x, self.start_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.current_rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2, dash=(4, 4))
    def on_drag(self, event): self.canvas.coords(self.current_rect, self.start_x, self.start_y, self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
    def on_release(self, event):
        end_x, end_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        if abs(end_x - self.start_x) > 10 and abs(end_y - self.start_y) > 10:
            self.canvas.itemconfig(self.current_rect, dash=())
            self.rectangles.append({'id': self.current_rect, 'rx1': min(self.start_x, end_x)/self.img_w, 'ry1': min(self.start_y, end_y)/self.img_h, 'rx2': max(self.start_x, end_x)/self.img_w, 'ry2': max(self.start_y, end_y)/self.img_h})
        else: self.canvas.delete(self.current_rect)
    def clear_rects(self):
        for r in self.rectangles: self.canvas.delete(r['id'])
        self.rectangles.clear()
    def save_and_close(self):
        global selected_crop_regions
        selected_crop_regions = [(r['rx1'], r['ry1'], r['rx2'], r['ry2']) for r in self.rectangles]
        btn_select_crop.config(text=f"抽出範囲を選択 (設定済: {len(selected_crop_regions)}か所)" if selected_crop_regions else "抽出範囲を選択")
        self.doc.close(); self.top.destroy()

def open_crop_selector():
    files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    if not pdf_files: return messagebox.showinfo("情報", "PDFファイルが選択されていません。")
    try: CropSelector(root, pdf_files[0])
    except Exception as e: messagebox.showerror("エラー", str(e))

def reset_crop_regions():
    global selected_crop_regions; selected_crop_regions = []
    btn_select_crop.config(text="抽出範囲を選択")

# ==============================
# UI操作関連関数 (ファイル選択など)
# ==============================
def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("すべての対応ファイル", "*.pdf;*.xlsx;*.xlsm;*.xls;*.csv;*.txt;*.json;*.md;*.docx"), ("PDF", "*.pdf")])
    if files: selected_files, selected_folder, current_mode = list(files), "", "file"; update_ui()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory(title="フォルダを選択")
    if folder: selected_folder, selected_files, current_mode = folder, [], "folder"; update_ui()

def select_save_dir():
    global preset_save_dir
    folder = filedialog.askdirectory()
    if folder: preset_save_dir = folder; save_label.config(text=preset_save_dir); save_option.set(2)

def on_save_mode_change():
    global preset_save_dir; preset_save_dir = ""
    save_label.config(text="同じフォルダ" if save_option.get() == 1 else "未選択")

format_radiobuttons = {}
def toggle_extraction_settings(*args):
    is_active = current_mode is not None
    for fmt, rb in format_radiobuttons.items(): rb.configure(state=tk.NORMAL)
    
    is_gemini = (engine_var.get() == "Gemini")
    state_gemini = tk.NORMAL if is_gemini else tk.DISABLED
    if hasattr(sys.modules[__name__], 'btn_api_settings'):
        btn_api_settings.configure(state=state_gemini)
        
    state_crop = tk.NORMAL if is_active else tk.DISABLED
    for child in crop_frame.winfo_children():
        if isinstance(child, ttk.Button) or isinstance(child, ttk.Label): 
            child.configure(state=state_crop)

def update_ui():
    path_label.config(text="\n".join(selected_files) if current_mode == "file" else (f"フォルダ: {selected_folder}" if selected_folder else "未選択"))
    is_active = current_mode is not None
    state_val = tk.NORMAL if is_active else tk.DISABLED
    btn_split.config(state=state_val); btn_rotate.config(state=state_val); btn_extract.config(state=state_val)
    btn_aggregate_local.config(state=state_val); btn_aggregate_gemini.config(state=state_val)
    btn_merge.config(state=tk.NORMAL if current_mode=="folder" else tk.DISABLED)
    if not is_active: reset_crop_regions()
    toggle_extraction_settings()

def show_text_window(title, content):
    win = tk.Toplevel(root); win.title(title); win.geometry("620x550"); win.configure(bg=BG_COLOR)
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 310; y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 275; win.geometry(f"+{x}+{y}")
    text_area = st.ScrolledText(win, wrap=tk.WORD, font=("Meiryo UI", 10), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.FLAT, padx=15, pady=15)
    text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10); text_area.insert(tk.END, content); text_area.configure(state=tk.DISABLED)

def show_version_info(): 
    messagebox.showinfo("バージョン情報", f"{APP_TITLE}\nバージョン: {VERSION}\n\nPython & Tkinter製 PDF編集ツール")

def show_history(): 
    show_text_window("バージョン履歴", VERSION_HISTORY.strip())

def show_readme():
    p = resource_path("README.md")
    content = open(p, "r", encoding="utf-8").read() if os.path.exists(p) else "README.mdが見つかりません。\nアプリと同じフォルダに配置してください。"
    show_text_window("Readme", content)

# ==============================
# グローバルマウスホイール制御
# ==============================
def _on_mousewheel(event):
    widget = root.winfo_containing(event.x_root, event.y_root)
    if widget:
        current = widget
        while current:
            if isinstance(current, tk.Canvas):
                try:
                    current.yview_scroll(int(-1*(event.delta/120)), "units")
                except: pass
                return
            current = current.master

# ==============================
# UI画面の構築 (レスポンシブ設計)
# ==============================
root = tk.Tk(); root.title(f"{APP_TITLE} {VERSION}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+0+0")
root.minsize(width=760, height=650) 
root.configure(bg=BG_COLOR)
root.bind_all("<MouseWheel>", _on_mousewheel)

style = ttk.Style(); style.theme_use("clam") if "clam" in style.theme_names() else None
style.configure(".", background=BG_COLOR, font=("Segoe UI", 10))
style.configure("Main.TFrame", background=BG_COLOR)
style.configure("Card.TFrame", background=CARD_BG)
style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1, bordercolor=BORDER_COLOR)
style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=PRIMARY, font=("Segoe UI", 11, "bold"))
style.configure("TButton", padding=6, font=("Segoe UI", 10), background="#E9ECEF", foreground=TEXT_COLOR, borderwidth=1)
style.map("TButton", background=[("active", "#DEE2E6")])
style.configure("Primary.TButton", background=PRIMARY, foreground="white", borderwidth=0)
style.map("Primary.TButton", background=[("active", PRIMARY_HOVER)])
style.configure("Warning.TButton", background=COLOR_WARNING, foreground="black", borderwidth=0)
style.map("Warning.TButton", background=[("active", COLOR_WARNING_HOVER)])
style.configure("Danger.TButton", background=COLOR_DANGER, foreground="white", borderwidth=0)
style.map("Danger.TButton", background=[("active", COLOR_DANGER_HOVER)])
style.configure("Purple.TButton", background=COLOR_PURPLE, foreground="white", borderwidth=0)
style.map("Purple.TButton", background=[("active", COLOR_PURPLE_HOVER)])
style.configure("TRadiobutton", background=CARD_BG, font=("Segoe UI", 10), foreground=TEXT_COLOR)
style.configure("TCheckbutton", background=CARD_BG, font=("Segoe UI", 10), foreground=TEXT_COLOR)

menubar = Menu(root)
help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="AI抽出の準備 (使い方)", command=lambda: show_text_window("AI抽出の準備 (使い方)", AI_HELP_TEXT.strip()))
help_menu.add_separator()
help_menu.add_command(label="Readmeを表示", command=show_readme)
help_menu.add_command(label="バージョン履歴", command=show_history)
help_menu.add_separator()
help_menu.add_command(label="バージョン情報", command=show_version_info)
menubar.add_cascade(label="ヘルプ", menu=help_menu)
root.config(menu=menubar)

rotate_option, save_option = tk.IntVar(value=270), tk.IntVar(value=1)
engine_var, output_format_var = tk.StringVar(value="Internal"), tk.StringVar(value="xlsx")

api_plan_var = tk.StringVar(value="free")
api_key_free_var = tk.StringVar(value=get_api_key() or "")
api_key_paid_var = tk.StringVar(value=get_api_key() or "")
gemini_model_free_var = tk.StringVar(value="gemini-2.5-flash")
gemini_model_paid_var = tk.StringVar(value="gemini-2.5-flash")
api_rpm_free_var = tk.IntVar(value=12)
api_rpm_paid_var = tk.IntVar(value=300)

temperature_free_var = tk.DoubleVar(value=0.0)
temperature_paid_var = tk.DoubleVar(value=0.0)
safety_free_var = tk.BooleanVar(value=True)
safety_paid_var = tk.BooleanVar(value=True)
max_tokens_free_var = tk.IntVar(value=8192)
max_tokens_paid_var = tk.IntVar(value=8192)
custom_prompt_free_var = tk.StringVar(value="")
custom_prompt_paid_var = tk.StringVar(value="")
threads_free_var = tk.IntVar(value=1)
threads_paid_var = tk.IntVar(value=5)

main_outer = ttk.Frame(root)
main_outer.pack(fill=tk.BOTH, expand=True)

canvas = tk.Canvas(main_outer, bg=BG_COLOR, highlightthickness=0)
scrollbar = ttk.Scrollbar(main_outer, orient=tk.VERTICAL, command=canvas.yview)

main_container = ttk.Frame(canvas, padding=10, style="Main.TFrame")
canvas_frame_id = canvas.create_window((0, 0), window=main_container, anchor="nw")

def on_canvas_configure(event): canvas.itemconfig(canvas_frame_id, width=event.width)
canvas.bind('<Configure>', on_canvas_configure)

def on_frame_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
main_container.bind('<Configure>', on_frame_configure)

canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# --- UI コンポーネント群 ---
title_frame = ttk.Frame(main_container, style="Main.TFrame"); title_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(title_frame, text=APP_TITLE, font=("Segoe UI", 20, "bold"), foreground=PRIMARY, background=BG_COLOR).pack(side=tk.LEFT)
ttk.Label(title_frame, text=f" {VERSION}", font=("Segoe UI", 12), foreground=MUTED_TEXT, background=BG_COLOR).pack(side=tk.LEFT, pady=(8, 0))

file_card = ttk.Frame(main_container, style="Card.TFrame", padding=5); file_card.pack(fill=tk.X, pady=2)
btn_frame = ttk.Frame(file_card, style="Card.TFrame"); btn_frame.pack()
ttk.Button(btn_frame, text="📄 ファイルを選択", command=select_files, width=22, style="Primary.TButton").grid(row=0, column=0, padx=8)
ttk.Button(btn_frame, text="📁 フォルダを選択", command=select_folder, width=22, style="Primary.TButton").grid(row=0, column=1, padx=8)
path_label = ttk.Label(file_card, text="未選択", background=CARD_BG, foreground=TEXT_COLOR, wraplength=580, justify="center"); path_label.pack(pady=(5, 0))

settings_grid = ttk.Frame(main_container, style="Main.TFrame"); settings_grid.pack(fill=tk.X, pady=2); settings_grid.columnconfigure(0, weight=1); settings_grid.columnconfigure(1, weight=1)
save_frame = ttk.LabelFrame(settings_grid, text=" 保存先設定 ", style="Card.TLabelframe", padding=5); save_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
ttk.Radiobutton(save_frame, text="元のファイルと同じフォルダ", variable=save_option, value=1, command=on_save_mode_change).pack(anchor="w", pady=2)
ttk.Radiobutton(save_frame, text="任意のフォルダを指定", variable=save_option, value=2, command=on_save_mode_change).pack(anchor="w", pady=2)
ttk.Button(save_frame, text="📂 フォルダ参照", command=select_save_dir).pack(pady=(4, 2))
save_label = ttk.Label(save_frame, text="同じフォルダ", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 9)); save_label.pack()

rotate_frame = ttk.LabelFrame(settings_grid, text=" 回転設定 ", style="Card.TLabelframe", padding=5); rotate_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
for t, v in [("左（270°）", 270), ("上下（180°）", 180), ("右（90°）", 90)]: ttk.Radiobutton(rotate_frame, text=t, variable=rotate_option, value=v).pack(anchor="w", pady=2)

extract_frame = ttk.LabelFrame(main_container, text=" ⚙️ データ抽出・変換設定 ", style="Card.TLabelframe", padding=5); extract_frame.pack(fill=tk.X, pady=5)

engine_frame = ttk.Frame(extract_frame, style="Card.TFrame"); engine_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(engine_frame, text="① エンジン:", width=12, background=CARD_BG, font=("Segoe UI", 10, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
engine_inner = ttk.Frame(engine_frame, style="Card.TFrame"); engine_inner.pack(anchor="w", fill=tk.X)
for text, val in [("Python標準ライブラリ (高速・オフライン)", "Internal"), ("Tesseract (ローカルOCR)", "Tesseract"), ("Gemini API (超高精度AI)", "Gemini")]:
    ttk.Radiobutton(engine_inner, text=text, variable=engine_var, value=val).pack(anchor="w", pady=1)

ttk.Separator(extract_frame, orient="horizontal").pack(fill=tk.X, pady=6)

format_frame = ttk.Frame(extract_frame, style="Card.TFrame"); format_frame.pack(fill=tk.X, pady=0)
ttk.Label(format_frame, text="② 出力形式:", width=12, background=CARD_BG, font=("Segoe UI", 10, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)

formats_row1 = [("Excel (.xlsx)", "xlsx"), ("CSV (.csv)", "csv"), ("Text (.txt)", "txt"), ("JSON (.json)", "json"), ("Markdown (.md)", "md"), ("Word (.docx)", "docx")]
formats_row2 = [("JPEG (.jpg)", "jpg"), ("PNG (.png)", "png"), ("SVG (.svg)", "svg"), ("TIFF (.tiff)", "tiff"), ("BMP (.bmp)", "bmp"), ("DXF (.dxf)", "dxf")]

format_inner1 = ttk.Frame(format_frame, style="Card.TFrame"); format_inner1.pack(anchor="w", fill=tk.X)
for text, val in formats_row1:
    rb = ttk.Radiobutton(format_inner1, text=text, variable=output_format_var, value=val); rb.pack(side=tk.LEFT, padx=(0, 10)); format_radiobuttons[val] = rb
format_inner2 = ttk.Frame(format_frame, style="Card.TFrame"); format_inner2.pack(anchor="w", fill=tk.X, pady=(2, 0))
for text, val in formats_row2:
    rb = ttk.Radiobutton(format_inner2, text=text, variable=output_format_var, value=val); rb.pack(side=tk.LEFT, padx=(0, 10)); format_radiobuttons[val] = rb

ttk.Separator(extract_frame, orient="horizontal").pack(fill=tk.X, pady=6)

api_settings_frame = ttk.Frame(extract_frame, style="Card.TFrame")
api_settings_frame.pack(fill=tk.X, pady=2)
ttk.Label(api_settings_frame, text="[AI用] API設定:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
btn_api_settings = ttk.Button(api_settings_frame, text="⚙️ 詳細設定 (APIキー / モデル / 制限)", command=open_api_settings_dialog, style="Primary.TButton")
btn_api_settings.pack(side=tk.LEFT)

crop_frame = ttk.Frame(extract_frame, style="Card.TFrame"); crop_frame.pack(fill=tk.X, pady=(5, 0))
ttk.Label(crop_frame, text="抽出範囲:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
btn_select_crop = ttk.Button(crop_frame, text="抽出範囲を選択", command=open_crop_selector); btn_select_crop.pack(side=tk.LEFT)
btn_reset_crop = ttk.Button(crop_frame, text="全体に戻す", command=reset_crop_regions, style="Warning.TButton"); btn_reset_crop.pack(side=tk.LEFT, padx=(5, 0))

save_settings_frame = ttk.Frame(extract_frame, style="Card.TFrame")
save_settings_frame.pack(fill=tk.X, pady=(10, 0))
btn_save_settings = ttk.Button(save_settings_frame, text="💾 現在の選択項目を保存", command=save_settings)
btn_save_settings.pack(side=tk.RIGHT)

action_container = ttk.Frame(main_container, style="Main.TFrame"); action_container.pack(fill=tk.BOTH, expand=True, pady=2)
action_container.columnconfigure(0, weight=1); action_container.columnconfigure(1, weight=1)

pdf_action_frame = ttk.LabelFrame(action_container, text=" ✂️ PDF編集 ", style="Card.TLabelframe", padding=5); pdf_action_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
btn_merge = ttk.Button(pdf_action_frame, text="結合 (フォルダ)", command=lambda: safe_run(merge_pdfs, "PDF結合"), style="Primary.TButton"); btn_merge.pack(fill=tk.X, pady=2)
btn_split = ttk.Button(pdf_action_frame, text="分割", command=lambda: safe_run(split_pdfs, "PDF分割"), style="Primary.TButton"); btn_split.pack(fill=tk.X, pady=2)
btn_rotate = ttk.Button(pdf_action_frame, text="回転", command=lambda: safe_run(rotate_pdfs, "PDF回転"), style="Primary.TButton"); btn_rotate.pack(fill=tk.X, pady=2)

data_action_frame = ttk.LabelFrame(action_container, text=" 📊 データ操作 ", style="Card.TLabelframe", padding=5); data_action_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
btn_extract = ttk.Button(data_action_frame, text="🚀 選択した抽出・変換を実行", command=run_selected_extraction, style="Danger.TButton"); btn_extract.pack(fill=tk.X, pady=(2, 6), ipady=6) 

aggregate_btn_frame = ttk.Frame(data_action_frame, style="Card.TFrame")
aggregate_btn_frame.pack(fill=tk.X, pady=2)
aggregate_btn_frame.columnconfigure(0, weight=1)
aggregate_btn_frame.columnconfigure(1, weight=1)

btn_aggregate_local = ttk.Button(aggregate_btn_frame, text="🧩 ローカルで集約", command=lambda: safe_run(aggregate_local_task, "ローカルデータ集約"), style="Purple.TButton")
btn_aggregate_local.grid(row=0, column=0, sticky="nsew", padx=(0, 2))

btn_aggregate_gemini = ttk.Button(aggregate_btn_frame, text="✨ Geminiで集約", command=lambda: safe_run(aggregate_gemini_task, "Geminiデータ集約"), style="Purple.TButton")
btn_aggregate_gemini.grid(row=0, column=1, sticky="nsew", padx=(2, 0))

status_frame = ttk.Frame(main_container, style="Main.TFrame")
status_frame.pack(fill=tk.X, pady=(2, 0))
status_label = ttk.Label(status_frame, text="ステータス: 待機中", font=("Segoe UI", 10), foreground=MUTED_TEXT, background=BG_COLOR)
status_label.pack(side=tk.LEFT, padx=5)

# 関数・コンポーネント定義がすべて終わった後で Trace を設定する
engine_var.trace("w", toggle_extraction_settings)

load_settings()
update_ui()

root.mainloop()