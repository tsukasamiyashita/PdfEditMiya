# -*- coding: utf-8 -*-
import os, sys, threading, re, json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter.scrolledtext as st
import fitz
from PIL import Image, ImageTk
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
APP_TITLE, VERSION = "PdfEditMiya", "v2.0.0"
WINDOW_WIDTH, WINDOW_HEIGHT = 900, 750  # すべてのボタンが収まる適切な高さに調整

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
SETTINGS_FILE = os.path.join(USER_HOME, ".pdfeditmiya_settings.json") # 設定保存用ファイル

# ==============================
# ヘルプ・履歴テキスト
# ==============================
VERSION_HISTORY = """
[ v2.0.0 ]
- 【UI改善】初期起動時にすべてのボタンが確実に表示されるよう、ウィンドウサイズを最適化しました。
- 【UI改善】「①エンジン」と「②出力形式」の文字がラジオボタンと重ならないようにレイアウトを修正しました。
- 【機能追加】処理中に安全に停止できる「処理を中止」ボタンを追加しました。
- 【機能追加】メイン画面下部に現在の処理状況がわかる「ステータス表示」を追加しました。
- 【UI改善】「①エンジン」と「②出力形式」の間に区切り線を追加し、設定項目の境界を分かりやすくしました。
- 【高速化】Gemini APIの課金枠選択時に最大10スレッドの完全な並行処理を実現し、劇的な処理速度の向上を達成しました。
- 【UI改善】メイン画面全体にスクロール機能を実装し、レスポンシブに対応しました。
- 【機能追加】Gemini API利用時に「無料枠」と「課金枠」のプラン選択機能を追加しました。
- 【UI改善】入力したAPIキーを表示して確認・コピーできるトグルボタンと、右クリックでの「貼り付け(ペースト)」に対応しました。
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
6. 本アプリの「APIキー」入力枠内で右クリックして貼り付け、「テスト」ボタンを押してください。

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

# ==============================
# 設定の保存・読み込み機能
# ==============================
def load_settings():
    global preset_save_dir
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
            
            if "rotate_option" in settings: rotate_option.set(settings["rotate_option"])
            if "engine_var" in settings: engine_var.set(settings["engine_var"])
            if "output_format_var" in settings: output_format_var.set(settings["output_format_var"])
            if "api_key_var" in settings and settings["api_key_var"]: api_key_var.set(settings["api_key_var"])
            if "api_plan_var" in settings: api_plan_var.set(settings["api_plan_var"])
            
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
            
            # ウィンドウサイズの復元 (最小サイズを下回らないように保護)
            if "window_width" in settings and "window_height" in settings:
                w = max(settings["window_width"], 760)
                h = max(settings["window_height"], 650)
                root.geometry(f"{w}x{h}")
                
        except Exception as e:
            print(f"Failed to load settings: {e}")

def save_settings():
    # 最新のウィンドウサイズを正しく取得するために更新
    root.update_idletasks()
    
    settings = {
        "rotate_option": rotate_option.get(),
        "save_option": save_option.get(),
        "preset_save_dir": preset_save_dir,
        "engine_var": engine_var.get(),
        "output_format_var": output_format_var.get(),
        "api_key_var": api_key_var.get().strip(),
        "api_plan_var": api_plan_var.get(),
        "window_width": root.winfo_width(),
        "window_height": root.winfo_height()
    }
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
            
        # 既存のAPIキーファイルにも書き込んでおく（互換性のため）
        if settings["api_key_var"]:
            with open(API_KEY_FILE, "w", encoding="utf-8") as f:
                f.write(settings["api_key_var"])
                
        messagebox.showinfo("保存完了", "現在の選択項目を保存しました。\n次回起動時もこの設定が適用されます。")
    except Exception as e:
        messagebox.showerror("エラー", f"設定の保存に失敗しました。\n{e}")

# ==============================
# UI コントローラー & ヘルパー関数
# ==============================
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

def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    return None

def toggle_api_key_show():
    if api_key_entry.cget('show') == '*':
        api_key_entry.configure(show='')
        btn_api_check.configure(text="隠す")
    else:
        api_key_entry.configure(show='*')
        btn_api_check.configure(text="確認")

def show_context_menu(event):
    menu = Menu(root, tearoff=0)
    menu.add_command(label="貼り付け", command=lambda: paste_to_entry(event.widget))
    menu.post(event.x_root, event.y_root)

def paste_to_entry(widget):
    try:
        text = root.clipboard_get()
        try: widget.delete("sel.first", "sel.last")
        except tk.TclError: pass
        widget.insert(tk.INSERT, text)
    except tk.TclError: pass

def test_api_key_ui():
    key = api_key_var.get().strip()
    if not key: return messagebox.showwarning("警告", "APIキーが入力されていません。")
    
    genai.configure(api_key=key)
    model_name = gemini_model_var.get()
    
    try:
        model = genai.GenerativeModel(model_name)
        model.generate_content("Test")
        
        with open(API_KEY_FILE, "w", encoding="utf-8") as f: 
            f.write(key)
        messagebox.showinfo("テスト成功", f"APIキーは正しく認識されました。\nAIモデル「{model_name}」による通信は正常です！")
        
    except Exception as e:
        err_str = str(e).lower()
        if "404" in err_str or "not found" in err_str:
            messagebox.showerror("モデル利用不可", f"エラー: モデル「{model_name}」が存在しないか、利用する権限がありません。\n\n詳細:\n{e}")
        elif "429" in err_str or "quota" in err_str:
            msg = f"エラー: APIの利用枠（クォータ）を超過しています。\n\n"
            m = re.search(r'retry in ([\d\.]+)s', err_str)
            if not m: m = re.search(r'seconds:\s*(\d+)', err_str)
            if m:
                wait_sec = int(float(m.group(1)))
                msg += f"⚠️ Googleのバースト制限です。約 {wait_sec} 秒後に利用枠が回復します。\n表示された秒数だけ待機してから再度テストしてください。\n\n"
            else:
                if "perday" in err_str.lower(): msg += "⚠️ 【1日の利用上限】に達した可能性があります。\n翌日になるまで待つか、課金設定（Paid Tier）を確認してください。\n\n"
                else: msg += "⚠️ APIの無料枠制限に達しました。\n約1分ほど待ってから再度テストするか、課金設定をご確認ください。\n\n"
            msg += f"詳細:\n{e}"
            messagebox.showerror("利用枠超過", msg)
        else:
            messagebox.showerror("通信エラー", f"APIキーまたは通信に問題が発生しました。\n\n詳細:\n{e}")

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
        model_name = gemini_model_var.get()
        engine_text = f"使用エンジン: Gemini API ( {model_name} )"
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
        
        files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
        if not files: return
        save_dir = os.path.dirname(files[0]) if save_option.get() == 1 else preset_save_dir
        options = {
            "rotate_deg": rotate_option.get(), "crop_regions": selected_crop_regions, "out_format": output_format_var.get(),
            "folder_name": os.path.basename(selected_folder) if selected_folder else "Merged",
            "api_key": api_key_var.get().strip(),
            "models_to_try": [gemini_model_var.get()] if engine_var.get() == "Gemini" else [],
            "api_plan": api_plan_var.get()
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
    files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
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
        if not api_key_var.get().strip(): return messagebox.showerror("エラー", "Gemini APIキーを入力してください。")
        safe_run(extract_gemini_task, task_name)

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
    files = selected_files if current_mode == "file" else ([os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".csv", ".txt", ".json", ".md", ".docx"))] if selected_folder else [])
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    if not pdf_files: return messagebox.showinfo("情報", "PDFファイルが選択されていません。")
    try: CropSelector(root, pdf_files[0])
    except Exception as e: messagebox.showerror("エラー", str(e))

def reset_crop_regions():
    global selected_crop_regions; selected_crop_regions = []
    btn_select_crop.config(text="抽出範囲を選択")

def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(filetypes=[("すべての対応ファイル", "*.pdf;*.xlsx;*.csv;*.txt;*.json;*.md;*.docx"), ("PDF", "*.pdf")])
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
    for fmt, rb in format_radiobuttons.items(): rb.configure(state=tk.NORMAL if is_active else tk.DISABLED)
    
    is_gemini = (is_active and engine_var.get() == "Gemini")
    state_gemini = tk.NORMAL if is_gemini else tk.DISABLED
    
    api_key_entry.configure(state=state_gemini)
    btn_api_check.configure(state=state_gemini)
    btn_api_test.configure(state=state_gemini)
    
    for child in api_plan_frame.winfo_children():
        if isinstance(child, ttk.Radiobutton) or isinstance(child, ttk.Label): 
            child.configure(state=state_gemini)
        
    state_crop = tk.NORMAL if is_active else tk.DISABLED
    for child in crop_frame.winfo_children():
        if isinstance(child, ttk.Button) or isinstance(child, ttk.Label): child.configure(state=state_crop)

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
# UI画面の構築 (レスポンシブ設計)
# ==============================
root = tk.Tk(); root.title(f"{APP_TITLE} {VERSION}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+0+0")
root.minsize(width=760, height=650) # ウィンドウの最小サイズを引き上げ、見切れを防止
root.configure(bg=BG_COLOR)

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
api_key_var = tk.StringVar(value=get_api_key() or "")
gemini_model_var = tk.StringVar(value="gemini-2.5-flash")
api_plan_var = tk.StringVar(value="free")

engine_var.trace("w", toggle_extraction_settings)

# 全体を覆うスクロール可能なキャンバス領域の構築
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

def _on_mousewheel(event):
    widget = root.winfo_containing(event.x_root, event.y_root)
    if widget and widget.winfo_toplevel() == root:
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
root.bind_all("<MouseWheel>", _on_mousewheel)

canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# --- 以下、既存のUIコンポーネント群を main_container 内に配置 ---
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

api_key_frame = ttk.Frame(extract_frame, style="Card.TFrame"); api_key_frame.pack(fill=tk.X, pady=2)
ttk.Label(api_key_frame, text="[AI用] APIキー:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
api_key_entry = ttk.Entry(api_key_frame, textvariable=api_key_var, width=45, show="*"); api_key_entry.pack(side=tk.LEFT, padx=(0, 8))
api_key_entry.bind("<Button-3>", show_context_menu)
btn_api_check = ttk.Button(api_key_frame, text="確認", command=toggle_api_key_show, width=6); btn_api_check.pack(side=tk.LEFT, padx=(0, 5))
btn_api_test = ttk.Button(api_key_frame, text="テスト", command=test_api_key_ui, width=6); btn_api_test.pack(side=tk.LEFT)

api_plan_frame = ttk.Frame(extract_frame, style="Card.TFrame"); api_plan_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(api_plan_frame, text="APIプラン:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
rb_free = ttk.Radiobutton(api_plan_frame, text="無料枠 (安全な通信間隔で実行)", variable=api_plan_var, value="free"); rb_free.pack(side=tk.LEFT, padx=(0, 15))
rb_paid = ttk.Radiobutton(api_plan_frame, text="課金枠 (制限なしで高速実行)", variable=api_plan_var, value="paid"); rb_paid.pack(side=tk.LEFT)

crop_frame = ttk.Frame(extract_frame, style="Card.TFrame"); crop_frame.pack(fill=tk.X, pady=(2, 0))
ttk.Label(crop_frame, text="抽出範囲:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
btn_select_crop = ttk.Button(crop_frame, text="抽出範囲を選択", command=open_crop_selector); btn_select_crop.pack(side=tk.LEFT)
btn_reset_crop = ttk.Button(crop_frame, text="全体に戻す", command=reset_crop_regions, style="Warning.TButton"); btn_reset_crop.pack(side=tk.LEFT, padx=(5, 0))

# --- 設定保存ボタンの追加 ---
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

# --- 起動時の設定読み込みとUI更新 ---
load_settings()
update_ui()

root.mainloop()