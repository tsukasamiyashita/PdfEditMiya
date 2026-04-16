# -*- coding: utf-8 -*-
import os, sys, json, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import tkinter.scrolledtext as st

from common import *
from dialogs import open_api_settings_dialog, open_crop_selector, reset_crop_regions, show_pdf_type_info
from gemini_engine import extract_gemini_task
from engines import (
    merge_pdfs, split_pdfs, rotate_pdfs,
    extract_text_internal, convert_to_excel_internal, convert_to_csv_internal,
    convert_to_image_jpg, convert_to_image_png, convert_to_dxf,
    convert_to_image_tiff, convert_to_image_bmp, convert_to_svg,
    extract_tesseract_task, aggregate_local_task, combine_local_task
)

# ==============================
# プログレス関連のローカル状態・タスク管理機能
# ==============================
processing_popup = None
overall_label = None
overall_progress = None
file_label = None
file_progress = None
btn_cancel = None
rb_table_mode = None
rb_text_mode = None

def show_message(msg, color=PRIMARY):
    def _task():
        win = tk.Toplevel(state.root)
        win.geometry("260x90"); win.configure(bg=CARD_BG); win.attributes("-topmost", True)
        x = state.root.winfo_x() + (WINDOW_WIDTH // 2) - 130; y = state.root.winfo_y() + (WINDOW_HEIGHT // 2) - 45
        win.geometry(f"+{x}+{y}"); win.overrideredirect(True)
        frame = tk.Frame(win, bg=CARD_BG, highlightbackground=color, highlightthickness=2); frame.pack(expand=True, fill=tk.BOTH)
        ttk.Label(frame, text=msg, foreground=color, font=("Segoe UI", 10, "bold"), background=CARD_BG, wraplength=240).pack(expand=True)
        win.after(2500, win.destroy)
    state.root.after(0, _task)

class UIController:
    def update_overall(self, step, max_val=None, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                if max_val is not None: overall_progress["maximum"] = max_val
                overall_progress["value"] = step
                if text: overall_label.config(text=text)
        state.root.after(0, _task)
    def set_indeterminate(self, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                file_progress.config(mode="indeterminate"); file_progress.start(15)
                if text: file_label.config(text=text)
        state.root.after(0, _task)
    def set_determinate(self, step, max_val=None, text=None):
        def _task():
            if processing_popup and processing_popup.winfo_exists():
                file_progress.stop(); file_progress.config(mode="determinate")
                if max_val is not None: file_progress["maximum"] = max_val
                file_progress["value"] = step
                if text: file_label.config(text=text)
        state.root.after(0, _task)
    def is_cancelled(self):
        return state.cancelled

def show_processing(total_files=1):
    global processing_popup, overall_label, overall_progress, file_label, file_progress, btn_cancel
    processing_popup = tk.Toplevel(state.root)
    processing_popup.title("処理を実行中...")
    processing_popup.geometry("440x280")
    processing_popup.configure(bg=CARD_BG)
    processing_popup.grab_set()
    x = state.root.winfo_x() + (WINDOW_WIDTH // 2) - 220
    y = state.root.winfo_y() + (WINDOW_HEIGHT // 2) - 140
    processing_popup.geometry(f"+{x}+{y}")
    
    engine_name = state.engine_var.get()
    if engine_name == "Gemini":
        plan = state.api_plan_var.get()
        model_name = state.gemini_model_free_var.get() if plan == "free" else state.gemini_model_paid_var.get()
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
        if messagebox.askyesno("確認", "処理を中止しますか？", parent=processing_popup):
            state.cancelled = True
            btn_cancel.config(text="中止処理中...", state=tk.DISABLED)
            
    btn_cancel = ttk.Button(processing_popup, text="処理を中止", command=cancel_processing, style="Warning.TButton")
    btn_cancel.pack(pady=(5, 10))

def close_processing():
    def _task():
        global processing_popup
        if processing_popup: processing_popup.destroy(); processing_popup = None
    state.root.after(0, _task)

def run_task(func, task_name):
    state.cancelled = False
    try:
        def _start_status(): 
            if state.status_label: state.status_label.config(text=f"ステータス: {task_name} を実行中...", foreground=PRIMARY)
        state.root.after(0, _start_status)
        
        files = state.selected_files if state.current_mode == "file" else ([os.path.join(state.selected_folder, f) for f in os.listdir(state.selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if state.selected_folder else [])
        if not files: return
        save_dir = os.path.dirname(files[0]) if state.save_option.get() == 1 else state.preset_save_dir
        
        plan = state.api_plan_var.get()
        is_free = (plan == "free")
        api_key = state.api_key_free_var.get().strip() if is_free else state.api_key_paid_var.get().strip()
        model = state.gemini_model_free_var.get() if is_free else state.gemini_model_paid_var.get()
        rpm = state.api_rpm_free_var.get() if is_free else state.api_rpm_paid_var.get()
        
        options = {
            "rotate_deg": state.rotate_option.get(), "crop_regions": state.selected_crop_regions, "out_format": state.output_format_var.get(),
            "folder_name": os.path.basename(state.selected_folder) if state.selected_folder else "Merged",
            "api_key": api_key,
            "models_to_try": [model] if state.engine_var.get() == "Gemini" else [],
            "api_plan": plan,
            "api_rpm": rpm,
            "temperature": state.temperature_free_var.get() if is_free else state.temperature_paid_var.get(),
            "disable_safety": state.safety_free_var.get() if is_free else state.safety_paid_var.get(),
            "max_tokens": state.max_tokens_free_var.get() if is_free else state.max_tokens_paid_var.get(),
            "custom_prompt": state.custom_prompt_free_var.get() if is_free else state.custom_prompt_paid_var.get(),
            "threads": state.threads_free_var.get() if is_free else state.threads_paid_var.get(),
            "extract_mode": state.extract_mode_var.get()
        }
        func(files, save_dir, options, UIController())
        close_processing()
        
        def _end_status():
            if state.cancelled:
                show_message("⚠️ 処理が中止されました", COLOR_WARNING)
                if state.status_label: state.status_label.config(text=f"ステータス: {task_name} は中止されました", foreground=COLOR_WARNING)
            else:
                show_message("✅ 処理が完了しました", SUCCESS)
                if state.status_label: state.status_label.config(text=f"ステータス: {task_name} が完了しました", foreground=SUCCESS)
        state.root.after(0, _end_status)
        
    except Exception as e:
        print(f"Error: {e}"); close_processing()
        err_msg = str(e)
        def _err_status(msg=err_msg):
            messagebox.showerror("実行エラー", f"{task_name} 中にエラーが発生しました。\n\n詳細:\n{msg}", parent=state.root)
            if state.status_label: state.status_label.config(text=f"ステータス: {task_name} でエラーが発生しました", foreground=ERROR)
        state.root.after(100, _err_status)

def safe_run(func, task_name="処理"):
    files = state.selected_files if state.current_mode == "file" else ([os.path.join(state.selected_folder, f) for f in os.listdir(state.selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if state.selected_folder else [])
    if not files: return
    if state.save_option.get() == 2 and not state.preset_save_dir:
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        if not folder: return
        state.preset_save_dir = folder
        if state.save_label: state.save_label.config(text=state.preset_save_dir)
    show_processing(len(files))
    threading.Thread(target=run_task, args=(func, task_name), daemon=True).start()

def run_selected_extraction():
    engine = state.engine_var.get(); fmt = state.output_format_var.get()
    
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
        plan = state.api_plan_var.get()
        api_key = state.api_key_free_var.get().strip() if plan == "free" else state.api_key_paid_var.get().strip()
        if not api_key: 
            return messagebox.showerror("エラー", f"Gemini APIキー({plan.capitalize()}枠用)を入力してください。\n「⚙️ 詳細設定」ボタンから設定できます。")
        safe_run(extract_gemini_task, task_name)

# ==============================
# 設定の保存・読み込み機能
# ==============================
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
            
            if "rotate_option" in settings: state.rotate_option.set(settings["rotate_option"])
            if "engine_var" in settings: state.engine_var.set(settings["engine_var"])
            if "output_format_var" in settings: state.output_format_var.set(settings["output_format_var"])
            if "extract_mode_var" in settings: state.extract_mode_var.set(settings["extract_mode_var"])
            if "api_plan_var" in settings: state.api_plan_var.set(settings["api_plan_var"])
            
            if "api_key_free_var" in settings: state.api_key_free_var.set(settings["api_key_free_var"])
            elif "api_key_var" in settings: state.api_key_free_var.set(settings["api_key_var"])
            
            if "api_key_paid_var" in settings: state.api_key_paid_var.set(settings["api_key_paid_var"])
            elif "api_key_var" in settings: state.api_key_paid_var.set(settings["api_key_var"])

            if "gemini_model_free_var" in settings: state.gemini_model_free_var.set(settings["gemini_model_free_var"])
            elif "gemini_model_var" in settings: state.gemini_model_free_var.set(settings["gemini_model_var"])

            if "gemini_model_paid_var" in settings: state.gemini_model_paid_var.set(settings["gemini_model_paid_var"])
            elif "gemini_model_var" in settings: state.gemini_model_paid_var.set(settings["gemini_model_var"])

            if "api_rpm_free_var" in settings: state.api_rpm_free_var.set(settings["api_rpm_free_var"])
            if "api_rpm_paid_var" in settings: state.api_rpm_paid_var.set(settings["api_rpm_paid_var"])

            if "temperature_free_var" in settings: state.temperature_free_var.set(settings["temperature_free_var"])
            if "temperature_paid_var" in settings: state.temperature_paid_var.set(settings["temperature_paid_var"])
            if "safety_free_var" in settings: state.safety_free_var.set(settings["safety_free_var"])
            if "safety_paid_var" in settings: state.safety_paid_var.set(settings["safety_paid_var"])
            if "max_tokens_free_var" in settings: state.max_tokens_free_var.set(settings["max_tokens_free_var"])
            if "max_tokens_paid_var" in settings: state.max_tokens_paid_var.set(settings["max_tokens_paid_var"])
            if "custom_prompt_free_var" in settings: state.custom_prompt_free_var.set(settings["custom_prompt_free_var"])
            if "custom_prompt_paid_var" in settings: state.custom_prompt_paid_var.set(settings["custom_prompt_paid_var"])
            if "threads_free_var" in settings: state.threads_free_var.set(settings["threads_free_var"])
            if "threads_paid_var" in settings: state.threads_paid_var.set(settings["threads_paid_var"])
            
            if "saved_custom_prompts" in settings:
                state.saved_custom_prompts = settings["saved_custom_prompts"]

            if "save_option" in settings: 
                state.save_option.set(settings["save_option"])
                if settings["save_option"] == 1:
                    state.save_label.config(text="同じフォルダ")
                elif settings["save_option"] == 2:
                    if "preset_save_dir" in settings and settings["preset_save_dir"]:
                        state.preset_save_dir = settings["preset_save_dir"]
                        state.save_label.config(text=state.preset_save_dir)
                    else:
                        state.save_label.config(text="未選択")
            
            if "window_width" in settings and "window_height" in settings:
                w = max(settings["window_width"], 720)
                h = max(settings["window_height"], 580)
                root.geometry(f"{w}x{h}")
                
        except Exception as e:
            print(f"Failed to load settings: {e}")

def save_settings():
    root.update_idletasks()
    settings = {
        "rotate_option": state.rotate_option.get(),
        "save_option": state.save_option.get(),
        "preset_save_dir": state.preset_save_dir,
        "engine_var": state.engine_var.get(),
        "output_format_var": state.output_format_var.get(),
        "extract_mode_var": state.extract_mode_var.get(),
        "api_plan_var": state.api_plan_var.get(),
        "api_key_free_var": state.api_key_free_var.get().strip(),
        "api_key_paid_var": state.api_key_paid_var.get().strip(),
        "gemini_model_free_var": state.gemini_model_free_var.get(),
        "gemini_model_paid_var": state.gemini_model_paid_var.get(),
        "api_rpm_free_var": state.api_rpm_free_var.get(),
        "api_rpm_paid_var": state.api_rpm_paid_var.get(),
        "temperature_free_var": state.temperature_free_var.get(),
        "temperature_paid_var": state.temperature_paid_var.get(),
        "safety_free_var": state.safety_free_var.get(),
        "safety_paid_var": state.safety_paid_var.get(),
        "max_tokens_free_var": state.max_tokens_free_var.get(),
        "max_tokens_paid_var": state.max_tokens_paid_var.get(),
        "custom_prompt_free_var": state.custom_prompt_free_var.get(),
        "custom_prompt_paid_var": state.custom_prompt_paid_var.get(),
        "threads_free_var": state.threads_free_var.get(),
        "threads_paid_var": state.threads_paid_var.get(),
        "saved_custom_prompts": state.saved_custom_prompts,
        "window_width": root.winfo_width(),
        "window_height": root.winfo_height()
    }
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
            
        plan = state.api_plan_var.get()
        current_key = state.api_key_free_var.get().strip() if plan == "free" else state.api_key_paid_var.get().strip()
        if current_key:
            with open(API_KEY_FILE, "w", encoding="utf-8") as f:
                f.write(current_key)
                
        messagebox.showinfo("保存完了", "現在の選択項目を保存しました。\n次回起動時もこの設定が適用されます。")
        update_ui() 
    except Exception as e:
        messagebox.showerror("エラー", f"設定の保存に失敗しました。\n{e}")

# ==============================
# UI操作関連関数
# ==============================
def select_files():
    files = filedialog.askopenfilenames(filetypes=[("すべての対応ファイル", "*.pdf;*.xlsx;*.xlsm;*.xls;*.csv;*.txt;*.json;*.md;*.docx"), ("PDF", "*.pdf")])
    if files: state.selected_files, state.selected_folder, state.current_mode = list(files), "", "file"; update_ui()

def select_folder():
    folder = filedialog.askdirectory(title="フォルダを選択")
    if folder: state.selected_folder, state.selected_files, state.current_mode = folder, [], "folder"; update_ui()

def select_save_dir():
    folder = filedialog.askdirectory()
    if folder: state.preset_save_dir = folder; state.save_label.config(text=state.preset_save_dir); state.save_option.set(2)

def on_save_mode_change():
    state.preset_save_dir = ""
    state.save_label.config(text="同じフォルダ" if state.save_option.get() == 1 else "未選択")

format_radiobuttons = {}
def toggle_extraction_settings(*args):
    is_active = state.current_mode is not None
    for fmt, rb in format_radiobuttons.items(): rb.configure(state=tk.NORMAL)
    
    fmt = state.output_format_var.get()
    is_image_fmt = fmt in ["jpg", "png", "tiff", "bmp", "svg", "dxf"]
    
    state_mode = tk.NORMAL if is_active and not is_image_fmt else tk.DISABLED
    if 'rb_table_mode' in globals() and 'rb_text_mode' in globals() and rb_table_mode and rb_text_mode:
        rb_table_mode.configure(state=state_mode)
        rb_text_mode.configure(state=state_mode)
    
    is_gemini = (state.engine_var.get() == "Gemini")
    state_gemini = tk.NORMAL if is_gemini else tk.DISABLED
    if state.btn_api_settings:
        state.btn_api_settings.configure(state=state_gemini)
        
    state_crop = tk.NORMAL if is_active else tk.DISABLED
    for child in crop_frame.winfo_children():
        if isinstance(child, ttk.Button) or isinstance(child, ttk.Label): 
            child.configure(state=state_crop)
            
    # データ確認ボタンも連動させる
    for child in pdf_check_frame.winfo_children():
        if isinstance(child, ttk.Button) or isinstance(child, ttk.Label): 
            child.configure(state=state_crop)

def update_ui():
    state.path_label.config(text="\n".join(state.selected_files) if state.current_mode == "file" else (f"フォルダ: {state.selected_folder}" if state.selected_folder else "未選択"))
    is_active = state.current_mode is not None
    state_val = tk.NORMAL if is_active else tk.DISABLED
    btn_split.config(state=state_val); btn_rotate.config(state=state_val); btn_extract.config(state=state_val)
    btn_aggregate_local.config(state=state_val)
    if 'btn_combine_local' in globals() and btn_combine_local: btn_combine_local.config(state=state_val)
    btn_merge.config(state=tk.NORMAL if state.current_mode=="folder" else tk.DISABLED)
    if not is_active: reset_crop_regions()
    toggle_extraction_settings()
    
    if state.plan_indicator:
        plan = state.api_plan_var.get()
        if plan == "free":
            state.plan_indicator.config(text="🟢 無料枠 (Free)", foreground=COLOR_SUCCESS)
        else:
            state.plan_indicator.config(text="🔵 課金枠 (Paid)", foreground=PRIMARY)

def show_text_window(title, content):
    win = tk.Toplevel(root); win.title(title); win.geometry("620x550"); win.configure(bg=BG_COLOR)
    x = root.winfo_x() + (WINDOW_WIDTH // 2) - 310; y = root.winfo_y() + (WINDOW_HEIGHT // 2) - 275; win.geometry(f"+{x}+{y}")
    text_area = st.ScrolledText(win, wrap=tk.WORD, font=("Meiryo UI", 9), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.FLAT, padx=15, pady=15)
    text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=10); text_area.insert(tk.END, content); text_area.configure(state=tk.DISABLED)

def show_version_info(): 
    messagebox.showinfo("バージョン情報", f"{APP_TITLE}\nバージョン: {VERSION}\n\nPython & Tkinter製 PDF編集ツール")

def show_history(): 
    show_text_window("バージョン履歴", VERSION_HISTORY.strip())

def show_readme():
    p = resource_path("README.md")
    content = open(p, "r", encoding="utf-8").read() if os.path.exists(p) else "README.mdが見つかりません。\nアプリと同じフォルダに配置してください。"
    show_text_window("Readme", content)

def _on_mousewheel(event):
    widget = root.winfo_containing(event.x_root, event.y_root)
    if widget:
        current = widget
        while current:
            if isinstance(current, tk.Canvas):
                try: current.yview_scroll(int(-1*(event.delta/120)), "units")
                except: pass
                return
            current = current.master

# ==============================
# アプリケーション初期化とUI構築
# ==============================
root = tk.Tk(); root.title(f"{APP_TITLE} {VERSION}")
icon_path = resource_path("icon.ico")
if os.path.exists(icon_path):
    try:
        root.iconphoto(True, tk.PhotoImage(file=icon_path))
    except Exception as e:
        try:
            root.iconbitmap(icon_path)
        except Exception as e2:
            print(f"Failed to set icon: {e}, {e2}")
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+0+0")
root.minsize(width=720, height=580) 
root.configure(bg=BG_COLOR)
root.bind_all("<MouseWheel>", _on_mousewheel)
state.root = root

style = ttk.Style(); style.theme_use("clam") if "clam" in style.theme_names() else None
style.configure(".", background=BG_COLOR, font=("Segoe UI", 9))
style.configure("Main.TFrame", background=BG_COLOR)
style.configure("Card.TFrame", background=CARD_BG)
style.configure("Card.TLabelframe", background=CARD_BG, borderwidth=1, bordercolor=BORDER_COLOR)
style.configure("Card.TLabelframe.Label", background=CARD_BG, foreground=PRIMARY, font=("Segoe UI", 10, "bold"))
style.configure("TButton", padding=4, font=("Segoe UI", 9), background="#E9ECEF", foreground=TEXT_COLOR, borderwidth=1)
style.map("TButton", background=[("active", "#DEE2E6")])
style.configure("Primary.TButton", background=PRIMARY, foreground="white", borderwidth=0)
style.map("Primary.TButton", background=[("active", PRIMARY_HOVER)])
style.configure("Warning.TButton", background=COLOR_WARNING, foreground="black", borderwidth=0)
style.map("Warning.TButton", background=[("active", COLOR_WARNING_HOVER)])
style.configure("Danger.TButton", background=COLOR_DANGER, foreground="white", borderwidth=0)
style.map("Danger.TButton", background=[("active", COLOR_DANGER_HOVER)])
style.configure("Purple.TButton", background=COLOR_PURPLE, foreground="white", borderwidth=0)
style.map("Purple.TButton", background=[("active", COLOR_PURPLE_HOVER)])
style.configure("TRadiobutton", background=CARD_BG, font=("Segoe UI", 9), foreground=TEXT_COLOR)
style.configure("TCheckbutton", background=CARD_BG, font=("Segoe UI", 9), foreground=TEXT_COLOR)

menubar = Menu(root)
help_menu = Menu(menubar, tearoff=0)
help_menu.add_command(label="PDFの内部データ構造について", command=lambda: show_text_window("PDFの内部データ構造について", PDF_TYPE_HELP_TEXT.strip()))
help_menu.add_command(label="AI抽出の準備 (使い方)", command=lambda: show_text_window("AI抽出の準備 (使い方)", AI_HELP_TEXT.strip()))
help_menu.add_separator()
help_menu.add_command(label="Readmeを表示", command=show_readme)
help_menu.add_command(label="バージョン履歴", command=show_history)
help_menu.add_separator()
help_menu.add_command(label="バージョン情報", command=show_version_info)
menubar.add_cascade(label="ヘルプ", menu=help_menu)
root.config(menu=menubar)

state.rotate_option, state.save_option = tk.IntVar(value=270), tk.IntVar(value=1)
state.engine_var, state.output_format_var = tk.StringVar(value="Internal"), tk.StringVar(value="xlsx")
state.extract_mode_var = tk.StringVar(value="table")
state.api_plan_var = tk.StringVar(value="free")
state.api_key_free_var = tk.StringVar(value=get_api_key() or "")
state.api_key_paid_var = tk.StringVar(value=get_api_key() or "")
state.gemini_model_free_var = tk.StringVar(value="gemini-2.5-flash")
state.gemini_model_paid_var = tk.StringVar(value="gemini-2.5-flash")
state.api_rpm_free_var = tk.IntVar(value=12)
state.api_rpm_paid_var = tk.IntVar(value=300)
state.temperature_free_var = tk.DoubleVar(value=0.0)
state.temperature_paid_var = tk.DoubleVar(value=0.0)
state.safety_free_var = tk.BooleanVar(value=True)
state.safety_paid_var = tk.BooleanVar(value=True)
state.max_tokens_free_var = tk.IntVar(value=8192)
state.max_tokens_paid_var = tk.IntVar(value=8192)
state.custom_prompt_free_var = tk.StringVar(value="")
state.custom_prompt_paid_var = tk.StringVar(value="")
state.threads_free_var = tk.IntVar(value=1)
state.threads_paid_var = tk.IntVar(value=5)

main_outer = ttk.Frame(root)
main_outer.pack(fill=tk.BOTH, expand=True)

canvas = tk.Canvas(main_outer, bg=BG_COLOR, highlightthickness=0)
scrollbar = ttk.Scrollbar(main_outer, orient=tk.VERTICAL, command=canvas.yview)

main_container = ttk.Frame(canvas, padding=8, style="Main.TFrame")
canvas_frame_id = canvas.create_window((0, 0), window=main_container, anchor="nw")

def on_canvas_configure(event): canvas.itemconfig(canvas_frame_id, width=event.width)
canvas.bind('<Configure>', on_canvas_configure)

def on_frame_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
main_container.bind('<Configure>', on_frame_configure)

canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

title_frame = ttk.Frame(main_container, style="Main.TFrame"); title_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(title_frame, text=APP_TITLE, font=("Segoe UI", 16, "bold"), foreground=PRIMARY, background=BG_COLOR).pack(side=tk.LEFT)
ttk.Label(title_frame, text=f" {VERSION}", font=("Segoe UI", 10), foreground=MUTED_TEXT, background=BG_COLOR).pack(side=tk.LEFT, pady=(6, 0))

file_card = ttk.Frame(main_container, style="Card.TFrame", padding=5); file_card.pack(fill=tk.X, pady=2)
btn_frame = ttk.Frame(file_card, style="Card.TFrame"); btn_frame.pack()
ttk.Button(btn_frame, text="📄 ファイルを選択", command=select_files, width=20, style="Primary.TButton").grid(row=0, column=0, padx=8)
ttk.Button(btn_frame, text="📁 フォルダを選択", command=select_folder, width=20, style="Primary.TButton").grid(row=0, column=1, padx=8)
state.path_label = ttk.Label(file_card, text="未選択", background=CARD_BG, foreground=TEXT_COLOR, wraplength=580, justify="center"); state.path_label.pack(pady=(4, 0))

settings_grid = ttk.Frame(main_container, style="Main.TFrame"); settings_grid.pack(fill=tk.X, pady=2); settings_grid.columnconfigure(0, weight=1); settings_grid.columnconfigure(1, weight=1)
save_frame = ttk.LabelFrame(settings_grid, text=" 保存先設定 ", style="Card.TLabelframe", padding=4); save_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
ttk.Radiobutton(save_frame, text="元のファイルと同じフォルダ", variable=state.save_option, value=1, command=on_save_mode_change).pack(anchor="w", pady=1)
ttk.Radiobutton(save_frame, text="任意のフォルダを指定", variable=state.save_option, value=2, command=on_save_mode_change).pack(anchor="w", pady=1)
ttk.Button(save_frame, text="📂 フォルダ参照", command=select_save_dir).pack(pady=(2, 2))
state.save_label = ttk.Label(save_frame, text="同じフォルダ", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 9)); state.save_label.pack()

rotate_frame = ttk.LabelFrame(settings_grid, text=" 回転設定 ", style="Card.TLabelframe", padding=4); rotate_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
for t, v in [("左（270°）", 270), ("上下（180°）", 180), ("右（90°）", 90)]: ttk.Radiobutton(rotate_frame, text=t, variable=state.rotate_option, value=v).pack(anchor="w", pady=1)

extract_frame = ttk.LabelFrame(main_container, text=" ⚙️ データ抽出・変換設定 ", style="Card.TLabelframe", padding=5); extract_frame.pack(fill=tk.X, pady=4)

engine_frame = ttk.Frame(extract_frame, style="Card.TFrame"); engine_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(engine_frame, text="① エンジン:", width=12, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
engine_inner = ttk.Frame(engine_frame, style="Card.TFrame"); engine_inner.pack(anchor="w", fill=tk.X)
for text, val in [("Python標準ライブラリ (高速・オフライン) - テキストデータ入りPDF向け（文字選択できるPDF用）", "Internal"), ("Tesseract (ローカルOCR) - ラスターデータ向け（スキャンされた印刷活字のPDF用）", "Tesseract"), ("Gemini API (超高精度AI) - ラスターデータ向け（スキャンされた手書き文字や複雑な表のPDF用）", "Gemini")]:
    ttk.Radiobutton(engine_inner, text=text, variable=state.engine_var, value=val).pack(anchor="w", pady=1)

ttk.Separator(extract_frame, orient="horizontal").pack(fill=tk.X, pady=4)

format_frame = ttk.Frame(extract_frame, style="Card.TFrame"); format_frame.pack(fill=tk.X, pady=0)
ttk.Label(format_frame, text="② 出力形式:", width=12, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)

formats_row1 = [("Excel (.xlsx)", "xlsx"), ("CSV (.csv)", "csv"), ("Text (.txt)", "txt"), ("JSON (.json)", "json"), ("Markdown (.md)", "md"), ("Word (.docx)", "docx")]
formats_row2 = [("JPEG (.jpg)", "jpg"), ("PNG (.png)", "png"), ("SVG (.svg)", "svg"), ("TIFF (.tiff)", "tiff"), ("BMP (.bmp)", "bmp"), ("DXF (.dxf)", "dxf")]

format_inner1 = ttk.Frame(format_frame, style="Card.TFrame"); format_inner1.pack(anchor="w", fill=tk.X)
for text, val in formats_row1:
    rb = ttk.Radiobutton(format_inner1, text=text, variable=state.output_format_var, value=val); rb.pack(side=tk.LEFT, padx=(0, 10)); format_radiobuttons[val] = rb
format_inner2 = ttk.Frame(format_frame, style="Card.TFrame"); format_inner2.pack(anchor="w", fill=tk.X, pady=(2, 0))
for text, val in formats_row2:
    rb = ttk.Radiobutton(format_inner2, text=text, variable=state.output_format_var, value=val); rb.pack(side=tk.LEFT, padx=(0, 10)); format_radiobuttons[val] = rb

ttk.Separator(extract_frame, orient="horizontal").pack(fill=tk.X, pady=4)

extract_mode_frame = ttk.Frame(extract_frame, style="Card.TFrame")
extract_mode_frame.pack(fill=tk.X, pady=(0, 2))
ttk.Label(extract_mode_frame, text="③ 抽出モード:", width=12, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
rb_table_mode = ttk.Radiobutton(extract_mode_frame, text="表として抽出 (セルごとに分割)", variable=state.extract_mode_var, value="table")
rb_table_mode.pack(side=tk.LEFT, padx=(0, 10))
rb_text_mode = ttk.Radiobutton(extract_mode_frame, text="テキストのみ抽出 (横に1行で出力)", variable=state.extract_mode_var, value="text")
rb_text_mode.pack(side=tk.LEFT)

ttk.Separator(extract_frame, orient="horizontal").pack(fill=tk.X, pady=4)

api_settings_frame = ttk.Frame(extract_frame, style="Card.TFrame")
api_settings_frame.pack(fill=tk.X, pady=2)
ttk.Label(api_settings_frame, text="[AI用] API設定:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
state.btn_api_settings = ttk.Button(api_settings_frame, text="⚙️ 詳細設定 (APIキー / モデル / 制限)", command=open_api_settings_dialog, style="Primary.TButton")
state.btn_api_settings.pack(side=tk.LEFT)

state.plan_indicator = ttk.Label(api_settings_frame, text="", font=("Segoe UI", 9, "bold"), background=CARD_BG)
state.plan_indicator.pack(side=tk.LEFT, padx=(12, 0))

global crop_frame, pdf_check_frame
crop_frame = ttk.Frame(extract_frame, style="Card.TFrame"); crop_frame.pack(fill=tk.X, pady=(4, 0))
ttk.Label(crop_frame, text="抽出範囲:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
state.btn_select_crop = ttk.Button(crop_frame, text="抽出範囲を選択", command=open_crop_selector); state.btn_select_crop.pack(side=tk.LEFT)
btn_reset_crop = ttk.Button(crop_frame, text="全体に戻す", command=reset_crop_regions, style="Warning.TButton"); btn_reset_crop.pack(side=tk.LEFT, padx=(5, 0))

pdf_check_frame = ttk.Frame(extract_frame, style="Card.TFrame")
pdf_check_frame.pack(fill=tk.X, pady=(4, 0))
ttk.Label(pdf_check_frame, text="データ確認:", width=14, background=CARD_BG, font=("Segoe UI", 9, "bold"), foreground=TEXT_COLOR).pack(side=tk.LEFT)
btn_check_pdf = ttk.Button(pdf_check_frame, text="🔍 選択中のPDFデータ構造を確認", command=show_pdf_type_info, style="Primary.TButton")
btn_check_pdf.pack(side=tk.LEFT)

save_settings_frame = ttk.Frame(extract_frame, style="Card.TFrame")
save_settings_frame.pack(fill=tk.X, pady=(6, 0))
btn_save_settings = ttk.Button(save_settings_frame, text="💾 現在の選択項目を保存", command=save_settings)
btn_save_settings.pack(side=tk.RIGHT)

action_container = ttk.Frame(main_container, style="Main.TFrame"); action_container.pack(fill=tk.BOTH, expand=True, pady=2)
action_container.columnconfigure(0, weight=1); action_container.columnconfigure(1, weight=1)

pdf_action_frame = ttk.LabelFrame(action_container, text=" ✂️ PDF編集 ", style="Card.TLabelframe", padding=4); pdf_action_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
btn_merge = ttk.Button(pdf_action_frame, text="結合 (フォルダ)", command=lambda: safe_run(merge_pdfs, "PDF結合"), style="Primary.TButton"); btn_merge.pack(fill=tk.X, pady=2)
btn_split = ttk.Button(pdf_action_frame, text="分割", command=lambda: safe_run(split_pdfs, "PDF分割"), style="Primary.TButton"); btn_split.pack(fill=tk.X, pady=2)
btn_rotate = ttk.Button(pdf_action_frame, text="回転", command=lambda: safe_run(rotate_pdfs, "PDF回転"), style="Primary.TButton"); btn_rotate.pack(fill=tk.X, pady=2)

data_action_frame = ttk.LabelFrame(action_container, text=" 📊 データ操作 ", style="Card.TLabelframe", padding=4); data_action_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
btn_extract = ttk.Button(data_action_frame, text="🚀 選択した抽出・変換を実行", command=run_selected_extraction, style="Danger.TButton"); btn_extract.pack(fill=tk.X, pady=(2, 4), ipady=4) 
btn_aggregate_local = ttk.Button(data_action_frame, text="🧩 データ集約", command=lambda: safe_run(aggregate_local_task, "データ集約"), style="Purple.TButton")
btn_aggregate_local.pack(fill=tk.X, pady=(2, 0))
ttk.Label(data_action_frame, text="※列名を自動認識して賢く名寄せ・結合します", font=("Segoe UI", 8), foreground=MUTED_TEXT, background=CARD_BG).pack(pady=(0, 4))

global btn_combine_local
btn_combine_local = ttk.Button(data_action_frame, text="🔗 データ単純結合", command=lambda: safe_run(combine_local_task, "データ単純結合"), style="Primary.TButton")
btn_combine_local.pack(fill=tk.X, pady=(2, 0))
ttk.Label(data_action_frame, text="※同じ列構成のファイルをそのまま縦に繋げます", font=("Segoe UI", 8), foreground=MUTED_TEXT, background=CARD_BG).pack(pady=(0, 4))

status_frame = ttk.Frame(main_container, style="Main.TFrame")
status_frame.pack(fill=tk.X, pady=(2, 0))
state.status_label = ttk.Label(status_frame, text="ステータス: 待機中", font=("Segoe UI", 9), foreground=MUTED_TEXT, background=BG_COLOR)
state.status_label.pack(side=tk.LEFT, padx=5)

state.engine_var.trace("w", toggle_extraction_settings)
state.output_format_var.trace("w", toggle_extraction_settings)

load_settings()
update_ui()

if __name__ == "__main__":
    root.mainloop()