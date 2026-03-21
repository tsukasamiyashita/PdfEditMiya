# -*- coding: utf-8 -*-
import os, sys, threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from config import *
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
# プログレス関連のローカル状態
# ==============================
processing_popup = None
overall_label = None
overall_progress = None
file_label = None
file_progress = None
btn_cancel = None

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

# ==============================
# バックグラウンド処理からのUI更新コントローラー
# ==============================
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

# ==============================
# 処理実行・制御系関数
# ==============================
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
            "threads": state.threads_free_var.get() if is_free else state.threads_paid_var.get()
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
        def _err_status():
            show_message(f"❌ エラーが発生しました\n{str(e)[:40]}...", ERROR)
            if state.status_label: state.status_label.config(text=f"ステータス: {task_name} でエラーが発生しました", foreground=ERROR)
        state.root.after(0, _err_status)

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