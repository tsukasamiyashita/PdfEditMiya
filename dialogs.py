# -*- coding: utf-8 -*-
import os, sys, re, warnings
import tkinter as tk
from tkinter import ttk, messagebox, Menu
import fitz
from PIL import Image, ImageTk

# FutureWarningを抑制
warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai

from common import *

# ==============================
# UI共通コンポーネント
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

# ==============================
# コンテキストメニューヘルパー
# ==============================
def show_context_menu(event, target_widget=None):
    widget = target_widget if target_widget else event.widget
    menu = Menu(state.root, tearoff=0)
    menu.add_command(label="貼り付け", command=lambda: paste_to_entry(widget))
    menu.post(event.x_root, event.y_root)

def paste_to_entry(widget):
    try:
        text = state.root.clipboard_get()
        try: widget.delete("sel.first", "sel.last")
        except tk.TclError: pass
        widget.insert(tk.INSERT, text)
    except tk.TclError: pass

def show_text_context_menu(event, text_widget):
    menu = Menu(state.root, tearoff=0)
    menu.add_command(label="コピー", command=lambda: text_widget.event_generate("<<Copy>>"))
    menu.add_command(label="切り取り", command=lambda: text_widget.event_generate("<<Cut>>"))
    menu.add_command(label="貼り付け", command=lambda: text_widget.event_generate("<<Paste>>"))
    menu.post(event.x_root, event.y_root)

# ==============================
# API詳細設定ダイアログ
# ==============================
def open_api_settings_dialog():
    dialog = tk.Toplevel(state.root)
    dialog.title("⚙️ AI詳細設定 (Gemini API)")
    
    screen_h = state.root.winfo_screenheight()
    screen_w = state.root.winfo_screenwidth()
    
    dialog_w = min(1050, screen_w - 40)
    dialog_h = min(780, screen_h - 80) 
    dialog.geometry(f"{dialog_w}x{dialog_h}") 
    dialog.configure(bg=BG_COLOR)
    dialog.grab_set()
    
    raw_x = state.root.winfo_x() + (WINDOW_WIDTH // 2) - (dialog_w // 2)
    raw_y = state.root.winfo_y() + (WINDOW_HEIGHT // 2) - (dialog_h // 2)
    
    x = max(10, min(raw_x, screen_w - dialog_w - 10))
    y = max(10, min(raw_y, screen_h - dialog_h - 40))
    
    dialog.geometry(f"+{x}+{y}")

    fav_lists = []

    def update_all_fav_lists():
        for f_list in fav_lists:
            f_list.set_items(state.saved_custom_prompts)

    original_values = {
        "plan": state.api_plan_var.get(),
        "key_free": state.api_key_free_var.get(), "key_paid": state.api_key_paid_var.get(),
        "model_free": state.gemini_model_free_var.get(), "model_paid": state.gemini_model_paid_var.get(),
        "rpm_free": state.api_rpm_free_var.get(), "rpm_paid": state.api_rpm_paid_var.get(),
        "temp_free": state.temperature_free_var.get(), "temp_paid": state.temperature_paid_var.get(),
        "safety_free": state.safety_free_var.get(), "safety_paid": state.safety_paid_var.get(),
        "tokens_free": state.max_tokens_free_var.get(), "tokens_paid": state.max_tokens_paid_var.get(),
        "prompt_free": state.custom_prompt_free_var.get(), "prompt_paid": state.custom_prompt_paid_var.get(),
        "threads_free": state.threads_free_var.get(), "threads_paid": state.threads_paid_var.get(),
        "saved_prompts": list(state.saved_custom_prompts)
    }

    def apply_and_close():
        if state.plan_indicator:
            plan = state.api_plan_var.get()
            if plan == "free":
                state.plan_indicator.config(text="🟢 無料枠 (Free)", foreground=COLOR_SUCCESS)
            else:
                state.plan_indicator.config(text="🔵 課金枠 (Paid)", foreground=PRIMARY)
        dialog.destroy()

    def has_changes():
        if state.api_plan_var.get() != original_values["plan"]: return True
        if state.api_key_free_var.get() != original_values["key_free"]: return True
        if state.api_key_paid_var.get() != original_values["key_paid"]: return True
        if state.gemini_model_free_var.get() != original_values["model_free"]: return True
        if state.gemini_model_paid_var.get() != original_values["model_paid"]: return True
        if state.api_rpm_free_var.get() != original_values["rpm_free"]: return True
        if state.api_rpm_paid_var.get() != original_values["rpm_paid"]: return True
        if state.temperature_free_var.get() != original_values["temp_free"]: return True
        if state.temperature_paid_var.get() != original_values["temp_paid"]: return True
        if state.safety_free_var.get() != original_values["safety_free"]: return True
        if state.safety_paid_var.get() != original_values["safety_paid"]: return True
        if state.max_tokens_free_var.get() != original_values["tokens_free"]: return True
        if state.max_tokens_paid_var.get() != original_values["tokens_paid"]: return True
        if state.custom_prompt_free_var.get() != original_values["prompt_free"]: return True
        if state.custom_prompt_paid_var.get() != original_values["prompt_paid"]: return True
        if state.threads_free_var.get() != original_values["threads_free"]: return True
        if state.threads_paid_var.get() != original_values["threads_paid"]: return True
        if state.saved_custom_prompts != original_values["saved_prompts"]: return True
        return False

    def cancel_and_close():
        if has_changes():
            if not messagebox.askyesno("確認", "変更が適用されていません。\n破棄して設定画面を閉じますか？", parent=dialog):
                return 
                
        state.api_plan_var.set(original_values["plan"])
        state.api_key_free_var.set(original_values["key_free"])
        state.api_key_paid_var.set(original_values["key_paid"])
        state.gemini_model_free_var.set(original_values["model_free"])
        state.gemini_model_paid_var.set(original_values["model_paid"])
        state.api_rpm_free_var.set(original_values["rpm_free"])
        state.api_rpm_paid_var.set(original_values["rpm_paid"])
        state.temperature_free_var.set(original_values["temp_free"])
        state.temperature_paid_var.set(original_values["temp_paid"])
        state.safety_free_var.set(original_values["safety_free"])
        state.safety_paid_var.set(original_values["safety_paid"])
        state.max_tokens_free_var.set(original_values["tokens_free"])
        state.max_tokens_paid_var.set(original_values["tokens_paid"])
        state.custom_prompt_free_var.set(original_values["prompt_free"])
        state.custom_prompt_paid_var.set(original_values["prompt_paid"])
        state.threads_free_var.set(original_values["threads_free"])
        state.threads_paid_var.set(original_values["threads_paid"])
        state.saved_custom_prompts.clear()
        state.saved_custom_prompts.extend(original_values["saved_prompts"])
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", cancel_and_close)

    # ==============================
    # ダイアログ全体のレイアウト構築
    # ==============================
    btn_action_frame = ttk.Frame(dialog, style="Main.TFrame")
    btn_action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 15))
    
    btn_cancel = ttk.Button(btn_action_frame, text="キャンセル", command=cancel_and_close, width=15)
    btn_cancel.pack(side=tk.LEFT, padx=10)
    
    btn_apply = ttk.Button(btn_action_frame, text="設定を適用して閉じる", command=apply_and_close, style="Primary.TButton", width=25)
    btn_apply.pack(side=tk.LEFT, padx=10)

    container = ttk.Frame(dialog, style="Main.TFrame")
    container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    
    canvas = tk.Canvas(container, bg=BG_COLOR, highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas, style="Main.TFrame")

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    
    def _on_canvas_configure(event):
        canvas.itemconfig(canvas_window, width=event.width)
    canvas.bind("<Configure>", _on_canvas_configure)

    def _on_mousewheel(event):
        if not dialog.winfo_exists(): return
        widget = dialog.winfo_containing(event.x_root, event.y_root)
        if widget:
            current = widget
            while current:
                if isinstance(current, tk.Canvas):
                    try: 
                        current.yview_scroll(int(-1*(event.delta/120)), "units")
                    except: pass
                    return
                current = current.master

    dialog.bind("<MouseWheel>", _on_mousewheel)

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # ==============================
    # コンテンツの配置
    # ==============================
    lbl_title = ttk.Label(scrollable_frame, text="Gemini API 詳細設定", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
    lbl_title.pack(pady=(10, 5))

    plan_frame = ttk.LabelFrame(scrollable_frame, text=" 実行プランの選択 ", style="Card.TLabelframe", padding=8)
    plan_frame.pack(fill=tk.X, padx=15, pady=(0, 5))
    
    plan_inner = ttk.Frame(plan_frame, style="Card.TFrame")
    plan_inner.pack(anchor="w", padx=5, pady=2)

    ttk.Label(plan_inner, text="実際に抽出で使用するプランを選んでください（下のタブとは連動しません）:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 15))
    rb_free = ttk.Radiobutton(plan_inner, text="無料枠 (Free Tier)", variable=state.api_plan_var, value="free")
    rb_free.pack(side=tk.LEFT, padx=(0, 15))
    rb_paid = ttk.Radiobutton(plan_inner, text="課金枠 (Paid Tier)", variable=state.api_plan_var, value="paid")
    rb_paid.pack(side=tk.LEFT)

    notebook = ttk.Notebook(scrollable_frame)
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
        key_var = state.api_key_free_var if is_free else state.api_key_paid_var
        model_var = state.gemini_model_free_var if is_free else state.gemini_model_paid_var
        rpm_var = state.api_rpm_free_var if is_free else state.api_rpm_paid_var
        temp_var = state.temperature_free_var if is_free else state.temperature_paid_var
        safety_var = state.safety_free_var if is_free else state.safety_paid_var
        tokens_var = state.max_tokens_free_var if is_free else state.max_tokens_paid_var
        prompt_var = state.custom_prompt_free_var if is_free else state.custom_prompt_paid_var
        threads_var = state.threads_free_var if is_free else state.threads_paid_var
        
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
            import time
            key = k_var.get().strip()
            if not key: return messagebox.showwarning("警告", "APIキーが入力されていません。", parent=dialog)
            genai.configure(api_key=key)
            model_name = m_var.get()
            try:
                model = genai.GenerativeModel(model_name)
                # キャッシュ回避のため毎回異なる文字列（タイムスタンプ付き）を送信
                model.generate_content(f"Connection Test: {time.time()}")
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
                        msg += f"⚠️ Googleの制限により、約 {wait_sec} 秒後に利用枠が回復すると報告されています。\n"
                        msg += "（※数分待っても回復しない場合は、1日あたりの利用上限に到達している可能性が高いです）\n"
                    else:
                        if "per day" in err_str or "perday" in err_str: msg += "⚠️ 【1日の利用上限】に達しました。明日以降に再度お試しください。\n"
                        else: msg += "⚠️ APIの制限に達しました。\n"
                    msg += f"\n詳細（生のエラー）:\n{e}"
                    messagebox.showerror("利用枠超過", msg, parent=dialog)
                else:
                    messagebox.showerror("通信エラー", f"APIキーまたは通信に問題が発生しました。\n詳細:\n{e}", parent=dialog)
        
        btn_test = ttk.Button(key_inner, text="テスト", command=test_key, width=6)
        btn_test.pack(side=tk.LEFT)

        middle_frame = ttk.Frame(parent_tab, style="Main.TFrame")
        middle_frame.pack(fill=tk.X, padx=10, pady=5)
        middle_frame.columnconfigure(0, weight=1)
        middle_frame.columnconfigure(1, weight=1)

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
            
            info_w = min(950, screen_w - 40)
            info_h = min(700, screen_h - 80)
            info_win.geometry(f"{info_w}x{info_h}") 
            info_win.configure(bg=BG_COLOR)
            info_win.grab_set()
            
            raw_x = dialog.winfo_x() + 30
            raw_y = dialog.winfo_y() + 30
            win_x = max(10, min(raw_x, screen_w - info_w - 10))
            win_y = max(10, min(raw_y, screen_h - info_h - 40))
            info_win.geometry(f"+{win_x}+{win_y}")
            
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
                if text not in state.saved_custom_prompts:
                    state.saved_custom_prompts.append(text)
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
                    if text in state.saved_custom_prompts:
                        state.saved_custom_prompts.remove(text)
                update_all_fav_lists()

        ttk.Button(right_actions, text="◀ 選択を左に追加", command=add_fav_to_current, style="Primary.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(right_actions, text="🗑 選択を削除", command=delete_fav_selected).pack(side=tk.LEFT)

        initial_prompts = [p for p in prompt_var.get().split('\n') if p.strip()]
        current_list.set_items(initial_prompts)

    build_tab(tab_free, "free")
    build_tab(tab_paid, "paid")
    
    update_all_fav_lists() 
    
    if state.api_plan_var.get() == "free":
        notebook.select(tab_free)
    else:
        notebook.select(tab_paid)

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
        state.selected_crop_regions = [(r['rx1'], r['ry1'], r['rx2'], r['ry2']) for r in self.rectangles]
        if state.btn_select_crop:
            state.btn_select_crop.config(text=f"抽出範囲を選択 (設定済: {len(state.selected_crop_regions)}か所)" if state.selected_crop_regions else "抽出範囲を選択")
        self.doc.close(); self.top.destroy()

def open_crop_selector():
    files = state.selected_files if state.current_mode == "file" else ([os.path.join(state.selected_folder, f) for f in os.listdir(state.selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if state.selected_folder else [])
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    if not pdf_files: return messagebox.showinfo("情報", "PDFファイルが選択されていません。")
    try: CropSelector(state.root, pdf_files[0])
    except Exception as e: messagebox.showerror("エラー", str(e))

def reset_crop_regions():
    state.selected_crop_regions = []
    if state.btn_select_crop:
        state.btn_select_crop.config(text="抽出範囲を選択")