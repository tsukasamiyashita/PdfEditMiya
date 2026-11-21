# -*- coding: utf-8 -*-
import os, sys, re, warnings, json, webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, Menu
import tkinter.scrolledtext as st
import fitz
from PIL import Image, ImageTk

# FutureWarningを抑制
warnings.filterwarnings("ignore", category=FutureWarning)
import google.generativeai as genai

from common import *
from gemini_engine import (
    test_azure_connection, fetch_gemini_models, fetch_azure_models,
    fetch_azure_models_with_meta,
    filter_mainstream_gemini_models, filter_mainstream_azure_models,
    get_recommended_gemini_settings, get_recommended_azure_settings
)

MODELS_FILE = os.path.join(APP_DIR, ".pdfeditmiya_models.json")
LEGACY_MODELS_FILE = os.path.join(USER_HOME, ".pdfeditmiya_models.json")
AZURE_MODELS_FILE = os.path.join(APP_DIR, ".pdfeditmiya_azure_models.json")

def load_azure_models_list():
    default_models = [
        "gpt-5.2", "gpt-5.1", "gpt-5",
        "gpt-5-mini", "gpt-5-nano", "gpt-5-chat",
        "o4-mini"
    ]
    if os.path.exists(AZURE_MODELS_FILE):
        try:
            with open(AZURE_MODELS_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            if isinstance(saved, list) and saved:
                return [str(m) for m in saved if str(m).strip()]
        except Exception:
            pass
    return default_models

def save_azure_models_list(models_list):
    try:
        with open(AZURE_MODELS_FILE, "w", encoding="utf-8") as f:
            json.dump(models_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def load_models_list():
    default_models = [
        ("Gemini 4 Hydrogen", "gemini-4-hydrogen"),
        ("Gemini 4 Argon", "gemini-4-argon"),
        ("Gemini 4 Neon", "gemini-4-neon"),
        ("Gemini 4 Carbon", "gemini-4-carbon"),
        ("Gemini 3.8 Flash", "gemini-3.8-flash"),
        ("Gemini 3.5 Flash-Lite", "gemini-3.5-flash-lite")
    ]
    target_file = MODELS_FILE if os.path.exists(MODELS_FILE) else (LEGACY_MODELS_FILE if os.path.exists(LEGACY_MODELS_FILE) else None)
    if target_file:
        try:
            with open(target_file, "r", encoding="utf-8") as f:
                saved_models = json.load(f)

                default_dict = {m[1]: m[0] for m in default_models}
                merged_models = []
                for m in saved_models:
                    model_id = m[1]
                    if model_id in default_dict:
                        merged_models.append((default_dict[model_id], model_id))
                        del default_dict[model_id]
                    else:
                        display_name = m[0].split(" - ")[0]
                        merged_models.append((display_name, model_id))
                for k, v in default_dict.items():
                    merged_models.append((v, k))
                return merged_models
        except Exception:
            pass
    return default_models

def save_models_list(models_list):
    try:
        with open(MODELS_FILE, "w", encoding="utf-8") as f:
            json.dump(models_list, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

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
        try:
            widget.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        widget.insert(tk.INSERT, text)
    except tk.TclError:
        pass

def show_text_context_menu(event, text_widget):
    menu = Menu(state.root, tearoff=0)
    menu.add_command(label="コピー", command=lambda: text_widget.event_generate("<<Copy>>"))
    menu.add_command(label="切り取り", command=lambda: text_widget.event_generate("<<Cut>>"))
    menu.add_command(label="貼り付け", command=lambda: text_widget.event_generate("<<Paste>>"))
    menu.post(event.x_root, event.y_root)

def center_window(win, parent, w, h):
    """ウィンドウを親ウィンドウの中央かつ画面内に収まる位置へ配置する"""
    try:
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        pw = parent.winfo_width() if parent and parent.winfo_width() > 100 else WINDOW_WIDTH
        ph = parent.winfo_height() if parent and parent.winfo_height() > 100 else WINDOW_HEIGHT
        px = parent.winfo_rootx() if parent else 0
        py = parent.winfo_rooty() if parent else 0
        raw_x = px + (pw // 2) - (w // 2)
        raw_y = py + (ph // 2) - (h // 2)
        x = max(10, min(raw_x, screen_w - w - 10))
        y = max(10, min(raw_y, screen_h - h - 40))
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        win.geometry(f"{w}x{h}")

def show_runtime_error_dialog(parent, task_name, err_msg):
    """処理実行中のエラーをスクロール・コピー可能な見やすいダイアログで表示する"""
    err_win = tk.Toplevel(parent)
    err_win.title(f"実行エラー - {task_name}")
    screen_w = err_win.winfo_screenwidth()
    screen_h = err_win.winfo_screenheight()
    win_w = min(620, screen_w - 40)
    win_h = min(460, screen_h - 80)
    err_win.configure(bg=BG_COLOR)
    center_window(err_win, parent, win_w, win_h)
    err_win.grab_set()

    banner = tk.Frame(err_win, bg=CARD_BG, padx=16, pady=12, highlightbackground=BORDER_COLOR, highlightthickness=1)
    banner.pack(fill=tk.X, side=tk.TOP)
    tk.Label(
        banner, text=" ❌ 実行エラーが発生しました ",
        font=("Segoe UI", 10, "bold"), bg="#F8D7DA", fg="#842029", padx=8, pady=3
    ).pack(anchor="w")
    tk.Label(
        banner, text=f"「{task_name}」の処理中にエラーが発生したため停止しました。",
        font=("Meiryo UI", 9, "bold"), bg=CARD_BG, fg=TEXT_COLOR, justify="left", anchor="w", wraplength=win_w - 48
    ).pack(fill=tk.X, pady=(6, 0))

    btn_bar = tk.Frame(err_win, bg=BG_COLOR, padx=16, pady=10)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _copy_err():
        try:
            err_win.clipboard_clear()
            err_win.clipboard_append(f"[{task_name} エラー]\n{err_msg}")
            btn_cp.config(text="✅ コピーしました")
            err_win.after(2000, lambda: btn_cp.winfo_exists() and btn_cp.config(text="📋 エラー内容をコピー"))
        except Exception:
            pass

    btn_cp = ttk.Button(btn_bar, text="📋 エラー内容をコピー", command=_copy_err)
    btn_cp.pack(side=tk.LEFT)
    ttk.Button(btn_bar, text="OK", command=err_win.destroy, style="Primary.TButton", width=14).pack(side=tk.RIGHT)

    body = tk.Frame(err_win, bg=BG_COLOR, padx=14, pady=8)
    body.pack(fill=tk.BOTH, expand=True)
    tk.Label(body, text="▼ 詳細内容", font=("Segoe UI", 9, "bold"), bg=BG_COLOR, fg=MUTED_TEXT, anchor="w").pack(fill=tk.X, pady=(0, 4))
    box = st.ScrolledText(body, wrap=tk.WORD, font=("Meiryo UI", 9), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.SOLID, borderwidth=1, padx=12, pady=10)
    box.pack(fill=tk.BOTH, expand=True)
    box.insert(tk.END, err_msg)
    box.configure(state=tk.DISABLED)

def show_api_error_dialog(parent, diag):
    """
    diagnose_api_error() の診断結果を、原因・対処法・エラー原文に分けて
    初心者にもひと目で分かりやすく表示するエラーダイアログ。
    """
    err_win = tk.Toplevel(parent)
    err_win.title(f"接続エラー - {diag.get('title', 'エラー')}")
    screen_w = err_win.winfo_screenwidth()
    screen_h = err_win.winfo_screenheight()
    win_w = min(640, screen_w - 40)
    win_h = min(540, screen_h - 80)
    err_win.configure(bg=BG_COLOR)
    center_window(err_win, parent, win_w, win_h)
    err_win.grab_set()

    kind = diag.get("kind", "unknown")
    is_wait_type = kind in ["init_quota", "rate_limit"]
    accent_col = COLOR_WARNING_HOVER if is_wait_type else COLOR_DANGER
    badge_bg = "#FFF3CD" if is_wait_type else "#F8D7DA"
    badge_fg = "#856404" if is_wait_type else "#842029"
    badge_icon = "⏳" if is_wait_type else "❌"

    # 上部タイトルバナー
    banner = tk.Frame(err_win, bg=CARD_BG, padx=16, pady=12, highlightbackground=BORDER_COLOR, highlightthickness=1)
    banner.pack(fill=tk.X, side=tk.TOP)

    title_row = tk.Frame(banner, bg=CARD_BG)
    title_row.pack(fill=tk.X)
    tk.Label(
        title_row, text=f" {badge_icon} {diag.get('title', 'エラー')} ",
        font=("Segoe UI", 11, "bold"), bg=badge_bg, fg=badge_fg, padx=8, pady=3
    ).pack(side=tk.LEFT)

    tk.Label(
        banner, text=diag.get("summary", ""),
        font=("Meiryo UI", 10, "bold"), bg=CARD_BG, fg=TEXT_COLOR,
        justify="left", anchor="w", wraplength=win_w - 48
    ).pack(fill=tk.X, pady=(8, 0))

    # 下部ボタンバー
    btn_bar = tk.Frame(err_win, bg=BG_COLOR, padx=16, pady=10)
    btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _copy_all():
        try:
            text = f"[{diag.get('title', '')}]\n" + format_api_error_message(diag)
            err_win.clipboard_clear()
            err_win.clipboard_append(text)
            btn_copy.config(text="✅ コピーしました")
            err_win.after(2000, lambda: btn_copy.winfo_exists() and btn_copy.config(text="📋 エラー内容をコピー"))
        except Exception:
            pass

    btn_copy = ttk.Button(btn_bar, text="📋 エラー内容をコピー", command=_copy_all)
    btn_copy.pack(side=tk.LEFT)
    ttk.Button(btn_bar, text="OK", command=err_win.destroy, style="Primary.TButton", width=14).pack(side=tk.RIGHT)

    # 本文エリア
    body = tk.Frame(err_win, bg=BG_COLOR, padx=14, pady=8)
    body.pack(fill=tk.BOTH, expand=True)

    def _add_section_card(title, items, bar_color):
        if not items:
            return
        card_outer = tk.Frame(body, bg=BORDER_COLOR, padx=1, pady=1)
        card_outer.pack(fill=tk.X, pady=(0, 8))
        card_row = tk.Frame(card_outer, bg=CARD_BG)
        card_row.pack(fill=tk.BOTH, expand=True)
        tk.Frame(card_row, bg=bar_color, width=5).pack(side=tk.LEFT, fill=tk.Y)
        content = tk.Frame(card_row, bg=CARD_BG, padx=12, pady=8)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(content, text=title, font=("Segoe UI", 9, "bold"), bg=CARD_BG, fg=bar_color, anchor="w").pack(fill=tk.X, pady=(0, 4))
        for item in items:
            tk.Label(
                content, text=f"・{item}", font=("Meiryo UI", 9),
                bg=CARD_BG, fg=TEXT_COLOR, justify="left", anchor="w", wraplength=win_w - 75
            ).pack(fill=tk.X, pady=1)

    _add_section_card("🔍 考えられる原因", diag.get("causes", []), accent_col)
    _add_section_card("💡 対処方法", diag.get("actions", []), PRIMARY)

    # エラー原文エリア
    raw_lbl = tk.Label(body, text="▼ 詳細情報（システムからのエラー原文）", font=("Segoe UI", 8, "bold"), bg=BG_COLOR, fg=MUTED_TEXT, anchor="w")
    raw_lbl.pack(fill=tk.X, pady=(2, 2))
    raw_box = st.ScrolledText(body, wrap=tk.WORD, font=("Consolas", 8), bg="#F8FAFC", fg=MUTED_TEXT, relief=tk.SOLID, borderwidth=1, height=6, padx=8, pady=6)
    raw_box.pack(fill=tk.BOTH, expand=True)
    raw_box.insert(tk.END, diag.get("raw", ""))
    raw_box.configure(state=tk.DISABLED)

# ==============================
# API詳細設定ダイアログ
# ==============================
def open_api_settings_dialog():
    dialog = tk.Toplevel(state.root)
    dialog.title("⚙️ AI詳細設定 (Gemini API / Azure OpenAI Service)")

    screen_h = state.root.winfo_screenheight()
    screen_w = state.root.winfo_screenwidth()

    dialog_w = min(1200, screen_w - 40)
    dialog_h = min(800, screen_h - 80)
    dialog.configure(bg=BG_COLOR)
    center_window(dialog, state.root, dialog_w, dialog_h)
    dialog.grab_set()

    fav_lists = []
    models = load_models_list()
    model_combos = []
    gemini_mainstream_only_var = tk.BooleanVar(value=True)
    gemini_count_labels = []

    def get_visible_gemini_models():
        if gemini_mainstream_only_var.get():
            return filter_mainstream_gemini_models(models)
        return models

    def refresh_gemini_combos():
        visible = get_visible_gemini_models()
        values = [m[0] for m in visible]
        for cb in model_combos:
            cb.config(values=values)
        count_text = f"({len(visible)} / {len(models)}件)"
        for lbl in gemini_count_labels:
            lbl.config(text=count_text)

    def update_all_fav_lists():
        for f_list in fav_lists:
            f_list.set_items(state.saved_custom_prompts)

    original_values = {
        "engine": state.engine_var.get(),
        "plan": state.api_plan_var.get(),
        "key_free": state.api_key_free_var.get(), "key_paid": state.api_key_paid_var.get(),
        "model_free": state.gemini_model_free_var.get(), "model_paid": state.gemini_model_paid_var.get(),
        "rpm_free": safe_int_var(state.api_rpm_free_var, 12, 1, 2000),
        "rpm_paid": safe_int_var(state.api_rpm_paid_var, 300, 1, 2000),
        "temp_free": safe_float_var(state.temperature_free_var, 0.0, 0.0, 2.0),
        "temp_paid": safe_float_var(state.temperature_paid_var, 0.0, 0.0, 2.0),
        "safety_free": state.safety_free_var.get(), "safety_paid": state.safety_paid_var.get(),
        "tokens_free": safe_int_var(state.max_tokens_free_var, 8192, 1024, 2097152),
        "tokens_paid": safe_int_var(state.max_tokens_paid_var, 8192, 1024, 2097152),
        "prompt_free": state.custom_prompt_free_var.get(), "prompt_paid": state.custom_prompt_paid_var.get(),
        "threads_free": safe_int_var(state.threads_free_var, 1, 1, 20),
        "threads_paid": safe_int_var(state.threads_paid_var, 5, 1, 20),
        "azure_key": state.azure_api_key_var.get(),
        "azure_endpoint": state.azure_endpoint_var.get(),
        "azure_deployment": state.azure_deployment_var.get(),
        "azure_api_version": state.azure_api_version_var.get(),
        "azure_rpm": safe_int_var(state.azure_rpm_var, 60, 1, 2000),
        "azure_temp": safe_float_var(state.azure_temperature_var, 0.0, 0.0, 2.0),
        "azure_tokens": safe_int_var(state.azure_max_tokens_var, 4096, 1024, 128000),
        "azure_prompt": state.azure_custom_prompt_var.get(),
        "azure_threads": safe_int_var(state.azure_threads_var, 3, 1, 20),
        "saved_prompts": list(state.saved_custom_prompts)
    }

    def apply_and_close():
        # 数値Spinboxが空欄・範囲外だった場合は安全な値へ補正
        state.api_rpm_free_var.set(safe_int_var(state.api_rpm_free_var, 12, 1, 2000))
        state.api_rpm_paid_var.set(safe_int_var(state.api_rpm_paid_var, 300, 1, 2000))
        state.temperature_free_var.set(safe_float_var(state.temperature_free_var, 0.0, 0.0, 2.0))
        state.temperature_paid_var.set(safe_float_var(state.temperature_paid_var, 0.0, 0.0, 2.0))
        state.max_tokens_free_var.set(safe_int_var(state.max_tokens_free_var, 8192, 1024, 2097152))
        state.max_tokens_paid_var.set(safe_int_var(state.max_tokens_paid_var, 8192, 1024, 2097152))
        state.threads_free_var.set(safe_int_var(state.threads_free_var, 1, 1, 20))
        state.threads_paid_var.set(safe_int_var(state.threads_paid_var, 5, 1, 20))
        state.azure_rpm_var.set(safe_int_var(state.azure_rpm_var, 60, 1, 2000))
        state.azure_temperature_var.set(safe_float_var(state.azure_temperature_var, 0.0, 0.0, 2.0))
        state.azure_max_tokens_var.set(safe_int_var(state.azure_max_tokens_var, 4096, 1024, 128000))
        state.azure_threads_var.set(safe_int_var(state.azure_threads_var, 3, 1, 20))

        # Azure エンドポイントがフルURLで入力されていた場合は正規化・補完
        raw_ep = state.azure_endpoint_var.get().strip()
        if raw_ep:
            base_ep, dep, ver = parse_azure_endpoint_info(
                raw_ep,
                state.azure_deployment_var.get().strip(),
                state.azure_api_version_var.get().strip()
            )
            state.azure_endpoint_var.set(base_ep)
            if dep:
                state.azure_deployment_var.set(dep)
            if ver:
                state.azure_api_version_var.set(ver)

        if state.plan_indicator:
            engine = state.engine_var.get()
            if engine == "AzureOpenAI":
                dep_name = state.azure_deployment_var.get().strip() or "未設定"
                state.plan_indicator.config(text=f"🔷 Azure OpenAI / {dep_name}", foreground=PRIMARY)
            elif engine == "Gemini":
                plan = state.api_plan_var.get()
                m_name = state.gemini_model_free_var.get() if plan == "free" else state.gemini_model_paid_var.get()
                if plan == "free":
                    state.plan_indicator.config(text=f"🟢 無料枠 (Free) / {m_name}", foreground=COLOR_SUCCESS)
                else:
                    state.plan_indicator.config(text=f"🔵 課金枠 (Paid) / {m_name}", foreground=PRIMARY)

        # 設定ファイルへも自動保存し、次回起動時に失われないようにする
        if state.save_settings_callback:
            try:
                state.save_settings_callback()
            except Exception:
                pass

        dialog.destroy()

    def has_changes():
        if state.engine_var.get() != original_values["engine"]: return True
        if state.api_plan_var.get() != original_values["plan"]: return True
        if state.api_key_free_var.get() != original_values["key_free"]: return True
        if state.api_key_paid_var.get() != original_values["key_paid"]: return True
        if state.gemini_model_free_var.get() != original_values["model_free"]: return True
        if state.gemini_model_paid_var.get() != original_values["model_paid"]: return True
        if safe_int_var(state.api_rpm_free_var, original_values["rpm_free"]) != original_values["rpm_free"]: return True
        if safe_int_var(state.api_rpm_paid_var, original_values["rpm_paid"]) != original_values["rpm_paid"]: return True
        if safe_float_var(state.temperature_free_var, original_values["temp_free"]) != original_values["temp_free"]: return True
        if safe_float_var(state.temperature_paid_var, original_values["temp_paid"]) != original_values["temp_paid"]: return True
        if state.safety_free_var.get() != original_values["safety_free"]: return True
        if state.safety_paid_var.get() != original_values["safety_paid"]: return True
        if safe_int_var(state.max_tokens_free_var, original_values["tokens_free"]) != original_values["tokens_free"]: return True
        if safe_int_var(state.max_tokens_paid_var, original_values["tokens_paid"]) != original_values["tokens_paid"]: return True
        if state.custom_prompt_free_var.get() != original_values["prompt_free"]: return True
        if state.custom_prompt_paid_var.get() != original_values["prompt_paid"]: return True
        if safe_int_var(state.threads_free_var, original_values["threads_free"]) != original_values["threads_free"]: return True
        if safe_int_var(state.threads_paid_var, original_values["threads_paid"]) != original_values["threads_paid"]: return True
        if state.azure_api_key_var.get() != original_values["azure_key"]: return True
        if state.azure_endpoint_var.get() != original_values["azure_endpoint"]: return True
        if state.azure_deployment_var.get() != original_values["azure_deployment"]: return True
        if state.azure_api_version_var.get() != original_values["azure_api_version"]: return True
        if safe_int_var(state.azure_rpm_var, original_values["azure_rpm"]) != original_values["azure_rpm"]: return True
        if safe_float_var(state.azure_temperature_var, original_values["azure_temp"]) != original_values["azure_temp"]: return True
        if safe_int_var(state.azure_max_tokens_var, original_values["azure_tokens"]) != original_values["azure_tokens"]: return True
        if state.azure_custom_prompt_var.get() != original_values["azure_prompt"]: return True
        if safe_int_var(state.azure_threads_var, original_values["azure_threads"]) != original_values["azure_threads"]: return True
        if state.saved_custom_prompts != original_values["saved_prompts"]: return True
        return False

    def cancel_and_close():
        if has_changes():
            if not messagebox.askyesno("確認", "変更が適用されていません。\n破棄して設定画面を閉じますか？", parent=dialog):
                return

        state.engine_var.set(original_values["engine"])
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
        state.azure_api_key_var.set(original_values["azure_key"])
        state.azure_endpoint_var.set(original_values["azure_endpoint"])
        state.azure_deployment_var.set(original_values["azure_deployment"])
        state.azure_api_version_var.set(original_values["azure_api_version"])
        state.azure_rpm_var.set(original_values["azure_rpm"])
        state.azure_temperature_var.set(original_values["azure_temp"])
        state.azure_max_tokens_var.set(original_values["azure_tokens"])
        state.azure_custom_prompt_var.set(original_values["azure_prompt"])
        state.azure_threads_var.set(original_values["azure_threads"])
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

    btn_apply = ttk.Button(btn_action_frame, text="設定を適用して保存・閉じる", command=apply_and_close, style="Primary.TButton", width=26)
    btn_apply.pack(side=tk.LEFT, padx=10)
    ttk.Label(btn_action_frame, text="※適用すると設定ファイルへ自動保存されます", background=BG_COLOR, foreground=MUTED_TEXT, font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(6, 0))

    tab_reset_handlers = {}
    def reset_current_tab_all():
        try:
            current_tab = notebook.select()
            handler = tab_reset_handlers.get(current_tab)
            if handler:
                handler()
        except Exception:
            pass

    btn_reset_tab_all = ttk.Button(btn_action_frame, text="🔄 このタブを推奨値に戻す", command=reset_current_tab_all)
    btn_reset_tab_all.pack(side=tk.RIGHT, padx=15)

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

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # ==============================
    # コンテンツの配置
    # ==============================
    lbl_title = ttk.Label(scrollable_frame, text="AI エンジン詳細設定 (Gemini API / Azure OpenAI Service)", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
    lbl_title.pack(pady=(10, 5))

    plan_frame = ttk.LabelFrame(scrollable_frame, text=" 実際に使用するAIエンジン / プランの選択 ", style="Card.TLabelframe", padding=8)
    plan_frame.pack(fill=tk.X, padx=15, pady=(0, 5))

    plan_inner = ttk.Frame(plan_frame, style="Card.TFrame")
    plan_inner.pack(anchor="w", padx=5, pady=2)

    initial_profile = "azure" if state.engine_var.get() == "AzureOpenAI" else ("gemini_paid" if state.api_plan_var.get() == "paid" else "gemini_free")
    active_ai_profile_var = tk.StringVar(value=initial_profile)

    def on_profile_change():
        prof = active_ai_profile_var.get()
        if prof == "azure":
            state.engine_var.set("AzureOpenAI")
        elif prof == "gemini_paid":
            state.engine_var.set("Gemini")
            state.api_plan_var.set("paid")
        else:
            state.engine_var.set("Gemini")
            state.api_plan_var.set("free")

    ttk.Label(plan_inner, text="抽出時に使用するAI設定を選んでください（下のタブとは連動しません）:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 12))
    ttk.Radiobutton(plan_inner, text="Gemini API (無料枠)", variable=active_ai_profile_var, value="gemini_free", command=on_profile_change).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(plan_inner, text="Gemini API (課金枠)", variable=active_ai_profile_var, value="gemini_paid", command=on_profile_change).pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(plan_inner, text="Azure OpenAI Service", variable=active_ai_profile_var, value="azure", command=on_profile_change).pack(side=tk.LEFT)

    notebook = ttk.Notebook(scrollable_frame)
    notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

    tab_free = ttk.Frame(notebook, style="Main.TFrame")
    tab_paid = ttk.Frame(notebook, style="Main.TFrame")
    tab_azure = ttk.Frame(notebook, style="Main.TFrame")
    notebook.add(tab_free, text="  Gemini 無料枠 (Free)  ")
    notebook.add(tab_paid, text="  Gemini 課金枠 (Paid)  ")
    notebook.add(tab_azure, text="  Azure OpenAI Service  ")

    def build_prompt_section(parent_tab, prompt_var):
        """カスタムプロンプト設定セクション（Gemini / Azure OpenAI 共通）"""
        prompt_frame = ttk.LabelFrame(parent_tab, text=" ④ 独自の追加指示 (カスタムプロンプト) - 任意 ", style="Card.TLabelframe", padding=8)
        prompt_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        input_inner = ttk.Frame(prompt_frame, style="Card.TFrame")
        input_inner.pack(fill=tk.X, pady=(0, 8))

        entry_new_prompt = ttk.Entry(input_inner)
        entry_new_prompt.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        entry_new_prompt.bind("<Button-3>", lambda e, widget=entry_new_prompt: show_context_menu(e, widget))

        def sync_current_to_var():
            prompt_var.set('\n'.join(current_list.get_all_items()))

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

        def delete_current_selected():
            current_list.remove_selected()
            sync_current_to_var()

        def save_selected_to_fav():
            sel = current_list.get_selected_items()
            if not sel:
                return
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
            if not sel:
                return
            for text in sel:
                current_list.add_item(text)
            sync_current_to_var()

        def delete_fav_selected():
            sel = fav_list.get_selected_items()
            if not sel:
                return
            if messagebox.askyesno("確認", "選択したお気に入りを削除しますか？", parent=dialog):
                for text in sel:
                    if text in state.saved_custom_prompts:
                        state.saved_custom_prompts.remove(text)
                update_all_fav_lists()

        ttk.Button(right_actions, text="◀ 選択を左に追加", command=add_fav_to_current, style="Primary.TButton").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(right_actions, text="🗑 選択を削除", command=delete_fav_selected).pack(side=tk.LEFT)

        initial_prompts = [p for p in prompt_var.get().split('\n') if p.strip()]
        current_list.set_items(initial_prompts)

    def build_gemini_tab(parent_tab, plan_type):
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

        entry_key = ttk.Entry(key_inner, textvariable=key_var, width=50, show="*")
        entry_key.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
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
            if not key:
                return messagebox.showwarning("警告", "APIキーが入力されていません。", parent=dialog)
            model_name = m_var.get()
            try:
                btn_test.config(text="確認中...", state=tk.DISABLED)
                dialog.update()
                genai.configure(api_key=key)
                model = genai.GenerativeModel(model_name)
                model.generate_content(f"Connection Test: {time.time()}")
                messagebox.showinfo("テスト成功", f"APIキーは正しく認識されました。\nAIモデル「{model_name}」による通信は正常です！", parent=dialog)
            except Exception as e:
                diag = diagnose_api_error("gemini", e, model_or_deployment=model_name)
                show_api_error_dialog(dialog, diag)
            finally:
                btn_test.config(text="テスト", state=tk.NORMAL)

        btn_test = ttk.Button(key_inner, text="テスト", command=test_key, width=8)
        btn_test.pack(side=tk.LEFT)

        middle_frame = ttk.Frame(parent_tab, style="Main.TFrame")
        middle_frame.pack(fill=tk.X, padx=10, pady=5)

        middle_frame.columnconfigure(0, weight=3)
        middle_frame.columnconfigure(1, weight=2)

        perf_frame = ttk.LabelFrame(middle_frame, text=" ② モデル・パフォーマンス設定 ", style="Card.TLabelframe", padding=8)
        perf_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        model_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        model_inner.pack(fill=tk.X, pady=(0, 2))

        ttk.Label(model_inner, text="使用モデル:", width=10, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        model_combo = ttk.Combobox(model_inner, values=[m[0] for m in get_visible_gemini_models()], width=70)
        model_combos.append(model_combo)

        current_val = model_var.get()
        for m in models:
            if m[1] == current_val:
                model_combo.set(m[0])
                break
        if not model_combo.get():
            model_combo.set(current_val)

        def on_model_select(event=None, cb=model_combo, m_var=model_var, r_var=rpm_var, t_var=threads_var, tok_var=tokens_var, is_f=is_free):
            selected_display = cb.get()
            matched = False
            for m in models:
                if m[0] == selected_display:
                    m_var.set(m[1])
                    matched = True
                    break
            if not matched:
                m_var.set(selected_display)

            rec = get_recommended_gemini_settings(m_var.get(), is_f)
            r_var.set(rec["rpm"])
            t_var.set(rec["threads"])
            tok_var.set(rec["max_tokens"])

        model_combo.bind("<<ComboboxSelected>>", on_model_select)
        model_combo.bind("<FocusOut>", on_model_select)
        model_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def fetch_models(k_var=key_var, m_var=model_var, r_var=rpm_var, t_var=threads_var, tp_var=temp_var, tok_var=tokens_var, s_var=safety_var, is_f=is_free):
            key = k_var.get().strip()
            if not key:
                messagebox.showwarning("警告", "最新モデルを取得するには、APIキーを入力してください。", parent=dialog)
                return
            try:
                btn_fetch_models.config(text="取得中...", state=tk.DISABLED)
                dialog.update()

                new_models = fetch_gemini_models(key)

                if new_models:
                    models.clear()
                    models.extend(new_models)
                    save_models_list(models)
                    refresh_gemini_combos()

                    for m_var_ref, cb_ref in [(state.gemini_model_free_var, model_combos[0]), (state.gemini_model_paid_var, model_combos[1])]:
                        current_model_id = m_var_ref.get()
                        for new_m in models:
                            if new_m[1] == current_model_id:
                                cb_ref.set(new_m[0])
                                break

                    vis_cnt = len(get_visible_gemini_models())
                    mode_note = f"（主力モデル {vis_cnt}件 を表示中。すべて見る場合は「主力モデルのみ表示」のチェックを外してください）" if gemini_mainstream_only_var.get() else f"（全 {vis_cnt}件 を表示中）"

                    # 現在選択中モデルの推奨設定を確認
                    rec = get_recommended_gemini_settings(m_var.get(), is_f)
                    diffs = []
                    cur_rpm = safe_int_var(r_var, rec["rpm"])
                    cur_th = safe_int_var(t_var, rec["threads"])
                    cur_tp = safe_float_var(tp_var, rec["temperature"])
                    cur_tok = safe_int_var(tok_var, rec["max_tokens"])
                    if cur_rpm != rec["rpm"]:
                        diffs.append(f"・RPM: {cur_rpm} → {rec['rpm']}")
                    if cur_th != rec["threads"]:
                        diffs.append(f"・スレッド: {cur_th} → {rec['threads']}")
                    if abs(cur_tp - rec["temperature"]) > 1e-6:
                        diffs.append(f"・Temp: {cur_tp} → {rec['temperature']}")
                    if cur_tok != rec["max_tokens"]:
                        diffs.append(f"・最大トークン: {cur_tok} → {rec['max_tokens']}")

                    if diffs:
                        ask_msg = (
                            f"最新のモデルリスト (全{len(models)}件) を Gemini API から取得しました！\n"
                            f"{mode_note}\n\n"
                            f"【現在のモデル（{m_var.get()}）の推奨設定を確認しました】\n"
                            + "\n".join(diffs)
                            + "\n\nこれらの推奨値を適用しますか？"
                        )
                        if messagebox.askyesno("更新完了・推奨値の確認", ask_msg, parent=dialog):
                            r_var.set(rec["rpm"])
                            t_var.set(rec["threads"])
                            tp_var.set(rec["temperature"])
                            tok_var.set(rec["max_tokens"])
                            s_var.set(rec["disable_safety"])
                    else:
                        messagebox.showinfo(
                            "更新完了",
                            f"最新のモデルリスト (全{len(models)}件) を Gemini API から取得しました！\n"
                            f"{mode_note}\n\n"
                            f"✅ 現在の設定（RPM / スレッド / パラメータ）は推奨値と一致しています。",
                            parent=dialog
                        )
                else:
                    messagebox.showinfo("情報", "取得可能なモデルが見つかりませんでした。", parent=dialog)
            except Exception as e:
                diag = diagnose_api_error("gemini", e)
                show_api_error_dialog(dialog, diag)
            finally:
                btn_fetch_models.config(text="🌐 最新モデルを更新", state=tk.NORMAL)

        btn_fetch_models = ttk.Button(model_inner, text="🌐 最新モデルを更新", width=18, command=fetch_models)
        btn_fetch_models.pack(side=tk.LEFT, padx=(5, 0))

        link_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        link_inner.pack(fill=tk.X, pady=(2, 5))
        chk_mainstream = ttk.Checkbutton(link_inner, text="主力モデルのみ表示", variable=gemini_mainstream_only_var, command=refresh_gemini_combos, style="TCheckbutton")
        chk_mainstream.pack(side=tk.LEFT)
        lbl_g_cnt = ttk.Label(link_inner, text=f"({len(get_visible_gemini_models())} / {len(models)}件)", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 8))
        lbl_g_cnt.pack(side=tk.LEFT, padx=(2, 0))
        gemini_count_labels.append(lbl_g_cnt)
        lbl_link = ttk.Label(link_inner, text="🔗 各モデルの特徴を確認する (公式)", cursor="hand2", foreground=PRIMARY, font=("Segoe UI", 9, "underline"), background=CARD_BG)
        lbl_link.pack(side=tk.RIGHT, padx=5)
        lbl_link.bind("<Button-1>", lambda e: webbrowser.open_new("https://ai.google.dev/gemini-api/docs/models/gemini"))

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

            c_info = tk.Canvas(info_win, bg=BG_COLOR, highlightthickness=0)
            sb_info = ttk.Scrollbar(info_win, orient="vertical", command=c_info.yview)
            sf_info = ttk.Frame(c_info, style="Main.TFrame")

            sf_info.bind("<Configure>", lambda e: c_info.configure(scrollregion=c_info.bbox("all")))
            cw_info = c_info.create_window((0, 0), window=sf_info, anchor="nw")

            def _on_c_config(e):
                c_info.itemconfig(cw_info, width=e.width)
            c_info.bind("<Configure>", _on_c_config)

            c_info.configure(yscrollcommand=sb_info.set)
            c_info.pack(side="left", fill="both", expand=True)
            sb_info.pack(side="right", fill="y")

            lbl_t = ttk.Label(sf_info, text="Gemini API 仕様・制限一覧", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY)
            lbl_t.pack(pady=(15, 5))

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

            ttk.Label(sf_info, text="▼ プラン比較", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(10, 0))
            headers_plan = ["比較項目", "無料枠 (Free Tier)", "課金枠 (Paid Tier)"]
            data_plan = [
                ["利用料金", "完全無料（クレジットカード登録不要）", "従量課金（トークンと呼ばれるデータ量に応じて支払い）"],
                ["利用できるモデル", "Gemini 4 Hydrogen / 3.8 Flash / 3.5 Flash-Lite など", "すべての最新モデル（Gemini 4 Carbon / Neon / Argon 含む）"],
                ["データの\nプライバシー", "入力データがGoogleのAI学習に利用される可能性がある", "入力データはAI学習に利用されない"]
            ]
            create_table(sf_info, headers_plan, data_plan, [150, 360, 360])

            KNOWN_MODEL_INFO = {
                "gemini-4-hydrogen": {
                    "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "高スループット\n[推奨: 300 RPM / 並列(5〜)]"],
                    "desc": ["Gemini 4 標準主力モデル", "汎用性・スピード・マルチモーダル精度の高いバランス", "一般的な図面管理台帳やPDFのテキスト・表抽出（デフォルト推奨）"]
                },
                "gemini-4-argon": {
                    "limit": ["制限目安は公式参照\n[推奨: 5 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 200 RPM / 並列(5)]"],
                    "desc": ["高効率フロンティア推論", "高い知能と高速応答・トークン効率の両立", "複雑な表構造の解析と処理速度を両立したい場合"]
                },
                "gemini-4-neon": {
                    "limit": ["制限目安は公式参照\n[推奨: 2 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 150 RPM / 並列(5)]"],
                    "desc": ["構造化抽出・高度推論モデル", "複雑な構造化データ抽出と段階的思考", "手書きや変則レイアウトの帳票を精密に構造化したい場合"]
                },
                "gemini-4-carbon": {
                    "limit": ["非常に厳しい\n[推奨: 1 RPM / 直列(1)]", "時期・枠により変動\n[推奨: 100 RPM / 並列(3〜5)]"],
                    "desc": ["最上位フロンティアモデル", "極めて複雑な表構造解析、かすれた手書き文字、高度推論", "絶対にミスが許されない複雑なPDFのデータ抽出"]
                },
                "gemini-3.8-flash": {
                    "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                    "desc": ["高速・軽量Flash最新世代", "低遅延・大量ページ処理・マルチモーダル対応", "大量のPDFページを高速に処理したい場合"]
                },
                "gemini-3.5-flash-lite": {
                    "limit": ["15 RPM目安\n[推奨: 12 RPM / 直列(1)]", "1000 RPM\n[推奨: 300 RPM / 並列(5〜)]"],
                    "desc": ["最軽量・低コストモデル", "圧倒的な処理スピードと低コスト", "画質が良いPDFの単純なテキスト抽出、大量データの安価な処理"]
                }
            }

            data_limit = []
            data_model = []
            for m_display, m_id in models:
                info = KNOWN_MODEL_INFO.get(m_id)
                if not info:
                    for known_id, known_data in KNOWN_MODEL_INFO.items():
                        if known_id in m_id:
                            info = known_data
                            break
                if info:
                    data_limit.append([m_display, info["limit"][0], info["limit"][1]])
                    data_model.append([m_display, info["desc"][0], info["desc"][1], info["desc"][2]])
                else:
                    data_limit.append([m_display, "詳細は公式ドキュメントを参照", "詳細は公式ドキュメントを参照"])
                    data_model.append([m_display, "APIから取得した追加モデル", "-", "最新機能を試したい場合"])

            ttk.Label(sf_info, text="▼ 各モデルの制限目安 (RPM と 推奨スレッド数)", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(15, 0))
            create_table(sf_info, ["モデル名", "無料枠の制限目安\n(RPM / スレッド数)", "課金枠の制限目安\n(RPM / スレッド数)"], data_limit, [200, 335, 335])

            ttk.Label(sf_info, text="▼ 各モデルの特徴と適した用途", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(15, 0))
            create_table(sf_info, ["モデル名", "特徴", "得意なこと", "適した用途"], data_model, [160, 140, 270, 290])

            current_plan = "無料枠 (Free Tier)" if is_f else "課金枠 (Paid Tier)"
            current_model = m_var.get()
            status_text = f"【現在、このタブで選択中の設定】\nプラン: {current_plan}　／　モデル: {current_model}"
            ttk.Label(sf_info, text=status_text, font=("Segoe UI", 10, "bold"), background=CARD_BG, foreground=PRIMARY, relief="solid", borderwidth=1, padding=10).pack(fill=tk.X, padx=20, pady=15)

            ttk.Button(sf_info, text="閉じる", command=info_win.destroy, width=15).pack(pady=(0, 20))

        btn_show_limit = ttk.Button(perf_action_inner, text="ℹ️ 制限と仕様を確認", command=lambda m=model_var, f=is_free: show_limit_info(m, f))
        btn_show_limit.pack(side=tk.LEFT)

        def reset_perf(cb=model_combo, m_var=model_var, r_var=rpm_var, t_var=threads_var, is_f=is_free):
            target_model_val = "gemini-4-hydrogen"
            m_var.set(target_model_val)
            cb.set(target_model_val)
            for m in models:
                if m[1] == target_model_val:
                    cb.set(m[0])
                    break
            rec = get_recommended_gemini_settings(target_model_val, is_f)
            r_var.set(rec["rpm"])
            t_var.set(rec["threads"])

        btn_reset_perf = ttk.Button(perf_action_inner, text="🔄 推奨値", command=lambda cb=model_combo, m=model_var, r=rpm_var, t=threads_var, f=is_free: reset_perf(cb, m, r, t, f))
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

        def reset_param(t_var=temp_var, tok_var=tokens_var, s_var=safety_var, m_var=model_var, is_f=is_free):
            rec = get_recommended_gemini_settings(m_var.get(), is_f)
            t_var.set(rec["temperature"])
            tok_var.set(rec["max_tokens"])
            s_var.set(rec["disable_safety"])

        btn_reset_param = ttk.Button(param_row2, text="🔄 推奨値", command=reset_param)
        btn_reset_param.pack(side=tk.RIGHT, pady=(10, 0))

        def _reset_gemini_tab_all(cb=model_combo, m=model_var, r=rpm_var, th=threads_var, tp=temp_var, tok=tokens_var, sf=safety_var, f=is_free):
            reset_perf(cb, m, r, th, f)
            reset_param(tp, tok, sf, m, f)
            messagebox.showinfo("推奨値に戻しました", "モデル・RPM・スレッド・パラメータ設定をすべて推奨値に戻しました。", parent=dialog)

        tab_reset_handlers[str(parent_tab)] = _reset_gemini_tab_all

        build_prompt_section(parent_tab, prompt_var)

    def build_azure_tab(parent_tab):
        """Azure OpenAI Service 用の詳細設定タブ"""
        conn_frame = ttk.LabelFrame(parent_tab, text=" ① Azure OpenAI 接続設定 (エンドポイント / APIキー / デプロイ名) ", style="Card.TLabelframe", padding=8)
        conn_frame.pack(fill=tk.X, padx=10, pady=5)

        row_ep = ttk.Frame(conn_frame, style="Card.TFrame")
        row_ep.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row_ep, text="エンドポイントURL:", width=16, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        entry_ep = ttk.Entry(row_ep, textvariable=state.azure_endpoint_var, width=48)
        entry_ep.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        entry_ep.bind("<Button-3>", lambda e, widget=entry_ep: show_context_menu(e, widget))
        ttk.Label(row_ep, text="例: https://your-resource.openai.azure.com/", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        row_key = ttk.Frame(conn_frame, style="Card.TFrame")
        row_key.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(row_key, text="APIキー:", width=16, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        entry_az_key = ttk.Entry(row_key, textvariable=state.azure_api_key_var, width=48, show="*")
        entry_az_key.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        entry_az_key.bind("<Button-3>", lambda e, widget=entry_az_key: show_context_menu(e, widget))

        btn_toggle_az = ttk.Button(row_key, text="確認", width=6)
        btn_toggle_az.pack(side=tk.LEFT, padx=(0, 5))

        def toggle_az_key():
            if entry_az_key.cget('show') == '*':
                entry_az_key.configure(show='')
                btn_toggle_az.configure(text="隠す")
            else:
                entry_az_key.configure(show='*')
                btn_toggle_az.configure(text="確認")
        btn_toggle_az.config(command=toggle_az_key)

        row_dep = ttk.Frame(conn_frame, style="Card.TFrame")
        row_dep.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(row_dep, text="デプロイ名:", width=16, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        azure_models_all = load_azure_models_list()
        azure_mainstream_only_var = tk.BooleanVar(value=True)
        detected_az_api_ver = {"value": DEFAULT_AZURE_API_VERSION}

        def get_visible_azure_models():
            if azure_mainstream_only_var.get():
                return filter_mainstream_azure_models(azure_models_all)
            return azure_models_all

        combo_dep = ttk.Combobox(row_dep, textvariable=state.azure_deployment_var, values=get_visible_azure_models(), width=26)
        combo_dep.pack(side=tk.LEFT, padx=(0, 5))
        combo_dep.bind("<Button-3>", lambda e, widget=combo_dep: show_context_menu(e, widget))

        def on_azure_dep_select(event=None):
            rec = get_recommended_azure_settings(state.azure_deployment_var.get(), detected_az_api_ver["value"])
            state.azure_rpm_var.set(rec["rpm"])
            state.azure_threads_var.set(rec["threads"])
            state.azure_max_tokens_var.set(rec["max_tokens"])

        combo_dep.bind("<<ComboboxSelected>>", on_azure_dep_select)

        def refresh_azure_combo():
            visible = get_visible_azure_models()
            combo_dep.config(values=visible)
            lbl_az_cnt.config(text=f"({len(visible)} / {len(azure_models_all)}件)")

        def fetch_azure_models_ui():
            raw_ep = state.azure_endpoint_var.get().strip()
            key = state.azure_api_key_var.get().strip()
            dep = state.azure_deployment_var.get().strip()
            ver = state.azure_api_version_var.get().strip()

            base_ep, resolved_dep, resolved_ver = parse_azure_endpoint_info(raw_ep, dep, ver)
            if base_ep:
                state.azure_endpoint_var.set(base_ep)
            if resolved_dep and not dep:
                state.azure_deployment_var.set(resolved_dep)
            if resolved_ver:
                state.azure_api_version_var.set(resolved_ver)

            if not base_ep:
                return messagebox.showwarning("警告", "最新モデルを取得するには、エンドポイントURLを入力してください。", parent=dialog)
            if not key:
                return messagebox.showwarning("警告", "最新モデルを取得するには、APIキーを入力してください。", parent=dialog)

            try:
                btn_fetch_az_models.config(text="取得中...", state=tk.DISABLED)
                dialog.update()

                new_az_models, best_ver = fetch_azure_models_with_meta(key, base_ep, resolved_ver)
                detected_az_api_ver["value"] = best_ver
                if new_az_models:
                    azure_models_all.clear()
                    azure_models_all.extend(new_az_models)
                    save_azure_models_list(azure_models_all)
                    refresh_azure_combo()
                    vis_models = get_visible_azure_models()
                    if not state.azure_deployment_var.get().strip() and vis_models:
                        state.azure_deployment_var.set(vis_models[0])

                    cur_dep = state.azure_deployment_var.get().strip()
                    rec = get_recommended_azure_settings(cur_dep, best_ver)
                    diffs = []
                    cur_ver = state.azure_api_version_var.get().strip()
                    cur_rpm = safe_int_var(state.azure_rpm_var, rec["rpm"])
                    cur_th = safe_int_var(state.azure_threads_var, rec["threads"])
                    cur_tp = safe_float_var(state.azure_temperature_var, rec["temperature"])
                    cur_tok = safe_int_var(state.azure_max_tokens_var, rec["max_tokens"])

                    if cur_ver != rec["api_version"]:
                        diffs.append(f"・APIバージョン: {cur_ver or '未設定'} → {rec['api_version']}")
                    if cur_rpm != rec["rpm"]:
                        diffs.append(f"・RPM: {cur_rpm} → {rec['rpm']}")
                    if cur_th != rec["threads"]:
                        diffs.append(f"・スレッド: {cur_th} → {rec['threads']}")
                    if abs(cur_tp - rec["temperature"]) > 1e-6:
                        diffs.append(f"・Temp: {cur_tp} → {rec['temperature']}")
                    if cur_tok != rec["max_tokens"]:
                        diffs.append(f"・最大トークン: {cur_tok} → {rec['max_tokens']}")

                    mode_note = f"（主力モデル {len(vis_models)}件 を表示中。すべて見る場合は「主力モデルのみ表示」のチェックを外してください）" if azure_mainstream_only_var.get() else f"（全 {len(vis_models)}件 を表示中）"

                    if diffs:
                        ask_msg = (
                            f"最新のモデルリスト (全{len(azure_models_all)}件) を Azure から取得しました！\n"
                            f"{mode_note}\n\n"
                            f"【現在のデプロイ（{cur_dep or '未選択'}）に対する推奨設定を確認しました】\n"
                            + "\n".join(diffs)
                            + "\n\nこれらの推奨値を適用しますか？\n"
                            f"（※Azure側で独自のデプロイ名を付けている場合はそのまま維持されます）"
                        )
                        if messagebox.askyesno("更新完了・推奨値の確認", ask_msg, parent=dialog):
                            state.azure_api_version_var.set(rec["api_version"])
                            state.azure_rpm_var.set(rec["rpm"])
                            state.azure_threads_var.set(rec["threads"])
                            state.azure_temperature_var.set(rec["temperature"])
                            state.azure_max_tokens_var.set(rec["max_tokens"])
                    else:
                        messagebox.showinfo(
                            "更新完了",
                            f"最新のモデルリスト (全{len(azure_models_all)}件) を Azure から取得しました！\n"
                            f"{mode_note}\n\n"
                            f"✅ 現在の設定（APIバージョン / RPM / スレッド / パラメータ）は推奨値と一致しています。",
                            parent=dialog
                        )
                else:
                    messagebox.showinfo("情報", "取得可能なチャット対応モデルが見つかりませんでした。", parent=dialog)
            except Exception as e:
                diag = diagnose_api_error("azure", e, model_or_deployment=resolved_dep, endpoint=base_ep)
                show_api_error_dialog(dialog, diag)
            finally:
                btn_fetch_az_models.config(text="🌐 最新モデルを更新", state=tk.NORMAL)

        btn_fetch_az_models = ttk.Button(row_dep, text="🌐 最新モデルを更新", width=18, command=fetch_azure_models_ui)
        btn_fetch_az_models.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row_dep, text="APIバージョン:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        combo_ver = ttk.Combobox(row_dep, textvariable=state.azure_api_version_var, values=AZURE_API_VERSION_CANDIDATES, width=18)
        combo_ver.pack(side=tk.LEFT, padx=(0, 8))

        row_dep_opt = ttk.Frame(conn_frame, style="Card.TFrame")
        row_dep_opt.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(row_dep_opt, text="", width=16, background=CARD_BG).pack(side=tk.LEFT)
        chk_az_mainstream = ttk.Checkbutton(row_dep_opt, text="主力モデルのみ表示", variable=azure_mainstream_only_var, command=refresh_azure_combo, style="TCheckbutton")
        chk_az_mainstream.pack(side=tk.LEFT)
        lbl_az_cnt = ttk.Label(row_dep_opt, text=f"({len(get_visible_azure_models())} / {len(azure_models_all)}件)", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 8))
        lbl_az_cnt.pack(side=tk.LEFT, padx=(2, 12))
        ttk.Label(row_dep_opt, text="※独自のデプロイ名を付けている場合は直接入力も可能です", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        def test_azure():
            raw_ep = state.azure_endpoint_var.get().strip()
            key = state.azure_api_key_var.get().strip()
            dep = state.azure_deployment_var.get().strip()
            ver = state.azure_api_version_var.get().strip()

            base_ep, resolved_dep, resolved_ver = parse_azure_endpoint_info(raw_ep, dep, ver)
            if base_ep:
                state.azure_endpoint_var.set(base_ep)
            if resolved_dep and not dep:
                state.azure_deployment_var.set(resolved_dep)
            if resolved_ver:
                state.azure_api_version_var.set(resolved_ver)

            if not base_ep:
                return messagebox.showwarning("警告", "エンドポイントURLを入力してください。", parent=dialog)
            if not key:
                return messagebox.showwarning("警告", "APIキーを入力してください。", parent=dialog)
            if not resolved_dep:
                return messagebox.showwarning("警告", "デプロイ名（モデルデプロイ名）を入力してください。", parent=dialog)

            try:
                btn_test_az.config(text="通信中...", state=tk.DISABLED)
                dialog.update()
                test_azure_connection(key, base_ep, resolved_dep, resolved_ver)
                messagebox.showinfo(
                    "テスト成功",
                    f"Azure OpenAI Service への接続に成功しました！\n\n"
                    f"・エンドポイント: {base_ep}\n"
                    f"・デプロイ名: {resolved_dep}\n"
                    f"・APIバージョン: {resolved_ver}",
                    parent=dialog
                )
            except Exception as e:
                diag = diagnose_api_error("azure", e, model_or_deployment=resolved_dep, endpoint=base_ep)
                show_api_error_dialog(dialog, diag)
            finally:
                btn_test_az.config(text="🔌 接続テスト", state=tk.NORMAL)

        btn_test_az = ttk.Button(row_dep, text="🔌 接続テスト", command=test_azure, style="Primary.TButton", width=14)
        btn_test_az.pack(side=tk.LEFT, padx=(5, 0))

        middle_frame = ttk.Frame(parent_tab, style="Main.TFrame")
        middle_frame.pack(fill=tk.X, padx=10, pady=5)
        middle_frame.columnconfigure(0, weight=3)
        middle_frame.columnconfigure(1, weight=2)

        perf_frame = ttk.LabelFrame(middle_frame, text=" ② パフォーマンス・並列処理設定 ", style="Card.TLabelframe", padding=8)
        perf_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        speed_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        speed_inner.pack(fill=tk.X, pady=4)

        ttk.Label(speed_inner, text="RPM (1分あたりの最大リクエスト数):", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        spin_az_rpm = ttk.Spinbox(speed_inner, from_=1, to=2000, textvariable=state.azure_rpm_var, width=6)
        spin_az_rpm.pack(side=tk.LEFT, padx=(5, 15))

        ttk.Label(speed_inner, text="スレッド (並列数):", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(5, 5))
        spin_az_threads = ttk.Spinbox(speed_inner, from_=1, to=20, textvariable=state.azure_threads_var, width=5)
        spin_az_threads.pack(side=tk.LEFT)

        link_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        link_inner.pack(fill=tk.X, pady=(2, 4))
        lbl_az_link = ttk.Label(link_inner, text="🔗 Azure OpenAI 公式ドキュメントを確認する", cursor="hand2", foreground=PRIMARY, font=("Segoe UI", 9, "underline"), background=CARD_BG)
        lbl_az_link.pack(side=tk.RIGHT, padx=5)
        lbl_az_link.bind("<Button-1>", lambda e: webbrowser.open_new("https://learn.microsoft.com/ja-jp/azure/ai-services/openai/"))

        perf_action_inner = ttk.Frame(perf_frame, style="Card.TFrame")
        perf_action_inner.pack(fill=tk.X, pady=(6, 0))

        def show_azure_guide():
            info_win = tk.Toplevel(dialog)
            info_win.title("Azure OpenAI Service 仕様・設定ガイド")

            info_w = min(950, screen_w - 40)
            info_h = min(680, screen_h - 80)
            info_win.geometry(f"{info_w}x{info_h}")
            info_win.configure(bg=BG_COLOR)
            info_win.grab_set()

            raw_x = dialog.winfo_x() + 30
            raw_y = dialog.winfo_y() + 30
            win_x = max(10, min(raw_x, screen_w - info_w - 10))
            win_y = max(10, min(raw_y, screen_h - info_h - 40))
            info_win.geometry(f"+{win_x}+{win_y}")

            c_info = tk.Canvas(info_win, bg=BG_COLOR, highlightthickness=0)
            sb_info = ttk.Scrollbar(info_win, orient="vertical", command=c_info.yview)
            sf_info = ttk.Frame(c_info, style="Main.TFrame")

            sf_info.bind("<Configure>", lambda e: c_info.configure(scrollregion=c_info.bbox("all")))
            cw_info = c_info.create_window((0, 0), window=sf_info, anchor="nw")

            def _on_c_config(e):
                c_info.itemconfig(cw_info, width=e.width)
            c_info.bind("<Configure>", _on_c_config)

            c_info.configure(yscrollcommand=sb_info.set)
            c_info.pack(side="left", fill="both", expand=True)
            sb_info.pack(side="right", fill="y")

            ttk.Label(sf_info, text="🔷 Azure OpenAI Service 設定・モデルガイド", font=("Segoe UI", 16, "bold"), background=BG_COLOR, foreground=PRIMARY).pack(pady=(15, 5))

            def create_az_table(parent, headers, data, col_widths):
                table_frame = tk.Frame(parent, bg=BORDER_COLOR)
                table_frame.pack(fill=tk.X, expand=True, padx=20, pady=5)
                for col_idx, w in enumerate(col_widths):
                    table_frame.columnconfigure(col_idx, weight=1, minsize=w)
                for col_idx, header_text in enumerate(headers):
                    lbl = tk.Label(table_frame, text=header_text, font=("Segoe UI", 9, "bold"), bg="#E7F1FF", fg=PRIMARY, padx=8, pady=8, wraplength=col_widths[col_idx]-16)
                    lbl.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
                for row_idx, row_data in enumerate(data, 1):
                    for col_idx, cell_text in enumerate(row_data):
                        lbl = tk.Label(table_frame, text=cell_text, font=("Segoe UI", 9), bg="white", fg=TEXT_COLOR, padx=8, pady=8, justify="left", anchor="nw", wraplength=col_widths[col_idx]-16)
                        lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)

            ttk.Label(sf_info, text="▼ 1. 接続設定に必要な項目（Azure Portal / Azure AI Foundry で確認）", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(10, 0))
            create_az_table(
                sf_info,
                ["設定項目", "入力内容の例", "確認場所・ポイント"],
                [
                    ["エンドポイントURL", "https://xxxx.openai.azure.com/", "Azure Portal の「キーとエンドポイント」画面に表示されるURLです。\n※誤って長いターゲットURI全体を貼り付けても自動でベースURLを抽出します。"],
                    ["APIキー", "32文字以上の英数字キー", "「KEY 1」または「KEY 2」のいずれかをコピーして貼り付けます。"],
                    ["デプロイ名", "gpt-5 / gpt-5-mini / gpt-5.2 など", "Azure AI Foundry の「デプロイ」画面で作成したモデルの名前です。"],
                    ["APIバージョン", f"{DEFAULT_AZURE_API_VERSION} (推奨)", f"通常は推奨値「{DEFAULT_AZURE_API_VERSION}」のままで問題ありません。「🌐 最新モデルを更新」を押すと利用可能な最新バージョンを自動確認できます。"]
                ],
                [150, 230, 480]
            )

            ttk.Label(sf_info, text="▼ 2. 画像抽出（Vision）に対応した推奨モデル一覧", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=TEXT_COLOR).pack(anchor="w", padx=20, pady=(15, 0))
            create_az_table(
                sf_info,
                ["モデル種別", "特徴", "おすすめの用途", "推奨設定目安 (RPM / 並列数)"],
                [
                    ["gpt-5.2 / gpt-5.1\ngpt-5 / gpt-6", "最新主力フラグシップ", "手書き文字、複雑な表構造、図面管理台帳の精密な抽出（最もおすすめ）", "60 RPM / 並列 3〜5"],
                    ["gpt-5-mini\ngpt-5-nano", "高速・低コスト・軽量主力", "活字メインのスキャンPDF、大量ページの高速処理・コスト重視の運用", "120 RPM / 並列 5〜10"],
                    ["o4-mini", "高度な推論モデル", "極めて複雑なレイアウトや文脈判断が必要な帳票（※本アプリ側で自動的に推論モデル用パラメータへ切り替わります）", "30 RPM / 並列 1〜3"]
                ],
                [150, 180, 330, 200]
            )

            ttk.Button(sf_info, text="閉じる", command=info_win.destroy, width=15).pack(pady=(15, 20))

        ttk.Button(perf_action_inner, text="ℹ️ 設定ガイドを確認", command=show_azure_guide).pack(side=tk.LEFT)

        def reset_az_perf():
            rec = get_recommended_azure_settings(state.azure_deployment_var.get(), detected_az_api_ver["value"])
            state.azure_api_version_var.set(rec["api_version"])
            state.azure_rpm_var.set(rec["rpm"])
            state.azure_threads_var.set(rec["threads"])

        ttk.Button(perf_action_inner, text="🔄 推奨値", command=reset_az_perf).pack(side=tk.RIGHT)

        param_frame = ttk.LabelFrame(middle_frame, text=" ③ AI抽出パラメータ設定 ", style="Card.TLabelframe", padding=8)
        param_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        param_row1 = ttk.Frame(param_frame, style="Card.TFrame")
        param_row1.pack(fill=tk.X, pady=(4, 5))

        ttk.Label(param_row1, text="Temp:", width=6, background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        spin_az_temp = ttk.Spinbox(param_row1, from_=0.0, to=2.0, increment=0.1, textvariable=state.azure_temperature_var, width=4)
        spin_az_temp.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(param_row1, text="最大トークン:", background=CARD_BG, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
        spin_az_tokens = ttk.Spinbox(param_row1, from_=1024, to=128000, increment=1024, textvariable=state.azure_max_tokens_var, width=8)
        spin_az_tokens.pack(side=tk.LEFT, padx=(0, 5))

        param_row2 = ttk.Frame(param_frame, style="Card.TFrame")
        param_row2.pack(fill=tk.X, pady=2)
        ttk.Label(param_row2, text="※推論モデル(GPT-5/6, o4等)は自動で最適パラメータに調整されます", background=CARD_BG, foreground=MUTED_TEXT, font=("Segoe UI", 8)).pack(side=tk.LEFT)

        def reset_az_param():
            rec = get_recommended_azure_settings(state.azure_deployment_var.get(), detected_az_api_ver["value"])
            state.azure_temperature_var.set(rec["temperature"])
            state.azure_max_tokens_var.set(rec["max_tokens"])

        ttk.Button(param_row2, text="🔄 推奨値", command=reset_az_param).pack(side=tk.RIGHT, pady=(10, 0))

        def _reset_azure_tab_all():
            reset_az_perf()
            reset_az_param()
            rec = get_recommended_azure_settings(state.azure_deployment_var.get(), detected_az_api_ver["value"])
            messagebox.showinfo(
                "推奨値に戻しました",
                f"Azure OpenAI Service の設定をすべて推奨値に戻しました。\n\n"
                f"・APIバージョン: {rec['api_version']}\n"
                f"・RPM: {rec['rpm']} / スレッド: {rec['threads']}\n"
                f"・Temp: {rec['temperature']} / 最大トークン: {rec['max_tokens']}",
                parent=dialog
            )

        tab_reset_handlers[str(parent_tab)] = _reset_azure_tab_all

        build_prompt_section(parent_tab, state.azure_custom_prompt_var)

    build_gemini_tab(tab_free, "free")
    build_gemini_tab(tab_paid, "paid")
    build_azure_tab(tab_azure)

    update_all_fav_lists()

    if state.engine_var.get() == "AzureOpenAI":
        notebook.select(tab_azure)
    elif state.api_plan_var.get() == "free":
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
        self.top.protocol("WM_DELETE_WINDOW", self.cancel_and_close)

        self.pdf_path = pdf_path
        self.zoom = 1.0
        self.zoom_mode = False
        self.is_line_mode = (state.extract_mode_var.get() == "text")

        # 1. ヘッダーエリア
        header_area = ttk.Frame(self.top, style="Card.TFrame", padding=(15, 12))
        header_area.pack(fill=tk.X)

        title_lbl = ttk.Label(header_area, text="✂️ 抽出範囲の編集", font=("Segoe UI", 12, "bold"), foreground=PRIMARY, background=CARD_BG)
        title_lbl.pack(side=tk.LEFT)

        self.help_lbl = ttk.Label(header_area, text="", foreground=MUTED_TEXT, font=("Segoe UI", 10), background=CARD_BG)
        self.help_lbl.pack(side=tk.LEFT, padx=(25, 0))

        self.status_lbl = ttk.Label(header_area, text="", foreground=COLOR_PURPLE, font=("Segoe UI", 9, "bold"), background=CARD_BG)
        self.status_lbl.pack(side=tk.RIGHT)

        # 2. ツールバーエリア
        toolbar = ttk.Frame(self.top, padding=(10, 8))
        toolbar.pack(fill=tk.X)

        edit_grp = ttk.LabelFrame(toolbar, text=" 編集ツール ", padding=8)
        edit_grp.pack(side=tk.LEFT, padx=5)
        ttk.Button(edit_grp, text="↩️ 戻る", command=self.undo, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(edit_grp, text="↪️ 進む", command=self.redo, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(edit_grp, text="🗑 全て消去", command=self.clear_rects, style="Warning.TButton", width=12).pack(side=tk.LEFT, padx=(10, 2))

        view_grp = ttk.LabelFrame(toolbar, text=" 表示・ズーム ", padding=8)
        view_grp.pack(side=tk.LEFT, padx=5)
        self.btn_zoom_range = ttk.Button(view_grp, text="🔍 範囲で拡大", command=self.toggle_zoom_mode, width=15)
        self.btn_zoom_range.pack(side=tk.LEFT, padx=2)
        ttk.Button(view_grp, text="🏠 全体表示", command=self.zoom_fit, width=10).pack(side=tk.LEFT, padx=2)

        if self.is_line_mode:
            self.line_dir = tk.StringVar(value="h")
            dir_grp = ttk.LabelFrame(toolbar, text=" 抽出設定 ", padding=8)
            dir_grp.pack(side=tk.LEFT, padx=5)
            ttk.Radiobutton(dir_grp, text="横書き (水平線)", variable=self.line_dir, value="h", command=self.update_help_text).pack(side=tk.LEFT, padx=5)
            ttk.Radiobutton(dir_grp, text="縦書き (垂直線)", variable=self.line_dir, value="v", command=self.update_help_text).pack(side=tk.LEFT, padx=5)

        app_grp = ttk.Frame(toolbar, padding=(5, 10))
        app_grp.pack(side=tk.RIGHT, padx=5)
        ttk.Button(app_grp, text="中止", command=self.cancel_and_close, width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(app_grp, text="✅ 設定を確定して閉じる", command=self.save_and_close, style="Primary.TButton", width=22).pack(side=tk.RIGHT, padx=5)

        ttk.Separator(self.top, orient="horizontal").pack(fill=tk.X, pady=(5, 10))

        self.update_help_text()

        # 3. 操作ガイド（フッター）
        footer = ttk.Frame(self.top, style="Card.TFrame", padding=(15, 5))
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        guide_text = "🖱️ 【マウス操作】 ドラッグ: 範囲選択 / Ctrl+ホイール: 拡大縮小 / Shift+ホイール: 横スクロール / ホイール: 縦スクロール　　⌨️ 【ショートカット】 Ctrl+Z: 戻る / Ctrl+Y: 進む"
        ttk.Label(footer, text=guide_text, font=("Segoe UI", 9), foreground=MUTED_TEXT, background=CARD_BG).pack(side=tk.LEFT)

        canvas_frame = ttk.Frame(self.top)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))
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

        self.start_x, self.start_y, self.current_rect, self.rectangles, self.redo_stack = None, None, None, [], []

        # 既存の選択範囲（プリセット読込済など）があれば復元
        if state.selected_crop_regions:
            for region in state.selected_crop_regions:
                if len(region) >= 4:
                    rx1, ry1, rx2, ry2 = region[:4]
                    is_vert = bool(region[4]) if len(region) > 4 else False
                    is_line = is_vert or abs(ry2 - ry1) < 0.03 or abs(rx2 - rx1) < 0.03
                    self.rectangles.append({
                        'id': None,
                        'rx1': rx1,
                        'ry1': ry1,
                        'rx2': rx2,
                        'ry2': ry2,
                        'is_line': is_line,
                        'is_vertical': is_vert
                    })

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        if sys.platform.startswith("win"):
            self.canvas.bind("<MouseWheel>", self.on_mousewheel_y)
            self.canvas.bind("<Shift-MouseWheel>", self.on_mousewheel_x)
            self.canvas.bind("<Control-MouseWheel>", self.on_mousewheel_zoom)
        self.top.bind("<Control-z>", lambda e: self.undo())
        self.top.bind("<Control-y>", lambda e: self.redo())

        try:
            self.set_status("📄 PDFファイルを読み込み中...")
            self.doc = fitz.open(pdf_path)
            self.page = self.doc[0]

            try:
                self.top.state('zoomed')
            except Exception:
                pass
            self.top.update()

            self.zoom_fit()
            self.set_status("")
        except Exception as e:
            if hasattr(self, 'doc') and self.doc:
                self.doc.close()
            self.top.destroy()
            raise Exception(f"プレビュー生成失敗: {e}")

    def set_status(self, text):
        if hasattr(self, 'status_lbl') and self.status_lbl:
            self.status_lbl.config(text=text)
            self.top.update_idletasks()

    def draw_image(self):
        self.set_status("🖼️ プレビュー画像を生成中...")
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = self.page.get_pixmap(matrix=mat)
        self.tk_image = ImageTk.PhotoImage(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_image)
        self.canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.img_w, self.img_h = pix.width, pix.height
        self.set_status("✏️ 選択範囲を再描画中...")
        for r in self.rectangles:
            if r.get('is_line'):
                if r.get('is_vertical'):
                    mid_x = r['rx1'] * self.img_w + (r['rx2'] - r['rx1']) * 0.5 * self.img_w
                    r['id'] = self.canvas.create_line(mid_x, r['ry1'] * self.img_h, mid_x, r['ry2'] * self.img_h, fill="red", width=2)
                else:
                    mid_y = r['ry1'] * self.img_h + (r['ry2'] - r['ry1']) * 0.5 * self.img_h
                    r['id'] = self.canvas.create_line(r['rx1'] * self.img_w, mid_y, r['rx2'] * self.img_w, mid_y, fill="red", width=2)
            else:
                r['id'] = self.canvas.create_rectangle(r['rx1'] * self.img_w, r['ry1'] * self.img_h, r['rx2'] * self.img_w, r['ry2'] * self.img_h, outline="red", width=2)
        self.set_status("")

    def toggle_zoom_mode(self):
        self.zoom_mode = not self.zoom_mode
        self.btn_zoom_range.config(style="Primary.TButton" if self.zoom_mode else "TButton")
        self.canvas.config(cursor="plus" if self.zoom_mode else "cross")
        self.update_help_text()

    def update_help_text(self):
        if self.zoom_mode:
            text = "【ズーム】拡大したい範囲をマウスで囲んでください。"
        elif self.is_line_mode:
            if self.line_dir.get() == "v":
                text = "【使い方】ドラッグで「垂直の線」を引き、抽出したい「列（縦書き）」の位置を指定します。"
            else:
                text = "【使い方】ドラッグで「水平の線」を引き、抽出したい「行（横書き）」の位置を指定します。"
        else:
            text = "【使い方】抽出したい範囲をマウスでドラッグして囲みます。"
        self.help_lbl.config(text=text)

    def zoom_in(self, event=None):
        x = event.x if event else self.canvas.winfo_width() // 2
        y = event.y if event else self.canvas.winfo_height() // 2
        self.zoom_at_pos(1.2, x, y)

    def zoom_out(self, event=None):
        x = event.x if event else self.canvas.winfo_width() // 2
        y = event.y if event else self.canvas.winfo_height() // 2
        self.zoom_at_pos(1 / 1.2, x, y)

    def zoom_fit(self):
        self.zoom_mode = False
        self.btn_zoom_range.config(style="TButton")
        self.canvas.config(cursor="cross")
        self.update_help_text()

        cw = self.canvas.winfo_width() - 40
        ch = self.canvas.winfo_height() - 40

        if cw < 100 or ch < 100:
            win_h = self.top.winfo_height()
            win_w = self.top.winfo_width()
            if win_h < 100:
                win_h = self.top.winfo_screenheight() * 0.8
                win_w = self.top.winfo_screenwidth() * 0.8
            cw, ch = win_w - 60, win_h - 180

        zoom_w = cw / self.page.rect.width
        zoom_h = ch / self.page.rect.height

        self.zoom = min(zoom_w, zoom_h)
        self.zoom = max(0.2, min(self.zoom, 3.0))

        self.draw_image()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def zoom_at_pos(self, factor, x, y):
        old_img_x = self.canvas.canvasx(x)
        old_img_y = self.canvas.canvasy(y)

        new_zoom = self.zoom * factor
        if new_zoom < 0.1:
            new_zoom = 0.1
        if new_zoom > 10.0:
            new_zoom = 10.0

        actual_factor = new_zoom / self.zoom
        self.zoom = new_zoom
        self.draw_image()

        new_img_x = old_img_x * actual_factor
        new_img_y = old_img_y * actual_factor

        self.canvas.xview_moveto((new_img_x - x) / self.img_w)
        self.canvas.yview_moveto((new_img_y - y) / self.img_h)

    def on_mousewheel_y(self, event):
        if event.state & 0x0001 or event.state & 0x0004:
            return "break"
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def on_mousewheel_x(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def on_mousewheel_zoom(self, event):
        if event.delta > 0:
            self.zoom_in(event)
        else:
            self.zoom_out(event)
        return "break"

    def on_press(self, event):
        self.start_x, self.start_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        if self.zoom_mode:
            self.current_rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline=PRIMARY, width=2, dash=(4, 4))
        elif self.is_line_mode:
            self.current_rect = self.canvas.create_line(self.start_x, self.start_y, self.start_x, self.start_y, fill="red", width=2, dash=(4, 4))
        else:
            self.current_rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red", width=2, dash=(4, 4))

    def on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        if self.zoom_mode:
            self.canvas.coords(self.current_rect, self.start_x, self.start_y, cur_x, cur_y)
        elif self.is_line_mode:
            if self.line_dir.get() == "v":
                self.canvas.coords(self.current_rect, self.start_x, self.start_y, self.start_x, cur_y)
            else:
                self.canvas.coords(self.current_rect, self.start_x, cur_y, cur_x, cur_y)
        else:
            self.canvas.coords(self.current_rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        end_x, end_y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

        if self.zoom_mode:
            self.canvas.delete(self.current_rect)
            rw, rh = abs(end_x - self.start_x), abs(end_y - self.start_y)
            if rw > 10 and rh > 10:
                cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
                factor = min(cw / rw, ch / rh)
                self.zoom *= factor
                self.draw_image()
                self.canvas.xview_moveto(min(self.start_x, end_x) * factor / self.img_w)
                self.canvas.yview_moveto(min(self.start_y, end_y) * factor / self.img_h)
            self.toggle_zoom_mode()
            return

        if self.is_line_mode:
            self.canvas.itemconfig(self.current_rect, dash=())
            is_vert = (self.line_dir.get() == "v")
            tolerance = 0.005

            if is_vert:
                rx = end_x / self.img_w
                self.rectangles.append({
                    'id': self.current_rect,
                    'rx1': max(0, rx - tolerance),
                    'ry1': min(self.start_y, end_y) / self.img_h,
                    'rx2': min(1.0, rx + tolerance),
                    'ry2': max(self.start_y, end_y) / self.img_h,
                    'is_line': True,
                    'is_vertical': True
                })
            else:
                ry = end_y / self.img_h
                self.rectangles.append({
                    'id': self.current_rect,
                    'rx1': min(self.start_x, end_x) / self.img_w,
                    'ry1': max(0, ry - tolerance),
                    'rx2': max(self.start_x, end_x) / self.img_w,
                    'ry2': min(1.0, ry + tolerance),
                    'is_line': True,
                    'is_vertical': False
                })
            self.redo_stack.clear()
        elif abs(end_x - self.start_x) > 10 and abs(end_y - self.start_y) > 10:
            self.canvas.itemconfig(self.current_rect, dash=())
            self.rectangles.append({
                'id': self.current_rect,
                'rx1': min(self.start_x, end_x) / self.img_w,
                'ry1': min(self.start_y, end_y) / self.img_h,
                'rx2': max(self.start_x, end_x) / self.img_w,
                'ry2': max(self.start_y, end_y) / self.img_h,
                'is_line': False
            })
            self.redo_stack.clear()
        else:
            self.canvas.delete(self.current_rect)

    def undo(self):
        if self.rectangles:
            rect = self.rectangles.pop()
            self.canvas.delete(rect['id'])
            self.redo_stack.append(rect)

    def redo(self):
        if self.redo_stack:
            rect = self.redo_stack.pop()
            if rect.get('is_line'):
                if rect.get('is_vertical'):
                    mid_x = rect['rx1'] * self.img_w + (rect['rx2'] - rect['rx1']) * 0.5 * self.img_w
                    rect['id'] = self.canvas.create_line(mid_x, rect['ry1'] * self.img_h, mid_x, rect['ry2'] * self.img_h, fill="red", width=2)
                else:
                    mid_y = rect['ry1'] * self.img_h + (rect['ry2'] - rect['ry1']) * 0.5 * self.img_h
                    rect['id'] = self.canvas.create_line(rect['rx1'] * self.img_w, mid_y, rect['rx2'] * self.img_w, mid_y, fill="red", width=2)
            else:
                rect['id'] = self.canvas.create_rectangle(rect['rx1'] * self.img_w, rect['ry1'] * self.img_h, rect['rx2'] * self.img_w, rect['ry2'] * self.img_h, outline="red", width=2)
            self.rectangles.append(rect)

    def clear_rects(self):
        for r in self.rectangles:
            self.canvas.delete(r['id'])
        self.rectangles.clear()
        self.redo_stack.clear()

    def save_and_close(self):
        state.selected_crop_regions = [(r['rx1'], r['ry1'], r['rx2'], r['ry2'], r.get('is_vertical', False)) for r in self.rectangles]
        state.loaded_preset_name = "カスタム" if state.selected_crop_regions else "未読込"
        if state.btn_select_crop:
            state.btn_select_crop.config(text=f"抽出範囲を選択 (設定済: {len(state.selected_crop_regions)}か所)" if state.selected_crop_regions else "抽出範囲を選択")
        if state.preset_filename_label:
            state.preset_filename_label.config(text=state.loaded_preset_name)
        if hasattr(self, 'doc') and self.doc:
            self.doc.close()
        self.top.destroy()

    def cancel_and_close(self):
        if hasattr(self, 'doc') and self.doc:
            self.doc.close()
        self.top.destroy()

def open_crop_selector():
    files = state.selected_files if state.current_mode == "file" else ([os.path.join(state.selected_folder, f) for f in os.listdir(state.selected_folder) if f.lower().endswith((".pdf", ".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".json", ".md", ".docx"))] if state.selected_folder else [])
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]
    if not pdf_files:
        return messagebox.showinfo("情報", "PDFファイルが選択されていません。")
    try:
        CropSelector(state.root, pdf_files[0])
    except Exception as e:
        messagebox.showerror("エラー", str(e))

def reset_crop_regions():
    state.selected_crop_regions = []
    state.loaded_preset_name = "未読込"
    if state.btn_select_crop:
        state.btn_select_crop.config(text="抽出範囲を選択")
    if state.preset_filename_label:
        state.preset_filename_label.config(text=state.loaded_preset_name)

# ==============================
# PDFデータ確認ツール (PDFアナライザー)
# ==============================
def show_pdf_type_info():
    files = state.selected_files if state.current_mode == "file" else ([os.path.join(state.selected_folder, f) for f in os.listdir(state.selected_folder) if f.lower().endswith(".pdf")] if state.selected_folder else [])
    pdf_files = [f for f in files if f.lower().endswith('.pdf')]

    if not pdf_files:
        messagebox.showinfo("情報", "確認するPDFファイルが選択されていません。")
        return

    win = tk.Toplevel(state.root)
    win.title("🔍 PDFの内部データ構造 確認結果")
    win.configure(bg=BG_COLOR)
    center_window(win, state.root, 750, 550)
    win.grab_set()

    lbl = ttk.Label(win, text="選択されたPDFのデータ構造を解析しています...", font=("Segoe UI", 11, "bold"), background=BG_COLOR, foreground=PRIMARY)
    lbl.pack(pady=10)

    text_area = st.ScrolledText(win, wrap=tk.WORD, font=("Meiryo UI", 10), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.FLAT, padx=15, pady=15)
    text_area.pack(expand=True, fill=tk.BOTH, padx=10, pady=(0, 10))
    text_area.tag_configure("header", font=("Meiryo UI", 10, "bold"), foreground=PRIMARY)
    text_area.tag_configure("fname", font=("Meiryo UI", 10, "bold"), foreground=TEXT_COLOR)
    text_area.tag_configure("struct", font=("Meiryo UI", 9), foreground=MUTED_TEXT)
    text_area.tag_configure("rec_std", font=("Meiryo UI", 9, "bold"), foreground=COLOR_SUCCESS)
    text_area.tag_configure("rec_ai", font=("Meiryo UI", 9, "bold"), foreground=COLOR_PURPLE)
    text_area.tag_configure("err", font=("Meiryo UI", 9, "bold"), foreground=ERROR)

    btn_frame_bottom = ttk.Frame(win, style="Main.TFrame")
    btn_frame_bottom.pack(pady=(0, 10))
    ttk.Button(btn_frame_bottom, text="📖 内部データ構造の詳しい解説を見る", command=lambda: [win.destroy(), open_help_window("pdf_type")]).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame_bottom, text="閉じる", command=win.destroy, width=15).pack(side=tk.LEFT, padx=5)

    def run_analysis():
        if not win.winfo_exists():
            return
        results = []
        max_files_to_check = 50

        for i, pdf_path in enumerate(pdf_files[:max_files_to_check]):
            if not win.winfo_exists():
                return
            filename = os.path.basename(pdf_path)
            lbl.config(text=f"解析中 ({i+1}/{len(pdf_files)}): {filename}")
            win.update_idletasks()
            try:
                with fitz.open(pdf_path) as doc:
                    if len(doc) == 0:
                        res = "不明 (空のPDF)"
                        engine = "👉 解析不能"
                        rec_tag = "err"
                    else:
                        text_score = 0
                        vector_score = 0
                        pages_to_check = min(3, len(doc))
                        for p_idx in range(pages_to_check):
                            page = doc[p_idx]
                            if len(page.get_text().strip()) > 30:
                                text_score += 1
                            if len(page.get_drawings()) > 5:
                                vector_score += 1

                        if text_score > 0 and vector_score > 0:
                            res = "テキストデータ ＆ ベクターデータ (CAD等からのデジタル生成PDF)"
                            engine = "👉 推奨エンジン: Python標準ライブラリ (高速・オフライン)"
                            rec_tag = "rec_std"
                        elif text_score > 0:
                            res = "テキストデータ (Word/Excel等からのデジタル生成PDF)"
                            engine = "👉 推奨エンジン: Python標準ライブラリ (高速・オフライン)"
                            rec_tag = "rec_std"
                        elif vector_score > 0:
                            res = "ベクターデータ (文字はアウトライン化されている可能性あり)"
                            engine = "👉 推奨エンジン: Python標準ライブラリ または AIエンジン (Gemini / Azure OpenAI)"
                            rec_tag = "rec_ai"
                        else:
                            res = "ラスターデータのみ (スキャンされたPDFや画像のみ)"
                            engine = "👉 推奨エンジン: Gemini API / Azure OpenAI API (超高精度AI) または Tesseract"
                            rec_tag = "rec_ai"
                results.append((filename, res, engine, rec_tag))
            except Exception as e:
                results.append((filename, f"解析エラー: {e}", "", "err"))

            if not win.winfo_exists():
                return
            text_area.delete(1.0, tk.END)
            text_area.insert(tk.END, f"📊 解析対象ファイル数: {len(pdf_files)} 件\n", "header")
            if len(pdf_files) > max_files_to_check:
                text_area.insert(tk.END, f"※ファイル数が多いため、最初の {max_files_to_check} 件のみをサンプリング解析しています。\n", "struct")
            text_area.insert(tk.END, "─" * 58 + "\n\n", "struct")
            for fn, r_str, eng_str, tag in results:
                text_area.insert(tk.END, f"📄 {fn}\n", "fname")
                text_area.insert(tk.END, f"   内部構造: {r_str}\n", "struct" if tag != "err" else "err")
                if eng_str:
                    text_area.insert(tk.END, f"   {eng_str}\n\n", tag)
                else:
                    text_area.insert(tk.END, "\n")
            text_area.see(tk.END)
            win.update_idletasks()

        if win.winfo_exists():
            lbl.config(text="✅ 解析が完了しました。")
            text_area.configure(state=tk.DISABLED)

    win.after(100, run_analysis)


# ==============================
# モダン・ヘルプ＆ガイドセンター / バージョン情報
# ==============================
def _create_scrollable_tab(notebook, tab_title):
    """Notebook内にスクロール可能なフレームを持つタブを作成して返す"""
    outer = ttk.Frame(notebook, style="Main.TFrame")
    notebook.add(outer, text=tab_title)

    canvas = tk.Canvas(outer, bg=BG_COLOR, highlightthickness=0)
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    content = ttk.Frame(canvas, style="Main.TFrame", padding=(18, 14))

    content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    win_id = canvas.create_window((0, 0), window=content, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(win_id, width=e.width))
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return outer, content


def _create_styled_table(parent, headers, rows, col_widths, header_bg="#E7F1FF", header_fg=PRIMARY):
    """カード風の見やすいテーブルを作成するヘルパー"""
    table_border = tk.Frame(parent, bg=BORDER_COLOR, bd=0)
    table_border.pack(fill=tk.X, pady=(4, 10))

    for idx, w in enumerate(col_widths):
        table_border.columnconfigure(idx, weight=1, minsize=w)

    for c_idx, h_text in enumerate(headers):
        lbl = tk.Label(
            table_border, text=h_text, font=("Segoe UI", 9, "bold"),
            bg=header_bg, fg=header_fg, padx=10, pady=8,
            justify="left", anchor="w", wraplength=col_widths[c_idx] - 20
        )
        lbl.grid(row=0, column=c_idx, sticky="nsew", padx=1, pady=1)

    for r_idx, r_data in enumerate(rows, 1):
        row_bg = "#FFFFFF" if r_idx % 2 == 1 else "#F8FAFC"
        for c_idx, val in enumerate(r_data):
            is_first_col = (c_idx == 0)
            lbl = tk.Label(
                table_border, text=val,
                font=("Segoe UI", 9, "bold") if is_first_col else ("Segoe UI", 9),
                bg=row_bg, fg=TEXT_COLOR, padx=10, pady=8,
                justify="left", anchor="nw", wraplength=col_widths[c_idx] - 20
            )
            lbl.grid(row=r_idx, column=c_idx, sticky="nsew", padx=1, pady=1)


def _create_accent_card(parent, accent_color, badge_text, title_text, body_text, extra_builder=None):
    """左側にカラーアクセントバーとバッジを備えたモダンな解説カードを作成する"""
    outer_card = tk.Frame(parent, bg=BORDER_COLOR, padx=1, pady=1)
    outer_card.pack(fill=tk.X, pady=6)

    inner_row = tk.Frame(outer_card, bg=CARD_BG)
    inner_row.pack(fill=tk.BOTH, expand=True)

    accent_bar = tk.Frame(inner_row, bg=accent_color, width=6)
    accent_bar.pack(side=tk.LEFT, fill=tk.Y)

    body_frame = tk.Frame(inner_row, bg=CARD_BG, padx=14, pady=12)
    body_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    top_row = tk.Frame(body_frame, bg=CARD_BG)
    top_row.pack(fill=tk.X, pady=(0, 6))

    if badge_text:
        badge = tk.Label(
            top_row, text=f" {badge_text} ", font=("Segoe UI", 8, "bold"),
            bg=accent_color, fg="white", padx=6, pady=2
        )
        badge.pack(side=tk.LEFT, padx=(0, 8))

    lbl_title = tk.Label(
        top_row, text=title_text, font=("Segoe UI", 11, "bold"),
        bg=CARD_BG, fg=TEXT_COLOR, anchor="w"
    )
    lbl_title.pack(side=tk.LEFT, fill=tk.X)

    if body_text:
        lbl_body = tk.Label(
            body_frame, text=body_text, font=("Meiryo UI", 9),
            bg=CARD_BG, fg=TEXT_COLOR, justify="left", anchor="w", wraplength=820
        )
        lbl_body.pack(fill=tk.X, pady=(2, 4))

    if extra_builder:
        extra_frame = tk.Frame(body_frame, bg=CARD_BG)
        extra_frame.pack(fill=tk.X, pady=(6, 0))
        extra_builder(extra_frame)


def open_help_window(initial_tab="guide"):
    """
    ビジュアルに優れた統合ヘルプ＆ガイドセンターを開く。
    initial_tab: 'guide' | 'pdf_type' | 'ai_help' | 'faq' | 'readme' | 'history'
    """
    win = tk.Toplevel(state.root)
    win.title(f"📖 {APP_TITLE} ヘルプ ＆ ガイドセンター")

    screen_w = state.root.winfo_screenwidth()
    screen_h = state.root.winfo_screenheight()
    win_w = min(980, screen_w - 40)
    win_h = min(740, screen_h - 80)
    win.configure(bg=BG_COLOR)
    center_window(win, state.root, win_w, win_h)
    win.grab_set()

    # --- 1. 上部ヘッダーバナー ---
    header_banner = tk.Frame(win, bg=CARD_BG, padx=20, pady=12, highlightbackground=BORDER_COLOR, highlightthickness=1)
    header_banner.pack(fill=tk.X, side=tk.TOP)

    header_left = tk.Frame(header_banner, bg=CARD_BG)
    header_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

    title_row = tk.Frame(header_left, bg=CARD_BG)
    title_row.pack( anchor="w")
    tk.Label(title_row, text="📖 PdfEditMiya ヘルプ ＆ 操作ガイド", font=("Segoe UI", 15, "bold"), bg=CARD_BG, fg=PRIMARY).pack(side=tk.LEFT)
    tk.Label(title_row, text=f" {VERSION} ", font=("Segoe UI", 9, "bold"), bg="#E7F1FF", fg=PRIMARY, padx=6, pady=2).pack(side=tk.LEFT, padx=(10, 0))

    tk.Label(
        header_left,
        text="初心者の方から業務利用の方まで、機能の特徴・操作手順・トラブル解決方法をひと目で確認できます。",
        font=("Meiryo UI", 9), bg=CARD_BG, fg=MUTED_TEXT
    ).pack(anchor="w", pady=(3, 0))

    header_right = tk.Frame(header_banner, bg=CARD_BG)
    header_right.pack(side=tk.RIGHT)
    ttk.Button(
        header_right, text="⚙️ AI詳細設定を開く",
        command=lambda: [win.destroy(), open_api_settings_dialog()],
        style="Primary.TButton"
    ).pack(side=tk.RIGHT, padx=4)

    # --- 2. 下部フッター ---
    footer_bar = tk.Frame(win, bg=BG_COLOR, padx=20, pady=10)
    footer_bar.pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(footer_bar, text="💡 ヒント: マウスホイールで画面を上下にスクロールできます。", font=("Segoe UI", 9), bg=BG_COLOR, fg=MUTED_TEXT).pack(side=tk.LEFT)
    ttk.Button(footer_bar, text="閉じる", command=win.destroy, width=16).pack(side=tk.RIGHT)

    # --- 3. タブ構成 (Notebook) ---
    notebook = ttk.Notebook(win)
    notebook.pack(fill=tk.BOTH, expand=True, padx=12, pady=(8, 0))

    tab_guide, frame_guide = _create_scrollable_tab(notebook, " 🚀 使い方ガイド ")
    tab_pdf, frame_pdf = _create_scrollable_tab(notebook, " 🔍 PDFの内部構造 ")
    tab_ai, frame_ai = _create_scrollable_tab(notebook, " 🤖 AI・OCRの準備 ")
    tab_faq, frame_faq = _create_scrollable_tab(notebook, " ❓ よくある質問 (FAQ) ")
    tab_readme, frame_readme = _create_scrollable_tab(notebook, " 📄 Readme ")
    tab_history, frame_history = _create_scrollable_tab(notebook, " 📜 バージョン履歴 ")

    # ==========================================================
    # タブ1: 🚀 使い方ガイド
    # ==========================================================
    tk.Label(frame_guide, text="かんたん5ステップ！基本的な使い方の流れ", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 6))

    steps_data = [
        ("STEP 1", PRIMARY, "対象のPDFファイル（またはフォルダ）を選択する",
         "画面上部の「📄 ファイルを選択」または「📁 フォルダを選択」ボタンを押して、処理したいPDFを読み込みます。\n複数ファイルをまとめて処理したい場合は「フォルダを選択」またはファイルの複数選択が便利です。"),
        ("STEP 2", "#0BACCE", "① エンジン と ② 出力形式・③ 抽出モード を選ぶ",
         "PDFの種類に合わせて抽出エンジンを選びます（迷ったら「🔍 選択中のPDFデータ構造を確認」ボタンで自動診断できます）。\n続いて Excel (.xlsx) や CSV (.csv)、画像 (JPEG/PNG/SVG)、CAD図面 (DXF) などの出力形式を選択します。"),
        ("STEP 3", COLOR_SUCCESS, "必要に応じて「抽出範囲を選択」で読み取り場所を囲む（任意）",
         "ページ全体ではなく、図面の表題欄や特定の項目だけを抜き出したい場合は「抽出範囲を選択」ボタンを押します。\nプレビュー画面上でマウスをドラッグして何箇所でも赤枠（または赤線）で指定でき、プリセットとして保存・再利用も可能です。"),
        ("STEP 4", COLOR_DANGER, "「🚀 選択した抽出・変換を実行」ボタンを押す",
         "右下の実行ボタンを押すと変換処理がスタートし、リアルタイムで進捗が表示されます。\n完了すると、指定した保存先フォルダに変換後のファイルが出力されます。"),
        ("STEP 5", COLOR_PURPLE, "複数ファイルの結果を1つにまとめる（データ集約 / データ単純結合）",
         "抽出した複数のExcelやCSVファイルを1つの一覧表にまとめたい場合は、「🧩 データ集約」（列名を自動認識して名寄せ＆『〃』記号も自動補完）または「🔗 データ単純結合」（そのまま縦に連結）を実行します。")
    ]
    for badge, color, title, desc in steps_data:
        _create_accent_card(frame_guide, color, badge, title, desc)

    tk.Label(frame_guide, text="▼ 用途で選べる4つの抽出エンジン比較表", font=("Segoe UI", 12, "bold"), bg=BG_COLOR, fg=TEXT_COLOR).pack(anchor="w", pady=(14, 4))
    _create_styled_table(
        frame_guide,
        ["エンジン名", "適したPDFの種類", "ネット接続", "特徴・おすすめのケース"],
        [
            ["Python標準ライブラリ\n(Internal)", "デジタル作成されたPDF\n（文字がマウス選択できるもの）", "不要\n(オフライン)", "処理が一瞬で終わり、文字コードから直接抜き出すため100%正確です。WordやExcelから直接出力したPDFに最適です。"],
            ["Gemini API\n(Google AI)", "スキャンされたPDF\n手書き書類・複雑な表", "必要", "Googleの最新AIが画像を見て人間のように文字と表構造を解読します。無料枠と課金枠を切り替え可能です。"],
            ["Azure OpenAI API\n(法人向けAI)", "スキャンされたPDF\n社外秘・機密帳票・手書き表", "必要", "Microsoft Azure上のセキュアなAI（GPT-5 / GPT-6 等）を使用します。入力データがAI学習に利用されないため企業利用に最適です。"],
            ["Tesseract\n(ローカルOCR)", "スキャンされたPDF\n（きれいな印刷活字メイン）", "不要\n(オフライン)", "インターネットに接続できない環境でスキャンPDFから文字を読み取りたい場合に使用します。"]
        ],
        [160, 200, 100, 380]
    )

    tk.Label(frame_guide, text="▼ 「抽出モード」と「データ結合機能」の使い分け", font=("Segoe UI", 12, "bold"), bg=BG_COLOR, fg=TEXT_COLOR).pack(anchor="w", pady=(14, 4))
    _create_styled_table(
        frame_guide,
        ["機能・モード名", "どんなときに使う？", "動作と出力の特徴"],
        [
            ["③ 抽出モード:\n表として抽出", "表全体のデータや、枠内の複数項目をセル分けして出力したいとき", "PDF内の縦横の罫線や空白を認識し、ExcelやCSVの別々のセル（行・列）に綺麗に分割して出力します。"],
            ["③ 抽出モード:\nテキストのみ抽出", "1行の文字列や縦書きの列をピンポイントで抜き出したいとき", "範囲選択画面で「水平線（横書き用）」または「垂直線（縦書き用）」を引くだけで、線に触れている文字だけを正確に抜き出して横1行に出力します。"],
            ["🧩 データ集約", "列の並び順が少しずつ違う表や、『〃』『同上』を含む表をまとめたいとき", "1行目の列名（見出し）を自動判別して同じ列に名寄せします。さらに『〃』『同上』などの省略記号を上の行のテキストで自動補完します。"],
            ["🔗 データ単純結合", "すでに同じ列構成になっている抽出ファイルをそのまま縦に繋げたいとき", "余計な加工や位置合わせを行わず、すべてのファイルの行をそのまま順番に縦へ連結して1つのファイルにまとめます。"]
        ],
        [160, 260, 420]
    )

    # ==========================================================
    # タブ2: 🔍 PDFの内部構造
    # ==========================================================
    tk.Label(frame_pdf, text="PDFの内部データ構造と最適なエンジンの選び方", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 4))
    tk.Label(
        frame_pdf,
        text="PDFファイルは画面上の見た目が同じに見えても、作成された方法によって内部の仕組みが大きく3つに分かれています。",
        font=("Meiryo UI", 9), bg=BG_COLOR, fg=TEXT_COLOR
    ).pack(anchor="w", pady=(0, 8))

    _create_accent_card(
        frame_pdf, COLOR_SUCCESS, "構造 1", "テキストデータ（文字情報） ─ デジタル作成PDF",
        "・作られ方：Word、Excel、PowerPointなどから直接「PDFとして保存（エクスポート）」したファイル。\n"
        "・見分け方：PDF閲覧ソフトで文字をマウスでなぞって選択・コピー＆ペーストやキーワード検索ができます。\n"
        "・特徴：背後に文字コード（あいうえお等）が直接埋め込まれているため、拡大しても文字が一切ぼやけません。\n"
        "・おすすめエンジン：「Python標準ライブラリ」を選べば、オフラインで超高速・正確に抽出できます。"
    )

    _create_accent_card(
        frame_pdf, PRIMARY, "構造 2", "ベクターデータ（図形・線画情報） ─ CAD図面・イラストPDF",
        "・作られ方：AutoCAD、Jw_cadなどのCADソフトや、Illustrator等のドローソフトから出力した図面ファイル。\n"
        "・特徴：図形を「点（ピクセル）」ではなく「ここからここまで直線を引く」「半径○の円を描く」という座標と数式で記録しています。\n"
        "・メリット：どれだけ拡大しても線がギザギザになりません。本アプリで「DXF」や「SVG」に変換すると、線の情報をそのままCADの図形要素として高精度に復元できます。\n"
        "・注意点：文字が「図形（アウトライン）」に変換されている場合は、文字抽出に「Gemini API」または「Azure OpenAI API」が必要です。"
    )

    _create_accent_card(
        frame_pdf, COLOR_PURPLE, "構造 3", "ラスターデータ（画像・ピクセル情報） ─ スキャンPDF",
        "・作られ方：紙の書類や図面を複合機・スキャナで読み取ったファイル、または写真を貼り付けたファイル。\n"
        "・見分け方：拡大すると文字や線がモザイク状にギザギザになり、マウスで文字を選択することができません。\n"
        "・特徴：ページ全体が「1枚の写真（画像の模様）」として保存されており、内部に文字コードが存在しません。\n"
        "・おすすめエンジン：画像から文字を解読する必要があるため、「Gemini API」「Azure OpenAI API」（手書きや複雑な表もOK）または「Tesseract」（オフライン活字用）を使用してください。"
    )

    def _build_pdf_check_btn(parent_f):
        ttk.Button(
            parent_f, text="🔍 今すぐ選択中のPDFデータ構造を診断する",
            command=lambda: [win.destroy(), show_pdf_type_info()],
            style="Primary.TButton"
        ).pack(side=tk.LEFT)

    _create_accent_card(
        frame_pdf, COLOR_WARNING, "便利ツール", "手元のPDFがどの構造かわからないときは？",
        "メイン画面でファイルまたはフォルダを選択したあと、「🔍 選択中のPDFデータ構造を確認」ボタンを押してください。\n"
        "アプリが自動的にPDFの中身を解析し、どのエンジンを使えばよいかその場でアドバイスします！",
        extra_builder=_build_pdf_check_btn
    )

    # ==========================================================
    # タブ3: 🤖 AI・OCRの準備
    # ==========================================================
    tk.Label(frame_ai, text="AI抽出エンジン・ローカルOCRの準備と設定方法", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 6))

    def _build_gemini_links(parent_f):
        ttk.Button(
            parent_f, text="🌐 Google AI Studio (APIキー取得ページ) を開く",
            command=lambda: webbrowser.open_new("https://aistudio.google.com/app/apikey"),
            style="Primary.TButton"
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            parent_f, text="⚙️ AI詳細設定画面を開く",
            command=lambda: [win.destroy(), open_api_settings_dialog()]
        ).pack(side=tk.LEFT)

    _create_accent_card(
        frame_ai, PRIMARY, "推奨 1: Gemini API", "Google Gemini API を使う場合（個人・一般利用にも最適 / 無料枠あり）",
        "Googleの最新マルチモーダルAIを利用して、かすれた手書き文字や複雑な表を高精度に読み取ります。\n\n"
        "【APIキーの取得と設定手順】\n"
        "1. 下のボタンから「Google AI Studio」にアクセスし、Googleアカウントでログインします。\n"
        "2. 画面の「Create API key（APIキーを作成）」ボタンを押します。\n"
        "3. 発行された長い英数字の文字列（APIキー）をコピーします。\n"
        "4. 本アプリの「⚙️ 詳細設定」を開き、「APIキー」欄で右クリック→「貼り付け」を行い、「テスト」ボタンを押して通信を確認します。",
        extra_builder=_build_gemini_links
    )

    def _build_azure_links(parent_f):
        ttk.Button(
            parent_f, text="🌐 Azure OpenAI 公式ドキュメントを開く",
            command=lambda: webbrowser.open_new("https://learn.microsoft.com/ja-jp/azure/ai-services/openai/"),
            style="Primary.TButton"
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            parent_f, text="⚙️ AI詳細設定 (Azure OpenAI) を開く",
            command=lambda: [state.engine_var.set("AzureOpenAI"), win.destroy(), open_api_settings_dialog()]
        ).pack(side=tk.LEFT)

    _create_accent_card(
        frame_ai, COLOR_PURPLE, "推奨 2: Azure OpenAI", "Azure OpenAI Service を使う場合（法人・機密データ・セキュア環境向け）",
        "Microsoft Azure 上に構築された企業向け OpenAI モデル（GPT-5 / GPT-6 等）を利用します。入力データが外部のAI学習に利用されないため、社外秘の図面や帳票でも安心して処理できます。\n\n"
        "【設定手順】\n"
        "1. Azure Portal または Azure AI Foundry (旧 Azure OpenAI Studio) を開きます。\n"
        "2. 対象リソースから以下の3点を確認します。\n"
        "   ・エンドポイントURL （例: https://xxxx.openai.azure.com/）\n"
        "   ・APIキー （KEY 1 または KEY 2）\n"
        "   ・デプロイ名 （画像認識に対応した gpt-5、gpt-5-mini、gpt-5.2 等のモデルデプロイ名）\n"
        "3. 本アプリの「⚙️ 詳細設定」→「🔷 Azure OpenAI Service の設定」タブに入力し、「🔌 接続テスト」ボタンを押します。",
        extra_builder=_build_azure_links
    )

    def _build_tesseract_links(parent_f):
        ttk.Button(
            parent_f, text="🌐 Tesseract 配布ページ (GitHub Wiki) を開く",
            command=lambda: webbrowser.open_new("https://github.com/UB-Mannheim/tesseract/wiki"),
            style="Primary.TButton"
        ).pack(side=tk.LEFT)

    _create_accent_card(
        frame_ai, COLOR_SUCCESS, "オフライン用: Tesseract", "Tesseract (ローカルOCR) を使う場合（インターネット接続不要・完全無料）",
        "インターネットに接続できないパソコン環境でも、PC内部で文字認識（OCR）を行える機能です。初回のみ以下の手順で無料ソフトをインストールしてください。\n\n"
        "【インストール手順と重要ポイント】\n"
        "1. 下のボタンから配布ページを開き、最新の64bit版インストーラー（tesseract-ocr-w64-setup...exe）をダウンロードして実行します。\n"
        "2. ⚠️【超重要】途中の「Choose Components（追加機能の選択）」画面で、「Additional language data (download)」の左の「+」を展開し、\n"
        "   『Japanese』と『Japanese (vertical)』の2箇所に必ずチェックを入れてください（チェックしないと日本語が読めません）。\n"
        "3. ⚠️【超重要】インストール先フォルダ（Destination Folder）は変更せず、初期設定（C:\\Program Files\\Tesseract-OCR）のまま完了させてください。",
        extra_builder=_build_tesseract_links
    )

    # ==========================================================
    # タブ4: ❓ よくある質問 (FAQ)
    # ==========================================================
    tk.Label(frame_faq, text="よくある質問・トラブル解決 (FAQ)", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 6))

    faqs = [
        ("Q1. 「標準ライブラリ」で実行すると「スキャンされた画像のため抽出できません」と表示されます",
         "対象のPDFが文字データを持たない「スキャン画像（ラスターデータ）」で作られています。\n"
         "「① エンジン」を『Gemini API』『Azure OpenAI API』または『Tesseract』に切り替えてから実行してください。"),
        ("Q2. Gemini API で「429 (利用枠超過 / Quota Exceeded)」エラーが出る場合の対処法は？",
         "短時間にリクエストを送りすぎた（バースト制限）か、無料枠の1日あたりの上限に達しています。\n"
         "・「⚙️ 詳細設定」で RPM を小さく（例: 12以下）、スレッド数を「1」に設定してください。\n"
         "・モデルを高速・軽量な『Gemini 4 Hydrogen』または『Gemini 3.8 Flash』『Gemini 3.5 Flash-Lite』に変更すると解消しやすくなります。"),
        ("Q3. Azure OpenAI で「404」や「401」エラーが出る場合はどこを確認すればよいですか？",
         "・401エラーの場合：APIキーが間違っているか、別のリソースのエンドポイントURLを入力している可能性があります。\n"
         "・404エラーの場合：「デプロイ名」が Azure AI Foundry 上で作成した名前と完全一致しているか確認してください（モデル名そのものではなく、ご自身で付けたデプロイ名を入力します）。"),
        ("Q4. AI抽出で表のセルがうまく分かれず、隣の列とくっついてしまう場合のコツは？",
         "・「③ 抽出モード」が『表として抽出 (セルごとに分割)』になっているか確認してください。\n"
         "・「⚙️ 詳細設定」の『④ 独自の追加指示 (カスタムプロンプト)』に、「日付と図面番号と品名は必ず別々の列に分割してください」のように具体的な列名を指示すると、分割精度が劇的に向上します。"),
        ("Q5. 縦書きの文章や、図面の1行テキストだけをピンポイントで綺麗に抜きたいときは？",
         "「③ 抽出モード」で『テキストのみ抽出 (横に1行で出力)』を選んでから「抽出範囲を選択」を開いてください。\n"
         "画面上部のツールバーで「横書き (水平線)」または「縦書き (垂直線)」を選び、文字の上を1本の線でなぞるだけで、上下や左右の余計な文字を混ぜずに狙った行・列だけを抜き出せます。"),
        ("Q6. 毎回同じレイアウトの図面から同じ場所を抽出するので、範囲選択を毎回やり直したくありません",
         "一度「抽出範囲を選択」で範囲を設定したあと、メイン画面の「抽出範囲:」の横にある『💾 保存』ボタンを押すと、範囲設定を JSON ファイル（プリセット）として保存できます。\n"
         "次回からは『📂 読込』ボタンでそのファイルを選ぶだけで、一瞬で同じ抽出範囲を呼び出せます。")
    ]
    for q_title, a_body in faqs:
        _create_accent_card(frame_faq, PRIMARY, "FAQ", q_title, f"【回答】\n{a_body}")

    # ==========================================================
    # タブ5: 📄 Readme
    # ==========================================================
    tk.Label(frame_readme, text="アプリケーション概要 (README)", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 6))

    readme_path = resource_path("README.md")
    if os.path.exists(readme_path):
        with open(readme_path, "r", encoding="utf-8") as f:
            readme_raw = f.read().strip()
    else:
        readme_raw = "README.md が見つかりません。"

    readme_card = tk.Frame(frame_readme, bg=BORDER_COLOR, padx=1, pady=1)
    readme_card.pack(fill=tk.BOTH, expand=True, pady=4)
    readme_text = st.ScrolledText(readme_card, wrap=tk.WORD, font=("Meiryo UI", 10), bg=CARD_BG, fg=TEXT_COLOR, relief=tk.FLAT, padx=20, pady=16, height=24)
    readme_text.pack(fill=tk.BOTH, expand=True)
    readme_text.tag_configure("h1", font=("Meiryo UI", 13, "bold"), foreground=PRIMARY, spacing1=4, spacing3=8)
    readme_text.tag_configure("h2", font=("Meiryo UI", 11, "bold"), foreground= COLOR_PURPLE, spacing1=10, spacing3=4)
    readme_text.tag_configure("warn", font=("Meiryo UI", 10, "bold"), foreground="#B45309", background="#FEF3C7")
    readme_text.tag_configure("body", font=("Meiryo UI", 10), foreground=TEXT_COLOR, spacing1=2, spacing3=2)

    for idx_line, line in enumerate(readme_raw.split("\n")):
        s_line = line.strip()
        if idx_line == 0:
            readme_text.insert(tk.END, line + "\n", "h1")
        elif s_line in ["主な機能", "動作環境", "インストール・実行方法", "基本的な使い方", "困ったときは"] or re.match(r'^\d+\.\s+', s_line):
            readme_text.insert(tk.END, line + "\n", "h2")
        elif "警告画面" in s_line or "SmartScreen" in s_line or "詳細情報" in s_line:
            readme_text.insert(tk.END, line + "\n", "warn")
        else:
            readme_text.insert(tk.END, line + "\n", "body")
    readme_text.configure(state=tk.DISABLED)

    # ==========================================================
    # タブ6: 📜 バージョン履歴
    # ==========================================================
    tk.Label(frame_history, text="バージョン更新履歴 (リリースノート)", font=("Segoe UI", 13, "bold"), bg=BG_COLOR, fg=PRIMARY).pack(anchor="w", pady=(0, 6))

    # VERSION_HISTORY をバージョンブロックごとにパースしてカード表示
    blocks = re.split(r'\n(?=\[\s*v[\d\.]+\s*\])', VERSION_HISTORY.strip())
    tag_color_map = {
        "大型機能追加": COLOR_PURPLE,
        "機能追加": PRIMARY,
        "機能刷新": PRIMARY,
        "機能改善": COLOR_SUCCESS,
        "精度向上": COLOR_SUCCESS,
        "データ処理改善": COLOR_SUCCESS,
        "AI連携強化": COLOR_PURPLE,
        "解析強化": COLOR_SUCCESS,
        "UI強化": "#0BACCE",
        "UI改善": "#0BACCE",
        "UIリデザイン": "#0BACCE",
        "視認性向上": "#0BACCE",
        "操作性向上": "#0BACCE",
        "操作性・安定性向上": "#0BACCE",
        "ズーム機能強化": "#0BACCE",
        "状況把握": "#0BACCE",
        "指示機能強化": PRIMARY,
        "不具合修正": COLOR_DANGER
    }

    for block in blocks:
        lines = [l.strip() for l in block.strip().split("\n") if l.strip()]
        if not lines:
            continue
        ver_header = lines[0].strip("[] ")
        items = lines[1:]
        is_latest = (ver_header == VERSION)

        v_card = tk.Frame(frame_history, bg=PRIMARY if is_latest else BORDER_COLOR, padx=1, pady=1)
        v_card.pack(fill=tk.X, pady=6)
        v_inner = tk.Frame(v_card, bg=CARD_BG, padx=16, pady=12)
        v_inner.pack(fill=tk.BOTH, expand=True)

        v_top = tk.Frame(v_inner, bg=CARD_BG)
        v_top.pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            v_top, text=f" 📌 {ver_header} ", font=("Segoe UI", 11, "bold"),
            bg=PRIMARY if is_latest else "#E9ECEF",
            fg="white" if is_latest else TEXT_COLOR,
            padx=8, pady=2
        ).pack(side=tk.LEFT)
        if is_latest:
            tk.Label(
                v_top, text=" LATEST (最新バージョン) ", font=("Segoe UI", 8, "bold"),
                bg=COLOR_SUCCESS, fg="white", padx=6, pady=2
            ).pack(side=tk.LEFT, padx=(8, 0))

        for item_line in items:
            clean_item = item_line.lstrip("-・ ").strip()
            m_tag = re.match(r'^【([^】]+)】\s*(.*)$', clean_item)
            row_f = tk.Frame(v_inner, bg=CARD_BG)
            row_f.pack(fill=tk.X, pady=3)
            if m_tag:
                tag_name = m_tag.group(1)
                rest_text = m_tag.group(2)
                badge_col = tag_color_map.get(tag_name, MUTED_TEXT)
                tk.Label(
                    row_f, text=f" {tag_name} ", font=("Meiryo UI", 8, "bold"),
                    bg=badge_col, fg="white", width=16, anchor="center", pady=1
                ).pack(side=tk.LEFT, anchor="nw", padx=(0, 8), pady=(2, 0))
                tk.Label(
                    row_f, text=rest_text, font=("Meiryo UI", 9),
                    bg=CARD_BG, fg=TEXT_COLOR, justify="left", anchor="w", wraplength=720
                ).pack(side=tk.LEFT, fill=tk.X, expand=True)
            else:
                tk.Label(
                    row_f, text="・" + clean_item, font=("Meiryo UI", 9),
                    bg=CARD_BG, fg=TEXT_COLOR, justify="left", anchor="w", wraplength=820
                ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # 初期表示タブの選択
    tab_map = {
        "guide": tab_guide,
        "pdf_type": tab_pdf,
        "ai_help": tab_ai,
        "faq": tab_faq,
        "readme": tab_readme,
        "history": tab_history
    }
    notebook.select(tab_map.get(initial_tab, tab_guide))


def open_about_dialog():
    """モダンなバージョン情報ダイアログを表示する"""
    win = tk.Toplevel(state.root)
    win.title(f"バージョン情報 - {APP_TITLE}")
    win_w, win_h = 480, 360
    win.configure(bg=BG_COLOR)
    win.resizable(False, False)
    center_window(win, state.root, win_w, win_h)
    win.grab_set()

    card = tk.Frame(win, bg=CARD_BG, highlightbackground=BORDER_COLOR, highlightthickness=1, padx=24, pady=20)
    card.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)

    title_row = tk.Frame(card, bg=CARD_BG)
    title_row.pack(pady=(4, 2))
    tk.Label(title_row, text=APP_TITLE, font=("Segoe UI", 20, "bold"), bg=CARD_BG, fg=PRIMARY).pack(side=tk.LEFT)
    tk.Label(title_row, text=f" {VERSION} ", font=("Segoe UI", 10, "bold"), bg="#E7F1FF", fg=PRIMARY, padx=8, pady=2).pack(side=tk.LEFT, padx=(10, 0))

    tk.Label(card, text="高性能 PDFデータ抽出・変換・編集ツール", font=("Meiryo UI", 10, "bold"), bg=CARD_BG, fg=TEXT_COLOR).pack(pady=(2, 10))

    badge_row = tk.Frame(card, bg=CARD_BG)
    badge_row.pack(pady=(0, 12))
    for b_txt, b_col in [("Python標準抽出", COLOR_SUCCESS), ("Gemini API", PRIMARY), ("Azure OpenAI", COLOR_PURPLE), ("Tesseract OCR", "#0BACCE")]:
        tk.Label(badge_row, text=f" {b_txt} ", font=("Segoe UI", 8, "bold"), bg=b_col, fg="white", padx=5, pady=2).pack(side=tk.LEFT, padx=3)

    info_box = tk.Frame(card, bg="#F8FAFC", highlightbackground=BORDER_COLOR, highlightthickness=1, padx=12, pady=10)
    info_box.pack(fill=tk.X, pady=(0, 14))
    tk.Label(info_box, text="・対応出力: Excel / CSV / Text / JSON / Markdown / Word / 画像 / SVG / DXF", font=("Meiryo UI", 8), bg="#F8FAFC", fg=TEXT_COLOR, anchor="w").pack(fill=tk.X)
    tk.Label(info_box, text=f"・設定保存先: {APP_DIR}", font=("Meiryo UI", 8), bg="#F8FAFC", fg=MUTED_TEXT, anchor="w").pack(fill=tk.X, pady=(3, 0))

    btn_row = tk.Frame(card, bg=CARD_BG)
    btn_row.pack(fill=tk.X, pady=(4, 0))

    def _open_app_dir():
        try:
            if os.path.exists(APP_DIR):
                os.startfile(APP_DIR)
        except Exception:
            pass

    ttk.Button(btn_row, text="📂 設定フォルダを開く", command=_open_app_dir).pack(side=tk.LEFT)
    ttk.Button(btn_row, text="📖 ヘルプを開く", command=lambda: [win.destroy(), open_help_window("guide")], style="Primary.TButton").pack(side=tk.LEFT, padx=8)
    ttk.Button(btn_row, text="閉じる", command=win.destroy, width=10).pack(side=tk.RIGHT)

