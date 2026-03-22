# -*- coding: utf-8 -*-
import os
import sys
import re
import ast
from openpyxl.utils import get_column_letter

# ==============================
# 基本設定 & カラーパレット
# ==============================
APP_TITLE, VERSION = "PdfEditMiya", "v2.2.2"
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
[ v2.2.2 ]
- 【UI改善】メイン画面のエンジン選択のラベルテキストを用途に合わせて分かりやすく修正しました。
- 【バグ修正】初期設定読み込み時の出力フォーマット変数の初期化エラーを修正しました。

[ v2.2.1 ]
- 【バグ修正】Geminiデータ集約時に発生していた「NameError」およびアプリがエラー表示なく無言で終了する問題を修正しました。
- 【バグ修正】処理中のエラー内容が画面上のポップアップ（メッセージボックス）で詳細に表示されるように改善しました。
- 【機能改善】Geminiデータ集約処理でも、詳細設定画面で設定した「モデル」や「カスタムプロンプト」等の詳細設定が正しく引き継がれるように修正しました。
- 【機能改善】Geminiデータ集約時にデータ量が多すぎてAIの出力が途切れた場合（JSONパースエラー）、その旨を分かりやすくユーザーに警告するよう改善しました。

[ v2.1.1 ]
- 【バグ修正】Geminiによるデータ集約機能で、特定のExcelデータが含まれる場合にJSONパースエラーが発生する問題を修正しました。
- 【UI改善】メイン画面に現在の「実行プラン（無料枠/課金枠）」を表示するインジケーターを追加し、設定状況を把握しやすくしました。

[ v2.1.0 ]
- 【UI改善】API詳細設定画面の「制限と仕様を確認」ボタンを強化し、プラン比較や各モデルの特徴に加え「モデルごとのRPM・スレッド数の制限と推奨値」を詳細なエクセル風テーブルで一目で確認できるよう改善しました。
- 【UI改善】API詳細設定画面のレイアウトを大幅に見直し、項目を左右に並べることで縦幅を削減。画面の低いノートPC等でも見切れずに全体が収まるように改善しました。
- 【機能改善】AI抽出のカスタムプロンプトを「左右分割のチェックボックス型UI」へと大幅に刷新しました。指示を1行ずつ追加でき、お気に入りの保存・復元もチェックボックスで直感的に行えるようになりました。
- 【機能追加】API詳細設定画面を実装し、Gemini APIの設定を「無料枠」と「課金枠」それぞれ独立して設定・保持できるようになりました。
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
■ Tesseract（ローカルOCR）を使う場合（オフライン環境向け）
───────────────────────────
インターネットに接続できない環境でもご利用いただける無料の文字認識機能です。
ご利用には、事前に専用ソフトのインストールが必要です。以下の手順でご準備ください。

[インストール手順]
1. ブラウザで以下の配布ページにアクセスします。
   https://github.com/UB-Mannheim/tesseract/wiki
2. ページ内から最新の64bit版インストーラー（tesseract-ocr-w64-setup...exe）をダウンロードして実行します。
3. 【重要】インストール中の「Choose Components（追加機能の選択）」画面で、
   「Additional language data (download)」の左側にある「+」を展開し、
   「Japanese」と「Japanese (vertical)」の2箇所に必ずチェックを入れてください。
4. 「Destination Folder（インストール先）」は絶対に場所を変更せず
   （初期設定の C:\Program Files\Tesseract-OCR のまま）最後まで完了させてください。
"""

# ==============================
# 共通ユーティリティ関数
# ==============================
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    return None

def sanitize_excel_text(text):
    """Excelに書き込む際にエラーとなる制御文字を除去する"""
    if text is None: return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(text))

def auto_adjust_excel_column_width(ws):
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    for line in str(cell.value).split('\n'):
                        length = sum(2 if ord(c) > 255 else 1 for c in line)
                        if length > max_length: max_length = length
            except: pass
        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

def analyze_column_profile(col_data):
    if not col_data: return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
    pure_num_cnt, fraction_cnt, total_num_len, total_cells = 0, 0, 0, 0
    for val in col_data:
        s = str(val).strip()
        if not s or s == "None": continue
        total_cells += 1
        if re.match(r'^\d+/\d+$', s):
            fraction_cnt += 1; total_num_len += len(s)
        else:
            s_clean = s.replace(",", "").replace(".", "", 1).replace("-", "", 1)
            if s_clean.isdigit(): pure_num_cnt += 1; total_num_len += len(s_clean)
    if total_cells == 0: return {"pure_num_ratio": 0.0, "fraction_ratio": 0.0, "avg_num_len": 0.0, "is_text": True}
    return {
        "pure_num_ratio": pure_num_cnt / total_cells,
        "fraction_ratio": fraction_cnt / total_cells,
        "avg_num_len": total_num_len / (pure_num_cnt + fraction_cnt) if (pure_num_cnt + fraction_cnt) > 0 else 0.0,
        "is_text": ((pure_num_cnt / total_cells) < 0.2 and (fraction_cnt / total_cells) < 0.2)
    }

def get_profile_similarity(p1, p2):
    diff_pure = abs(p1["pure_num_ratio"] - p2["pure_num_ratio"])
    diff_frac = abs(p1["fraction_ratio"] - p2["fraction_ratio"])
    max_len = max(p1["avg_num_len"], p2["avg_num_len"])
    diff_len = abs(p1["avg_num_len"] - p2["avg_num_len"]) / max_len if max_len > 0 else 0.0
    return max(0.0, 1.0 - (diff_pure * 0.4 + diff_frac * 0.4 + diff_len * 0.2))

def parse_row_data(row_data):
    if isinstance(row_data, list) and len(row_data) == 1: row_data = row_data[0]
    if isinstance(row_data, str):
        row_data = row_data.strip()
        if (row_data.startswith('(') and row_data.endswith(')')) or (row_data.startswith('[') and row_data.endswith(']')):
            try:
                parsed = ast.literal_eval(row_data)
                if isinstance(parsed, (list, tuple)): return [str(x) if x is not None else "" for x in parsed]
            except:
                return [x.strip().strip("'\"") for x in row_data.strip("()[]").split(",")]
        return [row_data]
    if isinstance(row_data, tuple): return [str(x) if x is not None else "" for x in row_data]
    if not isinstance(row_data, list): return [str(row_data)]
    return [str(x) if x is not None else "" for x in row_data]

def apply_text_inheritance(final_aggregated_data):
    if len(final_aggregated_data) <= 1: return
    def is_text_to_inherit(text):
        s = str(text).strip()
        if not s or s in ["〃", "”", "\"", "''", "””", "''", "同上"]: return False
        return bool(re.search(r'[a-zA-Zａ-ｚＡ-Ｚぁ-んァ-ン一-龥]', s))
    header = final_aggregated_data[0]
    skip_cols = {idx for idx, h in enumerate(header) if "備考" in str(h)}
    for col_idx in range(1, len(header)):
        if col_idx in skip_cols: continue
        last_text = ""
        for row_idx in range(1, len(final_aggregated_data)):
            cell_val = str(final_aggregated_data[row_idx][col_idx]).strip()
            if cell_val in ["", "None", "〃", "”", "\"", "''", "””", "''", "同上"]:
                if last_text: final_aggregated_data[row_idx][col_idx] = last_text
            else:
                last_text = cell_val if is_text_to_inherit(cell_val) else ""

def merge_2d_arrays_horizontally(arrays_list):
    if not arrays_list: return []
    max_rows = max((len(arr) for arr in arrays_list), default=0)
    merged = []
    region_max_cols = [max((len(row) for row in arr), default=0) if arr else 0 for arr in arrays_list]
    for r in range(max_rows):
        merged_row = []
        for i, arr in enumerate(arrays_list):
            max_c = region_max_cols[i]
            if arr and r < len(arr):
                row_data = list(arr[r])
                row_data += [""] * (max_c - len(row_data))
                merged_row.extend(row_data)
            else: merged_row.extend([""] * max_c)
        merged.append(merged_row)
    return merged

# ==============================
# アプリケーション全体で共有する状態（State）
# ==============================
class SharedState:
    """各モジュール間で循環参照を起こさずに変数やUIの参照を共有するためのコンテナ"""
    def __init__(self):
        self.root = None
        self.selected_files = []
        self.selected_folder = ""
        self.current_mode = None
        self.preset_save_dir = ""
        self.selected_crop_regions = []
        self.cancelled = False
        self.saved_custom_prompts = []
        
        # 外部から更新が必要なUIウィジェットへの参照
        self.path_label = None
        self.save_label = None
        self.btn_select_crop = None
        self.btn_api_settings = None
        self.plan_indicator = None # プラン表示用
        self.status_label = None

        # Tkinter Variables
        self.rotate_option = None
        self.save_option = None
        self.engine_var = None
        self.output_format_var = None
        
        self.api_plan_var = None
        self.api_key_free_var = None
        self.api_key_paid_var = None
        self.gemini_model_free_var = None
        self.gemini_model_paid_var = None
        self.api_rpm_free_var = None
        self.api_rpm_paid_var = None
        
        self.temperature_free_var = None
        self.temperature_paid_var = None
        self.safety_free_var = None
        self.safety_paid_var = None
        self.max_tokens_free_var = None
        self.max_tokens_paid_var = None
        self.custom_prompt_free_var = None
        self.custom_prompt_paid_var = None
        self.threads_free_var = None
        self.threads_paid_var = None

# グローバルな状態インスタンス
state = SharedState()