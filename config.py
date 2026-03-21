# -*- coding: utf-8 -*-
import os
import sys

# ==============================
# 基本設定 & カラーパレット
# ==============================
APP_TITLE, VERSION = "PdfEditMiya", "v2.1.1"
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
"""

# ==============================
# 共通ユーティリティ
# ==============================
def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_api_key():
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE, "r", encoding="utf-8") as f: return f.read().strip()
    return None

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