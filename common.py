# -*- coding: utf-8 -*-
import os
import sys
import re
import ast
import jaconv
from openpyxl.utils import get_column_letter

# ==============================
# 基本設定 & カラーパレット
# ==============================
APP_TITLE, VERSION = "PdfEditMiya", "v2.10.0"
WINDOW_WIDTH, WINDOW_HEIGHT = 840, 600

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
APP_DIR = os.path.join(USER_HOME, "PdfEditMiya")

# 保存用フォルダが存在しない場合は作成
if not os.path.exists(APP_DIR):
    os.makedirs(APP_DIR)

API_KEY_FILE = os.path.join(APP_DIR, ".pdfeditmiya_api_key.txt")
SETTINGS_FILE = os.path.join(APP_DIR, ".pdfeditmiya_settings.json") 

# ==============================
# ヘルプ・履歴テキスト
# ==============================
VERSION_HISTORY = """
[ v2.10.0 ]
- 【UIリデザイン】抽出範囲の選択画面を、よりモダンで機能的なレイアウトに大幅に刷新しました。
- 【機能追加】抽出範囲の選択画面に「中止」ボタンを追加。設定を保存せずに前の画面に戻ることが可能になりました。
- 【UI改善】抽出範囲（クロップ画面）でのPDF表示を、常に「画面幅に合わせる」最適サイズで開くように調整しました。
- 【ズーム機能強化】ドラッグした範囲を拡大する「範囲でズーム」や、マウス位置を基準にしたホイール拡大・縮小、1クリックでの全表示復帰を実装しました。
- 【操作性向上】画面下部にマウス操作やショートカットキーの「操作ガイド」を常時表示し、不慣れな方でも迷わず操作できるようになりました。
- 【機能追加】「データ単純結合」ボタンを実装。同じ列構成のファイルを、内容を加工せずにそのまま縦に繋げることができます。
- 【精度向上】「テキストのみ抽出（水平線選択）」において、文字の中心座標ではなく「線が文字に触れているか」で判定する高精度ロジックを採用。隣接する行の誤混入を大幅に削減しました。
- 【データ処理改善】全ての抽出テキストに対し「全角英数字・全角カナを半角に自動変換（正規化）」する機能を標準搭載。データ集権後の加工の手間を省きます。
- 【視認性向上】メイン画面の各種設定（エンジン・形式・モード）の選択状態を「青色・太字・背景強調」に更新し、一目で設定内容を把握できるようになりました。
- 【状況把握】処理中の「リアルタイム・モニター」を強化。PDF読み込み、AI解析、Excel保存などの詳細な状況がリアルタイムに表示されます。
- 【解析強化】PDFアナライザー（データ構造の確認）において、解析中のファイル名を個別に表示するように改善しました。

[ v2.9.0 ]
- 【機能追加】縦書きテキストの抽出に対応しました。「テキストのみ抽出」モードで「縦書き（垂直線）」を選択することで、縦書き文章をピンポイントで抽出可能です。
- 【UI改善】範囲選択画面で、水平線（横書き用）と垂直線（縦書き用）を自由に切り替えられる切り替えボタンを搭載しました。

[ v2.8.0 ]
- 【機能改善】PDFのテキスト抽出において、指定した枠に一部でも接している（またがっている）文字データも抽出対象に含まれるよう挙動を改善しました。
- 【AI連携強化】Gemini 3系列の最新モデルに対応。無料枠でのバースト制限（429エラー）を賢く回避する待機ロジックを刷新しました。

[ v2.5.0 ]
- 【機能追加】PDFの「内部データ構造（テキスト、ベクター、ラスター）」を自動判定し、最適な処理エンジンを提案してくれる「PDFデータ構造を確認」ボタンを追加しました。

[ v2.4.0 ]
- 【機能追加】設定画面からGeminiの最新モデル一覧をAPI経由で自動取得・更新する機能を追加しました。
- 【UI改善】モデル選択時に各モデルの制限（RPM）や推奨される並列実行数を一目で確認できる詳細テーブル表示を実装しました。

[ v2.1.0 ]
- 【機能刷新】API詳細設定画面を実装。Gemini APIの「無料枠」と「課金枠」を個別に設定・保存できるようにし、運用に合わせて柔軟に切り替えられるようになりました。
- 【指示機能強化】AIへの抽出指示（カスタムプロンプト）を、1行ずつチェックボックスで管理・保存・復元できる新UIへと進化させました。
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

PDF_TYPE_HELP_TEXT = """
【 PDFの内部データ構造について 】

PDFファイルは、画面上の見た目は同じでも、作成された方法（ソフトウェア出力か、スキャナ読み取りか等）によって内部に保持しているデータの種類が大きく異なります。
主に以下の3つのデータ要素で構成されています。

───────────────────────────
■ 1. テキストデータ（文字情報）
───────────────────────────
・特徴：
  Word、Excel、PowerPointなどのソフトウェアから直接「PDFとして保存（エクスポート）」した場合に生成されます。
  文字の形ではなく、背後に「文字コード（あいうえお等）」が直接埋め込まれている状態です。
・性質：
  PDF閲覧ソフト上で文字をマウスでなぞって「選択」や「コピー＆ペースト」、キーワード検索が可能です。
  フォント情報も保持されるため、どれだけ拡大しても文字がぼやけたり劣化したりしません。
  （本アプリの「標準ライブラリ」エンジンは、このデータを直接抽出します）

───────────────────────────
■ 2. ベクターデータ（図形・線画情報）
───────────────────────────
・特徴：
  CADソフト（AutoCAD、Jw_cadなど）やIllustratorなどのドローソフトから出力された図面PDFに多く含まれます。
  図形を「ピクセル（点）」ではなく、「ここからここまで直線を引く」「半径○の円を描く」といった座標と数式（ベクタ）で表現しています。
・性質：
  どれだけ拡大しても線がギザギザにならず、滑らかで鮮明なまま保たれます。
  （本アプリで「DXF」や「SVG」へ変換する際、このベクターデータが含まれていると、線の情報をそのままCAD等の図形要素として高精度に復元できます）

───────────────────────────
■ 3. ラスターデータ（画像・ピクセル情報）
───────────────────────────
・特徴：
  紙の書類を複合機やスキャナで読み取った場合（スキャンPDF）や、Word等に写真を貼り付けた場合に生成されます。
  色のついた小さな点（ピクセル）の集まりで構成されています。
・性質：
  拡大するとモザイク状に粗く（ギザギザに）なります。
  スキャナで読み取っただけのPDFは、ページ全体が「1枚の写真（ラスターデータ）」として保存されており、文字に見える部分も単なる「画像の模様」でしかありません。
  そのため、文字を選択したり検索したりできず、テキスト化するにはAIやOCR（画像認識技術）を使って画像から文字を推測・解読する必要があります。
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

def normalize_text(text):
    """全角英数字と全角カタカナを半角に変換する"""
    if not isinstance(text, str): return text
    return jaconv.z2h(text, kana=True, digit=True, ascii=True)

def sanitize_excel_text(text):
    """Excelに書き込む際にエラーとなる制御文字を除去し、全角・半角正規化を適用する"""
    if text is None: return ""
    text = normalize_text(str(text))
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

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
        if not s or s in ["〃", "”", "\"", "''", "””", "''", "同上", "...", "…"]: return False
        return bool(re.search(r'[a-zA-Zａ-ｚＡ-Ｚぁ-んァ-ン一-龥0-9０-９]', s))
    header = final_aggregated_data[0]
    skip_cols = {idx for idx, h in enumerate(header) if "備考" in str(h)}
    for col_idx in range(1, len(header)):
        if col_idx in skip_cols: continue
        last_text = ""
        for row_idx in range(1, len(final_aggregated_data)):
            cell_val = str(final_aggregated_data[row_idx][col_idx]).strip()
            
            if cell_val == "None":
                cell_val = ""
                final_aggregated_data[row_idx][col_idx] = ""
                
            if cell_val in ["〃", "”", "\"", "''", "””", "''", "同上", "...", "…"]:
                if last_text: final_aggregated_data[row_idx][col_idx] = last_text
            elif cell_val == "":
                pass
            else:
                last_text = cell_val if is_text_to_inherit(cell_val) else ""

def merge_2d_arrays_horizontally(arrays_list):
    if not arrays_list: return []
    max_rows = max((len(arr) for arr in arrays_list), default=0)
    if max_rows == 0:
        max_rows = 1
    merged = []
    # 抽出範囲にテキストデータが無い場合でも、最低1列の空のセル（列）を保証する
    region_max_cols = [max((max((len(row) for row in arr), default=0) if arr else 0), 1) for arr in arrays_list]
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
        self.loaded_preset_name = "未読込"
        
        # 外部から更新が必要なUIウィジェットへの参照
        self.path_label = None
        self.save_label = None
        self.btn_select_crop = None
        self.btn_api_settings = None
        self.plan_indicator = None # プラン表示用
        self.status_label = None
        self.preset_filename_label = None

        # Tkinter Variables
        self.rotate_option = None
        self.save_option = None
        self.engine_var = None
        self.output_format_var = None
        self.extract_mode_var = None # 抽出モード管理用変数を追加
        
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