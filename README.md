PdfEditMiya v1.12.0

PdfEditMiyaは、PDFファイルの編集（結合・分割・回転）から、AIや標準ライブラリを用いた高度なデータ抽出、そして抽出したデータの自動集約までをワンストップで行うデスクトップアプリケーションです。

🌟 主な特徴

直感的なUIとMVCアーキテクチャ: 画面操作（UI）と内部処理（Core）を分離したモダンな設計を採用。設定画面は「エンジン選択」→「出力形式」と流れるように操作でき、無効な設定は自動でロックされます。

高精度AIデータ抽出: Google Gemini APIを利用し、かすれた手書き文字や複雑な表の罫線を高精度に認識。また、ローカルで完結するTesseract OCRにも対応しています。

多彩な出力フォーマット: Excel、CSV、Text、DXF、画像（JPEG, PNG, TIFF, BMP）に加え、AI専用の出力としてJSON、Markdown、Word（.docx）、図面用のSVGベクター出力に完全対応しています。

全機能対応のクロップ（範囲抽出）機能: プレビュー画面からドラッグで指定した範囲（複数選択可）のみを抽出することが可能です。Ctrl+スクロールで拡大縮小、Shift+スクロールで横移動も可能です。

スマートデータ集約: バラバラに抽出された複数ページの表データ（Excel, CSV, JSON, Markdown, Word）を、列ごとの名前や位置関係を自動でマッピングし、レイアウトを崩さずに1つのマスターファイルに結合・集約します。

⚙️ システム要件・事前準備

本アプリをフルに活用するためには、以下の環境準備が必要です。

Python環境

Python 3.8 以上がインストールされていること。

Gemini APIキー (AI抽出を利用する場合 / 推奨)

Google AI Studio (https://aistudio.google.com/app/apikey) にて無料のAPIキーを取得し、アプリ内の入力枠に設定してください。

Tesseract OCR (ローカルAI抽出を利用する場合)

オフラインでOCR機能を使う場合はインストールが必要です。

Windows用インストーラー: https://github.com/UB-Mannheim/tesseract/wiki

※インストール時、必ず「Japanese」および「Japanese (vertical)」の言語データにチェックを入れてください。

📦 インストール方法

任意のフォルダに本アプリのファイル群（app.py, pdf_engine.py, ai_engine.py, utils.py, requirements.txt 等）を配置します。

コマンドプロンプトまたはターミナルを開き、配置したフォルダへ移動します。

以下のコマンドを実行し、必要なライブラリをインストールします。

pip install -r requirements.txt


アプリケーションを起動します。

python app.py
